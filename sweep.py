"""
Kingdom Structural automation sweep — Python implementation (v2.1.1).

Deterministic replacement for the LLM-driven scheduled task. Runs all 6 jobs
(A0, A02, A, C, D, B) using Notion's filter query API. A full sweep against
the current DB sizes completes in ~15 seconds instead of the ~15 minutes the
LLM loop takes.

Usage:
    python sweep.py                      # run all jobs
    python sweep.py --dry-run            # show what would be done, no writes
    python sweep.py --jobs A0,A,C        # run subset
    python sweep.py --verbose            # debug-level logging

Environment:
    NOTION_TOKEN    — Notion integration token (required)
    KS_EMAIL_MODE   — 'send' | 'draft' | 'stub' (default 'stub')

Exit codes:
    0 — every job ran cleanly (Failed rows logged but the script itself is ok)
    1 — unrecoverable error (missing token, Notion API down, etc.)

Architecture:
    This module is thin. All business constants live in config.py. All folder
    creation / contract rendering is delegated to render_proposal_package.py.
    Email sending is delegated to notifications.py (pluggable).

    The "cap 10 rows per job" rule is enforced here via page_size on the
    Notion query. Remaining rows pick up next sweep.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any

import urllib.request
import urllib.error

# Directory this script lives in. Used to anchor file-based paths (sweep.log,
# template, .env) so things still resolve when launched from a different
# working directory — most importantly when launched by pythonw.exe under
# Windows Task Scheduler with no console.
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# When launched by pythonw.exe under Windows Task Scheduler, sys.stdout and
# sys.stderr are either None or point at invalid/closed handles. ANY code
# in the import chain that calls print() will crash the process with
# 0x8007042B before our logging-to-file machinery gets a chance to run.
# Replace them with a no-op writable stream so prints become silent rather
# than fatal.
class _NullStream:
    def write(self, *a, **kw): return 0
    def flush(self):           return None
    def isatty(self):          return False
    def fileno(self):          raise OSError("no fileno on null stream")
    def close(self):           return None
    def writable(self):        return True
    def readable(self):        return False
if sys.stdout is None:
    sys.stdout = _NullStream()
if sys.stderr is None:
    sys.stderr = _NullStream()


# ────────────────────────────────────────────────────────────────────────
# .env auto-loader — zero dependencies. Loads KEY=VALUE lines from a file
# named `.env` next to this script into os.environ BEFORE anything else
# runs. Existing environment variables win (so Task Scheduler / shell can
# still override individual keys). Comments and blank lines are ignored.
# ────────────────────────────────────────────────────────────────────────
def _load_dotenv() -> None:
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if not os.path.exists(env_path):
        return
    try:
        with open(env_path, encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip()
                if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
                    val = val[1:-1]
                os.environ.setdefault(key, val)
    except OSError:
        pass


_load_dotenv()

# Local imports (must come AFTER _load_dotenv so config sees loaded vars)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import config  # noqa: E402

log = logging.getLogger("sweep")

NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2025-09-03"  # use the latest data-source API surface

# ========================================================================
# Notion HTTP layer
# ========================================================================


class NotionError(RuntimeError):
    pass


def _notion_request(method: str, path: str, body: dict | None = None,
                    *, _attempt: int = 1) -> dict:
    """Thin wrapper around Notion REST.

    Retries transient failures:
      - 429 Too Many Requests  → respect Retry-After header, up to 3 attempts
      - 5xx (502/503/504)      → exponential backoff, up to 3 attempts
      - urllib URLError        → 1s backoff, 1 retry (covers transient DNS)
    All other errors raise NotionError immediately.
    """
    token = os.environ.get("NOTION_TOKEN")
    if not token:
        raise NotionError("NOTION_TOKEN env var not set")

    url = f"{NOTION_API}{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Notion-Version", NOTION_VERSION)
    req.add_header("Content-Type", "application/json")

    MAX_ATTEMPTS = 3
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        # Try to retry 429 / 5xx; everything else is fatal.
        if e.code == 429 and _attempt < MAX_ATTEMPTS:
            # Notion sends Retry-After in seconds (int) per their docs.
            wait = float(e.headers.get("Retry-After") or 2 ** _attempt)
            log.warning("Notion 429; sleeping %.1fs (attempt %d/%d)",
                        wait, _attempt, MAX_ATTEMPTS)
            time.sleep(wait)
            return _notion_request(method, path, body, _attempt=_attempt + 1)
        if 500 <= e.code < 600 and _attempt < MAX_ATTEMPTS:
            wait = 2 ** _attempt
            log.warning("Notion %d; sleeping %ds (attempt %d/%d)",
                        e.code, wait, _attempt, MAX_ATTEMPTS)
            time.sleep(wait)
            return _notion_request(method, path, body, _attempt=_attempt + 1)
        payload = e.read().decode("utf-8", errors="replace")
        raise NotionError(f"{method} {path} → {e.code}: {payload}") from e
    except urllib.error.URLError as e:
        if _attempt < 2:
            log.warning("Notion network error: %s; retrying once", e)
            time.sleep(1)
            return _notion_request(method, path, body, _attempt=_attempt + 1)
        raise NotionError(f"{method} {path} → network error: {e}") from e


def query_data_source(ds_id: str, filter_obj: dict | None = None,
                      page_size: int = 10, sorts: list[dict] | None = None) -> list[dict]:
    """Query a data source with optional filter + sorts. Caps at page_size."""
    body: dict[str, Any] = {"page_size": page_size}
    if filter_obj is not None:
        body["filter"] = filter_obj
    if sorts is not None:
        body["sorts"] = sorts
    resp = _notion_request("POST", f"/data_sources/{ds_id}/query", body)
    return resp.get("results", [])


def create_page(parent: dict, properties: dict, children: list | None = None) -> dict:
    body: dict[str, Any] = {"parent": parent, "properties": properties}
    if children is not None:
        body["children"] = children
    return _notion_request("POST", "/pages", body)


# ────────────────────────────────────────────────────────────────────────
# Brief template block copying
# ────────────────────────────────────────────────────────────────────────
# Notion's REST API doesn't accept `template_id` on page creation (that's
# an MCP-only convention). To replicate template behavior, we fetch the
# template page's blocks once per sweep and pass them as `children` when
# creating each new Brief. Nested children (e.g. to-dos under a callout)
# are recursively unwrapped up to Notion's 2-level input limit.

# Block types that render_from_notion.parse_brief_body reads. Any block
# type not in this set gets dropped on the floor — safer than trying to
# copy a block type we don't understand.
_SAFE_BLOCK_TYPES = {
    "paragraph", "heading_1", "heading_2", "heading_3",
    "to_do", "bulleted_list_item", "numbered_list_item",
    "callout", "quote", "divider", "toggle",
}


def _strip_nulls(obj):
    """Recursively drop any key whose value is None.

    Notion's GET /blocks output includes null-valued keys (e.g. `icon: null`)
    on blocks that don't have optional objects set. The pages.create input
    API rejects those nulls — it wants the field either omitted or set to a
    proper object. Strip them uniformly to avoid death by a thousand cuts.
    """
    if isinstance(obj, dict):
        return {k: _strip_nulls(v) for k, v in obj.items() if v is not None}
    if isinstance(obj, list):
        return [_strip_nulls(x) for x in obj]
    return obj


def _block_output_to_input(block: dict) -> dict | None:
    """Convert a block from GET-output shape to pages.create-input shape.

    Strips read-only metadata (id, created_time, etc.) and null-valued
    fields that Notion's input API rejects. If the block has children,
    recursively fetch and embed them. Returns None for blocks we don't
    support, so the caller can filter them out.
    """
    bt = block.get("type")
    if bt not in _SAFE_BLOCK_TYPES:
        return None
    body = _strip_nulls(dict(block.get(bt, {}) or {}))
    out = {"object": "block", "type": bt, bt: body}
    if block.get("has_children"):
        try:
            kids_raw = get_block_children(block["id"])
            kids_in = [c for c in (_block_output_to_input(k) for k in kids_raw) if c]
            if kids_in:
                body["children"] = kids_in
        except NotionError as e:
            # Don't crash the whole Brief render over one missing block subtree,
            # but DO log it — silent failures here meant briefs were created
            # with empty bodies and engineers had to recreate them by hand.
            log.warning("Could not fetch children for block %s (%s): %s — "
                        "block will be created without nested content.",
                        block.get("id", "?")[:8],
                        block.get("type", "?"), e)
    return out


_template_children_cache: list[dict] | None = None


def _fetch_brief_template_children() -> list[dict]:
    """Return the Brief template's body blocks in pages.create input format.

    Cached for the lifetime of the process so multi-row A0 sweeps only pay
    the fetch cost once.
    """
    global _template_children_cache
    if _template_children_cache is not None:
        return _template_children_cache
    try:
        raw = get_block_children(BRIEF_TEMPLATE_ID)
        cleaned = [b for b in (_block_output_to_input(b) for b in raw) if b]
        _template_children_cache = cleaned
        return cleaned
    except NotionError as e:
        log.warning("Could not fetch Brief template children (%s). "
                    "Briefs will be created with empty body.", e)
        _template_children_cache = []
        return []


def update_page(page_id: str, properties: dict) -> dict:
    return _notion_request("PATCH", f"/pages/{page_id}", {"properties": properties})


def get_page(page_id: str) -> dict:
    return _notion_request("GET", f"/pages/{page_id}")


def get_block_children(block_id: str) -> list[dict]:
    """Paginated block-children fetch (used to check for existing Fee Schedule child DB)."""
    results: list[dict] = []
    cursor: str | None = None
    while True:
        path = f"/blocks/{block_id}/children?page_size=100"
        if cursor:
            path += f"&start_cursor={cursor}"
        resp = _notion_request("GET", path)
        results.extend(resp.get("results", []))
        if not resp.get("has_more"):
            break
        cursor = resp.get("next_cursor")
    return results


def create_child_database(parent_page_id: str, title: str, properties_schema: dict) -> dict:
    """Create an inline child database under a Notion page.

    Notion-Version 2025-09-03 restructured this endpoint: properties now
    live inside `initial_data_source` rather than at the top level. Sending
    them at the top level silently creates a DB with no custom properties,
    which is what caused the original 'Service is not a property that
    exists' failures on page-insert.
    """
    body = {
        "parent": {"type": "page_id", "page_id": parent_page_id},
        "title": [{"type": "text", "text": {"content": title}}],
        "initial_data_source": {
            "properties": properties_schema,
        },
        "is_inline": True,
    }
    return _notion_request("POST", "/databases", body)


def post_comment(page_id: str, rich_text: list[dict]) -> dict:
    body = {"parent": {"page_id": page_id}, "rich_text": rich_text}
    return _notion_request("POST", "/comments", body)


# ========================================================================
# Small helpers for Notion property shapes
# ========================================================================


def title_val(props: dict, name: str) -> str:
    arr = props.get(name, {}).get("title", [])
    return "".join(x.get("plain_text", "") for x in arr).strip()


def text_val(props: dict, name: str) -> str:
    arr = props.get(name, {}).get("rich_text", [])
    return "".join(x.get("plain_text", "") for x in arr).strip()


def select_val(props: dict, name: str) -> str | None:
    sel = props.get(name, {}).get("select")
    return sel.get("name") if sel else None


def status_val(props: dict, name: str) -> str | None:
    st = props.get(name, {}).get("status")
    return st.get("name") if st else None


def checkbox_val(props: dict, name: str) -> bool:
    """True if the property is a ticked checkbox OR a formula that evaluates
    to True. Tolerates both shapes so legacy DBs (Include as a real checkbox)
    and v2.1.7+ Fee Schedules (Include as a formula derived from Amount > 0)
    read the same."""
    p = props.get(name, {}) or {}
    if "checkbox" in p:
        return bool(p.get("checkbox", False))
    f = p.get("formula") or {}
    if f.get("type") == "boolean":
        return bool(f.get("boolean", False))
    return False


def relation_ids(props: dict, name: str) -> list[str]:
    return [r["id"] for r in props.get(name, {}).get("relation", [])]


def fetch_relation_ids(page_id: str, props: dict, name: str) -> list[str]:
    """Resolve a relation property, robust to the Notion 2025-09-03 behavior
    where GET /pages/{id} sometimes returns relations as an empty array even
    when the relation is populated.

    Strategy:
      1. Try the inline value from the page GET (fast path — works for most).
      2. If empty AND the property has a valid id, hit
         GET /pages/{id}/properties/{prop_id} which always returns the full
         (paginated) list of relation items.

    Returns the list of related page IDs.
    """
    prop = props.get(name) or {}
    inline = [r["id"] for r in prop.get("relation", [])]
    if inline:
        return inline
    prop_id = prop.get("id")
    if not prop_id:
        return []
    # URL-encode the property id — Notion uses short tokens that usually
    # don't need encoding, but belt-and-suspenders.
    from urllib.parse import quote
    prop_id_enc = quote(prop_id, safe="")
    ids: list[str] = []
    start_cursor: str | None = None
    try:
        while True:
            path = f"/pages/{page_id}/properties/{prop_id_enc}"
            if start_cursor:
                path += f"?start_cursor={start_cursor}"
            resp = _notion_request("GET", path)
            obj = resp.get("object")
            if obj == "list":
                for item in resp.get("results", []):
                    # Property-item shape for a relation:
                    #   {"object":"property_item", "type":"relation",
                    #    "relation": {"id": "<related_page_id>"}}
                    rel = item.get("relation") or {}
                    rid = rel.get("id") if isinstance(rel, dict) else None
                    if rid:
                        ids.append(rid)
                if not resp.get("has_more"):
                    break
                start_cursor = resp.get("next_cursor")
                if not start_cursor:
                    break
            elif obj == "property_item":
                # Single-item response (sometimes returned for small relations)
                rel = resp.get("relation")
                if isinstance(rel, list):
                    for r in rel:
                        if r.get("id"):
                            ids.append(r["id"])
                elif isinstance(rel, dict) and rel.get("id"):
                    ids.append(rel["id"])
                break
            else:
                log.warning("fetch_relation_ids(%s): unexpected response shape: %s",
                            name, str(resp)[:200])
                break
    except NotionError as e:
        log.warning("fetch_relation_ids(%s): property-endpoint error: %s", name, e)
        return []
    return ids


def people_ids(props: dict, name: str) -> list[str]:
    return [p["id"] for p in props.get(name, {}).get("people", [])]


def url_val(props: dict, name: str) -> str | None:
    return props.get(name, {}).get("url")


def rollup_val(props: dict, name: str, target_type: str) -> str:
    """Read a rollup property value and extract the first array item.

    target_type: one of 'title', 'rich_text', 'email', 'phone_number',
                 'number', 'select'.

    Returns empty string on miss. Used to pull contact data through rollups
    on the Project page without doing a second API call to the Contact page
    — which matters because Notion 2025-09-03's REST API returns one-way
    relations as empty even when the UI shows them populated (confirmed bug
    for Project → Project Contact → Contacts DB, since Contacts has no
    back-relation on Projects)."""
    p = props.get(name) or {}
    rollup = p.get("rollup") or {}
    rtype = rollup.get("type")

    # Single-value rollups (number, date) return the value directly
    if rtype == "number":
        v = rollup.get("number")
        return str(v) if v is not None else ""
    if rtype in ("date", "incomplete", "unsupported"):
        return ""

    # Array rollup — take the first element of the matching target type
    arr = rollup.get("array") or []
    if not arr:
        return ""
    first = arr[0]
    ftype = first.get("type")
    if ftype != target_type:
        # Allow a couple of safe coercions
        if target_type == "rich_text" and ftype == "title":
            ftype = "title"
            target_type = "title"
        else:
            return ""

    val = first.get(target_type)
    if target_type in ("title", "rich_text"):
        if isinstance(val, list):
            return "".join(t.get("plain_text", "") for t in val).strip()
        return ""
    if target_type in ("email", "phone_number"):
        return val or ""
    if target_type == "number":
        return str(val) if val is not None else ""
    if target_type == "select":
        return (val or {}).get("name", "") if isinstance(val, dict) else ""
    return ""


# ========================================================================
# Automation Log writer
# ========================================================================


def log_row(job: str, outcome: str, *, title: str, project_id: str | None = None,
            brief_id: str | None = None, error: str = "", details: str = "") -> None:
    """Write one row to the Automation Log DB."""
    props: dict[str, Any] = {
        "Log Entry": {"title": [{"type": "text", "text": {"content": title}}]},
        "Job": {"select": {"name": job}},
        "Outcome": {"select": {"name": outcome}},
        "Timestamp": {"date": {"start": datetime.now(timezone.utc).isoformat()}},
    }
    if project_id:
        props["Project"] = {"relation": [{"id": project_id}]}
    if brief_id:
        props["Brief"] = {"relation": [{"id": brief_id}]}
    if error:
        props["Error"] = {"rich_text": [{"type": "text", "text": {"content": error[:2000]}}]}
    if details:
        props["Details"] = {"rich_text": [{"type": "text", "text": {"content": details[:2000]}}]}

    try:
        create_page({"type": "data_source_id", "data_source_id": config.AUTOMATION_LOG_DS_ID}, props)
    except NotionError as e:
        log.error("Failed to write automation log row (%s / %s): %s", job, title, e)


# ========================================================================
# Job A0 — auto-create Proposal Brief
# ========================================================================


# v2.1.3 (2026-04-23): Tightened to inclusion-list semantics.
# Only projects whose Engineering Status EQUALS one of these values will
# ever be processed by Jobs A0, A, or C. This explicitly excludes every
# project that's already past the proposal stage (Design, CA, Complete,
# etc.) so 2024/2023 backlog projects don't get retroactively briefed.
#
# Add a status here if you want new projects in that state to auto-trigger
# Brief creation + folder creation + engineer notification.
PROPOSAL_TRIGGER_STATUSES = ["Proposal Requested"]

BRIEF_TEMPLATE_ID = "34ab73dc-460e-80bf-b97c-da13d53310e8"


def _proposal_status_filter() -> dict:
    """Return a Notion filter clause matching any PROPOSAL_TRIGGER_STATUSES."""
    return {
        "or": [
            {"property": config.ProjectProp.ENGINEERING_STATUS,
             "status": {"equals": s}}
            for s in PROPOSAL_TRIGGER_STATUSES
        ]
    }


def _intake_complete_filter() -> dict:
    """Return a Notion filter clause requiring Intake Complete == true.

    v2.1.4 guard: admin must explicitly tick 'Intake Complete' on the
    Project before Jobs A0/A/C will fire. Combined with the Intake Status
    formula in Notion that shows what's missing, this prevents automation
    from processing half-populated Projects.
    """
    return {"property": config.ProjectProp.INTAKE_COMPLETE,
            "checkbox": {"equals": True}}


@dataclass
class JobResult:
    touched: int = 0
    skipped: int = 0
    failed: int = 0
    errors: list[str] = field(default_factory=list)

    def line(self, name: str) -> str:
        return f"Job {name}: {self.touched} touched, {self.skipped} skipped, {self.failed} failed"


def job_a0(dry: bool) -> JobResult:
    r = JobResult()
    # v2.1.4 gate: Proposal Brief empty AND Engineering Status is Proposal
    # Requested AND admin has ticked Intake Complete. The Intake Complete
    # guard ensures we don't create a Brief until Client/Contact/etc. have
    # been fully filled in — otherwise downstream contract rendering would
    # leave <<FILL IN>> gaps.
    filt = {
        "and": [
            {"property": config.ProjectProp.PROPOSAL_BRIEF, "relation": {"is_empty": True}},
            _proposal_status_filter(),
            _intake_complete_filter(),
        ]
    }
    try:
        rows = query_data_source(config.PROJECTS_DS_ID, filt, page_size=10)
    except NotionError as e:
        r.failed += 1
        r.errors.append(f"A0 query failed: {e}")
        return r

    for project in rows:
        proj_id = project["id"]
        name = title_val(project["properties"], config.ProjectProp.NAME)
        if not name:
            r.skipped += 1
            log_row("A0", "Skipped", title=f"A0 · Project {proj_id[:8]} has empty Name",
                    project_id=proj_id, details="Project Name is empty; cannot name Brief.")
            continue

        if dry:
            log.info("[dry] A0 would create Brief for %s", name)
            r.touched += 1
            continue

        try:
            brief = create_page(
                {"type": "data_source_id", "data_source_id": config.BRIEF_DS_ID},
                {
                    config.BriefProp.NAME: {"title": [{"type": "text", "text": {"content": name}}]},
                    config.BriefProp.PROJECT: {"relation": [{"id": proj_id}]},
                    config.BriefProp.STATUS: {"select": {"name": config.BriefStatus.NOT_STARTED}},
                },
                children=_fetch_brief_template_children(),
            )
            brief_id = brief["id"]
            brief_url = brief.get("url", "")
            update_page(proj_id, {config.ProjectProp.PROPOSAL_BRIEF: {"relation": [{"id": brief_id}]}})
            r.touched += 1
            log_row("A0", "Success", title=f"A0 · Created Brief for {name}",
                    project_id=proj_id, brief_id=brief_id,
                    details=f"Brief URL: {brief_url}")
        except NotionError as e:
            r.failed += 1
            r.errors.append(str(e))
            log_row("A0", "Failed", title=f"A0 · Failed to create Brief for {name}",
                    project_id=proj_id, error=str(e))
    return r


# ========================================================================
# Job A02 — provision child Fee Schedule DB on each Brief
# ========================================================================


FEE_SCHEDULE_TITLE = "💵 Fee Schedule"
FEE_SCHEDULE_SEED = [
    "Schematic Design",
    "Design Development",
    "Construction Documents",
    "Construction Administration",
    "Special Structural Inspections",
    "Engineering Site Visits",
]
FEE_SCHEDULE_SCHEMA = {
    "Service": {"title": {}},
    # v2.1.7: Include is a formula derived from Amount, not a manual checkbox.
    # Empty Amount → unchecked; Amount > 0 → checked. A literal 0 stays
    # unchecked so it doesn't accidentally pull a $0 row into the contract.
    # Existing Fee Schedules provisioned before v2.1.7 keep Include as a
    # checkbox; checkbox_val() reads both shapes.
    "Include": {
        "formula": {
            "expression": 'if(empty(prop("Amount")), false, prop("Amount") > 0)'
        }
    },
    "Type": {
        "select": {
            "options": [
                {"name": "Fixed Fee", "color": "blue"},
                {"name": "Hourly", "color": "yellow"},
                {"name": "Not to Exceed", "color": "green"},
            ]
        }
    },
    "Amount": {"number": {"format": "dollar"}},
    "Order": {"number": {"format": "number"}},
}


def _brief_has_fee_schedule(brief_id: str) -> bool:
    """True only if a non-archived Fee Schedule DB with the 'Service' property
    (i.e. actually usable) exists as a child of the Brief.

    Skips:
      - Blocks moved to Notion Trash (archived / in_trash) — partners may
        have deleted a Fee Schedule on purpose; A02 re-provisions.
      - Blocks titled 'Fee Schedule' whose data source lacks the 'Service'
        property — these are the broken shells from the early sweep bug
        where properties were sent at the wrong JSON path. A02 will create
        a fresh proper one; the broken shell remains until partners
        manually delete it in Notion.
    """
    for block in get_block_children(brief_id):
        if block.get("archived") or block.get("in_trash"):
            continue
        if block.get("type") != "child_database":
            continue
        title = block.get("child_database", {}).get("title", "")
        if title != FEE_SCHEDULE_TITLE:
            continue
        # Verify the DB has the expected schema (i.e. it's usable).
        try:
            db = _notion_request("GET", f"/databases/{block['id']}")
            ds_list = db.get("data_sources") or []
            if not ds_list:
                continue
            ds = _notion_request("GET", f"/data_sources/{ds_list[0]['id']}")
            if "Service" in (ds.get("properties") or {}):
                return True
        except NotionError as e:
            # We previously returned True ("optimistic") on any error, but
            # that masked real schema-drift bugs (a Fee Schedule with the
            # wrong properties would silently get reused, then page.create
            # would fail downstream with a confusing "Service is not a
            # property" error). Log and treat as "not a valid Fee Schedule"
            # so A02 will provision a fresh one. Re-provisioning a missing
            # DB is cheaper than chasing the cryptic downstream failure.
            log.warning("Fee Schedule verification failed for block %s: %s — "
                        "treating as missing and re-provisioning.",
                        block.get("id", "?")[:8], e)
            return False
    return False


def job_a02(dry: bool) -> JobResult:
    r = JobResult()
    # Sort newest-first so newly-created Briefs (which always need Fee
    # Schedules) get seen ahead of old Briefs that already have them.
    # Scan up to 50 rows per sweep to handle backlog; we still only TOUCH
    # up to the implicit cap of whichever 10 need Fee Schedules first.
    try:
        briefs = query_data_source(
            config.BRIEF_DS_ID, filter_obj=None, page_size=50,
            sorts=[{"timestamp": "created_time", "direction": "descending"}],
        )
    except NotionError as e:
        r.failed += 1
        r.errors.append(f"A02 query failed: {e}")
        return r

    touched_count = 0
    for brief in briefs:
        if touched_count >= 10:
            break                         # enforce the 10-row job cap
        brief_id = brief["id"]
        name = title_val(brief["properties"], config.BriefProp.NAME)
        try:
            if _brief_has_fee_schedule(brief_id):
                continue  # idempotency: silent skip, no log row
            if dry:
                log.info("[dry] A02 would provision Fee Schedule for %s", name)
                r.touched += 1
                touched_count += 1
                continue
            db = create_child_database(brief_id, FEE_SCHEDULE_TITLE, FEE_SCHEDULE_SCHEMA)
            db_id = db["id"]
            # Notion's 2025-09-03 API distinguishes database_id (the parent
            # container) from data_source_id (the actual collection pages
            # live in). pages.create() inside a DB needs the data_source_id,
            # NOT the database_id. For single-source inline DBs the response
            # carries both; fall back to db_id only if the list is empty.
            data_sources = db.get("data_sources") or []
            ds_id = data_sources[0]["id"] if data_sources else db_id

            for idx, svc in enumerate(FEE_SCHEDULE_SEED, start=1):
                # v2.1.7: don't seed Include — it's a formula derived from
                # Amount. Rows start with no Amount, so Include reads as
                # unchecked until the engineer enters a value > 0.
                create_page(
                    {"type": "data_source_id", "data_source_id": ds_id},
                    {
                        "Service": {"title": [{"type": "text", "text": {"content": svc}}]},
                        "Order": {"number": idx},
                    },
                )
            r.touched += 1
            touched_count += 1
            log_row("A02", "Success", title=f"A02 · Provisioned Fee Schedule for {name}",
                    brief_id=brief_id, details=f"DB {db_id} / DS {ds_id}; seeded {len(FEE_SCHEDULE_SEED)} rows")
        except NotionError as e:
            r.failed += 1
            r.errors.append(str(e))
            log_row("A02", "Failed", title=f"A02 · Failed provisioning Fee Schedule for {name}",
                    brief_id=brief_id, error=str(e))
    return r


# ========================================================================
# Job A — OneDrive folder sync (v2.1.1 gate)
# ========================================================================


def job_a(dry: bool) -> JobResult:
    r = JobResult()
    # v2.1.4 gate: Folder=false AND City + State + ENGINEER + Admin present
    # AND Engineering Status is Proposal Requested AND Intake Complete ticked.
    filt = {
        "and": [
            {"property": config.ProjectProp.FOLDER, "checkbox": {"equals": False}},
            {"property": config.ProjectProp.CITY, "rich_text": {"is_not_empty": True}},
            {"property": config.ProjectProp.STATE, "rich_text": {"is_not_empty": True}},
            {"property": config.ProjectProp.ENGINEER, "people": {"is_not_empty": True}},
            {"property": config.ProjectProp.ADMIN, "people": {"is_not_empty": True}},
            _proposal_status_filter(),
            _intake_complete_filter(),
        ]
    }
    try:
        rows = query_data_source(config.PROJECTS_DS_ID, filt, page_size=10)
    except NotionError as e:
        r.failed += 1
        r.errors.append(f"A query failed: {e}")
        return r

    from render_proposal_package import populate_from_template, ONEDRIVE_ROOT  # lazy import

    for project in rows:
        proj_id = project["id"]
        props = project["properties"]
        name = title_val(props, config.ProjectProp.NAME)
        city = text_val(props, config.ProjectProp.CITY)
        state = text_val(props, config.ProjectProp.STATE)

        # Lint: Project Name must carry an 8-digit KS number. If not, log a
        # clean "rename needed" Automation Log row and move on to the next
        # candidate — do NOT raise or create a half-named folder.
        try:
            number, short = config.parse_project_name(name)
        except config.InvalidProjectNumberError as e:
            r.failed += 1
            r.errors.append(f"{name}: {e}")
            log_row("A", "Failed",
                    title=f"A · Rename needed — Project number is not 8 digits: {name!r}",
                    project_id=proj_id,
                    error=str(e),
                    details="Job A is gated on an 8-digit KS job-number prefix. "
                            "Fix the Project Name and the next sweep will pick it up.")
            continue

        try:
            year = config.build_year_from_number(number)
            folder_name = config.build_folder_name(number, short, city, state)
            folder_path = os.path.join(ONEDRIVE_ROOT, year, folder_name)

            if dry:
                log.info("[dry] A would mkdir %s", folder_path)
                r.touched += 1
                continue

            os.makedirs(folder_path, exist_ok=True)
            stats = populate_from_template(folder_path)
            if isinstance(stats, dict) and stats.get("error"):
                raise OSError(stats["error"])

            # Build a Windows-style file:// URL for the Notion stamp so
            # partners can click from Notion and it opens their local copy.
            # Match the format used by existing Notion Project entries, e.g.
            # file:///C:/Users/MichaelBrusnahan/OneDrive%20-%20Kingdom%20Structural%20LLC/_Projects/...
            import urllib.parse
            win_root = os.environ.get(
                "KS_ONEDRIVE_WIN_ROOT",
                "C:/Users/MichaelBrusnahan/OneDrive - Kingdom Structural LLC",
            )
            rel = os.path.relpath(folder_path, ONEDRIVE_ROOT).replace("\\", "/")
            folder_url = "file:///" + urllib.parse.quote(
                f"{win_root}/_Projects/{rel}", safe="/:"
            )

            update_page(proj_id, {
                config.ProjectProp.FOLDER: {"checkbox": True},
                config.ProjectProp.SHAREPOINT_FOLDER: {"url": folder_url},
                config.ProjectProp.STATUS: {"select": {"name": config.ProjectStatus.BRIEF_PENDING}},
            })
            r.touched += 1
            log_row("A", "Success", title=f"A · Folder created for {name}",
                    project_id=proj_id, details=f"Path: {folder_path}; {stats}")
        except (OSError, ValueError, NotionError) as e:
            err = str(e)
            r.failed += 1
            r.errors.append(err)
            # Failure-path per spec: stamp error, flip Status, post Admin comment.
            try:
                # v2.1.4: Last Automation Error property was removed from
                # Projects DB. Error detail still lands in sweep.log and
                # the Automation Log row; Status flips to FOLDER_FAILED so
                # it's visible on the Project row itself.
                update_page(proj_id, {
                    config.ProjectProp.STATUS: {"select": {"name": config.ProjectStatus.FOLDER_FAILED}},
                })
                admin_ids = people_ids(props, config.ProjectProp.ADMIN)
                mentions = [{"type": "mention", "mention": {"type": "user", "user": {"id": uid}}} for uid in admin_ids]
                joined = []
                for m in mentions:
                    joined.append(m)
                    joined.append({"type": "text", "text": {"content": " "}})
                joined.append({"type": "text", "text": {"content":
                    f"— ⚠ Folder Creation Failed for {name}. Error: {err[:400]}"}})
                post_comment(proj_id, joined)
            except NotionError as e2:
                log.error("Failure-path itself failed: %s", e2)
            log_row("A", "Failed", title=f"A · Folder creation failed for {name}",
                    project_id=proj_id, error=err)
    return r


# ========================================================================
# Job C — Engineer notification
# ========================================================================


def job_c(dry: bool) -> JobResult:
    import notifications  # lazy import so sweep runs even if module missing

    r = JobResult()
    # v2.1.4 gate: Folder=true AND Engineer Notified=false AND Engineering
    # Status is Proposal Requested AND Intake Complete ticked.
    filt = {
        "and": [
            {"property": config.ProjectProp.FOLDER, "checkbox": {"equals": True}},
            {"property": config.ProjectProp.ENGINEER_NOTIFIED, "checkbox": {"equals": False}},
            _proposal_status_filter(),
            _intake_complete_filter(),
        ]
    }
    try:
        rows = query_data_source(config.PROJECTS_DS_ID, filt, page_size=10)
    except NotionError as e:
        r.failed += 1
        r.errors.append(f"C query failed: {e}")
        return r

    for project in rows:
        proj_id = project["id"]
        props = project["properties"]
        name = title_val(props, config.ProjectProp.NAME)
        folder_url = url_val(props, config.ProjectProp.SHAREPOINT_FOLDER) or ""
        engineer_uids = people_ids(props, config.ProjectProp.ENGINEER)
        brief_rel = fetch_relation_ids(proj_id, props, config.ProjectProp.PROPOSAL_BRIEF)

        try:
            number, short = config.parse_project_name(name)
            brief_url = ""
            if brief_rel:
                brief = get_page(brief_rel[0])
                brief_url = brief.get("url", "")

            # When KS_EMAIL_MODE=notion, the user's own Notion automations
            # pick up the Engineer Notified flag flip and send their own
            # notification. We skip both the email and the in-sweep @mention
            # comment so we don't double-notify.
            email_mode = os.environ.get("KS_EMAIL_MODE", "stub").lower()
            notion_only = email_mode == "notion"

            sent_any = notion_only  # in notion mode, we're always "done" per project
            if not notion_only:
                for uid in engineer_uids:
                    email = config.engineer_email(uid)
                    if not email:
                        log.warning("Engineer %s not in ENGINEER_ROSTER; skipping email", uid)
                        continue
                    ctx = {
                        "project_number": number,
                        "project_name": short,
                        "engineer_first_name": config.engineer_first_name(uid),
                        "folder_url": folder_url,
                        "brief_url": brief_url,
                    }
                    subj, body = config.format_engineer_email(ctx)
                    if dry:
                        log.info("[dry] C would email %s subj=%r", email, subj)
                    else:
                        notifications.send_email(to=[email], subject=subj, body=body)
                    sent_any = True

            if dry:
                r.touched += 1
                continue

            if not sent_any:
                raise RuntimeError("no engineer addresses resolved from ENGINEER_ROSTER")

            # In-sweep @mention comment — only when NOT in notion mode.
            # (In notion mode, the user's own Notion automations handle it.)
            if not notion_only:
                mentions = [{"type": "mention", "mention": {"type": "user", "user": {"id": uid}}} for uid in engineer_uids]
                rt: list[dict] = []
                for m in mentions:
                    rt.append(m)
                    rt.append({"type": "text", "text": {"content": " "}})
                rt.append({"type": "text", "text": {"content":
                    f"— folders are ready for {number} {short}. OneDrive: {folder_url} · Brief: {brief_url}. "
                    "When done, flip the Brief Status to 'Ready for Admin Review'."}})
                post_comment(proj_id, rt)

            update_page(proj_id, {config.ProjectProp.ENGINEER_NOTIFIED: {"checkbox": True}})
            r.touched += 1
            detail = f"Engineers: {engineer_uids} (notion-only mode)" if notion_only \
                     else f"Engineers: {engineer_uids}"
            log_row("C", "Success", title=f"C · Engineer notified for {name}",
                    project_id=proj_id, details=detail)
        except (NotionError, RuntimeError, ValueError) as e:
            r.failed += 1
            r.errors.append(str(e))
            log_row("C", "Failed", title=f"C · Engineer notify failed for {name}",
                    project_id=proj_id, error=str(e))
    return r


# ========================================================================
# Job D — Admin notification
# ========================================================================


def job_d(dry: bool) -> JobResult:
    import notifications

    r = JobResult()
    filt = {
        "and": [
            {"property": config.BriefProp.STATUS, "select": {"equals": config.BriefStatus.READY_FOR_ADMIN_REVIEW}},
            {"property": config.BriefProp.ADMINS_NOTIFIED, "checkbox": {"equals": False}},
        ]
    }
    try:
        briefs = query_data_source(config.BRIEF_DS_ID, filt, page_size=10)
    except NotionError as e:
        r.failed += 1
        r.errors.append(f"D query failed: {e}")
        return r

    for brief in briefs:
        brief_id = brief["id"]
        bprops = brief["properties"]
        brief_name = title_val(bprops, config.BriefProp.NAME)
        project_rel = fetch_relation_ids(brief_id, bprops, config.BriefProp.PROJECT)

        try:
            if not project_rel:
                raise RuntimeError("Brief has no linked Project")
            project = get_page(project_rel[0])
            pprops = project["properties"]
            proj_name = title_val(pprops, config.ProjectProp.NAME)
            folder_url = url_val(pprops, config.ProjectProp.SHAREPOINT_FOLDER) or ""
            admin_uids = people_ids(pprops, config.ProjectProp.ADMIN)
            engineer_uids = people_ids(pprops, config.ProjectProp.ENGINEER)

            number, short = config.parse_project_name(proj_name)
            engineer_name = (config.ENGINEER_ROSTER.get(engineer_uids[0], ("Engineer", ""))[0]
                             if engineer_uids else "Engineer")

            email_mode = os.environ.get("KS_EMAIL_MODE", "stub").lower()
            notion_only = email_mode == "notion"

            if not notion_only:
                for uid in admin_uids:
                    email = config.engineer_email(uid)
                    if not email:
                        continue
                    ctx = {
                        "project_number": number,
                        "project_name": short,
                        "engineer_name": engineer_name,
                        "brief_url": brief.get("url", ""),
                        "folder_url": folder_url,
                    }
                    subj, body = config.format_admin_email(ctx)
                    if dry:
                        log.info("[dry] D would email %s subj=%r", email, subj)
                    else:
                        notifications.send_email(to=[email], subject=subj, body=body)

            if dry:
                r.touched += 1
                continue

            # In-sweep @mention comment — only when NOT in notion mode.
            if not notion_only:
                mentions = [{"type": "mention", "mention": {"type": "user", "user": {"id": uid}}} for uid in admin_uids]
                rt: list[dict] = []
                for m in mentions:
                    rt.append(m)
                    rt.append({"type": "text", "text": {"content": " "}})
                rt.append({"type": "text", "text": {"content":
                    f"— Brief for {number} {short} is ready for your review. Brief: {brief.get('url','')} · "
                    f"OneDrive: {folder_url}. Flip Brief Status → 'Approved' to kick off contract render."}})
                post_comment(brief_id, rt)

            update_page(brief_id, {config.BriefProp.ADMINS_NOTIFIED: {"checkbox": True}})
            update_page(project_rel[0], {
                config.ProjectProp.STATUS: {"select": {"name": config.ProjectStatus.ADMIN_REVIEW}},
            })
            r.touched += 1
            detail = f"Admins: {admin_uids} (notion-only mode)" if notion_only \
                     else f"Admins: {admin_uids}"
            log_row("D", "Success", title=f"D · Admins notified for {brief_name}",
                    project_id=project_rel[0], brief_id=brief_id,
                    details=detail)
        except (NotionError, RuntimeError, ValueError) as e:
            r.failed += 1
            r.errors.append(str(e))
            log_row("D", "Failed", title=f"D · Admin notify failed for {brief_name}",
                    brief_id=brief_id, error=str(e))
    return r


# ========================================================================
# Job B — Contract Render
# ========================================================================


# ----------------------------------------------------------------------
# Relation resolvers used by Job B
# ----------------------------------------------------------------------
# These convert Notion API pages into the flat dicts render_contract_package
# expects. Keys are intentionally matched to render_from_notion.py's shapes
# (lowercase for engineer; capitalized for client/contact/project dicts).


def _resolve_project_dict(project_page: dict) -> dict:
    """Flatten a Notion Project page into the dict render_from_notion expects.

    v2.1.4: Project Type / Scope Description / Approx. Structural SF were
    removed from the Projects DB. Project Type now lives on the Brief only;
    scope and SF flow through the Brief body + Fee Schedule. The empty
    strings returned below keep the merge dict shape stable so render_from_notion
    doesn't have to change.
    """
    p = project_page["properties"]
    return {
        "Project Name":            title_val(p, config.ProjectProp.NAME),
        "City":                    text_val(p, config.ProjectProp.CITY),
        "State":                   text_val(p, config.ProjectProp.STATE),
        "Project Street":          text_val(p, config.ProjectProp.PROJECT_STREET),
        "Project Type":            "",          # overlaid from Brief later
        "Scope Description":       "",          # legacy field, now empty
        "SharePoint Folder":       url_val(p, config.ProjectProp.SHAREPOINT_FOLDER) or "",
        "Approx. Structural SF":   0,            # legacy field, now zero
        "Engineering Status":      status_val(p, config.ProjectProp.ENGINEERING_STATUS) or "",
    }


def _resolve_client_dict(project_page: dict) -> dict:
    """Follow Project.Client → Client page. Returns {Name, Address, The Guy}.

    Falls back to empty strings (renderer will stamp <<FILL IN>>) when no
    Client is linked. Reads the 'Address' TEXT property — not 'Place'."""
    pprops = project_page["properties"]
    cids = fetch_relation_ids(project_page["id"], pprops, config.ProjectProp.CLIENT)
    if not cids:
        return {"Name": "", "Address": "", "The Guy": ""}
    try:
        page = get_page(cids[0])
    except NotionError:
        return {"Name": "", "Address": "", "The Guy": ""}
    cp = page["properties"]
    # Client DB uses a title property named "Name" (per collection schema)
    return {
        "Name":     title_val(cp, "Name"),
        "Address":  text_val(cp, "Address"),
        "The Guy":  select_val(cp, "The Guy") or "",
    }


def _resolve_contact_dict(project_page: dict) -> dict:
    """Get the Project's linked contact as {Contact Name, Email, Phone,
    Phone (Mobile), Role}.

    Reads the four Contact rollups on the Project page (v2.1.4 property
    additions) rather than fetching Project.Project Contact → Contact page.
    Why: Notion's 2025-09-03 REST API returns one-way relation properties
    as empty even when the UI shows them populated. 'Project Contact' is
    one-way (Contacts DB has no Projects back-relation), so both inline
    and /pages/{id}/properties/{prop_id} return relation: []. Rollups
    compute on Notion's side from the underlying relation graph and return
    the populated data normally, so we use them as the authoritative source.

    Falls back to a direct Contact-page fetch if the relation is two-way or
    if rollups are empty (e.g., someone cleared them on a specific project).
    """
    pprops = project_page["properties"]

    # Fast path: read rollups directly. These were added in v2.1.4 precisely
    # to flatten contact data onto the Project page.
    name   = rollup_val(pprops, "Client Contact",  "title")
    email  = rollup_val(pprops, "Contact Email",   "email")
    phone  = rollup_val(pprops, "Contact Phone",   "phone_number")
    mobile = rollup_val(pprops, "Contact Mobile",  "phone_number")
    role   = rollup_val(pprops, "Contact Role",    "rich_text")

    if name or email or phone or mobile:
        return {
            "Contact Name":   name,
            "Email":          email,
            "Phone":          phone,
            "Phone (Mobile)": mobile,
            "Role":           role,
        }

    # Fallback: relation lookup (works for two-way relations; kept for safety)
    cids = fetch_relation_ids(project_page["id"], pprops, config.ProjectProp.PROJECT_CONTACT)
    if not cids:
        log.warning("Project %s: Contact resolution returned empty (no rollups, "
                    "no Project Contact relation). Contract will render with "
                    "<<FILL IN: Contact Name>> markers. Check that the integration "
                    "is connected to the Contacts database and that this Project "
                    "has a Project Contact set.",
                    project_page.get("id", "?")[:8])
        return {"Contact Name": "", "Email": "", "Phone": "",
                "Phone (Mobile)": "", "Role": ""}
    try:
        page = get_page(cids[0])
    except NotionError as e:
        log.warning("Project %s: Could not fetch linked Contact page %s: %s",
                    project_page.get("id", "?")[:8], cids[0][:8], e)
        return {"Contact Name": "", "Email": "", "Phone": "",
                "Phone (Mobile)": "", "Role": ""}
    cp = page["properties"]
    return {
        "Contact Name":   title_val(cp, "Contact Name"),
        "Email":          (cp.get("Email", {}) or {}).get("email") or "",
        "Phone":          (cp.get("Phone", {}) or {}).get("phone_number") or "",
        "Phone (Mobile)": (cp.get("Phone (Mobile)", {}) or {}).get("phone_number") or "",
        "Role":           text_val(cp, "Role"),
    }


def _resolve_engineer_dict(project_page: dict) -> dict | None:
    """Map Project.ENGINEER[0] → ENGINEER_ROSTER → lowercase-keyed dict.

    Returns the shape resolve_engineer() expects:
      {code, name, email, phone}
    or None if no engineer / not in roster (renderer falls back to
    Client.The Guy → PARTNER_BLOCKS → <<FILL IN>>)."""
    pprops = project_page["properties"]
    eids = people_ids(pprops, config.ProjectProp.ENGINEER)
    if not eids:
        return None
    uid = eids[0]
    # Map Notion user UUID → PARTNER_BLOCKS code via the roster.
    # Roster key = UUID; the code embedded in PARTNER_BLOCKS is based on
    # initials, so we do a simple name-match against render_from_notion's
    # PARTNER_BLOCKS. This is cheap; <10 entries total.
    from render_from_notion import PARTNER_BLOCKS
    roster_entry = config.ENGINEER_ROSTER.get(uid)
    if not roster_entry:
        return None
    display_name, email = roster_entry
    for code, blk in PARTNER_BLOCKS.items():
        if blk["name"].lower() == display_name.lower():
            return {
                "code":  code,
                "name":  blk["name"],
                "email": blk["email"],
                "phone": blk["phone"],
            }
    # Roster hit but no PARTNER_BLOCKS match — synthesize minimal record.
    return {"code": "", "name": display_name, "email": email, "phone": ""}


def _fetch_brief_body_markdown(brief_id: str) -> str:
    """Reconstruct markdown from a Brief page's block tree (recursive).

    We need the checkbox lines + category headers for parse_brief_body().
    The tree is walked recursively because Notion often nests to-do items
    as CHILDREN of a callout block (instead of siblings). The previous
    top-level-only walker missed those — which caused Basic Services to
    come up empty even when every box was checked.

    Bold annotations are preserved (`**text**`) so parse_brief_body's
    `**CATEGORY**` matcher still finds the section headers when those
    headers live inside a rich_text with bold formatting.
    """
    lines: list[str] = []

    def extract_rich_text(rt_list) -> str:
        out: list[str] = []
        for t in (rt_list or []):
            text = t.get("plain_text", "") or ""
            ann = t.get("annotations") or {}
            if ann.get("bold"):
                text = f"**{text}**"
            out.append(text)
        return "".join(out)

    def emit(block: dict) -> None:
        bt = block.get("type")
        if bt == "to_do":
            checked = block["to_do"].get("checked", False)
            text = extract_rich_text(block["to_do"].get("rich_text", []))
            lines.append(f"- [{'x' if checked else ' '}] {text}")
        elif bt == "paragraph":
            lines.append(extract_rich_text(block["paragraph"].get("rich_text", [])))
        elif bt == "callout":
            lines.append("> " + extract_rich_text(block["callout"].get("rich_text", [])))
        elif bt == "quote":
            lines.append("> " + extract_rich_text(block["quote"].get("rich_text", [])))
        elif bt in ("heading_1", "heading_2", "heading_3"):
            level = int(bt[-1])
            lines.append("#" * level + " " + extract_rich_text(block[bt].get("rich_text", [])))
        elif bt == "bulleted_list_item":
            lines.append("- " + extract_rich_text(block["bulleted_list_item"].get("rich_text", [])))
        elif bt == "numbered_list_item":
            lines.append("1. " + extract_rich_text(block["numbered_list_item"].get("rich_text", [])))
        elif bt == "toggle":
            lines.append(extract_rich_text(block["toggle"].get("rich_text", [])))
        elif bt == "divider":
            lines.append("---")
        # Other block types (images, embeds, child_database, synced_block,
        # etc.) are skipped — they don't contribute to scope/basic-services/
        # reimbursables parsing.

        # Recurse into children. This is what catches to-dos nested under
        # a callout: the callout emits its header line above, then we walk
        # its children here.
        if block.get("has_children"):
            try:
                for child in get_block_children(block["id"]):
                    emit(child)
            except NotionError:
                # If we can't fetch children (transient), skip — the partial
                # reconstruction is still useful.
                pass

    for blk in get_block_children(brief_id):
        emit(blk)
    return "\n".join(lines)


def _fetch_fee_lines(brief_id: str) -> list[dict]:
    """Find the 💵 Fee Schedule child DB on the Brief and return its rows
    as fee-line dicts: {service, include, type, amount, order}.

    Notion 2025-09-03 API: child_database blocks carry the DATABASE id, but
    /data_sources/{id}/query requires the DATA SOURCE id (different entity).
    Previous versions treated these as interchangeable, which silently made
    every fee query return [] — so the Fee Schedule table in rendered
    contracts stayed as template placeholders. Fixed 2026-04-24 by resolving
    the data_source_id off the database page before querying.
    """
    fee_db_id = None
    for block in get_block_children(brief_id):
        if block.get("type") == "child_database":
            title = block.get("child_database", {}).get("title", "")
            if title == FEE_SCHEDULE_TITLE:
                fee_db_id = block["id"]
                break
    if not fee_db_id:
        return []
    try:
        db = _notion_request("GET", f"/databases/{fee_db_id}")
        data_sources = db.get("data_sources") or []
        fee_ds_id = data_sources[0]["id"] if data_sources else fee_db_id
        rows = query_data_source(fee_ds_id, filter_obj=None, page_size=100)
    except NotionError:
        return []
    out: list[dict] = []
    for row in rows:
        p = row["properties"]
        out.append({
            "service": title_val(p, "Service"),
            "include": checkbox_val(p, "Include"),
            "type":    select_val(p, "Type"),
            "amount":  (p.get("Amount", {}) or {}).get("number") or 0.0,
            "order":   (p.get("Order",  {}) or {}).get("number") or 0,
        })
    out.sort(key=lambda x: x.get("order") or 0)
    return out


def _resolve_brief_overlay_dict(brief_page: dict) -> dict:
    """Extract the proposal-only fields from a Brief page to overlay on the
    Project dict in render_from_notion.merge_brief_into_project().

    v2.1.8: 'Project Street' on the merged dict now sources from the Brief's
    'Project Address' property. The renderer still reads project.get(
    'Project Street') downstream — the rename only happens in the Notion
    schema, not in the renderer's vocabulary."""
    bp = brief_page["properties"]
    return {
        "City":            text_val(bp, config.BriefProp.CITY),
        "State":           text_val(bp, config.BriefProp.STATE),
        "Jurisdiction":    text_val(bp, config.BriefProp.JURISDICTION),
        "ICC Code Year":   select_val(bp, config.BriefProp.ICC_CODE_YEAR) or "",
        "Project Type":    select_val(bp, config.BriefProp.PROJECT_TYPE) or "",
        # v2.1.8: address now sourced from Brief.Project Address, not
        # Project.Project Street. Brief is canonical for proposal context.
        "Project Street":  text_val(bp, config.BriefProp.PROJECT_ADDRESS),
    }


def job_b(dry: bool) -> JobResult:
    r = JobResult()
    filt = {
        "and": [
            {"property": config.BriefProp.STATUS, "select": {"equals": config.BriefStatus.APPROVED}},
            {"property": config.BriefProp.RENDERED_AT, "date": {"is_empty": True}},
        ]
    }
    try:
        briefs = query_data_source(config.BRIEF_DS_ID, filt, page_size=10)
    except NotionError as e:
        r.failed += 1
        r.errors.append(f"B query failed: {e}")
        return r

    from render_proposal_package import render_contract_package  # lazy

    for brief in briefs:
        brief_id = brief["id"]
        bprops = brief["properties"]
        brief_name = title_val(bprops, config.BriefProp.NAME)
        project_rel = fetch_relation_ids(brief_id, bprops, config.BriefProp.PROJECT)
        try:
            if not project_rel:
                raise RuntimeError("Brief has no linked Project")
            project_page = get_page(project_rel[0])
            if not checkbox_val(project_page["properties"], config.ProjectProp.FOLDER):
                r.skipped += 1
                log_row("B", "Skipped",
                        title=f"B · {brief_name} skipped (Project.Folder=false)",
                        project_id=project_rel[0], brief_id=brief_id,
                        details="Folder=false on Project; cannot render.")
                continue

            if dry:
                log.info("[dry] B would render contract for %s", brief_name)
                r.touched += 1
                continue

            # ── Resolve every relation the renderer needs ───────────────
            project_data = _resolve_project_dict(project_page)
            client_data  = _resolve_client_dict(project_page)
            contact_data = _resolve_contact_dict(project_page)
            engineer     = _resolve_engineer_dict(project_page)    # lowercase keys
            brief_overlay = _resolve_brief_overlay_dict(brief)
            brief_body   = _fetch_brief_body_markdown(brief_id)    # verbatim
            fee_lines    = _fetch_fee_lines(brief_id)

            result = render_contract_package(
                project_data=project_data,
                client_data=client_data,
                contact_data=contact_data,
                engineer=engineer,
                brief=brief_overlay,
                brief_body=brief_body,
                fee_lines=fee_lines,
            )
            if isinstance(result, dict) and result.get("error"):
                raise RuntimeError(result["error"])
            contract_path = result.get("contract") if isinstance(result, dict) else result

            now = datetime.now(timezone.utc).isoformat()
            update_page(brief_id, {
                config.BriefProp.RENDERED_AT: {"date": {"start": now}},
                config.BriefProp.STATUS: {"select": {"name": config.BriefStatus.RENDERED}},
            })
            update_page(project_rel[0], {
                config.ProjectProp.STATUS: {"select": {"name": config.ProjectStatus.CONTRACT_RENDERED}},
            })
            r.touched += 1
            log_row("B", "Success",
                    title=f"B · Contract rendered for {brief_name}",
                    project_id=project_rel[0], brief_id=brief_id,
                    details=f"Contract: {contract_path}")
        except Exception as e:
            # v2.1.8: catch ALL exceptions, not just the previously-listed
            # subset. A bug in the render pipeline (e.g., a NameError in a
            # helper) used to bypass this handler, which meant the Brief
            # stayed Status=Approved + Rendered At=empty AND the Automation
            # Log got no failure row. The next sweep would re-fire B, the
            # render would fail again, leaving versioned-but-blank contract
            # files accumulating in OneDrive. Catch broadly so failures are
            # always logged and the loop continues to the next Brief.
            r.failed += 1
            r.errors.append(f"{type(e).__name__}: {e}")
            log_row("B", "Failed",
                    title=f"B · Contract render failed for {brief_name}",
                    brief_id=brief_id, error=f"{type(e).__name__}: {e}")
    return r


# ========================================================================
# Orchestrator
# ========================================================================


JOB_FUNCS = {
    "A0": job_a0,
    "A02": job_a02,
    "A": job_a,
    "C": job_c,
    "D": job_d,
    "B": job_b,
}
JOB_ORDER = ["A0", "A02", "A", "C", "D", "B"]


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Kingdom Structural automation sweep v2.1.1")
    ap.add_argument("--dry-run", action="store_true", help="Show work without mutating anything")
    ap.add_argument("--jobs", default=",".join(JOB_ORDER),
                    help="Comma-separated subset of jobs to run (default: all, in order)")
    ap.add_argument("--skip-preflight", action="store_true",
                    help="Skip the preflight health check at startup. Use only "
                         "when you know the environment is good and want to "
                         "shave a few seconds off a one-off run.")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    # Logging: always to sweep.log (so pythonw.exe / windowless launches
    # still capture output), plus to stdout when one exists (manual runs
    # in PowerShell). pythonw.exe sets sys.stdout = None, so we test for it.
    handlers: list[logging.Handler] = []
    log_path = os.path.join(_SCRIPT_DIR, "sweep.log")
    try:
        handlers.append(logging.FileHandler(log_path, encoding="utf-8"))
    except OSError:
        pass  # if log file is locked, just skip — better than crashing
    if sys.stdout is not None:
        handlers.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=handlers,
    )

    # Preflight: validate token, integration identity, DB access, and template
    # before doing any expensive work. Bails out loudly if anything's broken.
    # Bypassable with --skip-preflight or env KS_SKIP_PREFLIGHT=1.
    if not args.skip_preflight:
        try:
            from preflight import run_preflight, PreflightError
            run_preflight(deep=False, raise_on_fail=True)
        except PreflightError as e:
            print(f"\nABORTING SWEEP: {e}", file=sys.stderr)
            return 2
        except Exception as e:
            log.warning("preflight raised unexpected exception (continuing): %s", e)

    requested = [j.strip() for j in args.jobs.split(",") if j.strip()]
    unknown = [j for j in requested if j not in JOB_FUNCS]
    if unknown:
        print(f"Unknown jobs: {unknown}. Valid: {list(JOB_FUNCS)}", file=sys.stderr)
        return 1

    # Preserve canonical order regardless of how --jobs was specified.
    plan = [j for j in JOB_ORDER if j in requested]

    t0 = time.time()
    summary: list[str] = []
    total_failed = 0
    for job in plan:
        try:
            result = JOB_FUNCS[job](dry=args.dry_run)
        except Exception as e:  # last-ditch safety net
            log.exception("Job %s crashed", job)
            result = JobResult(failed=1, errors=[str(e)])
        summary.append(result.line(job))
        total_failed += result.failed
        if result.errors:
            for err in result.errors[:3]:
                summary.append(f"    ! {err[:200]}")

    dt = time.time() - t0
    # Emit summary via logging (works under pythonw.exe where stdout is None)
    # AND via print() when a console exists (so manual runs see it inline).
    summary_block = "\n".join(summary)
    final_line = (f"--- sweep complete in {dt:.1f}s "
                  f"({'dry-run' if args.dry_run else 'live'}; "
                  f"{total_failed} failed row(s)) ---")
    log.info("Sweep summary:\n%s\n%s", summary_block, final_line)
    if sys.stdout is not None:
        try:
            print(summary_block)
            print(final_line)
        except (OSError, ValueError):
            pass  # console disappeared mid-run; logging above already captured it
    # rc=1 when at least one row failed, so Task Scheduler flags the run and
    # verify.py reports accurately. Non-failure runs exit 0 as normal.
    return 1 if total_failed > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
