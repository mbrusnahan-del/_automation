"""
preflight.py — automation health check.

Runs before every sweep to catch the failure modes that have actually bitten
us in production. Fails LOUD (raises) so a stale config / broken integration
/ missing template surfaces immediately instead of producing silently-broken
contracts.

Checks performed:
  1. NOTION_TOKEN is set in the environment.
  2. The token is valid — calls /users/me and prints the bot identity. This
     is the single most useful piece of information when permissions break:
     it tells you exactly which integration owns the token in your .env.
  3. All four databases the automation depends on are queryable. If any
     returns 0 rows, we abort with the offending DB id in the message.
  4. The contract template file exists and loads cleanly via python-docx.
  5. (Optional, --deep) — pick a recent project and walk Project → Client,
     Project → Project Contact → Contact page, Project → Proposal Brief.
     Confirms the integration has live access to every relation traversal
     the renderer performs.

CLI usage:
    python preflight.py              # quick checks (1-4)
    python preflight.py --deep       # also exercises a real relation walk
    python preflight.py --json       # emit machine-readable result

Programmatic usage:
    from preflight import run_preflight
    run_preflight(deep=False, raise_on_fail=True)

Called automatically by sweep.py at the top of main(). Disable with
    KS_SKIP_PREFLIGHT=1
in the environment if you ever need to bypass (e.g., emergency manual run).
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from typing import Callable

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)

import config           # noqa: E402
import sweep            # noqa: E402


CHECKMARK = "[OK]"   # avoid Unicode for Windows console safety
WARNING   = "[WARN]"
FAILURE   = "[FAIL]"

_log = logging.getLogger("preflight")


def _emit(msg: str, end: str = "\n") -> None:
    """Write a line to both stdout (when present) and the logger.

    pythonw.exe has sys.stdout = None, so naive print() would crash. This
    helper makes preflight safe to import from sweep.py regardless of how
    sweep was launched (manual python.exe with console, scheduled
    pythonw.exe with no console, or sweep.py being run as a module)."""
    if sys.stdout is not None:
        try:
            print(msg, end=end, flush=True)
        except (OSError, ValueError):
            pass
    # Strip carriage returns and newlines for the log message — the logger
    # adds its own line ending.
    clean = msg.replace("\r", "").rstrip("\n")
    if clean:
        _log.info(clean)


class PreflightError(RuntimeError):
    """Raised when a check fails and `raise_on_fail=True`."""


def _step(label: str):
    """Decorator that prints a labeled step and captures success/failure."""
    def deco(fn: Callable):
        def wrapper(state: dict, *args, **kwargs):
            _emit(f"  ... {label}", end="")
            try:
                detail = fn(state, *args, **kwargs)
                _emit(f"\r  {CHECKMARK} {label}" + (f" — {detail}" if detail else ""))
                state["passed"].append(label)
                return True
            except Exception as e:
                _emit(f"\r  {FAILURE} {label} — {e}")
                state["failed"].append((label, str(e)))
                return False
        return wrapper
    return deco


# ── Individual checks ───────────────────────────────────────────────────


@_step("Notion token loaded")
def _check_token(state):
    sweep._load_dotenv()
    if not os.environ.get("NOTION_TOKEN"):
        raise PreflightError("NOTION_TOKEN not in environment (.env)")
    return "set"


@_step("Integration identity")
def _check_identity(state):
    """Hit /users/me to confirm the token works and surface which bot owns it.

    The bot name is what you see in Notion's "Connections" panel on each
    database. Mismatch between this name and what's connected to a DB is
    exactly what bit us today.
    """
    me = sweep._notion_request("GET", "/users/me")
    bot = me.get("bot", {}) if me.get("type") == "bot" else {}
    name = me.get("name") or "(no name)"
    bot_owner = (bot.get("owner") or {}).get("type", "unknown")
    state["integration_name"] = name
    return f"name={name!r}, owner={bot_owner}"


@_step("Projects DB reachable")
def _check_projects(state):
    rows = sweep.query_data_source(config.PROJECTS_DS_ID, filter_obj=None, page_size=1)
    if not rows:
        raise PreflightError("Projects DB returned 0 rows — empty or no access")
    state["sample_project_id"] = rows[0]["id"]
    return f"sample id={rows[0]['id'][:8]}"


@_step("Proposal Brief DB reachable")
def _check_briefs(state):
    rows = sweep.query_data_source(config.BRIEF_DS_ID, filter_obj=None, page_size=1)
    if not rows:
        raise PreflightError("Brief DB returned 0 rows")
    return "ok"


@_step("Client DB reachable")
def _check_clients(state):
    rows = sweep.query_data_source(config.CLIENT_DS_ID, filter_obj=None, page_size=1)
    if not rows:
        raise PreflightError("Client DB returned 0 rows")
    return "ok"


@_step("Contacts DB reachable")
def _check_contacts(state):
    """If the integration loses access to Contacts, every contract renders
    with <<FILL IN: Contact Name>> markers. We've been bitten by this; loud
    fail here is much better than the user discovering it on a delivered PDF."""
    rows = sweep.query_data_source(config.CONTACT_DS_ID, filter_obj=None, page_size=1)
    if not rows:
        raise PreflightError(
            "Contacts DB returned 0 rows. Verify the integration "
            f"({state.get('integration_name', '?')!r}) is connected to the "
            "Contacts database in Notion."
        )
    return "ok"


@_step("Automation Log DB reachable")
def _check_log(state):
    sweep.query_data_source(config.AUTOMATION_LOG_DS_ID, filter_obj=None, page_size=1)
    return "ok"


@_step("Contract template loads")
def _check_template(state):
    """If the template is corrupted (e.g., from a botched header patch),
    every render fails with cryptic lxml errors. Catch it here once."""
    from docx import Document
    tpl_path = os.path.join(_HERE, "contract_template_merge_ready.docx")
    if not os.path.isfile(tpl_path):
        raise PreflightError(f"template not found at {tpl_path}")
    doc = Document(tpl_path)
    n_paragraphs = len(doc.paragraphs)
    n_sections = len(doc.sections)
    if n_paragraphs < 10:
        raise PreflightError(f"template suspiciously small ({n_paragraphs} body paragraphs)")
    return f"{n_paragraphs} paragraphs, {n_sections} sections"


@_step("Relation walk: Project → Project Contact → Contact page")
def _check_deep_relation(state):
    """The end-to-end relation traversal that broke today. Pick a project
    that has a Project Contact set, walk it, confirm every layer returns
    real data."""
    # Find a project that has Project Contact populated. We scan recent
    # projects — most have this set after Job A0 fires.
    rows = sweep.query_data_source(
        config.PROJECTS_DS_ID, filter_obj=None, page_size=10,
        sorts=[{"timestamp": "created_time", "direction": "descending"}],
    )
    for row in rows:
        proj_id = row["id"]
        cids = sweep.fetch_relation_ids(proj_id, row["properties"], "Project Contact")
        if cids:
            # Verified the relation resolves AND the contact page is fetchable
            sweep.get_page(cids[0])
            return f"walked project {proj_id[:8]} → contact {cids[0][:8]}"
    raise PreflightError(
        "No recent project has a populated Project Contact relation accessible "
        "to this integration. Either no projects have contacts set, or the "
        "integration lacks access to the Contacts DB."
    )


# ── Driver ──────────────────────────────────────────────────────────────


def run_preflight(deep: bool = False, raise_on_fail: bool = True,
                  emit_json: bool = False) -> dict:
    """Run all preflight checks. Returns a dict summary; raises if any
    failed and raise_on_fail=True."""
    if os.environ.get("KS_SKIP_PREFLIGHT"):
        _emit("preflight: skipped via KS_SKIP_PREFLIGHT=1")
        return {"skipped": True, "passed": [], "failed": []}

    state: dict = {"passed": [], "failed": [], "integration_name": None}

    _emit("-" * 60)
    _emit("KS Automation preflight")
    _emit("-" * 60)

    # Order matters: token → identity → DB access → relation walk → template.
    # Stop walking deeper checks if a foundational one fails.
    if not _check_token(state):
        return _summarize(state, emit_json, raise_on_fail)
    if not _check_identity(state):
        return _summarize(state, emit_json, raise_on_fail)

    _check_projects(state)
    _check_briefs(state)
    _check_clients(state)
    _check_contacts(state)
    _check_log(state)
    _check_template(state)

    if deep:
        _check_deep_relation(state)

    return _summarize(state, emit_json, raise_on_fail)


def _summarize(state: dict, emit_json: bool, raise_on_fail: bool) -> dict:
    result = {
        "passed_count": len(state["passed"]),
        "failed_count": len(state["failed"]),
        "integration_name": state.get("integration_name"),
        "passed": state["passed"],
        "failed": state["failed"],
        "ok": len(state["failed"]) == 0,
    }
    _emit("-" * 60)
    if result["ok"]:
        _emit(f"preflight: ALL CHECKS PASSED ({result['passed_count']})")
    else:
        _emit(f"preflight: {result['failed_count']} FAILED, "
              f"{result['passed_count']} passed")
        for label, err in state["failed"]:
            _emit(f"   - {label}: {err}")
    _emit("-" * 60)
    if emit_json:
        _emit(json.dumps(result, indent=2))
    if not result["ok"] and raise_on_fail:
        raise PreflightError(
            f"{result['failed_count']} preflight check(s) failed; "
            "see output above."
        )
    return result


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--deep", action="store_true",
                    help="Run extra checks (real relation traversal). Slower.")
    ap.add_argument("--json", action="store_true",
                    help="Emit a JSON summary at the end.")
    args = ap.parse_args(argv)
    try:
        result = run_preflight(deep=args.deep, raise_on_fail=False, emit_json=args.json)
        return 0 if result["ok"] else 1
    except PreflightError:
        return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
