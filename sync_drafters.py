"""
sync_drafters.py — keep Excel Drafter column in sync with Notion DRAFTER property.

Reads every project in the Notion Projects DB. For each, extracts:
  - 8-digit job number (from Project Name)
  - DRAFTER (Notion People property — list of user IDs)

Resolves each user ID to a human name via Notion's /users API.

Then opens Master-Business Plan.xlsm, finds the Jobs 26KS sheet, auto-detects
the Drafter column (by scanning row 5 headers for "drafter"), and updates
each row:
  - 0 drafters in Notion → leave Excel cell alone
  - 1 drafter in Notion  → write that name to Excel if cell is empty or different
  - 2+ drafters in Notion → leave Excel cell BLANK + log to skipped list

After running, a CSV is written to drafter_sync_skipped.csv listing every
project that got skipped (multiple drafters, no Notion match, etc.) so
Michael can review and decide manually.

Usage:
    cd "C:\\Users\\MichaelBrusnahan\\OneDrive - Kingdom Structural LLC\\_Projects\\_automation"
    python sync_drafters.py            # dry-run — shows plan, no writes
    python sync_drafters.py --apply    # actually writes Excel + notification CSV

Requires Excel CLOSED during run.
"""
from __future__ import annotations
import csv
import json
import os
import re
import sys
import time
import urllib.request
import urllib.error
from collections import defaultdict
from datetime import datetime

from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
SKIPPED_CSV = os.path.join(HERE, "drafter_sync_skipped.csv")

PROJECTS_DS_ID = "262b73dc-460e-8137-b3bb-000b62103b15"
JOBS_SHEET_NAME = "Jobs 26KS"
HEADER_ROW = 5  # Row with column headers in Jobs 26KS

# Business Plan folder is a SIBLING of _Projects, so we go up two levels
# from _automation (which sits inside _Projects).
_PROJECTS_ROOT = os.path.dirname(HERE)
_ONEDRIVE_ROOT = os.path.dirname(_PROJECTS_ROOT)
EXCEL_CANDIDATES = [
    os.path.join(_ONEDRIVE_ROOT, "Business Plan", "Master-Business Plan.xlsm"),
    # Legacy fallback
    os.path.join(_PROJECTS_ROOT, "Business Plan", "Master-Business Plan.xlsm"),
]


def load_env() -> dict:
    if not os.path.exists(ENV_PATH):
        sys.exit(f"ERROR: {ENV_PATH} not found.")
    env = {}
    with open(ENV_PATH) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def notion(token: str, method: str, path: str, body=None):
    url = "https://api.notion.com/v1" + path
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {
        "Authorization": "Bearer " + token,
        "Notion-Version": "2025-09-03",
        "Accept": "application/json",
    }
    if data is not None:
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8")
            return resp.status, json.loads(raw) if raw else {}
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8")
        try:
            return e.code, json.loads(raw)
        except Exception:
            return e.code, raw


def fetch_users(token: str) -> dict[str, str]:
    """Return {user_id: full_name}."""
    body = {"page_size": 100}
    out = {}
    while True:
        # /users uses GET, not POST
        params = "?page_size=100"
        if "start_cursor" in body:
            params += "&start_cursor=" + body["start_cursor"]
        code, data = notion(token, "GET", "/users" + params)
        if code != 200:
            sys.exit(f"ERROR fetching users: {code} {data}")
        for u in data.get("results", []):
            uid = u.get("id")
            name = (u.get("name") or "").strip()
            if uid and name:
                out[uid] = name
        if not data.get("has_more"):
            break
        body["start_cursor"] = data.get("next_cursor")
    return out


def fetch_projects(token: str) -> list[dict]:
    """Return all Project pages with their DRAFTER + Project Name."""
    body = {
        "filter": {"property": "Project Name", "title": {"is_not_empty": True}},
        "page_size": 100,
    }
    out = []
    while True:
        code, data = notion(token, "POST",
                            f"/data_sources/{PROJECTS_DS_ID}/query", body)
        if code != 200:
            code, data = notion(token, "POST",
                                f"/databases/{PROJECTS_DS_ID}/query", body)
        if code != 200:
            sys.exit(f"ERROR querying projects: {code} {data}")
        out.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        body["start_cursor"] = data.get("next_cursor")
    return out


def parse_job_number(name: str) -> str | None:
    name = (name or "").strip()
    m = re.match(r"^(\d{8})", name)
    return m.group(1) if m else None


def to_initials(full_name: str) -> str:
    """Convert 'Marshall Foschini' -> 'MF'. First letter of first word +
    first letter of last word. Single-word names return that single letter
    uppercased. Empty / unknown returns ''."""
    if not full_name:
        return ""
    parts = [p for p in re.split(r"\s+", full_name.strip()) if p]
    if not parts:
        return ""
    if len(parts) == 1:
        return parts[0][0].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def find_excel():
    for p in EXCEL_CANDIDATES:
        if os.path.isfile(p):
            return p
    sys.exit("ERROR: Master-Business Plan.xlsm not found.")


def find_drafter_column(ws) -> tuple[int, str] | None:
    """Scan header row for any cell containing 'drafter' (case-insensitive).
    Returns (column_index, header_text) or None."""
    for col_idx in range(1, ws.max_column + 1):
        val = ws.cell(HEADER_ROW, col_idx).value
        if val and "drafter" in str(val).lower():
            return col_idx, str(val).strip()
    return None


def main() -> int:
    apply_changes = "--apply" in sys.argv
    env = load_env()
    token = env.get("NOTION_TOKEN")
    if not token:
        sys.exit("ERROR: NOTION_TOKEN missing from .env")

    print(f"Mode: {'APPLY' if apply_changes else 'DRY-RUN (use --apply to write)'}")
    print()

    # Step 1: pull Notion users (id -> name)
    print("Fetching Notion users...")
    users = fetch_users(token)
    print(f"  Loaded {len(users)} users.")

    # Step 2: pull all Notion projects with DRAFTER property
    print("Fetching projects from Notion...")
    projects = fetch_projects(token)
    print(f"  Loaded {len(projects)} projects.")

    # Build {job_number: [drafter_names]}
    notion_drafters: dict[str, list[str]] = {}
    notion_pages: dict[str, dict] = {}
    for page in projects:
        props = page["properties"]
        title_arr = props.get("Project Name", {}).get("title", [])
        name = "".join(t.get("plain_text", "") for t in title_arr).strip()
        job = parse_job_number(name)
        if not job:
            continue
        people = props.get("DRAFTER", {}).get("people", [])
        drafter_names = [users.get(p.get("id"), "(unknown user)") for p in people]
        notion_drafters[job] = drafter_names
        notion_pages[job] = {"name": name, "drafters": drafter_names}

    print(f"  {len(notion_drafters)} projects have parseable 8-digit job numbers.\n")

    # Step 3: open Excel and find Drafter column
    excel_path = find_excel()
    print(f"Loading workbook: {excel_path}")
    wb = load_workbook(excel_path, keep_vba=True)
    if JOBS_SHEET_NAME not in wb.sheetnames:
        sys.exit(f"ERROR: '{JOBS_SHEET_NAME}' sheet not found.")
    ws = wb[JOBS_SHEET_NAME]
    drafter_col_info = find_drafter_column(ws)
    if not drafter_col_info:
        # Fall back: print all headers and exit
        print("\nERROR: Could not auto-detect a Drafter column in row 5. "
              "Found these column headers:")
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(HEADER_ROW, col_idx).value
            if val:
                from openpyxl.utils import get_column_letter
                print(f"  Column {get_column_letter(col_idx)}: {val}")
        sys.exit("Add a column named 'Drafter' (or similar) in row 5 and re-run.")
    drafter_col, header_text = drafter_col_info
    from openpyxl.utils import get_column_letter
    print(f"  Drafter column found: {get_column_letter(drafter_col)} ('{header_text}').")

    # Find Job column too (always B per existing convention)
    job_col = 2  # Column B

    # Step 4: walk Excel rows and decide per-row action
    skipped = []   # for notification CSV
    will_update = []
    no_change = 0
    no_notion_match = 0
    multi_drafter = 0

    for row_idx in range(HEADER_ROW + 1, ws.max_row + 1):
        job_raw = ws.cell(row_idx, job_col).value
        if job_raw is None:
            continue
        job = str(job_raw).strip().split(".")[0]  # strip .1/.2 suffix
        if not re.match(r"^\d{8}$", job):
            continue

        excel_drafter = (ws.cell(row_idx, drafter_col).value or "").strip() if ws.cell(row_idx, drafter_col).value else ""
        notion_list = notion_drafters.get(job)

        if notion_list is None:
            no_notion_match += 1
            continue

        if len(notion_list) == 0:
            # No drafter in Notion — leave Excel alone
            no_change += 1
            continue

        if len(notion_list) >= 2:
            multi_drafter += 1
            skipped.append({
                "job": job,
                "project_name": notion_pages.get(job, {}).get("name", ""),
                "row": row_idx,
                "current_excel": excel_drafter,
                "notion_drafters": " | ".join(notion_list),
                "reason": "multiple drafters in Notion",
            })
            # Per spec: leave Excel BLANK
            if excel_drafter:
                will_update.append({
                    "job": job, "row": row_idx,
                    "old": excel_drafter, "new": "(BLANK)",
                    "reason": "multi-drafter — clearing per policy",
                })
            continue

        # Exactly 1 drafter in Notion. Excel uses initials (e.g. 'MF') and
        # Notion has the full name ('Marshall Foschini'). Convert before
        # comparing so we don't churn every row.
        full_name = notion_list[0]
        new_value = to_initials(full_name)
        excel_norm = excel_drafter.upper()
        if excel_norm == new_value:
            no_change += 1
        elif excel_drafter == "":
            will_update.append({
                "job": job, "row": row_idx,
                "old": "(blank)", "new": new_value,
                "reason": f"filling missing drafter ({full_name})",
            })
        else:
            # Mismatch — Excel has different initials than Notion
            will_update.append({
                "job": job, "row": row_idx,
                "old": excel_drafter, "new": new_value,
                "reason": f"Excel disagrees with Notion ({full_name})",
            })

    # Step 5: report plan
    print()
    print("=" * 70)
    print(f"PLAN: {'APPLY' if apply_changes else 'DRY-RUN'}")
    print("=" * 70)
    print(f"Excel rows reviewed:         {ws.max_row - HEADER_ROW}")
    print(f"  No matching Notion project: {no_notion_match}")
    print(f"  No change needed:           {no_change}")
    print(f"  Multi-drafter (skipped):    {multi_drafter}")
    print(f"  Will update:                {len(will_update)}")
    print()
    if will_update:
        print("UPDATES:")
        for u in will_update[:30]:
            print(f"  Row {u['row']:<5} job {u['job']}: '{u['old']}' -> '{u['new']}'  ({u['reason']})")
        if len(will_update) > 30:
            print(f"  ... and {len(will_update) - 30} more")
    print()
    if skipped:
        print(f"NOTIFICATION: {len(skipped)} projects skipped due to multiple drafters in Notion.")
        for s in skipped[:10]:
            print(f"  {s['job']} {s['project_name'][:50]} -> {s['notion_drafters']}")
        if len(skipped) > 10:
            print(f"  ... and {len(skipped) - 10} more (full list in {os.path.basename(SKIPPED_CSV)})")

    if not apply_changes:
        print("\nDRY-RUN — no changes written. Re-run with --apply to write Excel + notification CSV.")
        return 0

    # Step 6: apply changes
    print("\nWriting changes to Excel...")
    for u in will_update:
        if u["new"] == "(BLANK)":
            ws.cell(u["row"], drafter_col).value = None
        else:
            ws.cell(u["row"], drafter_col).value = u["new"]

    print(f"Saving workbook: {excel_path}")
    wb.save(excel_path)

    # Write notification CSV
    if skipped:
        with open(SKIPPED_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["job", "project_name", "row",
                                                    "current_excel", "notion_drafters", "reason"])
            writer.writeheader()
            writer.writerows(skipped)
        print(f"Notification CSV: {SKIPPED_CSV}")

    print(f"\nDone. {len(will_update)} updated, {len(skipped)} flagged for review.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
