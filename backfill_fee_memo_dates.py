"""
backfill_fee_memo_dates.py — one-off admin script.

Finds Projects where:
  - Fee Memo Generated At is empty
  - Folder = true
  - The PDF actually exists in the project's OneDrive folder

For each match, stamps Fee Memo Generated At with today's date so the project
drops out of future Job F sweeps. Idempotent — re-running just no-ops on
projects that already have the date.

Use this once after deploying v2.1.11 to clean up the back-catalog of
projects whose PDFs were generated before the date-stamp logic was added.

Usage:
    cd "C:\\Users\\...\\_automation"
    python backfill_fee_memo_dates.py            # dry-run, prints plan
    python backfill_fee_memo_dates.py --apply    # actually writes
"""
from __future__ import annotations
import json
import os
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")

PROJECTS_DS_ID = "262b73dc-460e-8137-b3bb-000b62103b15"

# Mirror config.py constants — kept inline so this script has no internal deps
ONEDRIVE_ROOT = os.path.join(
    os.environ.get("KS_ONEDRIVE_WIN_ROOT",
                   r"C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC"),
    "_Projects",
)


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


def notion(token: str, method: str, path: str, body=None) -> tuple[int, dict | str]:
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


def query_eligible_projects(token: str) -> list[dict]:
    """Return all Projects where Fee Memo Generated At is empty and Folder is true."""
    body = {
        "filter": {
            "and": [
                {"property": "Folder", "checkbox": {"equals": True}},
                {"property": "Fee Memo Generated At", "date": {"is_empty": True}},
            ]
        },
        "page_size": 100,
    }
    results = []
    next_cursor = None
    while True:
        if next_cursor:
            body["start_cursor"] = next_cursor
        code, data = notion(token, "POST",
                            f"/data_sources/{PROJECTS_DS_ID}/query", body)
        if code != 200:
            # Fall back to legacy endpoint
            code, data = notion(token, "POST",
                                f"/databases/{PROJECTS_DS_ID}/query", body)
        if code != 200:
            sys.exit(f"ERROR querying projects: {code} {data}")
        results.extend(data.get("results", []))
        if not data.get("has_more"):
            break
        next_cursor = data.get("next_cursor")
    return results


def parse_project_name(name: str) -> tuple[str, str] | None:
    """Return (8-digit number, short name) or None if format is bad."""
    name = name.strip()
    if len(name) < 9 or not name[:8].isdigit():
        return None
    return name[:8], name[8:].lstrip(" -").strip()


def find_pdf(number: str) -> str | None:
    """Return the path to {number} Fee Analysis Memo.pdf if it exists, else None.
    Uses prefix-scan to handle historical folder name drift."""
    year = "20" + number[:2]
    year_dir = os.path.join(ONEDRIVE_ROOT, year)
    if not os.path.isdir(year_dir):
        return None
    for entry in os.listdir(year_dir):
        if not entry.startswith(number):
            continue
        folder = os.path.join(year_dir, entry)
        if not os.path.isdir(folder):
            continue
        pdf_path = os.path.join(folder, f"{number} Fee Analysis Memo.pdf")
        if os.path.exists(pdf_path):
            return pdf_path
    return None


def stamp_project(token: str, page_id: str, today_iso: str) -> bool:
    body = {
        "properties": {
            "Fee Memo Generated At": {"date": {"start": today_iso}},
        }
    }
    code, _ = notion(token, "PATCH", f"/pages/{page_id}", body)
    return code == 200


def main() -> int:
    apply_changes = "--apply" in sys.argv
    env = load_env()
    token = env.get("NOTION_TOKEN")
    if not token:
        sys.exit("ERROR: NOTION_TOKEN missing from .env")

    print(f"OneDrive root: {ONEDRIVE_ROOT}")
    print(f"Mode: {'APPLY' if apply_changes else 'DRY-RUN (use --apply to write)'}")
    print()

    print("Querying Projects where Fee Memo Generated At is empty + Folder=true...")
    projects = query_eligible_projects(token)
    print(f"Found {len(projects)} candidate projects.\n")

    today_iso = datetime.now(timezone.utc).date().isoformat()
    will_stamp = []
    no_pdf = []
    bad_name = []

    for p in projects:
        props = p["properties"]
        title_arr = props.get("Project Name", {}).get("title", [])
        name = "".join(t.get("plain_text", "") for t in title_arr).strip()
        if not name:
            continue

        parsed = parse_project_name(name)
        if not parsed:
            bad_name.append(name)
            continue
        number, _short = parsed

        pdf_path = find_pdf(number)
        if pdf_path is None:
            no_pdf.append(name)
            continue

        will_stamp.append({
            "page_id": p["id"],
            "name": name,
            "number": number,
            "pdf": pdf_path,
        })

    print(f"Will stamp:           {len(will_stamp)} projects")
    print(f"No PDF found:         {len(no_pdf)} (left untouched)")
    print(f"Bad project name:     {len(bad_name)} (left untouched)")
    print()

    if will_stamp:
        print("Stamping plan:")
        for row in will_stamp:
            print(f"  {row['number']}  {row['name'][:60]}")
        print()

    if no_pdf:
        print("Projects without PDFs (Job F will pick these up later):")
        for n in no_pdf[:10]:
            print(f"  - {n}")
        if len(no_pdf) > 10:
            print(f"  ... and {len(no_pdf) - 10} more")
        print()

    if bad_name:
        print("Projects with bad names (manual fix needed):")
        for n in bad_name[:5]:
            print(f"  - {n}")
        if len(bad_name) > 5:
            print(f"  ... and {len(bad_name) - 5} more")
        print()

    if not apply_changes:
        print("DRY RUN — no changes written. Re-run with --apply to stamp dates.")
        return 0

    if not will_stamp:
        print("Nothing to stamp.")
        return 0

    print(f"Stamping {len(will_stamp)} projects with date {today_iso}...")
    succeeded = 0
    failed = 0
    for row in will_stamp:
        ok = stamp_project(token, row["page_id"], today_iso)
        if ok:
            succeeded += 1
            print(f"  [OK] {row['number']} stamped")
        else:
            failed += 1
            print(f"  [ERR] {row['number']} stamp failed")
        time.sleep(0.15)  # gentle on Notion's rate limit

    print()
    print(f"Stamped:  {succeeded}")
    print(f"Failed:   {failed}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
