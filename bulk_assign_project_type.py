"""
bulk_assign_project_type.py — one-off admin script.

Reads project_type_assignments.csv (which sits next to this file) and
for each row whose `proposed_type` is a real Project Type (not UNCONFIRMED
or SKIP), looks up the matching Project page in Notion by project_number
and sets its Project Type select.

Run this from the laptop, NOT from a sandbox — it talks directly to
api.notion.com using NOTION_TOKEN from .env.

Usage:
    cd "C:\\Users\\MichaelBrusnahan\\OneDrive - Kingdom Structural LLC\\_Projects\\_automation"
    python bulk_assign_project_type.py            # dry-run, prints plan
    python bulk_assign_project_type.py --apply    # actually writes

Skips any project whose Project Type is already set, so safe to re-run.
"""
from __future__ import annotations
import csv
import json
import os
import sys
import time
import urllib.request
import urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
CSV_PATH = os.path.join(HERE, "project_type_assignments.csv")

PROJECTS_DS_ID = "262b73dc-460e-8137-b3bb-000b62103b15"


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


def notion_request(token: str, method: str, path: str, body=None) -> tuple[int, dict | str]:
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


def query_projects_by_title(token: str, title_fragment: str) -> list[dict]:
    """Query the Projects data source filtering by title contains fragment."""
    body = {
        "filter": {
            "property": "Project Name",
            "title": {"contains": title_fragment},
        },
        "page_size": 5,
    }
    # 2025-09-03 endpoint is /data_sources/{id}/query
    code, data = notion_request(token, "POST",
                                f"/data_sources/{PROJECTS_DS_ID}/query", body)
    if code != 200:
        # fall back to legacy /databases/{id}/query
        code, data = notion_request(token, "POST",
                                    f"/databases/{PROJECTS_DS_ID}/query", body)
    if code != 200:
        print(f"  [ERR] query failed ({code}): {data}")
        return []
    return data.get("results", [])


def get_project_type(page: dict) -> str | None:
    sel = page.get("properties", {}).get("Project Type", {}).get("select")
    return sel.get("name") if sel else None


def set_project_type(token: str, page_id: str, type_name: str) -> bool:
    body = {
        "properties": {
            "Project Type": {"select": {"name": type_name}}
        }
    }
    code, data = notion_request(token, "PATCH", f"/pages/{page_id}", body)
    if code != 200:
        print(f"  [ERR] update failed ({code}): {data}")
        return False
    return True


SKIP_TYPES = {"UNCONFIRMED", "SKIP", ""}


def main() -> int:
    apply_changes = "--apply" in sys.argv
    env = load_env()
    token = env.get("NOTION_TOKEN")
    if not token:
        sys.exit("ERROR: NOTION_TOKEN missing from .env")

    if not os.path.exists(CSV_PATH):
        sys.exit(f"ERROR: {CSV_PATH} not found.")

    rows = []
    with open(CSV_PATH, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    print(f"Loaded {len(rows)} rows from CSV.")
    if not apply_changes:
        print("DRY RUN — pass --apply to actually write.\n")

    applied = 0
    skipped_unconfirmed = 0
    skipped_already_set = 0
    skipped_not_found = 0
    skipped_notes_only = 0
    errors = 0

    for row in rows:
        number = (row.get("project_number") or "").strip()
        ptype = (row.get("proposed_type") or "").strip()
        notes = (row.get("notes") or "").strip()
        if not number:
            continue
        if ptype in SKIP_TYPES:
            skipped_unconfirmed += 1
            print(f"  [SKIP] {number}: {ptype} — {notes}")
            continue
        if "APPLIED" in notes:
            skipped_notes_only += 1
            print(f"  [DONE] {number} — already applied this session")
            continue

        # Find the project
        results = query_projects_by_title(token, number)
        if not results:
            skipped_not_found += 1
            print(f"  [MISS] {number}: not found in Notion")
            continue

        # Pick best match by exact title prefix
        match = None
        for p in results:
            title_arr = p.get("properties", {}).get("Project Name", {}).get("title", [])
            full_title = "".join(t.get("plain_text", "") for t in title_arr).strip()
            if full_title.startswith(number):
                match = p
                break
        if match is None and len(results) == 1:
            match = results[0]
        if match is None:
            skipped_not_found += 1
            print(f"  [AMBIG] {number}: {len(results)} matches, no exact prefix")
            continue

        existing = get_project_type(match)
        if existing:
            skipped_already_set += 1
            print(f"  [SET ] {number}: already '{existing}', skipping")
            continue

        # Apply
        if apply_changes:
            ok = set_project_type(token, match["id"], ptype)
            if ok:
                applied += 1
                print(f"  [WRITE] {number} -> {ptype}")
            else:
                errors += 1
            time.sleep(0.15)  # gentle pace under Notion's 3 req/sec limit
        else:
            applied += 1
            print(f"  [PLAN] {number} -> {ptype}")

    print()
    print(f"Summary:")
    print(f"  {'Applied' if apply_changes else 'Would apply'}: {applied}")
    print(f"  Skipped (unconfirmed/skip):  {skipped_unconfirmed}")
    print(f"  Skipped (already set):       {skipped_already_set}")
    print(f"  Skipped (already done):      {skipped_notes_only}")
    print(f"  Not found in Notion:         {skipped_not_found}")
    print(f"  Errors:                      {errors}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
