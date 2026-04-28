"""
One-time cleanup: archive broken Fee Schedule child databases.

Context
-------
During the launch shakedown, several Briefs ended up with a Fee Schedule
child database that had an empty schema (no Service / Include / Order /
Type / Amount properties). That happened because early sweep.py runs
used the wrong JSON shape for the Notion 2025-09-03 POST /databases
endpoint.

After the `initial_data_source` fix, the next sweep created a *second*,
working Fee Schedule on each affected Brief — alongside the broken one.
This script finds and archives (moves to Notion Trash) the broken ones
so every Brief has exactly one usable Fee Schedule DB.

Usage
-----
    python cleanup_fee_schedules.py            # dry-run — list what would be archived
    python cleanup_fee_schedules.py --live     # actually archive them

The script is idempotent; re-running after a live pass produces an empty
report.

What counts as "broken"
-----------------------
A child_database block titled "💵 Fee Schedule" is flagged when its
first data source's property schema does NOT include a "Service" column.
Those are the empty-schema shells from the bugged create_child_database
call. Working Fee Schedules (whether created by the old MCP runs or by
the post-fix sweep) always have `Service` and are left alone.

Already-archived / already-in-trash blocks are skipped.
"""
from __future__ import annotations

import argparse
import os
import sys

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)

import config  # noqa: E402
import sweep   # noqa: E402


def _iter_all_briefs():
    """Page through every Brief in the Proposal Brief DB."""
    cursor: str | None = None
    while True:
        body: dict = {"page_size": 100}
        if cursor:
            body["start_cursor"] = cursor
        resp = sweep._notion_request(
            "POST", f"/data_sources/{config.BRIEF_DS_ID}/query", body
        )
        for row in resp.get("results", []):
            yield row
        if not resp.get("has_more"):
            return
        cursor = resp.get("next_cursor")


def _is_fee_schedule_broken(block: dict) -> tuple[bool, str]:
    """Given a child_database block, verify its data source has `Service`.

    Returns (is_broken, reason).
    """
    db_id = block["id"]
    try:
        db = sweep._notion_request("GET", f"/databases/{db_id}")
    except sweep.NotionError as e:
        return False, f"GET /databases/{db_id[:8]} failed: {e}"
    ds_list = db.get("data_sources") or []
    if not ds_list:
        return True, "no data_sources on database response"
    ds_id = ds_list[0]["id"]
    try:
        ds = sweep._notion_request("GET", f"/data_sources/{ds_id}")
    except sweep.NotionError as e:
        return False, f"GET /data_sources/{ds_id[:8]} failed: {e}"
    props = ds.get("properties") or {}
    if "Service" in props:
        return False, "has Service property (usable)"
    return True, f"missing Service property; keys: {sorted(props.keys())}"


def _find_broken_fee_schedules_on(brief: dict) -> list[tuple[str, str]]:
    """Return list of (block_id, reason) for broken Fee Schedule DBs on a Brief."""
    brief_id = brief["id"]
    try:
        blocks = sweep.get_block_children(brief_id)
    except sweep.NotionError as e:
        print(f"  ! could not list blocks on brief {brief_id[:8]}: {e}")
        return []
    out: list[tuple[str, str]] = []
    for b in blocks:
        if b.get("archived") or b.get("in_trash"):
            continue
        if b.get("type") != "child_database":
            continue
        title = b.get("child_database", {}).get("title", "")
        if title != sweep.FEE_SCHEDULE_TITLE:
            continue
        broken, reason = _is_fee_schedule_broken(b)
        if broken:
            out.append((b["id"], reason))
    return out


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(
        description="Archive broken Fee Schedule child DBs on Briefs."
    )
    ap.add_argument("--live", action="store_true",
                    help="Actually archive (default is dry-run).")
    args = ap.parse_args(argv)

    mode = "LIVE" if args.live else "dry-run"
    print(f"Fee Schedule cleanup — {mode}\n")

    total_briefs = 0
    affected: list[tuple[str, str, str]] = []  # (brief_name, block_id, reason)

    for brief in _iter_all_briefs():
        total_briefs += 1
        brief_name = sweep.title_val(brief["properties"], config.BriefProp.NAME)
        broken = _find_broken_fee_schedules_on(brief)
        for block_id, reason in broken:
            affected.append((brief_name, block_id, reason))

    print(f"Scanned {total_briefs} Briefs; found {len(affected)} broken Fee Schedule DB(s).\n")

    if not affected:
        print("Nothing to clean up. ✓")
        return 0

    print("Broken DBs found:")
    for name, block_id, reason in affected:
        print(f"  • {name}")
        print(f"      block {block_id}")
        print(f"      reason: {reason}")

    if not args.live:
        print(f"\n(dry-run) Re-run with --live to archive the {len(affected)} broken DB(s).")
        return 0

    print(f"\nArchiving {len(affected)} broken DB(s)...")
    ok, fail = 0, 0
    for name, block_id, _ in affected:
        try:
            sweep._notion_request(
                "PATCH", f"/blocks/{block_id}", {"archived": True}
            )
            print(f"  ✓ archived on {name}")
            ok += 1
        except sweep.NotionError as e:
            print(f"  ✗ failed on {name}: {e}")
            fail += 1

    print(f"\nDone: {ok} archived, {fail} failed.")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
