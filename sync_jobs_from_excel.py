"""
sync_jobs_from_excel.py
========================

Excel "Jobs 26KS" sheet → Notion "Jobs 26KS" database.

Scope (v2.1.7): TOTALS ONLY.
  - Reads Job, Name, Client, TOTAL_SIGNED from each Excel row.
  - Matches Excel Job# to Notion Job# (string match, exact — no rollup
    of .1/.2 add-service rows; those are first-class rows in Notion).
  - For matches: fill TOTAL SIGNED ONLY when the Notion field is blank.
    Notion always wins on populated fields ("fill-blanks-only" policy).
  - For Excel rows with no Notion match: CREATE a new Jobs 26KS row
    (Name, Job, Client, Y=true, TOTAL SIGNED). This is what catches
    the "lots of new projects funneling through" backlog.

What this script does NOT touch:
  - The fee-phase columns (CD, Bid & Permit, Reimb, Design, CA) —
    out of scope per "Totals only" choice.
  - The invoice-tracking columns (UNPAID, First Invoice, Inv 1 Amount,
    Inv 1 Paid, Inv 1 Date, etc.) — driven by accounting, not Excel.
  - Existing populated Notion fields — never overwritten.

Excel layout ('Jobs 26KS' sheet)
--------------------------------
Row 5 = header, rows 6+ = data. Relevant columns for this sync:
  B: Job          (text — "24008", "240038.2")
  E: Client       (number — internal client ID)
  F: Name         (project name)
  V: TOTAL_SIGNED (number, dollar)

Notion mapping
--------------
  Excel "Job"          → Notion "Job"          (rich_text, match key)
  Excel "Name"         → Notion "Name"         (title)
  Excel "Client"       → Notion "Client"       (number)
  Excel "TOTAL_SIGNED" → Notion "TOTAL SIGNED" (number, dollar)
  (new rows only)      → Notion "Y"            (checkbox = True)

Usage
-----
    python sync_jobs_from_excel.py                          # dry-run
    python sync_jobs_from_excel.py --live                   # apply
    python sync_jobs_from_excel.py --excel /path/to.xlsm    # custom path
    python sync_jobs_from_excel.py --job 25018032           # single job
    python sync_jobs_from_excel.py --no-create              # only fill blanks
    python sync_jobs_from_excel.py --force                  # ignore mtime cache
    python sync_jobs_from_excel.py --verbose
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)

import config   # noqa: E402
import sweep    # noqa: E402
sweep._load_dotenv()

from openpyxl import load_workbook  # noqa: E402

log = logging.getLogger("jobs_sync")


# ────────────────────────────────────────────────────────────────────────
# Configuration
# ────────────────────────────────────────────────────────────────────────

JOBS_SHEET    = "Jobs 26KS"
HEADER_ROW    = 5
DATA_START    = 6

# Excel column letters
COL_JOB           = "B"
COL_CLIENT        = "E"
COL_NAME          = "F"
COL_TOTAL_SIGNED  = "V"

# Notion property names (case-sensitive — must match the Jobs 26KS DB)
NP_NAME          = "Name"          # title
NP_JOB           = "Job"           # rich_text (match key)
NP_CLIENT        = "Client"        # number
NP_TOTAL_SIGNED  = "TOTAL SIGNED"  # number (dollar)
NP_Y             = "Y"             # checkbox (set True on new rows)

EXCEL_PATH_CANDIDATES = [
    os.path.join(os.path.dirname(os.path.dirname(_SCRIPT_DIR)),
                 "Business Plan", "Master-Business Plan.xlsm"),
]

STATE_PATH = os.path.join(_SCRIPT_DIR, "jobs_sync_state.json")


# ────────────────────────────────────────────────────────────────────────
# Data classes
# ────────────────────────────────────────────────────────────────────────


@dataclass
class ParsedJob:
    row_idx: int
    job: str                      # "24008", "240038.2" — text, no formatting
    name: str
    client: float | None = None   # nullable number
    total_signed: float | None = None


@dataclass
class SyncStats:
    jobs_scanned: int = 0
    jobs_created: int = 0
    jobs_updated: int = 0      # TOTAL SIGNED filled on existing job
    jobs_skipped: int = 0      # nothing to do
    errors: list[str] = field(default_factory=list)


# ────────────────────────────────────────────────────────────────────────
# Excel parsing
# ────────────────────────────────────────────────────────────────────────


def _col_index(letter: str) -> int:
    """Excel column letter → 0-based row index."""
    letter = letter.upper()
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _stringify_job(v) -> str:
    """Excel job cells can be int/float/text. Normalize to string with
    no trailing .0 on whole numbers."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return str(v)
    return str(v).strip()


def _to_float(v) -> float | None:
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def read_jobs_sheet(excel_path: str) -> list[ParsedJob]:
    wb = load_workbook(excel_path, data_only=True, read_only=True, keep_vba=False)
    if JOBS_SHEET not in wb.sheetnames:
        raise RuntimeError(
            f"'{JOBS_SHEET}' sheet not found in {excel_path}. "
            f"Available: {wb.sheetnames}"
        )
    ws = wb[JOBS_SHEET]

    job_i    = _col_index(COL_JOB)
    client_i = _col_index(COL_CLIENT)
    name_i   = _col_index(COL_NAME)
    total_i  = _col_index(COL_TOTAL_SIGNED)

    jobs: list[ParsedJob] = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx < DATA_START:
            continue

        def _cell(i):
            return row[i] if i < len(row) else None

        job  = _stringify_job(_cell(job_i))
        name = str(_cell(name_i) or "").strip()
        if not job and not name:
            continue   # blank row

        jobs.append(ParsedJob(
            row_idx=row_idx,
            job=job,
            name=name,
            client=_to_float(_cell(client_i)),
            total_signed=_to_float(_cell(total_i)),
        ))
    return jobs


# ────────────────────────────────────────────────────────────────────────
# Notion side
# ────────────────────────────────────────────────────────────────────────


def _job_text(props: dict) -> str:
    """Read the 'Job' rich_text property as a plain string."""
    rich = props.get(NP_JOB, {}).get("rich_text") or []
    return "".join(r.get("plain_text", "") for r in rich).strip()


def _number(props: dict, name: str) -> float | None:
    return (props.get(name, {}) or {}).get("number")


def load_existing_jobs() -> dict[str, dict]:
    """Index every Jobs 26KS row by its Job# (case-insensitive, trimmed).
    Returns {job#: page}. Last-write-wins on duplicate Job#s — surfaces a
    warning so partners can clean up the dupe in Notion."""
    out: dict[str, dict] = {}
    dupes: list[str] = []
    rows = sweep.query_data_source(
        config.JOBS_26KS_DS_ID, filter_obj=None, page_size=100,
    )
    for page in rows:
        job = _job_text(page["properties"]).lower()
        if not job:
            continue
        if job in out:
            dupes.append(job)
        out[job] = page
    if dupes:
        log.warning("Jobs 26KS has %d duplicate Job#: %s", len(dupes), dupes[:5])
    return out


def _build_create_payload(p: ParsedJob) -> dict:
    payload = {
        NP_NAME: {"title": [{"type": "text",
                             "text": {"content": p.name or p.job}}]},
        NP_JOB:  {"rich_text": [{"type": "text",
                                  "text": {"content": p.job}}]},
        NP_Y:    {"checkbox": True},
    }
    if p.client is not None:
        payload[NP_CLIENT] = {"number": p.client}
    if p.total_signed is not None:
        payload[NP_TOTAL_SIGNED] = {"number": p.total_signed}
    return payload


def sync_one_job(p: ParsedJob, existing: dict[str, dict],
                 *, allow_create: bool, live: bool, stats: SyncStats):
    stats.jobs_scanned += 1
    key = p.job.lower().strip()
    page = existing.get(key)

    if page is None:
        # No Notion row → create one (unless --no-create)
        if not allow_create:
            log.debug("[skip] No Notion row for Job#%s; --no-create set", p.job)
            stats.jobs_skipped += 1
            return
        if not live:
            log.info("[dry] would CREATE Job#%s (%s) total_signed=%s",
                     p.job, p.name, p.total_signed)
            stats.jobs_created += 1
            return
        try:
            sweep.create_page(
                {"type": "data_source_id", "data_source_id": config.JOBS_26KS_DS_ID},
                _build_create_payload(p),
            )
            log.info("  + CREATED Job#%s  %s", p.job, p.name)
            stats.jobs_created += 1
        except sweep.NotionError as e:
            stats.errors.append(f"create Job#{p.job}: {e}")
            log.error("  ! create Job#%s failed: %s", p.job, e)
        return

    # Existing row → fill blank TOTAL SIGNED only.
    notion_total = _number(page["properties"], NP_TOTAL_SIGNED)
    if notion_total is not None:
        log.debug("[skip] Job#%s TOTAL SIGNED already populated (%s)",
                  p.job, notion_total)
        stats.jobs_skipped += 1
        return
    if p.total_signed is None:
        log.debug("[skip] Job#%s blank in Notion AND blank in Excel",
                  p.job)
        stats.jobs_skipped += 1
        return

    if not live:
        log.info("[dry] would FILL Job#%s TOTAL SIGNED ← $%.2f",
                 p.job, p.total_signed)
        stats.jobs_updated += 1
        return
    try:
        sweep.update_page(page["id"], {
            NP_TOTAL_SIGNED: {"number": p.total_signed},
        })
        log.info("  ~ Job#%s  TOTAL SIGNED ← $%.2f", p.job, p.total_signed)
        stats.jobs_updated += 1
    except sweep.NotionError as e:
        stats.errors.append(f"fill Job#{p.job}: {e}")
        log.error("  ! fill Job#%s failed: %s", p.job, e)


# ────────────────────────────────────────────────────────────────────────
# State & file lookup
# ────────────────────────────────────────────────────────────────────────


def _load_state() -> dict:
    if not os.path.exists(STATE_PATH):
        return {}
    try:
        with open(STATE_PATH) as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_state(state: dict) -> None:
    try:
        with open(STATE_PATH, "w") as f:
            json.dump(state, f)
    except OSError:
        pass


def _find_excel(override: str | None) -> str | None:
    if override:
        return override if os.path.exists(override) else None
    for p in EXCEL_PATH_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


# ────────────────────────────────────────────────────────────────────────
# CLI
# ────────────────────────────────────────────────────────────────────────


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Sync Jobs 26KS financials from Excel into Notion (totals only).",
    )
    parser.add_argument("--live", action="store_true",
                        help="Apply changes (default: dry-run).")
    parser.add_argument("--excel", help="Path to Master-Business Plan.xlsm.")
    parser.add_argument("--job", help="Sync only this Job#.")
    parser.add_argument("--no-create", action="store_true",
                        help="Don't create new Notion rows; only fill blanks.")
    parser.add_argument("--force", action="store_true",
                        help="Ignore Excel mtime cache.")
    parser.add_argument("--verbose", "-v", action="store_true")
    args = parser.parse_args(argv)

    log_path = os.path.join(_SCRIPT_DIR, "jobs_sync.log")
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
    )
    log.info("----- jobs sync run starting (live=%s, job=%s) -----",
             args.live, args.job or "all")

    excel_path = _find_excel(args.excel)
    if not excel_path:
        print("ERROR: Master-Business Plan.xlsm not found. Tried:", file=sys.stderr)
        for p in EXCEL_PATH_CANDIDATES:
            print(f"  {p}", file=sys.stderr)
        print("Pass --excel to specify an explicit path.", file=sys.stderr)
        return 1

    mtime = os.path.getmtime(excel_path)
    state = _load_state()
    if (state.get("last_mtime") == mtime and not args.force and not args.job):
        print(f"No change since last sync (mtime={datetime.fromtimestamp(mtime)}).",
              "Use --force to re-run.")
        return 0

    if not os.environ.get("NOTION_TOKEN") and args.live:
        print("ERROR: NOTION_TOKEN not set. Check your .env file.", file=sys.stderr)
        return 1

    print(f"Excel: {excel_path}")
    print(f"Mode:  {'LIVE' if args.live else 'DRY-RUN'}")
    if args.job:
        print(f"Filter: only Job#'{args.job}'")
    print()

    try:
        jobs = read_jobs_sheet(excel_path)
    except Exception as e:
        print(f"ERROR parsing Excel: {type(e).__name__}: {e}", file=sys.stderr)
        return 1
    print(f"Parsed {len(jobs)} job rows from Excel.")

    if args.job:
        jobs = [j for j in jobs if j.job == args.job]
        if not jobs:
            print(f"No Excel row with Job#='{args.job}'.")
            return 1

    print("Loading existing Notion Jobs 26KS rows...")
    try:
        existing = load_existing_jobs()
    except sweep.NotionError as e:
        print(f"ERROR querying Notion: {e}", file=sys.stderr)
        return 1
    print(f"  Found {len(existing)} existing rows in Notion.\n")

    # Plan summary up front
    new_jobs   = [j for j in jobs if j.job.lower().strip() not in existing]
    blank_fill = [j for j in jobs
                  if j.job.lower().strip() in existing
                  and _number(existing[j.job.lower().strip()]["properties"],
                              NP_TOTAL_SIGNED) is None
                  and j.total_signed is not None]
    print("Plan")
    print("─" * 60)
    print(f"  Excel rows scanned:          {len(jobs)}")
    print(f"  Existing in Notion (match):  {len(jobs) - len(new_jobs)}")
    print(f"  Missing from Notion:         {len(new_jobs)}"
          + ("  (will CREATE)" if not args.no_create else "  (--no-create: SKIP)"))
    print(f"  Existing w/ blank Total Signed → fill candidates: {len(blank_fill)}")
    print()

    stats = SyncStats()
    for p in jobs:
        try:
            sync_one_job(p, existing,
                         allow_create=not args.no_create,
                         live=args.live, stats=stats)
        except Exception as e:
            stats.errors.append(f"Job#{p.job}: {type(e).__name__}: {e}")
            log.exception("Unhandled error syncing Job#%s", p.job)

    print()
    print("Summary")
    print("─" * 60)
    print(f"  Scanned:              {stats.jobs_scanned}")
    print(f"  Created:              {stats.jobs_created}")
    print(f"  Updated (blank-fill): {stats.jobs_updated}")
    print(f"  Skipped:              {stats.jobs_skipped}")
    print(f"  Errors:               {len(stats.errors)}")
    if stats.errors:
        for e in stats.errors[:5]:
            print(f"    ! {e}")

    if args.live and not stats.errors:
        state["last_mtime"] = mtime
        state["last_run"]   = datetime.now().isoformat()
        _save_state(state)

    return 0 if not stats.errors else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
