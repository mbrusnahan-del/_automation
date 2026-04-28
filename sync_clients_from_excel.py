"""
============================================================================
  STANDALONE TOOL — not part of the KS Automation Sweep.
  Run this manually (or via its own Task Scheduler entry) whenever the
  master Excel client list changes. It has no effect on sweep.py and is
  not invoked by run_sweep.bat.
============================================================================

Client List → Notion sync.

Reads the 'Client List' sheet from Master-Business Plan.xlsm and merges the
data into Notion's Client Database + Contacts database.

Runs independently of sweep.py. Shares only the low-level Notion HTTP
wrapper and env loader (because rewriting those would duplicate 100 lines
of well-tested plumbing for no gain). Everything else — schedule, log file,
state file, run command — is separate.

Run commands
------------
    python sync_clients_from_excel.py            # dry-run
    python sync_clients_from_excel.py --live     # apply changes
    python sync_clients_from_excel.py --client 42 --live   # one client
    run_excel_sync.bat                           # windowless wrapper
                                                 # (used by Task Scheduler)

Task Scheduler setup (optional, run daily)
------------------------------------------
    Program:   C:\\Users\\MichaelBrusnahan\\OneDrive - Kingdom Structural LLC\\_Projects\\_automation\\run_excel_sync.bat
    Start in:  C:\\Users\\MichaelBrusnahan\\OneDrive - Kingdom Structural LLC\\_Projects\\_automation
    Trigger:   Daily at 2:00 AM (or whenever suits you)
    Settings:  Run whether user is logged on or not; Configure for Windows 10

This is separate from the KS Automation Sweep task. Disable/enable it
independently. The sync only does work when the Excel file has changed
(mtime cache), so running it daily is effectively free when nothing moved.

Policy (v2.1.7 — 2026-04-27)
---------------------------
- FILL-BLANKS-ONLY across the board. The sync NEVER overwrites a populated
  Notion value. Once any Client-level field is set in Notion, Excel can't
  touch it — the team owns Notion edits and Excel is treated as a backstop.
- NEW CLIENTS get created with Name + Number + Address + The Guy populated
  from Excel.
- EXISTING CLIENTS get blank Address and blank The Guy filled from Excel
  where Excel has data. Name and Number are never touched (Number is the
  match key; Name is the title and is assumed correct).
- CONTACTS: new Contact rows are created for any Excel contact that doesn't
  exist on the Client yet (matched case-insensitively by Contact Name).
  Existing Contacts get blank Email / Phone fields filled from Excel
  (never overwrites populated values).
- MATCH BY NUMBER: Client match is strictly by Client.Number (Excel Col B)
  against Notion's Client.Number text field, leading zeros stripped.
- IDEMPOTENT: same Excel → same result. Safe to re-run.
- MTIME-CACHED: skips work when the Excel hasn't changed since last run
  (unless --force).
- DRY-RUN BY DEFAULT: prints the planned writes without touching Notion.
  Pass --live to actually apply changes.

Excel layout ('Client List' sheet)
----------------------------------
Row 3 = header, rows 4+ = data. Relevant columns:
  B: Number         (text, no padding — "1", "20", "147")
  C: Sub Client     (not synced)
  D: The Guy        (select: JB / CV / GO)
  E: Name           (Client company / person name — required)
  F: Address        (multi-line OK)
  G–R: Contact 1..12  (free-text: "Name", "Name email", or "Name email phone"
                       in various delimiter styles)

Contact cell parsing
--------------------
Regex-based:
  - Email: standard RFC-ish pattern
  - Phone: 10 digits with optional punctuation (`.`, `-`, `()`, space)
  - Name: everything else (whitespace-collapsed, trailing separators stripped)

Usage
-----
    python sync_clients_from_excel.py                          # dry-run
    python sync_clients_from_excel.py --live                   # apply
    python sync_clients_from_excel.py --excel /path/to.xlsm    # custom path
    python sync_clients_from_excel.py --client 42              # single client
    python sync_clients_from_excel.py --no-contacts            # skip contacts
    python sync_clients_from_excel.py --force                  # ignore mtime cache
    python sync_clients_from_excel.py --verbose                # debug output
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Iterable

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)

import config   # noqa: E402
import sweep    # noqa: E402
sweep._load_dotenv()

from openpyxl import load_workbook  # noqa: E402

log = logging.getLogger("client_sync")


# ────────────────────────────────────────────────────────────────────────
# Configuration
# ────────────────────────────────────────────────────────────────────────

CLIENT_LIST_SHEET = "Client List"
HEADER_ROW = 3          # 'Number' / 'Name' / ... live here
DATA_START_ROW = 4

# Column letters on the Client List sheet
COL_NUMBER   = "B"
COL_THE_GUY  = "D"
COL_NAME     = "E"
COL_ADDRESS  = "F"
# v2.1.7 layout: each contact slot has 3 dedicated columns (Name | Email | Phone).
# 12 slots × 3 columns = 36 contact columns starting at G.
# Slot 1 → G/H/I, Slot 2 → J/K/L, ..., Slot 12 → AN/AO/AP.
CONTACT_SLOTS = [
    ("G",  "H",  "I"),   # Slot 1
    ("J",  "K",  "L"),   # Slot 2
    ("M",  "N",  "O"),   # Slot 3
    ("P",  "Q",  "R"),   # Slot 4
    ("S",  "T",  "U"),   # Slot 5
    ("V",  "W",  "X"),   # Slot 6
    ("Y",  "Z",  "AA"),  # Slot 7
    ("AB", "AC", "AD"),  # Slot 8
    ("AE", "AF", "AG"),  # Slot 9
    ("AH", "AI", "AJ"),  # Slot 10
    ("AK", "AL", "AM"),  # Slot 11
    ("AN", "AO", "AP"),  # Slot 12
]

# Where to look for the xlsm. Override with --excel. These paths reflect the
# typical Cowork mount + the local Windows OneDrive layout.
EXCEL_PATH_CANDIDATES = [
    os.path.join(os.path.dirname(os.path.dirname(_SCRIPT_DIR)),
                 "Business Plan", "Master-Business Plan.xlsm"),
    "/sessions/nice-compassionate-pasteur/mnt/OneDrive - Kingdom Structural LLC/Business Plan/Master-Business Plan.xlsm",
    "/sessions/nice-compassionate-pasteur/mnt/uploads/Master-Business Plan.xlsm",
]

# State file — mtime cache so we don't re-parse an unchanged workbook
STATE_PATH = os.path.join(_SCRIPT_DIR, "client_sync_state.json")

# Valid values for The Guy select on Client DB (keep in sync with Notion)
VALID_THE_GUYS = {"JB", "CV", "GO"}

# Regex
_EMAIL_RE = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b')
_PHONE_RE = re.compile(r'(?<!\d)(\(?\d{3}\)?[.\-\s]?\d{3}[.\-\s]?\d{4})(?!\d)')


# ────────────────────────────────────────────────────────────────────────
# Data classes
# ────────────────────────────────────────────────────────────────────────


@dataclass
class ParsedContact:
    name: str
    email: str = ""
    phone: str = ""

    def is_empty(self) -> bool:
        return not (self.name or self.email or self.phone)


@dataclass
class ParsedClient:
    row_idx: int
    number: str             # "1", "20", "147" — text, no zero padding
    name: str
    the_guy: str = ""       # "JB" / "CV" / "GO" / ""
    address: str = ""
    contacts: list[ParsedContact] = field(default_factory=list)


@dataclass
class SyncStats:
    clients_scanned: int = 0
    clients_created: int = 0
    clients_updated: int = 0     # fields filled on existing client
    clients_skipped: int = 0     # all fields already populated
    contacts_created: int = 0
    contacts_updated: int = 0
    contacts_skipped: int = 0
    errors: list[str] = field(default_factory=list)


# ────────────────────────────────────────────────────────────────────────
# Excel parsing
# ────────────────────────────────────────────────────────────────────────


def _col_index(letter: str) -> int:
    """Excel column letter → 0-based index."""
    letter = letter.upper()
    if len(letter) == 1:
        return ord(letter) - 65
    return (ord(letter[0]) - 64) * 26 + (ord(letter[1]) - 65)


def parse_contact_cell(raw: str) -> ParsedContact | None:
    """Extract (name, email, phone) from a free-text contact cell.

    Handles all three delimiter styles found in the master list:
      - 'Patrick Wolf patwolf17@gmail.com 602.989.7797'
      - 'Richard Hernandez, richard@erginc.net, 480.220.7959'
      - 'David Garcia; david@gardelengineering.com; 480.217.5853'

    Falls back to name-only when only a name is present.
    """
    if raw is None:
        return None
    text = str(raw).replace("\n", " ").strip()
    if not text:
        return None

    email_m = _EMAIL_RE.search(text)
    phone_m = _PHONE_RE.search(text)
    email = email_m.group(0).lower() if email_m else ""
    phone = _format_phone(phone_m.group(0)) if phone_m else ""

    # Remove email/phone from text to recover the name
    name = text
    if email_m:
        name = name.replace(email_m.group(0), " ")
    if phone_m:
        name = name.replace(phone_m.group(0), " ")
    # Strip delimiters + whitespace
    name = re.sub(r'[\s,;:|]+', " ", name).strip().strip(",;:|").strip()

    return ParsedContact(name=name, email=email, phone=phone)


def _format_phone(raw: str) -> str:
    """Normalize a phone to XXX.XXX.XXXX."""
    digits = re.sub(r'\D', '', raw)
    if len(digits) == 10:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:]}"
    if len(digits) == 11 and digits.startswith("1"):
        return f"{digits[1:4]}.{digits[4:7]}.{digits[7:]}"
    return raw.strip()


def read_client_list(excel_path: str) -> list[ParsedClient]:
    """Parse the 'Client List' sheet into ParsedClient dicts."""
    wb = load_workbook(excel_path, data_only=True, read_only=True, keep_vba=False)
    if CLIENT_LIST_SHEET not in wb.sheetnames:
        raise RuntimeError(
            f"'{CLIENT_LIST_SHEET}' sheet not found in {excel_path}. "
            f"Available: {wb.sheetnames}"
        )
    ws = wb[CLIENT_LIST_SHEET]

    num_i   = _col_index(COL_NUMBER)
    guy_i   = _col_index(COL_THE_GUY)
    name_i  = _col_index(COL_NAME)
    addr_i  = _col_index(COL_ADDRESS)
    # Each slot is (name_idx, email_idx, phone_idx) — 0-based column indexes.
    slot_indexes = [
        (_col_index(n), _col_index(e), _col_index(p))
        for (n, e, p) in CONTACT_SLOTS
    ]

    def _cell(row, i):
        return row[i] if i < len(row) else None

    clients: list[ParsedClient] = []
    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx < DATA_START_ROW:
            continue

        number_raw = row[num_i] if num_i < len(row) else None
        name_raw   = row[name_i] if name_i < len(row) else None
        if not name_raw and not number_raw:
            continue  # blank row

        number = _stringify_number(number_raw)
        name = str(name_raw).strip() if name_raw else ""
        the_guy = str(row[guy_i] or "").strip().upper() if guy_i < len(row) else ""
        if the_guy and the_guy not in VALID_THE_GUYS:
            # Silently drop unknown values (e.g. "TBD", "—")
            the_guy = ""
        address = str(row[addr_i] or "").strip() if addr_i < len(row) else ""

        contacts: list[ParsedContact] = []
        for (n_i, e_i, p_i) in slot_indexes:
            n_val = _cell(row, n_i)
            e_val = _cell(row, e_i)
            p_val = _cell(row, p_i)
            # Empty slot — skip entirely.
            if not (n_val or e_val or p_val):
                continue
            # Parse the name cell with the regex (still tolerates legacy
            # rows where someone types "Name email phone" all into the
            # name column). The explicit email / phone columns then WIN
            # over whatever the regex extracted.
            parsed = parse_contact_cell(n_val) or ParsedContact()
            if e_val:
                parsed.email = str(e_val).strip().lower()
            if p_val:
                parsed.phone = _format_phone(str(p_val))
            if parsed.is_empty():
                continue
            contacts.append(parsed)

        clients.append(ParsedClient(
            row_idx=row_idx, number=number, name=name,
            the_guy=the_guy, address=address, contacts=contacts,
        ))
    return clients


def _stringify_number(v) -> str:
    """Excel number cells come in as int/float; the Client.Number field is
    text without padding. 1.0 → '1', '20' → '20'."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, int):
        return str(v)
    return str(v).strip()


# ────────────────────────────────────────────────────────────────────────
# Notion side: load existing state, match against Excel
# ────────────────────────────────────────────────────────────────────────


def load_existing_clients() -> dict[str, dict]:
    """Return {client_number: client_page} for every Client in the Notion DB
    that has a non-empty Number."""
    index: dict[str, dict] = {}
    cursor: str | None = None
    while True:
        body: dict = {"page_size": 100}
        if cursor:
            body["start_cursor"] = cursor
        resp = sweep._notion_request(
            "POST", f"/data_sources/{config.CLIENT_DS_ID}/query", body
        )
        for page in resp.get("results", []):
            num = sweep.text_val(page["properties"], "Number")
            if num:
                index[num.strip()] = page
        if not resp.get("has_more"):
            break
        cursor = resp.get("next_cursor")
    return index


def load_client_contacts(client_page_id: str) -> list[dict]:
    """All Contacts whose Company relation includes this Client."""
    try:
        resp = sweep._notion_request(
            "POST", f"/data_sources/{config.CONTACT_DS_ID}/query",
            {"page_size": 100,
             "filter": {"property": "Company", "relation": {"contains": client_page_id}}},
        )
        return resp.get("results", [])
    except sweep.NotionError as e:
        log.warning("Could not list contacts for client %s: %s", client_page_id[:8], e)
        return []


# ────────────────────────────────────────────────────────────────────────
# v2.1.6 one-time backfill: fill empty Address on existing Notion Clients
# from the Excel Client List. Strictly fill-blanks-only — never overwrites
# a populated Address. Bypasses the broader read-only-on-existing-clients
# policy because the team's address data is already in Excel and just
# hasn't been mirrored into Notion.
# ────────────────────────────────────────────────────────────────────────


def address_backfill(parsed_clients: list[ParsedClient],
                     existing_by_number: dict[str, dict],
                     *, live: bool) -> dict:
    """Walk Notion Clients with an empty Address; fill from Excel where
    we have a Number-matched address.

    Returns a stats dict suitable for printing.
    """
    excel_by_number = {pc.number: pc for pc in parsed_clients if pc.number}

    notion_total      = len(existing_by_number)
    notion_empty      = 0       # Notion client has no Address
    would_fill        = 0       # Notion empty + Excel has address
    no_excel_match    = 0       # Notion empty, no row in Excel with this Number
    no_excel_address  = 0       # Notion empty, Excel row exists, but its Address is also empty
    notion_already    = 0       # Notion already has Address — left alone
    filled            = 0       # Actually written (live mode)
    errors: list[str] = []
    plan: list[dict]  = []      # per-row preview

    for number, page in sorted(existing_by_number.items(), key=_natural_num):
        notion_addr = sweep.text_val(page["properties"], "Address") or ""
        notion_name = sweep.title_val(page["properties"], "Name") or "(unnamed)"
        if notion_addr.strip():
            notion_already += 1
            continue
        notion_empty += 1
        excel = excel_by_number.get(number)
        if not excel:
            no_excel_match += 1
            plan.append({
                "number": number, "name": notion_name,
                "action": "SKIP — no Excel row with this Number",
                "address": "",
            })
            continue
        if not (excel.address or "").strip():
            no_excel_address += 1
            plan.append({
                "number": number, "name": notion_name,
                "action": "SKIP — Excel row has no Address either",
                "address": "",
            })
            continue
        would_fill += 1
        plan.append({
            "number": number, "name": notion_name,
            "action": "FILL" if live else "would FILL",
            "address": excel.address.strip(),
            "excel_name": excel.name,
            "page_id": page["id"],
        })
        if live:
            try:
                sweep.update_page(page["id"], {
                    "Address": {"rich_text": [{"type": "text",
                                               "text": {"content": excel.address.strip()}}]},
                })
                filled += 1
            except Exception as e:
                errors.append(f"#{number} {notion_name}: {type(e).__name__}: {e}")

    return {
        "notion_total":     notion_total,
        "notion_already":   notion_already,
        "notion_empty":     notion_empty,
        "would_fill":       would_fill,
        "no_excel_match":   no_excel_match,
        "no_excel_address": no_excel_address,
        "filled":           filled,
        "errors":           errors,
        "plan":             plan,
    }


def _natural_num(item):
    """Sort-key for clients keyed by Number — numeric where possible."""
    n = item[0]
    try:
        return (0, int(n))
    except (TypeError, ValueError):
        return (1, n)


# ────────────────────────────────────────────────────────────────────────
# Sync logic — fill-blanks-only
# ────────────────────────────────────────────────────────────────────────


def sync_one_client(parsed: ParsedClient,
                    existing_by_number: dict[str, dict],
                    *, do_contacts: bool, live: bool,
                    stats: SyncStats) -> None:
    stats.clients_scanned += 1

    existing = existing_by_number.get(parsed.number)

    if existing is None:
        # Brand-new client. Create with everything we have.
        payload: dict = {
            "Name":   {"title": [{"type": "text", "text": {"content": parsed.name}}]},
            "Number": {"rich_text": [{"type": "text", "text": {"content": parsed.number}}]},
        }
        if parsed.address:
            payload["Address"] = {"rich_text": [{"type": "text", "text": {"content": parsed.address}}]}
        if parsed.the_guy:
            payload["The Guy"] = {"select": {"name": parsed.the_guy}}

        if live:
            try:
                created = sweep.create_page(
                    {"type": "data_source_id", "data_source_id": config.CLIENT_DS_ID},
                    payload,
                )
                client_page_id = created["id"]
                stats.clients_created += 1
                log.info("+ Client CREATED: #%s %s", parsed.number, parsed.name)
            except sweep.NotionError as e:
                stats.errors.append(f"create client #{parsed.number}: {e}")
                return
        else:
            client_page_id = None
            stats.clients_created += 1
            log.info("[dry] would CREATE client #%s %s "
                     "(address=%s, the_guy=%s)",
                     parsed.number, parsed.name,
                     bool(parsed.address), parsed.the_guy or "-")
    else:
        # Existing Client — fill BLANK Client-level fields from Excel where
        # we have data. NEVER overwrite a populated Notion value (the team
        # may have refined it in Notion; Excel is not authoritative for
        # filled values).
        # v2.1.7 (2026-04-27) — broadened from contacts-only to also
        # cover blank Address and The Guy on existing clients. Keeps the
        # safety guarantee that a populated Notion field is never touched.
        client_page_id = existing["id"]
        cp = existing["properties"]
        updates: dict = {}
        fields_filled: list[str] = []

        if parsed.address:
            notion_addr = sweep.text_val(cp, "Address") or ""
            if not notion_addr.strip():
                updates["Address"] = {"rich_text": [{"type": "text",
                                                      "text": {"content": parsed.address}}]}
                fields_filled.append("Address")

        if parsed.the_guy:
            notion_guy = (cp.get("The Guy", {}) or {}).get("select") or {}
            if not (notion_guy.get("name") or "").strip():
                updates["The Guy"] = {"select": {"name": parsed.the_guy}}
                fields_filled.append("The Guy")

        if updates:
            if live:
                try:
                    sweep.update_page(client_page_id, updates)
                    stats.clients_updated += 1
                    log.info("~ Client UPDATED: #%s %s  fields=%s",
                             parsed.number, parsed.name, fields_filled)
                except sweep.NotionError as e:
                    stats.errors.append(f"update client #{parsed.number}: {e}")
            else:
                stats.clients_updated += 1
                log.info("[dry] would UPDATE client #%s %s  fields=%s",
                         parsed.number, parsed.name, fields_filled)
        else:
            stats.clients_skipped += 1

    if do_contacts:
        # client_page_id may be None in dry-run for newly-created clients;
        # _sync_contacts_for_client handles that case (logs what would be
        # created without attempting a Notion write).
        _sync_contacts_for_client(parsed, client_page_id, live=live, stats=stats)


def _sync_contacts_for_client(parsed: ParsedClient, client_page_id: str,
                              *, live: bool, stats: SyncStats) -> None:
    """For each Excel contact, create-or-fill in Notion Contacts DB."""
    # In dry-run mode for a brand-new client, client_page_id is None; skip.
    if not client_page_id:
        for pc in parsed.contacts:
            stats.contacts_created += 1
            log.info("[dry]   would CREATE contact %r (email=%s, phone=%s) "
                     "on NEW client #%s",
                     pc.name, bool(pc.email), bool(pc.phone), parsed.number)
        return

    # Always fetch existing contacts — dry-run needs this to accurately
    # distinguish "would CREATE new contact" from "would FILL blanks on
    # existing contact" from "nothing to do." Costs one query per Client.
    existing_contacts = load_client_contacts(client_page_id)
    # Build a case-insensitive name index for the client's existing contacts
    existing_by_name: dict[str, dict] = {}
    for c in existing_contacts:
        cname = sweep.title_val(c["properties"], "Contact Name").lower().strip()
        if cname:
            existing_by_name[cname] = c

    for pc in parsed.contacts:
        key = pc.name.lower().strip()
        existing = existing_by_name.get(key)

        if existing is None:
            # New contact
            payload: dict = {
                "Contact Name": {"title": [{"type": "text", "text": {"content": pc.name}}]},
                "Company":      {"relation": [{"id": client_page_id}]},
            }
            if pc.email:
                payload["Email"] = {"email": pc.email}
            if pc.phone:
                payload["Phone"] = {"phone_number": pc.phone}
            if live:
                try:
                    sweep.create_page(
                        {"type": "data_source_id", "data_source_id": config.CONTACT_DS_ID},
                        payload,
                    )
                    stats.contacts_created += 1
                    log.info("  + Contact CREATED: %r  (email=%s, phone=%s)",
                             pc.name, bool(pc.email), bool(pc.phone))
                except sweep.NotionError as e:
                    stats.errors.append(f"create contact '{pc.name}' on client #{parsed.number}: {e}")
            else:
                stats.contacts_created += 1
                log.info("[dry]   would CREATE contact %r  (email=%s, phone=%s)",
                         pc.name, bool(pc.email), bool(pc.phone))
        else:
            # Existing contact — fill blanks only
            cp = existing["properties"]
            updates: dict = {}
            existing_email = (cp.get("Email", {}) or {}).get("email") or ""
            existing_phone = (cp.get("Phone", {}) or {}).get("phone_number") or ""
            if pc.email and not existing_email:
                updates["Email"] = {"email": pc.email}
            if pc.phone and not existing_phone:
                updates["Phone"] = {"phone_number": pc.phone}
            if updates:
                if live:
                    try:
                        sweep.update_page(existing["id"], updates)
                        stats.contacts_updated += 1
                        log.info("  ~ Contact UPDATED: %r  fields=%s",
                                 pc.name, list(updates.keys()))
                    except sweep.NotionError as e:
                        stats.errors.append(f"update contact '{pc.name}': {e}")
                else:
                    stats.contacts_updated += 1
                    log.info("[dry]   would UPDATE contact %r  fields=%s",
                             pc.name, list(updates.keys()))
            else:
                stats.contacts_skipped += 1


# ────────────────────────────────────────────────────────────────────────
# mtime cache
# ────────────────────────────────────────────────────────────────────────


def _load_state() -> dict:
    if os.path.isfile(STATE_PATH):
        try:
            with open(STATE_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_state(state: dict) -> None:
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)


def _find_excel(override: str | None) -> str | None:
    if override:
        return override if os.path.isfile(override) else None
    for p in EXCEL_PATH_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


# ────────────────────────────────────────────────────────────────────────
# Main
# ────────────────────────────────────────────────────────────────────────


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Sync Client List Excel → Notion Client DB + Contacts")
    ap.add_argument("--live", action="store_true",
                    help="Actually write to Notion. Default is dry-run.")
    ap.add_argument("--excel", help="Path to Master-Business Plan.xlsm (overrides defaults)")
    ap.add_argument("--client", type=str,
                    help="Only sync the Client row with this Number (e.g. '42'). "
                         "Useful for testing.")
    ap.add_argument("--no-contacts", action="store_true",
                    help="Sync Client rows only; skip per-contact syncing.")
    ap.add_argument("--fill-addresses", action="store_true",
                    help="One-time backfill: walk Notion Clients with an empty "
                         "Address and fill from Excel by Number match. Strictly "
                         "fill-blanks-only — never overwrites a populated "
                         "Address. Skips contacts and the create-new-client path.")
    ap.add_argument("--force", action="store_true",
                    help="Ignore the mtime cache and process the file even if unchanged.")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    # Log to BOTH stdout AND excel_sync.log (separate from sweep.log) so
    # interactive runs show output live while scheduled runs leave a
    # persistent trace you can review later.
    log_path = os.path.join(_SCRIPT_DIR, "excel_sync.log")
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
    )
    log.info("----- sync run starting (live=%s, client=%s) -----",
             args.live, args.client or "all")

    excel_path = _find_excel(args.excel)
    if not excel_path:
        print("ERROR: Master-Business Plan.xlsm not found. Tried:",
              file=sys.stderr)
        for p in EXCEL_PATH_CANDIDATES:
            print(f"  {p}", file=sys.stderr)
        print("Pass --excel to specify an explicit path.", file=sys.stderr)
        return 1

    mtime = os.path.getmtime(excel_path)
    state = _load_state()
    # mtime cache is for the recurring contact-sync flow only. For
    # --fill-addresses (a one-time Notion backfill) and --client (a single
    # row override), always run regardless of whether Excel has changed.
    if (state.get("last_mtime") == mtime and not args.force
            and not args.client and not args.fill_addresses):
        print(f"No change since last sync (mtime={datetime.fromtimestamp(mtime)}).",
              "Use --force to re-run.")
        return 0

    if not os.environ.get("NOTION_TOKEN") and args.live:
        print("ERROR: NOTION_TOKEN not set. Check your .env file.", file=sys.stderr)
        return 1

    print(f"Excel: {excel_path}")
    print(f"Mode:  {'LIVE' if args.live else 'DRY-RUN'}")
    if args.client:
        print(f"Filter: only Client Number '{args.client}'")
    print()

    try:
        clients = read_client_list(excel_path)
    except Exception as e:
        print(f"ERROR parsing Excel: {type(e).__name__}: {e}", file=sys.stderr)
        return 1
    print(f"Parsed {len(clients)} client rows from Excel.")

    if args.client:
        clients = [c for c in clients if c.number == args.client]
        if not clients:
            print(f"No client with Number='{args.client}'.")
            return 1

    # ── --fill-addresses branch (v2.1.6) ──────────────────────────────────
    # One-time backfill of Notion Address from Excel, fill-blanks-only.
    # Skips the create-new-client path and the per-contact sync.
    if args.fill_addresses:
        print("Loading existing Notion Clients...")
        existing = load_existing_clients()
        print(f"  Found {len(existing)} existing Clients with Numbers in Notion.\n")

        result = address_backfill(clients, existing, live=args.live)

        print("Address backfill report")
        print("─" * 70)
        print(f"  Notion Clients indexed by Number:    {result['notion_total']}")
        print(f"  Already have Address (left alone):   {result['notion_already']}")
        print(f"  Empty Address in Notion:             {result['notion_empty']}")
        print(f"    └ would-fill from Excel:           {result['would_fill']}")
        print(f"    └ no Excel row with this Number:   {result['no_excel_match']}")
        print(f"    └ Excel row also has empty Addr:   {result['no_excel_address']}")
        if args.live:
            print(f"  Filled (live writes):                {result['filled']}")
        print()

        if result["plan"]:
            verb = "FILL" if args.live else "would FILL"
            print(f"Per-row plan ({verb} actions first, then SKIPs):")
            print("─" * 70)
            fill_rows = [r for r in result["plan"] if r["action"].startswith(("FILL", "would FILL"))]
            skip_rows = [r for r in result["plan"] if not r["action"].startswith(("FILL", "would FILL"))]
            for r in fill_rows:
                tag = "+" if args.live else "·"
                print(f"  {tag} #{r['number']:>5}  {r['name']}")
                print(f"           ← {r['address']}")
            if skip_rows:
                print()
                print(f"  Skipped ({len(skip_rows)}):")
                for r in skip_rows[:30]:
                    print(f"    · #{r['number']:>5}  {r['name']:<35.35s}  {r['action']}")
                if len(skip_rows) > 30:
                    print(f"    ...and {len(skip_rows) - 30} more")
        print("─" * 70)

        if result["errors"]:
            print(f"\nErrors ({len(result['errors'])}):")
            for e in result["errors"][:15]:
                print(f"  ! {e}")

        return 0 if not result["errors"] else 1

    # Load existing Notion clients (always — dry-run needs this to accurately
    # categorize each Excel row as MATCH vs CREATE). API cost is ~2 paginated
    # queries, trivial compared to per-contact writes.
    print("Loading existing Notion Clients...")
    existing = load_existing_clients()
    print(f"  Found {len(existing)} existing Clients with Numbers in Notion.\n")

    # ── Pre-flight match report ────────────────────────────────────────────
    # Show which Excel rows will MATCH an existing Notion Client by Number
    # vs which would CREATE new ones. User can eyeball this before going live.
    matches: list[ParsedClient] = []
    creates: list[ParsedClient] = []
    for pc in clients:
        if pc.number and pc.number in existing:
            matches.append(pc)
        else:
            creates.append(pc)

    print("Pre-flight match report")
    print("─" * 60)
    print(f"  MATCH  (existing Notion Client, contacts-only sync): {len(matches)}")
    print(f"  CREATE (no Notion Client with this Number):          {len(creates)}")
    if creates:
        print("\n  Rows that would CREATE new Clients:")
        for pc in creates[:30]:
            num = pc.number or "(no number)"
            print(f"    • #{num:>5}  {pc.name}")
        if len(creates) > 30:
            print(f"    ...and {len(creates) - 30} more")
    print("─" * 60 + "\n")

    stats = SyncStats()
    for pc in clients:
        try:
            sync_one_client(pc, existing,
                            do_contacts=not args.no_contacts,
                            live=args.live, stats=stats)
        except Exception as e:
            stats.errors.append(f"client #{pc.number} {pc.name}: {type(e).__name__}: {e}")
            log.exception("unhandled error on %s", pc.name)

    print("\n" + "=" * 60)
    print(f"Clients scanned:   {stats.clients_scanned}")
    print(f"Clients created:   {stats.clients_created}")
    print(f"Clients updated:   {stats.clients_updated}")
    print(f"Clients skipped:   {stats.clients_skipped}  (already had every field)")
    print(f"Contacts created:  {stats.contacts_created}")
    print(f"Contacts updated:  {stats.contacts_updated}")
    print(f"Contacts skipped:  {stats.contacts_skipped}")
    print(f"Errors:            {len(stats.errors)}")
    if stats.errors:
        print("\nFirst 10 errors:")
        for e in stats.errors[:10]:
            print(f"  ! {e}")
    print("=" * 60)

    if args.live and not stats.errors:
        _save_state({
            "last_mtime": mtime,
            "last_run": datetime.now(timezone.utc).isoformat(),
            "clients_seen": stats.clients_scanned,
        })
        print("\n(mtime cache updated)")

    return 0 if not stats.errors else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
