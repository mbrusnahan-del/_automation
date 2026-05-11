"""
refresh_client_insights_sheet.py — refresh the "Client Insights" sheet in
Master-Business Plan.xlsm with the latest aggregated client revenue data.

Pulls:
  - All clients from the Notion Clients DB (number → name)
  - Jobs 25KS + Jobs 26KS sheets from Master-Business Plan.xlsm
  - Filters out placeholder/test rows (same logic as import_excel_comps.py)

Writes:
  - A "Client Insights" sheet (creates if missing, replaces if exists) with:
      Code | Client Name | Projects | Won | Lost | Win % | Total Signed |
      Total Proposed | Avg Signed/Won | Sample Project Names
  - Conditional formatting: green/blue/orange/red on Win % column
  - Sorted by Total Signed descending

The script preserves all other sheets and macros (workbook saved as .xlsm).
Safe to re-run; the Client Insights sheet gets fully rebuilt each time.

Usage:
    cd "C:\\Users\\MichaelBrusnahan\\OneDrive - Kingdom Structural LLC\\_Projects\\_automation"
    python refresh_client_insights_sheet.py

Requires the workbook to be CLOSED in Excel during the run.
"""
from __future__ import annotations
import json
import os
import re
import sys
import urllib.request
import urllib.error
from collections import defaultdict
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")

# Excel candidates — same as import_excel_comps.py
EXCEL_CANDIDATES = [
    os.path.join(os.path.dirname(HERE), "Business Plan", "Master-Business Plan.xlsm"),
    "/sessions/admiring-friendly-brown/mnt/Business Plan/Master-Business Plan.xlsm",
]

CLIENTS_DS_ID = "2efb73dc-460e-80fe-860c-000bdfd01348"
SHEET_NAME = "Client Insights"

# Sheet column mappings (mirror import_excel_comps.py)
HEADERS_26KS = {
    "B": "Job", "D": "Y", "E": "Client", "F": "Name",
    "R": "Design", "L": "CD", "S": "CA",
    "V": "TOTAL_SIGNED", "X": "UNSIGNED",
}
HEADERS_25KS = {
    "B": "Job", "C": "Y", "D": "Client", "E": "Name",
    "O": "Design", "L": "CD", "P": "CA",
    "R": "TOTAL_SIGNED", "T": "UNSIGNED",
}

# Test/placeholder markers (mirror import_excel_comps.py)
TEST_MARKERS = ("_test", "automation test", "n8n sandbox",
                "sandbox", "placeholder", "do not use", "delete me")


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


def fetch_clients(token: str) -> dict[str, str]:
    """Return {3-digit-code: client_name}."""
    body = {"page_size": 100}
    out = {}
    while True:
        code, data = notion(token, "POST",
                            f"/data_sources/{CLIENTS_DS_ID}/query", body)
        if code != 200:
            code, data = notion(token, "POST",
                                f"/databases/{CLIENTS_DS_ID}/query", body)
        if code != 200:
            sys.exit(f"ERROR querying clients: {code} {data}")
        for page in data.get("results", []):
            props = page["properties"]
            name_arr = props.get("Name", {}).get("title", [])
            name = "".join(t.get("plain_text", "") for t in name_arr).strip()
            number_arr = props.get("Number", {}).get("rich_text", [])
            number_raw = "".join(t.get("plain_text", "") for t in number_arr).strip()
            if not number_raw or not name:
                continue
            digits = re.sub(r"\D", "", number_raw)
            if not digits:
                continue
            code_3 = digits.zfill(3)[-3:]
            if code_3 not in out:
                out[code_3] = name
        if not data.get("has_more"):
            break
        body["start_cursor"] = data.get("next_cursor")
    return out


def col_to_index(letter):
    letter = letter.upper()
    if len(letter) == 1:
        return ord(letter) - 64
    return (ord(letter[0]) - 64) * 26 + (ord(letter[1]) - 64)


def to_num(v):
    try:
        return float(v) if v not in (None, "") else 0
    except (TypeError, ValueError):
        return 0


def is_meaningful_row(rec):
    name = (rec.get("Name") or "").strip()
    if not name:
        return False
    if not any(ch.isalpha() for ch in name):
        return False
    low = name.lower()
    if any(m in low for m in TEST_MARKERS):
        return False
    return True


def extract_sheet(ws, col_map):
    recs = []
    for row_idx in range(6, ws.max_row + 1):
        rec = {}
        for col_letter, key in col_map.items():
            rec[key] = ws.cell(row_idx, col_to_index(col_letter)).value
        if rec.get("Job"):
            recs.append(rec)
    return recs


def find_excel():
    for p in EXCEL_CANDIDATES:
        if os.path.isfile(p):
            return p
    sys.exit("ERROR: Master-Business Plan.xlsm not found.")


def main():
    print("Loading clients from Notion...")
    env = load_env()
    code_to_name = fetch_clients(env["NOTION_TOKEN"])
    print(f"  Loaded {len(code_to_name)} clients with numbers.")

    excel_path = find_excel()
    print(f"Loading workbook: {excel_path}")
    wb = load_workbook(excel_path, keep_vba=True)

    # Read job rows
    rows = []
    if "Jobs 25KS" in wb.sheetnames:
        rows += extract_sheet(wb["Jobs 25KS"], HEADERS_25KS)
    if "Jobs 26KS" in wb.sheetnames:
        rows += extract_sheet(wb["Jobs 26KS"], HEADERS_26KS)
    print(f"  Read {len(rows)} job rows from Excel.")

    # Aggregate by client code (last 3 digits of base job number)
    by_code = defaultdict(lambda: {
        "won": 0, "lost": 0,
        "total_signed": 0.0, "total_proposed": 0.0,
        "sample_names": [],
    })
    skipped_placeholders = 0
    for r in rows:
        if not is_meaningful_row(r):
            skipped_placeholders += 1
            continue
        job = str(r.get("Job") or "").strip()
        base = job.split(".")[0]
        digits = re.sub(r"\D", "", base)
        if len(digits) < 3:
            continue
        code_3 = digits[-3:]
        bucket = by_code[code_3]
        won = (str(r.get("Y") or "")).upper().startswith("Y")
        signed = to_num(r.get("TOTAL_SIGNED"))
        if won:
            bucket["won"] += 1
            bucket["total_signed"] += signed
        else:
            bucket["lost"] += 1
            bucket["total_proposed"] += signed
        if len(bucket["sample_names"]) < 3:
            name = (r.get("Name") or "").strip()
            if name:
                bucket["sample_names"].append(name[:50])
    print(f"  Filtered out {skipped_placeholders} placeholder rows.")
    print(f"  Aggregated to {len(by_code)} client codes.")

    # Build output rows, sort by signed revenue
    out_rows = []
    for code, info in by_code.items():
        total = info["won"] + info["lost"]
        win_pct = round(info["won"] / total * 100, 1) if total else 0
        avg_signed = round(info["total_signed"] / info["won"], 0) if info["won"] else 0
        out_rows.append({
            "code": code,
            "client": code_to_name.get(code, "(unknown — add to Notion)"),
            "projects": total,
            "won": info["won"],
            "lost": info["lost"],
            "win_pct": win_pct,
            "total_signed": round(info["total_signed"], 2),
            "total_proposed": round(info["total_proposed"], 2),
            "avg_signed_per_won": avg_signed,
            "samples": " | ".join(info["sample_names"]),
        })
    out_rows.sort(key=lambda r: -r["total_signed"])

    # Replace existing Client Insights sheet
    if SHEET_NAME in wb.sheetnames:
        del wb[SHEET_NAME]
    ws = wb.create_sheet(SHEET_NAME)

    # Header row
    headers = ["Code", "Client Name", "Projects", "Won", "Lost", "Win %",
               "Total Signed", "Total Proposed", "Avg $/Won", "Sample Project Names"]
    bold_white = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="1F3A5F")
    thin_border = Border(left=Side("thin", color="CCCCCC"),
                         right=Side("thin", color="CCCCCC"),
                         top=Side("thin", color="CCCCCC"),
                         bottom=Side("thin", color="CCCCCC"))
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = bold_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # Data rows
    fills = {
        "high":   PatternFill("solid", fgColor="D4F4E2"),  # green
        "mid":    PatternFill("solid", fgColor="DCE9F7"),  # blue
        "low":    PatternFill("solid", fgColor="FBE9D2"),  # orange
        "poor":   PatternFill("solid", fgColor="F8D5D5"),  # red
    }
    for i, row in enumerate(out_rows, start=2):
        ws.cell(row=i, column=1, value=row["code"])
        ws.cell(row=i, column=2, value=row["client"])
        ws.cell(row=i, column=3, value=row["projects"])
        ws.cell(row=i, column=4, value=row["won"])
        ws.cell(row=i, column=5, value=row["lost"])
        win_cell = ws.cell(row=i, column=6, value=row["win_pct"] / 100)
        win_cell.number_format = "0%"
        if row["win_pct"] >= 75:
            win_cell.fill = fills["high"]
        elif row["win_pct"] >= 50:
            win_cell.fill = fills["mid"]
        elif row["win_pct"] >= 25:
            win_cell.fill = fills["low"]
        else:
            win_cell.fill = fills["poor"]
        for col, value in [(7, row["total_signed"]), (8, row["total_proposed"]),
                           (9, row["avg_signed_per_won"])]:
            c = ws.cell(row=i, column=col, value=value)
            c.number_format = '"$"#,##0'
        ws.cell(row=i, column=10, value=row["samples"])
        for col_idx in range(1, 11):
            ws.cell(row=i, column=col_idx).border = thin_border

    # Column widths
    widths = [8, 38, 10, 8, 8, 10, 16, 16, 14, 60]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Footer summary row
    total_row = len(out_rows) + 3
    total_revenue = sum(r["total_signed"] for r in out_rows)
    top5_revenue = sum(r["total_signed"] for r in out_rows[:5])
    top10_revenue = sum(r["total_signed"] for r in out_rows[:10])

    ws.cell(row=total_row, column=1, value="Refreshed:").font = Font(bold=True)
    ws.cell(row=total_row, column=2,
            value=datetime.now().strftime("%Y-%m-%d %H:%M"))
    ws.cell(row=total_row + 1, column=1, value="Total signed:").font = Font(bold=True)
    c = ws.cell(row=total_row + 1, column=2, value=total_revenue)
    c.number_format = '"$"#,##0'
    ws.cell(row=total_row + 2, column=1, value="Top 5 share:").font = Font(bold=True)
    c = ws.cell(row=total_row + 2, column=2,
                value=top5_revenue / total_revenue if total_revenue else 0)
    c.number_format = "0%"
    ws.cell(row=total_row + 3, column=1, value="Top 10 share:").font = Font(bold=True)
    c = ws.cell(row=total_row + 3, column=2,
                value=top10_revenue / total_revenue if total_revenue else 0)
    c.number_format = "0%"
    ws.cell(row=total_row + 4, column=1, value="Unmatched codes:").font = Font(bold=True)
    ws.cell(row=total_row + 4, column=2,
            value=sum(1 for r in out_rows if r["client"].startswith("(unknown")))

    # Freeze top row
    ws.freeze_panes = "A2"

    # Save workbook
    print(f"Saving workbook with refreshed '{SHEET_NAME}' sheet...")
    wb.save(excel_path)
    print(f"Done. {len(out_rows)} client rows written.")
    print(f"\nOpen Master-Business Plan.xlsm and click the '{SHEET_NAME}' tab.")


if __name__ == "__main__":
    main()
