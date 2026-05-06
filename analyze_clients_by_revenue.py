"""
analyze_clients_by_revenue.py — one-off business analysis script.

Pulls all Clients from Notion, builds a {number → name} map, then joins
against the comps_cache.json comp database to aggregate signed/proposed
revenue and project count by actual client name.

Outputs:
  - Top-clients table to stdout
  - CSV at ./client_revenue_summary.csv (sorted by signed revenue)

Run:
    python analyze_clients_by_revenue.py
"""
from __future__ import annotations
import csv
import json
import os
import re
import sys
import urllib.request
import urllib.error
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
COMPS_PATH = os.path.join(HERE, "comps_cache.json")
OUT_CSV = os.path.join(HERE, "client_revenue_summary.csv")

CLIENTS_DS_ID = "2efb73dc-460e-80fe-860c-000bdfd01348"


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


def fetch_all_clients(token: str) -> dict[str, str]:
    """Return {client_number_3digit: client_name}. Pads numbers to 3 digits
    so '1' (Speedwell) becomes '001' to match project number suffixes."""
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
            # Multiple clients sharing a code (rare) — keep first
            if code_3 not in out:
                out[code_3] = name
        if not data.get("has_more"):
            break
        body["start_cursor"] = data.get("next_cursor")
    return out


def main() -> int:
    env = load_env()
    token = env.get("NOTION_TOKEN")
    if not token:
        sys.exit("ERROR: NOTION_TOKEN missing from .env")

    if not os.path.exists(COMPS_PATH):
        sys.exit(f"ERROR: {COMPS_PATH} not found.")

    print("Fetching all clients from Notion...")
    code_to_name = fetch_all_clients(token)
    print(f"Loaded {len(code_to_name)} clients with numbers.\n")

    print("Loading comp database...")
    comps = json.load(open(COMPS_PATH))["comps"]
    print(f"Loaded {len(comps)} historical comps.\n")

    # Aggregate by client code
    by_code = defaultdict(lambda: {
        "won": 0, "lost": 0,
        "total_signed": 0.0, "total_proposed": 0.0,
        "projects": [],
    })
    for c in comps:
        job = c.get("job", "")
        # Strip any sub-project suffix (".1", ".2", etc.) before extracting digits.
        # Job numbers like "25018032.1" should map to client code "032" not "321".
        base = job.split(".")[0]
        digits = re.sub(r"\D", "", base)
        if len(digits) < 3:
            continue
        code_3 = digits[-3:]
        bucket = by_code[code_3]
        if c.get("won"):
            bucket["won"] += 1
            bucket["total_signed"] += c.get("total_signed") or 0
        else:
            bucket["lost"] += 1
            bucket["total_proposed"] += c.get("total_proposed") or 0
        bucket["projects"].append(c.get("name", ""))

    # Build merged rows with client names
    rows = []
    for code, info in by_code.items():
        name = code_to_name.get(code, "(unknown — no client matching this code)")
        rows.append({
            "code": code,
            "client_name": name,
            "projects": info["won"] + info["lost"],
            "won": info["won"],
            "lost": info["lost"],
            "win_rate_pct": round(info["won"] / max(1, info["won"] + info["lost"]) * 100, 1),
            "total_signed": round(info["total_signed"], 2),
            "total_proposed": round(info["total_proposed"], 2),
            "sample_projects": " | ".join(info["projects"][:3]),
        })

    # Sort by signed revenue descending
    rows.sort(key=lambda r: -r["total_signed"])

    # Print top 25 to console
    print(f"{'Code':<6}{'Client':<40}{'Proj':<6}{'W':<4}{'L':<4}{'Win%':<7}{'Signed Rev':<14}")
    print("-" * 90)
    for r in rows[:25]:
        client_display = r["client_name"][:38]
        print(f"{r['code']:<6}{client_display:<40}{r['projects']:<6}{r['won']:<4}{r['lost']:<4}{r['win_rate_pct']:<7.0f}${r['total_signed']:>11,.0f}")

    # Top-line summary
    total = sum(r["total_signed"] for r in rows)
    top5 = sum(r["total_signed"] for r in rows[:5])
    top10 = sum(r["total_signed"] for r in rows[:10])
    print()
    print(f"Total signed revenue: ${total:,.0f}")
    print(f"Top 5 clients:        ${top5:,.0f} ({top5/total*100:.0f}%)")
    print(f"Top 10 clients:       ${top10:,.0f} ({top10/total*100:.0f}%)")
    print(f"Unique client codes in comps: {len(rows)}")
    print(f"  ... matched to Notion clients: {sum(1 for r in rows if not r['client_name'].startswith('(unknown'))}")
    print(f"  ... unmatched: {sum(1 for r in rows if r['client_name'].startswith('(unknown'))}")

    # Write CSV
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "code", "client_name", "projects", "won", "lost",
            "win_rate_pct", "total_signed", "total_proposed", "sample_projects",
        ])
        writer.writeheader()
        writer.writerows(rows)
    print(f"\nFull table written to: {OUT_CSV}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
