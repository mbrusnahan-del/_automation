"""
show_unmatched_clients.py — quick one-shot diagnostic.

Reads client_revenue_summary.csv (output of analyze_clients_by_revenue.py),
filters to only the codes that didn't match any Notion client, and prints
them sorted by signed revenue with sample project names.

Helps identify which 3-digit codes need real Client records added to Notion
vs. which are placeholders or junk.
"""
import csv
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(HERE, "client_revenue_summary.csv")

if not os.path.exists(CSV_PATH):
    sys.exit(f"ERROR: {CSV_PATH} not found. Run analyze_clients_by_revenue.py first.")

rows = list(csv.DictReader(open(CSV_PATH, encoding="utf-8")))
unmatched = [r for r in rows if "unknown" in r["client_name"]]
unmatched.sort(key=lambda r: -float(r["total_signed"]))

print(f"{'Code':<6}{'Proj':<6}{'W':<4}{'L':<4}{'Signed Rev':<14}Sample project names")
print("-" * 110)
for r in unmatched:
    signed = float(r["total_signed"])
    print(f"{r['code']:<6}{r['projects']:<6}{r['won']:<4}{r['lost']:<4}${signed:>11,.0f}  {r['sample_projects'][:65]}")

print()
print(f"Total unmatched codes:  {len(unmatched)}")
print(f"Unmatched signed rev:   ${sum(float(r['total_signed']) for r in unmatched):,.0f}")
print(f"Unmatched projects:     {sum(int(r['projects']) for r in unmatched)}")
