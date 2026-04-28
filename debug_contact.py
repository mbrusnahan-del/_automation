"""
Standalone diagnostic: fetch a specific Project from Notion and walk the
_resolve_contact_dict() code path with full tracing.

Usage:
  python debug_contact.py 34cb73dc460e803f8c85f29e91050c55
  python debug_contact.py https://www.notion.so/34cb73dc460e803f8c85f29e91050c55
"""
from __future__ import annotations
import sys, os, json, re
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)

import config
import sweep
sweep._load_dotenv()


def normalize(s):
    """Extract bare UUID from URL or accept a raw UUID."""
    m = re.search(r'[0-9a-f]{32}|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', s)
    return m.group(0) if m else s


def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_contact.py <project_page_id_or_url>")
        return 1

    pid = normalize(sys.argv[1])
    print(f"Fetching Project page: {pid}")
    project_page = sweep.get_page(pid)
    title = sweep.title_val(project_page["properties"], "Project Name")
    print(f"  Title: {title}")

    pprops = project_page["properties"]
    print(f"\nKeys on Project page: {sorted(pprops.keys())}")

    # ── Compare Client (worked) vs Project Contact (failed) ─────────────
    for prop_name in ("Client", "Project Contact", "Proposal Brief"):
        raw = pprops.get(prop_name)
        print(f"\n── Property: {prop_name} ──")
        print(f"  Inline GET /pages response:")
        print(f"  {json.dumps(raw, indent=2)[:400]}")
        # Try the property-specific endpoint (2025-09-03 API)
        if raw and raw.get("id"):
            prop_id = raw["id"]
            try:
                prop_resp = sweep._notion_request("GET", f"/pages/{pid}/properties/{prop_id}")
                print(f"  Property endpoint GET /pages/.../properties/{prop_id}:")
                print(f"  {json.dumps(prop_resp, indent=2)[:800]}")
            except Exception as e:
                print(f"  Property endpoint error: {e}")

    cids = sweep.relation_ids(pprops, config.ProjectProp.PROJECT_CONTACT)
    print(f"\n  relation_ids(Project Contact) via inline → {cids}")
    if not cids:
        print("  ⚠ relation_ids returned empty via inline fetch.")
        print("  (Continuing to rollup-based fallback test below.)")

    if cids:
        contact_id = cids[0]
        print(f"\n── Fetching Contact page {contact_id} ──")
        try:
            cp_full = sweep.get_page(contact_id)
            cp = cp_full["properties"]
            print(f"  Keys on Contact page: {sorted(cp.keys())}")
            print(f"  Raw 'Contact Name' property:")
            print(f"  {json.dumps(cp.get('Contact Name'), indent=2)[:400]}")
            print(f"\n  Raw 'Email' property:")
            print(f"  {json.dumps(cp.get('Email'), indent=2)[:200]}")
            print(f"\n  Raw 'Phone' property:")
            print(f"  {json.dumps(cp.get('Phone'), indent=2)[:200]}")
        except sweep.NotionError as e:
            print(f"  ⚠ get_page FAILED: {e}")

    # ── Extraction via the same helpers the sweep uses ──────────────────
    print("\n── _resolve_contact_dict() output ──")
    result = sweep._resolve_contact_dict(project_page)
    print(f"  {json.dumps(result, indent=2)}")

    # Also fire off what build_merge_data would see
    print("\n── What the renderer would store for each placeholder ──")
    contact = result
    contact_name = contact.get("Contact Name", "") if contact else ""
    email = (contact or {}).get("Email", "")
    phone = (contact or {}).get("Phone", "") or (contact or {}).get("Phone (Mobile)", "")
    def fill(val, label):
        return val if val else f"<<FILL IN: {label}>>"
    print(f"  contact_name  = {fill(contact_name, 'Contact Name')!r}")
    print(f"  client_email  = {fill(email, 'Contact Email')!r}")
    print(f"  client_phone  = {fill(phone, 'Contact Phone')!r}")

    # ── Raw rollup property dumps ────────────────────────────────────────
    print("\n── Raw rollup values on the Project page ──")
    for rollup_name in ("Client Contact", "Contact Email", "Contact Phone",
                         "Contact Mobile", "Contact Role"):
        raw = pprops.get(rollup_name)
        print(f"\n  '{rollup_name}':")
        print(f"  {json.dumps(raw, indent=2)[:500]}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
