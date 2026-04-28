"""
Test whether the integration token can directly read Carlos Dávalos's
contact page. If this fails with 'object_not_found' or similar permission
error, the KS Automation integration is not connected to the Contacts
database (or to that specific page).
"""
import os, sys, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sweep
sweep._load_dotenv()

# Carlos Dávalos contact page id
CARLOS_ID = "34cb73dc-460e-8086-8aaf-dbeb9aa5fef8"

print(f"Attempting to fetch Carlos Dávalos contact page directly...")
print(f"Page ID: {CARLOS_ID}\n")

try:
    page = sweep.get_page(CARLOS_ID)
    print("✓ SUCCESS — integration CAN read this contact page")
    print(f"  Title returned: {sweep.title_val(page['properties'], 'Contact Name')!r}")
    print(f"  Email property: {page['properties'].get('Email')}")
    print(f"  Phone property: {page['properties'].get('Phone')}")
    print()
    print("If get_page works but the Project Contact relation is empty, then")
    print("the issue is page-level permissions or relation cache on Notion's side.")
except sweep.NotionError as e:
    print(f"✗ FAILED — {e}")
    print()
    print("The integration does NOT have access to this contact page.")
    print("Fix: open the Contacts database in Notion, click ••• → Connections,")
    print("and add KS Automation.")

# Also try querying the Contacts DB directly
print("\n" + "─" * 60)
print("Querying Contacts data source for Carlos Dávalos...\n")
try:
    resp = sweep._notion_request(
        "POST", f"/data_sources/{sweep.config.CONTACT_DS_ID}/query",
        {"page_size": 5,
         "filter": {"property": "Contact Name", "title": {"contains": "Carlos"}}}
    )
    results = resp.get("results", [])
    print(f"✓ Query returned {len(results)} result(s)")
    for r in results:
        name = sweep.title_val(r["properties"], "Contact Name")
        print(f"  - {name}  (id: {r['id']})")
except sweep.NotionError as e:
    print(f"✗ Query failed: {e}")
