"""
Builder for the v2.1.8 Kingdom Structural n8n workflow JSONs (WEBHOOK mode).

Run: python build_workflows.py
Output: n8n/workflows_v2/01_..._.json through 06_..._.json

All workflows use Notion outgoing-webhook triggers (real-time, ~2s latency)
instead of polling. Workflow 05 (Job A) is now CLOUD-NATIVE — uses
Microsoft Graph API directly, no PC dependency.

Notion credential reference: "Notion account" — change CREDENTIAL_NAME
below if your n8n credential is named differently.
"""
from __future__ import annotations
import json, os, uuid

# ────────────────────────────────────────────────────────────────────────
# Configuration
# ────────────────────────────────────────────────────────────────────────
CREDENTIAL_NAME = "Notion account"
NOTION_VERSION  = "2025-09-03"

PROJECTS_DS   = "262b73dc-460e-8137-b3bb-000b62103b15"
BRIEF_DS      = "6b64658f-5fb3-4329-b02c-3ac3cb8a0828"
AUTOMATION_DS = "04d7aa20-eb8d-4d51-b302-25fb4eaaf26e"
BRIEF_TEMPLATE_PAGE = "34ab73dc-460e-80bf-b97c-da13d53310e8"

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "workflows_v2")
os.makedirs(OUT_DIR, exist_ok=True)


# ────────────────────────────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────────────────────────────
def _uid():
    return str(uuid.uuid4())

def notion_creds():
    return {"notionApi": {"name": CREDENTIAL_NAME}}

def webhook_trigger(name, path_slug):
    return {
        "parameters": {
            "httpMethod": "POST",
            "path": path_slug,
            "authentication": "none",
            "responseMode": "lastNode",
            "options": {},
        },
        "id": _uid(),
        "name": name,
        "type": "n8n-nodes-base.webhook",
        "typeVersion": 2,
        "position": [240, 300],
        "webhookId": _uid(),
    }

def get_page_node(name, position):
    return http_node(
        name=name, method="GET",
        url="=https://api.notion.com/v1/pages/{{ $json.body.data.id }}",
        body=None, position=position,
    )

def http_node(name, method, url, body, position,
              headers_extra=None, auth_notion=True, timeout_ms=None):
    headers = [{"name": "Notion-Version", "value": NOTION_VERSION}] if auth_notion else []
    if headers_extra:
        headers.extend(headers_extra)
    params = {
        "method": method, "url": url,
        "sendHeaders": True if headers else False,
        "headerParameters": {"parameters": headers} if headers else {},
    }
    if auth_notion:
        params["authentication"] = "predefinedCredentialType"
        params["nodeCredentialType"] = "notionApi"
    if body is not None:
        params["sendBody"] = True
        params["specifyBody"] = "json"
        params["jsonBody"] = body
    params["options"] = {"timeout": timeout_ms} if timeout_ms is not None else {}
    node = {
        "parameters": params, "id": _uid(), "name": name,
        "type": "n8n-nodes-base.httpRequest", "typeVersion": 4.2, "position": position,
    }
    if auth_notion:
        node["credentials"] = notion_creds()
    return node

def if_node(name, conditions, position):
    return {
        "parameters": {
            "conditions": {
                "combinator": "and", "conditions": conditions,
                "options": {"caseSensitive": True, "typeValidation": "loose"},
            },
            "options": {},
        },
        "id": _uid(), "name": name,
        "type": "n8n-nodes-base.if", "typeVersion": 2.2, "position": position,
    }

def code_node(name, code, position):
    return {
        "parameters": {"jsCode": code},
        "id": _uid(), "name": name,
        "type": "n8n-nodes-base.code", "typeVersion": 2, "position": position,
    }

def cond_bool(left, value):
    return {"leftValue": left, "rightValue": value,
            "operator": {"type": "boolean", "operation": "equals"}}

def cond_string_eq(left, value):
    return {"leftValue": left, "rightValue": value,
            "operator": {"type": "string", "operation": "equals"}}

def cond_string_neq(left, value):
    return {"leftValue": left, "rightValue": value,
            "operator": {"type": "string", "operation": "notEquals"}}

def cond_number_gt(left, value):
    return {"leftValue": left, "rightValue": value,
            "operator": {"type": "number", "operation": "gt"}}

def cond_array_length_eq(left, value):
    return {"leftValue": left, "rightValue": value,
            "operator": {"type": "array", "operation": "lengthEquals", "rightType": "number"}}

def cond_object_empty(left):
    return {"leftValue": left, "rightValue": "",
            "operator": {"type": "object", "operation": "empty", "singleValue": True}}

def build_connections(edges):
    out = {}
    for f, t, port in edges:
        out.setdefault(f, {"main": []})
        while len(out[f]["main"]) <= port:
            out[f]["main"].append([])
        out[f]["main"][port].append({"node": t, "type": "main", "index": 0})
    return out

def write_workflow(filename, name, nodes, edges, tags=None):
    workflow = {
        "name": name, "nodes": nodes,
        "connections": build_connections(edges),
        "active": False,
        "settings": {"executionOrder": "v1"},
        "versionId": _uid(),
        "tags": [{"name": t} for t in (tags or [])],
    }
    path = os.path.join(OUT_DIR, filename)
    with open(path, "w") as f:
        json.dump(workflow, f, indent=2)
    print(f"  wrote {filename}  ({os.path.getsize(path)} bytes)")


# ────────────────────────────────────────────────────────────────────────
# Workflow 01 — A0: Create Proposal Brief on new Project
# ────────────────────────────────────────────────────────────────────────
def workflow_01():
    trig = webhook_trigger("Webhook", "ks-a0-create-brief")
    get_page = get_page_node("Get Project Page", [460, 300])
    excluded_statuses = ["Dead", "IGNNROE FO RNOW", "On Hold"]
    gate = if_node(
        "Gate: no Brief AND active AND Intake Complete",
        [
            cond_array_length_eq("={{ $json.properties[\"Proposal Brief\"].relation }}", 0),
            *[cond_string_neq(
                "={{ $json.properties[\"Engineering Status\"].status?.name || \"\" }}", v,
            ) for v in excluded_statuses],
            cond_bool("={{ $json.properties[\"Intake Complete\"].checkbox }}", True),
        ],
        [680, 300],
    )
    fetch_template = http_node(
        "Fetch Brief template blocks", "GET",
        f"https://api.notion.com/v1/blocks/{BRIEF_TEMPLATE_PAGE}/children?page_size=100",
        None, [900, 240],
    )
    clean_code = """const incoming = $input.all();
const dropFields = new Set([
  "id","object","created_time","created_by","last_edited_time","last_edited_by",
  "parent","archived","in_trash","has_children","request_id","developer_survey",
]);
const supportedTypes = new Set([
  "paragraph","heading_1","heading_2","heading_3","bulleted_list_item",
  "numbered_list_item","to_do","toggle","callout","quote","divider","code",
]);
function cleanBlock(b) {
  const out = {};
  for (const k of Object.keys(b)) {
    if (dropFields.has(k)) continue;
    let v = b[k];
    if (v === null) continue;
    if (typeof v === "object" && !Array.isArray(v)) v = cleanBlock(v);
    if (Array.isArray(v)) v = v.map(x => (x && typeof x === "object" ? cleanBlock(x) : x));
    out[k] = v;
  }
  return out;
}
const allBlocks = [];
for (const item of incoming) {
  const data = item.json;
  if (data && Array.isArray(data.results)) allBlocks.push(...data.results);
  else if (Array.isArray(data)) allBlocks.push(...data);
}
const children = allBlocks.filter(b => supportedTypes.has(b.type)).map(cleanBlock);
const project = $("Get Project Page").item.json;
return [{ json: { project, children } }];"""
    clean = code_node("Clean template blocks", clean_code, [1120, 240])

    create_brief_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + BRIEF_DS + '''" },
  properties: {
    "Name": { title: [{ type: "text", text: { content: $json.project.properties["Project Name"].title[0]?.plain_text || "Untitled" } }] },
    "Project": { relation: [{ id: $json.project.id }] },
    "Status": { select: { name: "Not Started" } }
  },
  children: $json.children
}) }}'''
    create_brief = http_node("Create Brief", "POST",
                             "https://api.notion.com/v1/pages",
                             create_brief_body, [1340, 240])

    stamp_body = '''={{ JSON.stringify({
  properties: {
    "Proposal Brief": { relation: [{ id: $json.id }] }
  }
}) }}'''
    stamp = http_node("Stamp Brief on Project", "PATCH",
                      '=https://api.notion.com/v1/pages/{{ $("Get Project Page").item.json.id }}',
                      stamp_body, [1560, 240])

    log_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "A0 · Created Brief for " + ($("Get Project Page").item.json.properties["Project Name"].title[0]?.plain_text || "Untitled") } }] },
    "Job": { select: { name: "A0" } },
    "Outcome": { select: { name: "Success" } },
    "Project": { relation: [{ id: $("Get Project Page").item.json.id }] },
    "Brief": { relation: [{ id: $("Create Brief").item.json.id }] }
  }
}) }}'''
    logn = http_node("Log to Automation Log", "POST",
                     "https://api.notion.com/v1/pages", log_body, [1780, 240])

    nodes = [trig, get_page, gate, fetch_template, clean, create_brief, stamp, logn]
    edges = [
        ("Webhook", "Get Project Page", 0),
        ("Get Project Page", "Gate: no Brief AND active AND Intake Complete", 0),
        ("Gate: no Brief AND active AND Intake Complete", "Fetch Brief template blocks", 0),
        ("Fetch Brief template blocks", "Clean template blocks", 0),
        ("Clean template blocks", "Create Brief", 0),
        ("Create Brief", "Stamp Brief on Project", 0),
        ("Stamp Brief on Project", "Log to Automation Log", 0),
    ]
    write_workflow("01_A0_create_brief.json",
                   "KS A0 — Create Proposal Brief on new Project (v2.1.8 webhook)",
                   nodes, edges, tags=["ks-automation", "v2.1.8", "webhook"])


# ────────────────────────────────────────────────────────────────────────
# Workflow 02 — A02: Fee Schedule (database_id parent for seed rows)
# ────────────────────────────────────────────────────────────────────────
def workflow_02():
    trig = webhook_trigger("Webhook", "ks-a02-fee-schedule")
    get_page = get_page_node("Get Brief Page", [460, 300])

    create_db_body = '''={{ JSON.stringify({
  parent: { type: "page_id", page_id: $json.id },
  title: [{ type: "text", text: { content: "💵 Fee Schedule" } }],
  is_inline: true,
  properties: {
    "Service": { title: {} },
    "Include": { formula: { expression: "if(empty(prop(\\"Amount\\")), false, prop(\\"Amount\\") > 0)" } },
    "Type": { select: { options: [
      { name: "Fixed Fee", color: "blue" },
      { name: "Hourly", color: "yellow" },
      { name: "Not to Exceed", color: "green" }
    ] } },
    "Amount": { number: { format: "dollar" } },
    "Order": { number: { format: "number" } }
  }
}) }}'''
    create_db = http_node("Create Fee Schedule child DB", "POST",
                          "https://api.notion.com/v1/databases",
                          create_db_body, [680, 300])

    fanout_code = """const db = $input.first().json;
const databaseId = db.id;
const seeds = [
  "Schematic Design",
  "Design Development",
  "Construction Documents",
  "Construction Administration",
  "Special Structural Inspections",
  "Engineering Site Visits",
];
return seeds.map((service, i) => ({ json: { databaseId, service, order: i + 1 } }));"""
    fanout = code_node("Fan out 6 seed rows", fanout_code, [900, 300])

    seed_body = '''={{ JSON.stringify({
  parent: { type: "database_id", database_id: $json.databaseId },
  properties: {
    "Service": { title: [{ type: "text", text: { content: $json.service } }] },
    "Order": { number: $json.order }
  }
}) }}'''
    create_seed = http_node("Create seed row", "POST",
                            "https://api.notion.com/v1/pages",
                            seed_body, [1120, 300])

    log_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "A02 · Provisioned Fee Schedule (6 rows)" } }] },
    "Job": { select: { name: "A02" } },
    "Outcome": { select: { name: "Success" } },
    "Brief": { relation: [{ id: $("Get Brief Page").item.json.id }] }
  }
}) }}'''
    logn = http_node("Log to Automation Log", "POST",
                     "https://api.notion.com/v1/pages", log_body, [1340, 300])

    nodes = [trig, get_page, create_db, fanout, create_seed, logn]
    edges = [
        ("Webhook", "Get Brief Page", 0),
        ("Get Brief Page", "Create Fee Schedule child DB", 0),
        ("Create Fee Schedule child DB", "Fan out 6 seed rows", 0),
        ("Fan out 6 seed rows", "Create seed row", 0),
        ("Create seed row", "Log to Automation Log", 0),
    ]
    write_workflow("02_A02_provision_fee_schedule.json",
                   "KS A02 — Provision Fee Schedule on new Brief (v2.1.8 webhook)",
                   nodes, edges, tags=["ks-automation", "v2.1.8", "webhook"])


# ────────────────────────────────────────────────────────────────────────
# Workflow 03 — C: Engineer Notified on Folder = true
# ────────────────────────────────────────────────────────────────────────
def workflow_03():
    trig = webhook_trigger("Webhook", "ks-c-engineer-notified")
    get_page = get_page_node("Get Project Page", [460, 300])
    gate = if_node(
        "Gate: Folder=true AND Engineer Notified=false",
        [
            cond_bool("={{ $json.properties.Folder.checkbox }}", True),
            cond_bool("={{ $json.properties[\"Engineer Notified\"].checkbox }}", False),
        ],
        [680, 300],
    )
    flip_body = '={{ JSON.stringify({ properties: { "Engineer Notified": { checkbox: true } } }) }}'
    flip = http_node("Flip Engineer Notified = true", "PATCH",
                     '=https://api.notion.com/v1/pages/{{ $("Get Project Page").item.json.id }}',
                     flip_body, [900, 240])
    log_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "C · Engineer notified for " + ($("Get Project Page").item.json.properties["Project Name"].title[0]?.plain_text || "Project") } }] },
    "Job": { select: { name: "C" } },
    "Outcome": { select: { name: "Success" } },
    "Project": { relation: [{ id: $("Get Project Page").item.json.id }] }
  }
}) }}'''
    logn = http_node("Log to Automation Log", "POST",
                     "https://api.notion.com/v1/pages", log_body, [1120, 240])
    nodes = [trig, get_page, gate, flip, logn]
    edges = [
        ("Webhook", "Get Project Page", 0),
        ("Get Project Page", "Gate: Folder=true AND Engineer Notified=false", 0),
        ("Gate: Folder=true AND Engineer Notified=false", "Flip Engineer Notified = true", 0),
        ("Flip Engineer Notified = true", "Log to Automation Log", 0),
    ]
    write_workflow("03_C_engineer_notified.json",
                   "KS C — Engineer Notified on Folder Created (v2.1.8 webhook)",
                   nodes, edges, tags=["ks-automation", "v2.1.8", "webhook"])


# ────────────────────────────────────────────────────────────────────────
# Workflow 04 — D: Admins Notified on Brief Status = Ready for Admin Review
# ────────────────────────────────────────────────────────────────────────
def workflow_04():
    trig = webhook_trigger("Webhook", "ks-d-admins-notified")
    get_page = get_page_node("Get Brief Page", [460, 300])
    gate = if_node(
        "Gate: Status=Ready for Admin Review AND Admins Notified=false",
        [
            cond_string_eq("={{ $json.properties.Status.select?.name || \"\" }}",
                           "Ready for Admin Review"),
            cond_bool("={{ $json.properties[\"Admins Notified\"].checkbox }}", False),
        ],
        [680, 300],
    )
    flip_body = '={{ JSON.stringify({ properties: { "Admins Notified": { checkbox: true } } }) }}'
    flip = http_node("Flip Admins Notified = true", "PATCH",
                     '=https://api.notion.com/v1/pages/{{ $("Get Brief Page").item.json.id }}',
                     flip_body, [900, 240])
    bump_body = '={{ JSON.stringify({ properties: { "Status": { select: { name: "Admin Review" } } } }) }}'
    bump = http_node("Bump Project.Status = Admin Review", "PATCH",
                     '=https://api.notion.com/v1/pages/{{ $("Get Brief Page").item.json.properties.Project.relation[0].id }}',
                     bump_body, [1120, 240])
    log_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "D · Admins notified for " + ($("Get Brief Page").item.json.properties.Name.title[0]?.plain_text || "Brief") } }] },
    "Job": { select: { name: "D" } },
    "Outcome": { select: { name: "Success" } },
    "Brief": { relation: [{ id: $("Get Brief Page").item.json.id }] }
  }
}) }}'''
    logn = http_node("Log to Automation Log", "POST",
                     "https://api.notion.com/v1/pages", log_body, [1340, 240])
    nodes = [trig, get_page, gate, flip, bump, logn]
    edges = [
        ("Webhook", "Get Brief Page", 0),
        ("Get Brief Page", "Gate: Status=Ready for Admin Review AND Admins Notified=false", 0),
        ("Gate: Status=Ready for Admin Review AND Admins Notified=false", "Flip Admins Notified = true", 0),
        ("Flip Admins Notified = true", "Bump Project.Status = Admin Review", 0),
        ("Bump Project.Status = Admin Review", "Log to Automation Log", 0),
    ]
    write_workflow("04_D_admins_notified.json",
                   "KS D — Admins Notified on Brief Ready for Review (v2.1.8 webhook)",
                   nodes, edges, tags=["ks-automation", "v2.1.8", "webhook"])


# ────────────────────────────────────────────────────────────────────────
# Workflow 05 — A: Folder creation CLOUD-NATIVE via Microsoft Graph
# Notion trigger: Intake Complete = checked (in Projects DB)
# Requires n8n env vars:
#   KS_AZURE_TENANT_ID
#   KS_AZURE_CLIENT_ID
#   KS_AZURE_CLIENT_SECRET
#   KS_GRAPH_DRIVE_ID  (optional; defaults to known drive)
# ────────────────────────────────────────────────────────────────────────
def workflow_05():
    trig = webhook_trigger("Webhook", "ks-a-folder-creation")
    get_page = get_page_node("Get Project Page", [380, 300])

    gate = if_node(
        "Gate: v2.1.8 Job A",
        [
            cond_bool("={{ $json.properties.Folder.checkbox }}", False),
            cond_bool("={{ $json.properties[\"Intake Complete\"].checkbox }}", True),
            cond_number_gt(
                "={{ ($json.properties.City?.rich_text?.[0]?.plain_text || \"\").length }}", 0,
            ),
            cond_number_gt(
                "={{ ($json.properties.State?.rich_text?.[0]?.plain_text || \"\").length }}", 0,
            ),
            cond_number_gt("={{ $json.properties.ENGINEER?.people?.length || 0 }}", 0),
            cond_number_gt("={{ $json.properties.Admin?.people?.length || 0 }}", 0),
        ],
        [580, 300],
    )

    token_body_params = [
        {"name": "client_id",     "value": "={{ $env.KS_AZURE_CLIENT_ID }}"},
        {"name": "client_secret", "value": "={{ $env.KS_AZURE_CLIENT_SECRET }}"},
        {"name": "scope",         "value": "https://graph.microsoft.com/.default"},
        {"name": "grant_type",    "value": "client_credentials"},
    ]
    get_token = {
        "parameters": {
            "method": "POST",
            "url": "=https://login.microsoftonline.com/{{ $env.KS_AZURE_TENANT_ID }}/oauth2/v2.0/token",
            "sendBody": True,
            "contentType": "form-urlencoded",
            "bodyParameters": {"parameters": token_body_params},
            "options": {},
        },
        "id": _uid(),
        "name": "Get Graph Token",
        "type": "n8n-nodes-base.httpRequest",
        "typeVersion": 4.2,
        "position": [780, 240],
    }

    create_folders_code = """const project = $("Get Project Page").item.json;
const tokenJson = $("Get Graph Token").item.json;
const token = tokenJson.access_token;
if (!token) throw new Error("No access_token from Microsoft OAuth");

const DRIVE_ID = $env.KS_GRAPH_DRIVE_ID
  || "b!6BuR8AeAzE-i8tkyElJZj4W5gEAFUoFLgaF5LKROL2PkDU4xFZP6Sq_6155-5AYr";
const ROOT_PATH = "_Projects";
const SUBFOLDERS = [
  "Calculations","City Comments","Construction Administration",
  "Contracts","Correspondence","Drafting","Drawings",
  "Geotech Report","Notes","QAQC","Redlines","Sent to Client","Site Visit",
];

const projectName = (project.properties["Project Name"]?.title?.[0]?.plain_text || "").trim();
const m = projectName.match(/^(\\d{8})\\s+(.+)$/);
if (!m) throw new Error("Project Name lacks 8-digit number prefix: " + JSON.stringify(projectName));
const number = m[1];
const shortName = m[2].trim();
const year = "20" + number.substring(0, 2);

const city = (project.properties.City?.rich_text?.[0]?.plain_text || "").trim();
const state = (project.properties.State?.rich_text?.[0]?.plain_text || "").trim();
if (!city || !state) throw new Error("City and State must both be populated");

const folderName = number + " " + shortName + " - " + city + ", " + state;
const yearPath = ROOT_PATH + "/" + year;
const projectPath = yearPath + "/" + folderName;

async function createFolder(parentPath, name) {
  const url = "https://graph.microsoft.com/v1.0/drives/" + DRIVE_ID
            + "/root:/" + encodeURI(parentPath) + ":/children";
  const resp = await fetch(url, {
    method: "POST",
    headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({
      name: name, folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    }),
  });
  if (resp.status === 201) return await resp.json();
  if (resp.status === 409) {
    const itemPath = parentPath + "/" + name;
    const getUrl = "https://graph.microsoft.com/v1.0/drives/" + DRIVE_ID
                 + "/root:/" + encodeURI(itemPath);
    const getResp = await fetch(getUrl, { headers: { "Authorization": "Bearer " + token } });
    if (!getResp.ok) throw new Error("Conflict but GET failed (" + getResp.status + ") for " + itemPath);
    return await getResp.json();
  }
  const text = await resp.text();
  throw new Error("Folder create " + resp.status + " for " + parentPath + "/" + name + ": " + text);
}

await createFolder(ROOT_PATH, year);
const projectFolder = await createFolder(yearPath, folderName);
for (const sub of SUBFOLDERS) {
  await createFolder(projectPath, sub);
}

return [{ json: {
  projectId: project.id,
  projectName: projectName,
  folderName: folderName,
  webUrl: projectFolder.webUrl,
  year: year,
} }];"""
    create_folders = code_node("Create folder tree", create_folders_code, [980, 240])

    update_body = '''={{ JSON.stringify({
  properties: {
    "Folder": { checkbox: true },
    "SharePoint Folder": { url: $json.webUrl },
    "Status": { select: { name: "Proposal Brief Pending" } }
  }
}) }}'''
    update_project = http_node("Update Notion Project", "PATCH",
                               '=https://api.notion.com/v1/pages/{{ $json.projectId }}',
                               update_body, [1180, 240])

    log_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "A · Folder created (cloud) for " + ($("Create folder tree").item.json.folderName || "Project") } }] },
    "Job": { select: { name: "A" } },
    "Outcome": { select: { name: "Success" } },
    "Project": { relation: [{ id: $("Create folder tree").item.json.projectId }] },
    "Details": { rich_text: [{ type: "text", text: { content: $("Create folder tree").item.json.webUrl || "" } }] }
  }
}) }}'''
    logn = http_node("Log to Automation Log", "POST",
                     "https://api.notion.com/v1/pages", log_body, [1380, 240])

    nodes = [trig, get_page, gate, get_token, create_folders, update_project, logn]
    edges = [
        ("Webhook", "Get Project Page", 0),
        ("Get Project Page", "Gate: v2.1.8 Job A", 0),
        ("Gate: v2.1.8 Job A", "Get Graph Token", 0),
        ("Get Graph Token", "Create folder tree", 0),
        ("Create folder tree", "Update Notion Project", 0),
        ("Update Notion Project", "Log to Automation Log", 0),
    ]
    write_workflow("05_A_folder_creation_cloud.json",
                   "KS A — Folder Creation cloud-native (v2.1.8 webhook, Microsoft Graph)",
                   nodes, edges, tags=["ks-automation", "v2.1.8", "webhook", "cloud-native"])


# ────────────────────────────────────────────────────────────────────────
# Workflow 06 — B: Contract Render (still hybrid — calls PC)
# ────────────────────────────────────────────────────────────────────────
def workflow_06():
    trig = webhook_trigger("Webhook", "ks-b-contract-render")
    get_page = get_page_node("Get Brief Page", [460, 300])
    gate = if_node(
        "Gate: Status=Approved AND not yet rendered",
        [
            cond_string_eq("={{ $json.properties.Status.select?.name || \"\" }}", "Approved"),
            cond_object_empty("={{ $json.properties[\"Rendered At\"].date }}"),
        ],
        [680, 300],
    )
    dispatch = http_node(
        "POST /render-contract to PC", "POST",
        "={{ $env.KS_WEBHOOK_BASE_URL }}/render-contract",
        '={{ JSON.stringify({ brief_id: $json.id }) }}',
        [900, 240],
        headers_extra=[{"name": "X-KS-Token", "value": "={{ $env.KS_WEBHOOK_TOKEN }}"}],
        auth_notion=False,
        timeout_ms=120000,
    )
    succ_check = if_node("Render succeeded?",
                         [cond_bool("={{ $json.ok }}", True)],
                         [1120, 240])
    log_succ_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "B · Contract rendered" } }] },
    "Job": { select: { name: "B" } },
    "Outcome": { select: { name: "Success" } },
    "Brief": { relation: [{ id: $("Get Brief Page").item.json.id }] },
    "Details": { rich_text: [{ type: "text", text: { content: ($json.contract || "") } }] }
  }
}) }}'''
    log_succ = http_node("Log success", "POST",
                         "https://api.notion.com/v1/pages",
                         log_succ_body, [1340, 180])
    log_fail_body = '''={{ JSON.stringify({
  parent: { type: "data_source_id", data_source_id: "''' + AUTOMATION_DS + '''" },
  properties: {
    "Log Entry": { title: [{ type: "text", text: { content: "B · Contract render FAILED" } }] },
    "Job": { select: { name: "B" } },
    "Outcome": { select: { name: "Failed" } },
    "Brief": { relation: [{ id: $("Get Brief Page").item.json.id }] },
    "Error": { rich_text: [{ type: "text", text: { content: String($json.error || "unknown") } }] }
  }
}) }}'''
    log_fail = http_node("Log failure", "POST",
                         "https://api.notion.com/v1/pages",
                         log_fail_body, [1340, 320])
    nodes = [trig, get_page, gate, dispatch, succ_check, log_succ, log_fail]
    edges = [
        ("Webhook", "Get Brief Page", 0),
        ("Get Brief Page", "Gate: Status=Approved AND not yet rendered", 0),
        ("Gate: Status=Approved AND not yet rendered", "POST /render-contract to PC", 0),
        ("POST /render-contract to PC", "Render succeeded?", 0),
        ("Render succeeded?", "Log success", 0),
        ("Render succeeded?", "Log failure", 1),
    ]
    write_workflow("06_B_dispatch_contract_render.json",
                   "KS B — Dispatch Contract Render (v2.1.8 webhook, calls PC)",
                   nodes, edges, tags=["ks-automation", "v2.1.8", "webhook", "hybrid-pc"])


if __name__ == "__main__":
    print(f"Generating v2.1.8 webhook workflow JSONs into {OUT_DIR}")
    workflow_01()
    workflow_02()
    workflow_03()
    workflow_04()
    workflow_05()
    workflow_06()
    print("Done.")
