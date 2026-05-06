"""
n8n_admin.py — local admin utility for the Kingdom Structural n8n workflows.

Reads N8N_BASE_URL and N8N_API_KEY from ../.env (gitignored).
Uses Python stdlib only (no pip dependencies).

Commands:
  python n8n_admin.py list                       # list all workflows
  python n8n_admin.py get <id-or-name-fragment>  # dump a workflow's JSON
  python n8n_admin.py deploy <path.json>         # create-or-update a workflow
                                                  from a JSON file, then
                                                  activate it. --no-activate
                                                  to skip activation.
  python n8n_admin.py fix-a02                    # patch workflow 02's
                                                  Create Fee Schedule body
  python n8n_admin.py activate <id-or-name>      # turn a workflow on
  python n8n_admin.py deactivate <id-or-name>    # turn a workflow off
"""
from __future__ import annotations
import http.client
import json
import os
import sys
import urllib.parse

_HERE = os.path.dirname(os.path.abspath(__file__))
_ENV_PATH = os.path.normpath(os.path.join(_HERE, "..", ".env"))


def load_env() -> dict:
    if not os.path.exists(_ENV_PATH):
        sys.exit(f"ERROR: {_ENV_PATH} not found.")
    env = {}
    with open(_ENV_PATH) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def request(env: dict, method: str, path: str, body=None) -> tuple[int, dict | str]:
    """Use http.client directly — urllib lowercases header names ('X-N8N-API-KEY'
    becomes 'X-n8n-api-key') which n8n's auth middleware rejects with 401.
    http.client preserves the case we send."""
    base = env["N8N_BASE_URL"].rstrip("/")
    key  = env["N8N_API_KEY"]
    parsed = urllib.parse.urlparse(base)
    if parsed.scheme != "https":
        sys.exit(f"ERROR: N8N_BASE_URL must be https, got {parsed.scheme!r}")

    full_path = f"/api/v1{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    headers = {
        "X-N8N-API-KEY": key,
        "Accept": "application/json",
    }
    if data is not None:
        headers["Content-Type"] = "application/json"

    conn = http.client.HTTPSConnection(parsed.hostname, parsed.port or 443, timeout=30)
    try:
        conn.request(method, full_path, body=data, headers=headers)
        resp = conn.getresponse()
        raw = resp.read().decode("utf-8")
        status = resp.status
    finally:
        conn.close()

    if not raw:
        return status, ""
    try:
        return status, json.loads(raw)
    except json.JSONDecodeError:
        return status, raw


def list_workflows(env: dict) -> list[dict]:
    code, data = request(env, "GET", "/workflows?limit=250")
    if code != 200:
        sys.exit(f"ERROR listing workflows: {code} {data}")
    return data.get("data", [])


# n8n's create/update workflow endpoints accept only a narrow allowlist of
# top-level fields; extras like "active", "tags", "versionId", "id",
# "createdAt", "updatedAt", "pinData", "meta", "triggerCount" cause
# "must NOT have additional properties" 400s. Same for nested 'settings' —
# fields like binaryMode, availableInMCP, timeSavedMode are read-only.
ALLOWED_TOP = {"name", "nodes", "connections", "settings", "staticData"}
ALLOWED_SETTINGS = {
    "saveExecutionProgress",
    "saveManualExecutions",
    "saveDataErrorExecution",
    "saveDataSuccessExecution",
    "executionOrder",
    "executionTimeout",
    "errorWorkflow",
    "timezone",
    "callerPolicy",
}


def _filter_workflow_body(payload: dict) -> dict:
    """Strip a payload down to fields n8n's create/update endpoints accept."""
    body = {k: v for k, v in payload.items() if k in ALLOWED_TOP}
    if "settings" in body and isinstance(body["settings"], dict):
        body["settings"] = {
            k: v for k, v in body["settings"].items() if k in ALLOWED_SETTINGS
        }
    return body


def find_workflow(env: dict, query: str) -> dict:
    """Find a workflow by ID or by name fragment (case-insensitive). Exits
    if zero or multiple matches."""
    match = find_workflow_optional(env, query)
    if match is None:
        sys.exit(f"ERROR: no workflow matching {query!r}.")
    return match


def find_workflow_optional(env: dict, query: str) -> dict | None:
    """Like find_workflow but returns None on no match. Still exits on
    multiple ambiguous matches (so callers don't deploy into the wrong one)."""
    workflows = list_workflows(env)
    # Direct ID match first
    for w in workflows:
        if w.get("id") == query:
            return w
    # Exact name match before fuzzy — protects against substring collisions
    # like "A0" matching "A02" when deploying.
    exact = [w for w in workflows if (w.get("name") or "") == query]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        names = "\n  ".join(f"{w['id']}  {w['name']}" for w in exact)
        sys.exit(f"ERROR: multiple workflows named exactly {query!r}:\n  {names}")
    matches = [w for w in workflows if query.lower() in (w.get("name") or "").lower()]
    if not matches:
        return None
    if len(matches) > 1:
        names = "\n  ".join(f"{w['id']}  {w['name']}" for w in matches)
        sys.exit(f"ERROR: multiple matches for {query!r}:\n  {names}")
    return matches[0]


def get_workflow(env: dict, wf_id: str) -> dict:
    code, data = request(env, "GET", f"/workflows/{wf_id}")
    if code != 200:
        sys.exit(f"ERROR getting workflow {wf_id}: {code} {data}")
    return data


def create_workflow(env: dict, payload: dict) -> dict:
    body = _filter_workflow_body(payload)
    code, data = request(env, "POST", "/workflows", body)
    if code not in (200, 201):
        sys.exit(f"ERROR creating workflow: {code} {data}")
    return data


def update_workflow(env: dict, wf_id: str, payload: dict) -> dict:
    body = _filter_workflow_body(payload)
    code, data = request(env, "PUT", f"/workflows/{wf_id}", body)
    if code not in (200, 201):
        sys.exit(f"ERROR updating workflow {wf_id}: {code} {data}")
    return data


def set_active(env: dict, wf_id: str, active: bool) -> dict:
    verb = "activate" if active else "deactivate"
    code, data = request(env, "POST", f"/workflows/{wf_id}/{verb}")
    if code not in (200, 201):
        sys.exit(f"ERROR {verb} workflow {wf_id}: {code} {data}")
    return data


# ────────────────────────────────────────────────────────────────────────
# Specific fix: patch the Create Fee Schedule child DB node body in A02.
# ────────────────────────────────────────────────────────────────────────

CORRECTED_A02_BODY = (
    '={{ JSON.stringify({\n'
    '  parent: { type: "page_id", page_id: $json.id },\n'
    '  title: [{ type: "text", text: { content: "💵 Fee Schedule" } }],\n'
    '  is_inline: true,\n'
    '  properties: {\n'
    '    "Service": { title: {} },\n'
    '    "Include": { formula: { expression: "if(empty(prop(\\"Amount\\")), false, prop(\\"Amount\\") > 0)" } },\n'
    '    "Type": { select: { options: [\n'
    '      { name: "Fixed Fee", color: "blue" },\n'
    '      { name: "Hourly", color: "yellow" },\n'
    '      { name: "Not to Exceed", color: "green" }\n'
    '    ] } },\n'
    '    "Amount": { number: { format: "dollar" } },\n'
    '    "Order": { number: { format: "number" } }\n'
    '  }\n'
    '}) }}'
)


def fix_a02(env: dict) -> None:
    wf = find_workflow(env, "A02")
    wf_id = wf["id"]
    print(f"Found workflow:  {wf_id}  '{wf['name']}'")

    full = get_workflow(env, wf_id)
    nodes = full.get("nodes") or []
    target = next((n for n in nodes
                   if n.get("name") == "Create Fee Schedule child DB"), None)
    if not target:
        sys.exit("ERROR: 'Create Fee Schedule child DB' node not found in workflow.")

    current = target["parameters"].get("jsonBody", "")
    if "initial_data_source" not in current and "properties: {" in current:
        print("Body already in correct shape — no change needed.")
        return

    print("Patching 'Create Fee Schedule child DB' jsonBody…")
    target["parameters"]["jsonBody"] = CORRECTED_A02_BODY

    print("Pushing updated workflow back to n8n…")
    update_workflow(env, wf_id, full)

    print("Activating workflow…")
    set_active(env, wf_id, True)

    print("Done. Verify by triggering a fresh test cascade in Notion.")


# ────────────────────────────────────────────────────────────────────────
# CLI
# ────────────────────────────────────────────────────────────────────────


def cmd_list(env, args):
    workflows = list_workflows(env)
    print(f"{'ID':<22}  {'ACTIVE':<7}  NAME")
    print("─" * 80)
    for w in workflows:
        wid = w.get("id", "")
        active = "✓" if w.get("active") else " "
        name = w.get("name", "")
        print(f"{wid:<22}  {active:<7}  {name}")


def cmd_get(env, args):
    if not args:
        sys.exit("Usage: get <id-or-name-fragment>")
    wf = find_workflow(env, args[0])
    full = get_workflow(env, wf["id"])
    print(json.dumps(full, indent=2))


def cmd_deploy(env, args):
    """Create-or-update a workflow from a JSON file, then activate it.

    Matches an existing workflow by exact name (the JSON's top-level "name").
    If found, deactivates → PUTs the new definition → reactivates.
    If not found, POSTs a fresh workflow.

    Usage: deploy <path-to-workflow.json> [--no-activate]
    """
    if not args:
        sys.exit("Usage: deploy <path-to-workflow.json> [--no-activate]")

    path = args[0]
    activate_after = "--no-activate" not in args[1:]

    if not os.path.exists(path):
        sys.exit(f"ERROR: file not found: {path}")

    with open(path, encoding="utf-8") as f:
        payload = json.load(f)

    name = payload.get("name")
    if not name:
        sys.exit("ERROR: workflow JSON has no top-level 'name' field.")

    print(f"Deploying {name!r} from {path}")

    existing = find_workflow_optional(env, name)
    if existing is not None:
        wf_id = existing["id"]
        was_active = existing.get("active", False)
        print(f"  Found existing workflow {wf_id} (active={was_active}).")
        # n8n's PUT /workflows/{id} rejects updates to active workflows on
        # some versions ("Workflow is active: deactivate before updating").
        # Always deactivate first; we'll reactivate at the end if requested.
        if was_active:
            print("  Deactivating before update…")
            try:
                set_active(env, wf_id, False)
            except SystemExit:
                pass
        print("  PUT /workflows/{id}…")
        update_workflow(env, wf_id, payload)
    else:
        print("  No existing workflow with that exact name — creating new.")
        created = create_workflow(env, payload)
        wf_id = created.get("id")
        if not wf_id:
            sys.exit(f"ERROR: create response missing 'id': {created!r}")
        print(f"  Created workflow {wf_id}.")

    if activate_after:
        print("  Activating…")
        set_active(env, wf_id, True)
        print(f"Done. Workflow {wf_id} is now active.")
    else:
        print(f"Done. Workflow {wf_id} deployed (not activated).")


def cmd_activate(env, args):
    if not args:
        sys.exit("Usage: activate <id-or-name>")
    wf = find_workflow(env, args[0])
    set_active(env, wf["id"], True)
    print(f"Activated {wf['id']}  '{wf['name']}'.")


def cmd_deactivate(env, args):
    if not args:
        sys.exit("Usage: deactivate <id-or-name>")
    wf = find_workflow(env, args[0])
    set_active(env, wf["id"], False)
    print(f"Deactivated {wf['id']}  '{wf['name']}'.")


def cmd_fix_a02(env, args):
    fix_a02(env)


def cmd_executions(env, args):
    """List recent executions for a workflow."""
    if not args:
        sys.exit("Usage: executions <id-or-name>")
    wf = find_workflow(env, args[0])
    code, data = request(env, "GET", f"/executions?workflowId={wf['id']}&limit=10&includeData=true")
    if code != 200:
        sys.exit(f"ERROR fetching executions: {code} {data}")
    rows = data.get("data", [])
    print(f"{'STARTED':<22}  {'STATUS':<10}  {'MODE':<10}  ID")
    print("─" * 80)
    for ex in rows:
        started = ex.get("startedAt", "")[:19]
        status = ex.get("status", "?")
        mode = ex.get("mode", "?")
        eid = ex.get("id", "")
        print(f"{started:<22}  {status:<10}  {mode:<10}  {eid}")


def cmd_execution(env, args):
    """Dump a single execution's full data (use to inspect failures)."""
    if not args:
        sys.exit("Usage: execution <execution-id>")
    code, data = request(env, "GET", f"/executions/{args[0]}?includeData=true")
    if code != 200:
        sys.exit(f"ERROR fetching execution {args[0]}: {code} {data}")
    print(json.dumps(data, indent=2))


def cmd_patch_a02_seed_parent(env, args):
    """Force A02 seed-row parent to database_id (instead of data_source_id).
    Notion's 2025-09-03 API returns data_sources[] in the create response,
    but using those IDs immediately as a parent often returns 404 (the
    propagation isn't instant). database_id parent works reliably right
    after creation. The Python sweep also uses this shape successfully."""
    wf = find_workflow(env, "A02")
    full = get_workflow(env, wf["id"])
    nodes = full.get("nodes") or []

    fanout = next((n for n in nodes if n.get("name") == "Fan out 6 seed rows"), None)
    seed   = next((n for n in nodes if n.get("name") == "Create seed row"), None)
    if not fanout or not seed:
        sys.exit("ERROR: required nodes not found.")

    new_fanout = (
        "const db = $input.first().json;\n"
        "// Use database_id parent for seed rows (more reliable than\n"
        "// data_source_id immediately after DB creation — Notion 2025-09-03\n"
        "// briefly returns 404 on the just-issued data_source ID).\n"
        "const databaseId = db.id;\n"
        "const seeds = [\n"
        "  \"Schematic Design\",\n"
        "  \"Design Development\",\n"
        "  \"Construction Documents\",\n"
        "  \"Construction Administration\",\n"
        "  \"Special Structural Inspections\",\n"
        "  \"Engineering Site Visits\",\n"
        "];\n"
        "return seeds.map((service, i) => ({\n"
        "  json: { databaseId, service, order: i + 1 }\n"
        "}));"
    )
    fanout["parameters"]["jsCode"] = new_fanout

    seed["parameters"]["jsonBody"] = (
        '={{ JSON.stringify({\n'
        '  parent: { type: "database_id", database_id: $json.databaseId },\n'
        '  properties: {\n'
        '    "Service": { title: [{ type: "text", text: { content: $json.service } }] },\n'
        '    "Order": { number: $json.order }\n'
        '  }\n'
        '}) }}'
    )

    print("Deactivating A02 first (so PUT replaces cleanly)…")
    try:
        set_active(env, wf["id"], False)
    except SystemExit:
        pass  # already inactive is fine

    print("Pushing patched A02 (Fan out + Create seed row)…")
    update_workflow(env, wf["id"], full)

    # Verify the PUT actually took.
    refetched = get_workflow(env, wf["id"])
    seed_node = next((n for n in refetched.get("nodes") or []
                      if n.get("name") == "Create seed row"), None)
    if seed_node and "data_source_id" in seed_node["parameters"].get("jsonBody", ""):
        print("WARN: PUT didn't fully replace the seed body; data_source_id still present.")
        print("      Live body:")
        print(seed_node["parameters"]["jsonBody"])
    else:
        print("OK: live seed body now uses database_id parent.")

    print("Re-activating A02…")
    set_active(env, wf["id"], True)
    print("Done. Re-trigger to verify.")


COMMANDS = {
    "list":       cmd_list,
    "get":        cmd_get,
    "deploy":     cmd_deploy,
    "activate":   cmd_activate,
    "deactivate": cmd_deactivate,
    "fix-a02":    cmd_fix_a02,
    "executions": cmd_executions,
    "execution":  cmd_execution,
    "patch-a02-seed": cmd_patch_a02_seed_parent,
}


def main(argv):
    if not argv or argv[0] in ("-h", "--help", "help"):
        print(__doc__.strip())
        return 0
    cmd = argv[0]
    if cmd not in COMMANDS:
        print(f"Unknown command: {cmd}\n")
        print(__doc__.strip())
        return 1
    env = load_env()
    if not env.get("N8N_BASE_URL") or not env.get("N8N_API_KEY"):
        sys.exit("ERROR: N8N_BASE_URL and N8N_API_KEY must be set in .env")
    COMMANDS[cmd](env, argv[1:])
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
