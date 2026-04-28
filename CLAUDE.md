# CLAUDE.md — Kingdom Structural Notion Automation

Claude Code reads this file on startup. It gives you (Claude) the project
context and the guardrails the user cares about so you can be useful on the
first prompt instead of spelunking for ten minutes.

## What this is

Event-driven replacement for a 10-minute Cowork scheduled task that used to
poll the Notion Projects DB and sync OneDrive project folders + render
proposal packages. Architecture now:

```
Notion DB automation ──▶ n8n webhook ──▶ FastAPI /process  (Cloud Run)
                                               ├─▶ Graph API (copy template, upload docx)
                                               └─▶ Notion (patch Folder=true, SharePoint URL)
```

n8n is intentionally thin (6 nodes, almost no expressions). All real logic
lives in the Python service so there's one place to fix bugs.

## File map

```
_automation_n8n/
├── README.md                                     ← ops/deploy guide (humans)
├── CLAUDE.md                                     ← you are here
├── kingdom-structural-notion-automation.json     ← n8n workflow (import as-is)
└── render_service/
    ├── main.py                                   ← FastAPI app, all endpoints
    ├── render_proposal_package.py                ← top-level pipeline (vendored)
    ├── render_from_notion.py                     ← vendored
    ├── fee_memo.py                               ← vendored
    ├── scope_library.py                          ← vendored
    ├── contract_template_merge_ready.docx        ← asset (vendored)
    ├── comps_cache.json                          ← asset (vendored, refreshed by a separate task)
    ├── requirements.txt
    └── Dockerfile
```

The four `*.py` pipeline files and the two asset files are **vendored copies**
from `C:/Users/MichaelBrusnahan/OneDrive - Kingdom Structural LLC/_Projects/_automation/`.
That folder is the source of truth; when it changes, re-copy into
`render_service/` and redeploy.

## Endpoints (main.py)

- `GET  /healthz` — liveness, returns `{"ok": true}`
- `POST /process` — accepts `{"page_id": "...", "source": "..."}`; the entry
  point n8n calls. Decides whether to create the folder, render the proposal,
  or both, based on current Notion state.
- `POST /process/raw` — accepts a raw Notion outbound-webhook payload
  (`body.data.id` etc.) for manual/testing use.

Auth: optional `X-Kingdom-Auth: <SHARED_SECRET>` header. n8n's HTTP node sends
it via the `Render Service Shared Secret` credential.

## Env vars the service needs

| Var                   | Purpose                                     |
| --------------------- | ------------------------------------------- |
| `NOTION_TOKEN`        | Internal integration secret                 |
| `GRAPH_TENANT_ID`     | Azure AD tenant GUID                        |
| `GRAPH_CLIENT_ID`     | App registration client ID                  |
| `GRAPH_CLIENT_SECRET` | App registration secret value               |
| `GRAPH_USER_UPN`      | `michaelbrusnahan@kingdomstructural.com`    |
| `SHARED_SECRET`       | Matches the `X-Kingdom-Auth` in n8n         |
| `PORT`                | Set automatically by Cloud Run              |

Module constants in `main.py` (edit if paths change):
`TEMPLATE_DRIVE_PATH`, `PROJECTS_DRIVE_PATH`, `ONEDRIVE_HOST_ROOT`.

## Notion data sources (canonical IDs)

```
Projects:         262b73dc-460e-8137-b3bb-000b62103b15
Clients:          2efb73dc-460e-80fe-860c-000bdfd01348
Contacts:         c88189fd-5bdf-46ad-a02e-9f426fd58527
Jobs 26KS:        c58ca341-4985-45bf-a17f-19df8553a132
Scope Templates:  5480dd16-0c7d-4aaa-9664-e7e6b3a0fb4a
Proposal Brief:   6b64658f-5fb3-4329-b02c-3ac3cb8a0828
Contract Fee Lines: e05398ef-c999-477b-9420-3ccd4976730a
People:           72ab13bc-f91f-47c4-9759-e392e0f79e1e
```

## Schema rework (2026-04-22)

These fields moved **off** Projects and **onto** Proposal Brief. Any code that
still reads them from the Project page is wrong:

- ICC Code Year, Jurisdiction, City, State, Project Type, Project Street
- Reimbursables Treatment (parsed from Brief body checkbox)

Projects keeps: Project Name, Place, Approx. SF, Scope Description, DB
relations, people, `Engineering Status`, `Folder` (checkbox),
`SharePoint Folder` (url), `Proposal Brief` (relation).

## Workflow gates — do not render prematurely

- Folder sync runs only when `Project.Folder == false`.
- Proposal render runs only when:
  1. `Project.Folder == true` (folder must exist first)
  2. `Project.Engineering Status == "Proposal Requested"`
  3. Linked Proposal Brief has `Ready to Render == true`

Always log per-project status on skip ("awaiting Brief link", "Ready to Render
not checked", "already rendered") so runs stay debuggable.

## Idempotency rules (non-negotiable)

- Never delete or overwrite files in an existing project folder.
- Before rendering, check `{project}/Contracts/` for files already prefixed
  with `{job#} Contract` or `{job#} Fee Analysis`. If found, log
  "already rendered" and skip.
- `populate_from_template()` walks the template and only copies files that
  don't exist in the destination.
- Individual project errors must not stop processing of others (in the batch
  path — the webhook path is single-project so this is about future batch
  endpoints).

## Windows MAX_PATH

`project_folder_name(job_number, project_name_raw, city, state, client_name)`
in `render_proposal_package.py` truncates so the longest deliverable path
stays under 259 chars. If you change filename prefixes in the contract/fee
memo templates, re-check this helper.

## Deploy / iterate loop

Local smoke test:
```bash
cd render_service
pip install -r requirements.txt
export NOTION_TOKEN=... GRAPH_TENANT_ID=... GRAPH_CLIENT_ID=... \
       GRAPH_CLIENT_SECRET=... GRAPH_USER_UPN=michaelbrusnahan@kingdomstructural.com
uvicorn main:app --reload --port 8080
curl -X POST localhost:8080/process -H 'Content-Type: application/json' \
     -d '{"page_id":"<known-project-page-id>"}'
```

Deploy (Cloud Build, no local Docker needed):
```bash
gcloud run deploy kingdom-render --source render_service \
  --region us-central1 --allow-unauthenticated \
  --cpu 1 --memory 1Gi --timeout 300 \
  --min-instances 0 --max-instances 3
# env vars: set via Secret Manager in prod, `--set-env-vars` for quick tests
```

Validate the n8n JSON before committing anything that touches it:
```bash
python - <<'PY'
import json
spec = json.load(open('kingdom-structural-notion-automation.json'))
names = {n['name'] for n in spec['nodes']}
for src, out in spec['connections'].items():
    assert src in names
    for b in out['main']:
        for e in b: assert e['node'] in names
print('n8n JSON OK')
PY
```

## Related but separate

- `excel-comps-sync` — different scheduled task, owns `comps_cache.json`. Do
  **not** touch the cache from this repo's automation. When comps refresh,
  re-copy the file from `_Projects/_automation/` and redeploy.

## Things the user cares about (observed)

- **Don't introduce n8n-specific expression soup.** The last n8n build had
  errors on nearly every node. Keep the workflow boring.
- **Minimize new env knobs.** Reuse what's in `main.py` before adding a new
  `os.getenv`.
- **Ask before running destructive git operations.** Normal commits fine;
  force-push, hard-reset, deleting branches — confirm first.
