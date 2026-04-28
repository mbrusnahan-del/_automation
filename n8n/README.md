# Kingdom Structural — Notion → n8n → Render Service

Event-driven replacement for the `notion-project-folder-sync` scheduled task.
Instead of polling every 10 minutes, Notion fires a webhook the moment a
project changes; n8n forwards it to a small Python service (Cloud Run) that
creates the OneDrive folder from the template and/or renders the contract +
fee memo. End‑to‑end latency is seconds instead of minutes, and nothing
depends on the Cowork app being open on your laptop.

Design goal: **keep n8n thin.** The workflow is 6 nodes (webhook → extract id
→ guard → HTTP → respond). All real logic lives in the Python service, so
there is almost no n8n surface area to debug.

```
Notion DB automation ──▶ n8n webhook ──▶ FastAPI /process ──▶ Graph API (OneDrive)
                                                         └──▶ Notion patch (Folder=true, URL)
```

---

## Files

```
_automation_n8n/
├── kingdom-structural-notion-automation.json    ← import into n8n
├── render_service/
│   ├── main.py                                  ← FastAPI app
│   ├── render_proposal_package.py               ← vendored pipeline
│   ├── render_from_notion.py                    ← vendored
│   ├── fee_memo.py                              ← vendored
│   ├── scope_library.py                         ← vendored
│   ├── contract_template_merge_ready.docx       ← vendored asset
│   ├── comps_cache.json                         ← vendored asset
│   ├── requirements.txt
│   └── Dockerfile
└── README.md                                    ← you are here
```

---

## 1. Azure AD app (Graph API, one time)

The service needs app-only access to your OneDrive to copy the template and
upload rendered docs.

1. Azure Portal → **Microsoft Entra ID** → **App registrations** → **New registration**.
   - Name: `Kingdom Render Service`
   - Supported account types: *Single tenant*
   - Redirect URI: leave blank.
2. Copy the **Application (client) ID** and **Directory (tenant) ID** — you'll
   need these as env vars.
3. **Certificates & secrets** → **New client secret** → 24 months → copy the
   **Value** (not the Secret ID). This is `GRAPH_CLIENT_SECRET`.
4. **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Application permissions** → add `Files.ReadWrite.All`. Then click
   **Grant admin consent for Kingdom Structural**.
5. Note the UPN of the OneDrive that owns `_Projects`
   (`michaelbrusnahan@kingdomstructural.com`). That is `GRAPH_USER_UPN`.

---

## 2. Notion integration (one time)

1. https://www.notion.so/my-integrations → **New integration** →
   type *Internal* → name `Kingdom Render Service`. Copy the secret; that is
   `NOTION_TOKEN`.
2. Capabilities: Read content, Update content, Insert content (default).
3. Open each database and **Connections → Add connection → Kingdom Render
   Service**. Minimum required:
   - Projects (`262b73dc-460e-8137-b3bb-000b62103b15`)
   - Clients (`2efb73dc-460e-80fe-860c-000bdfd01348`)
   - Contacts (`c88189fd-5bdf-46ad-a02e-9f426fd58527`)
   - Proposal Brief (`6b64658f-5fb3-4329-b02c-3ac3cb8a0828`)
   - Contract Fee Lines (`e05398ef-c999-477b-9420-3ccd4976730a`)
   - Jobs 26KS (`c58ca341-4985-45bf-a17f-19df8553a132`)
   - Scope Templates (`5480dd16-0c7d-4aaa-9664-e7e6b3a0fb4a`)
   - People (`72ab13bc-f91f-47c4-9759-e392e0f79e1e`)

---

## 3. Deploy the render service to Cloud Run

Prereqs: `gcloud` CLI authenticated against the Kingdom GCP project, Docker
running locally (or use Cloud Build — preferred, no Docker needed).

From `_automation_n8n/render_service/`:

```bash
# Choose a project + region
PROJECT=kingdom-structural                 # or whatever GCP project you use
REGION=us-central1
SERVICE=kingdom-render

gcloud config set project $PROJECT

# Build + deploy in one shot (Cloud Build builds the image, no local Docker)
gcloud run deploy $SERVICE \
  --source . \
  --region $REGION \
  --platform managed \
  --allow-unauthenticated \
  --cpu 1 --memory 1Gi \
  --timeout 300 \
  --min-instances 0 --max-instances 3 \
  --set-env-vars "NOTION_TOKEN=secret-xxx" \
  --set-env-vars "GRAPH_TENANT_ID=<tenant-guid>" \
  --set-env-vars "GRAPH_CLIENT_ID=<client-guid>" \
  --set-env-vars "GRAPH_CLIENT_SECRET=<secret-value>" \
  --set-env-vars "GRAPH_USER_UPN=michaelbrusnahan@kingdomstructural.com" \
  --set-env-vars "SHARED_SECRET=<generate-a-random-32-char-string>"
```

For real production use secrets via **Secret Manager**:

```bash
printf '%s' 'secret-xxx' | gcloud secrets create notion-token --data-file=-
gcloud run services update $SERVICE --region $REGION \
  --update-secrets NOTION_TOKEN=notion-token:latest,\
GRAPH_CLIENT_SECRET=graph-client-secret:latest,\
SHARED_SECRET=render-shared-secret:latest
```

Grab the service URL — something like
`https://kingdom-render-xxxx-uc.a.run.app` — and test:

```bash
curl $SERVICE_URL/healthz
# → {"ok":true}
```

---

## 4. n8n: import workflow + credential

1. n8n UI → **Workflows → Import from File** →
   `kingdom-structural-notion-automation.json`.
2. **Credentials → New → Header Auth**
   - Name: `Render Service Shared Secret`
   - Header name: `X-Kingdom-Auth`
   - Header value: *(same `SHARED_SECRET` you set on Cloud Run)*
3. Open the **Call Render Service** node → pick the new credential in the
   Authentication dropdown (replaces the `REPLACE_WITH_RENDER_SERVICE_CREDENTIAL_ID`
   placeholder the first time you save).
4. Either set a workflow env var `RENDER_SERVICE_URL` in n8n Settings or
   hardcode the Cloud Run URL in that node's URL field.
5. **Activate** the workflow. Copy the production webhook URL — it looks like
   `https://n8n.yourhost.com/webhook/kingdom-notion-project`.

Sanity test (simulates a Notion webhook):

```bash
curl -X POST https://n8n.yourhost.com/webhook/kingdom-notion-project \
  -H 'Content-Type: application/json' \
  -d '{"data":{"id":"262b73dc-460e-8137-b3bb-000b62103b15"}}'
```

You should see a JSON response with `ok: true` and a `service_result`
block from the render service.

---

## 5. Notion database automations (wire the triggers)

Open the **Projects** database → **···** menu → **Automations** → **New
automation**. Create two automations (both point at the same n8n URL — the
service decides what to do based on state):

### Automation 1 — New project created

- **Trigger:** Page added to Projects
- **Action:** Send webhook
  - URL: `https://n8n.yourhost.com/webhook/kingdom-notion-project`
  - Method: POST

### Automation 2 — Proposal requested

- **Trigger:** Page properties change → *Engineering Status* becomes
  `Proposal Requested`
- **Action:** Send webhook (same URL)

Notion's outbound payload wraps the page in `body.data.id` — the n8n **Extract
Page ID** node already handles that shape plus a few fallbacks.

> Optional third automation: trigger on **Proposal Brief → Ready to Render =
> true**, with the webhook targeting the *Project* the Brief links to. If you
> want this, we'll need a second Notion automation that resolves the relation
> first; easiest is to have the service accept a brief ID and do the lookup
> itself. Say the word and we'll wire it.

---

## 6. What happens on each event

The render service `/process` endpoint is one call that does the right thing
based on project state:

1. Fetch the Notion project page.
2. If **Folder = unchecked** → copy the canonical template to
   `_Projects/{YEAR}/{job# project - city, STATE}` via Graph API → patch
   Notion: `Folder = true`, `SharePoint Folder = <web URL>`.
3. If **Engineering Status = Proposal Requested** → fetch the Client and
   primary Contact pages → render contract `.docx` + fee memo `.docx` using
   the existing Python pipeline → upload both to the project's
   `Contracts/` folder.
4. Idempotency: before rendering, the service looks in `Contracts/` for files
   already prefixed with `{job#} Contract` or `{job#} Fee Analysis`. If found,
   it logs "already rendered" and skips — so firing the webhook twice is safe.

Everything is logged as structured JSON to Cloud Run stdout; check the Cloud
Run console "Logs" tab if a project doesn't appear where expected.

---

## 7. What about the old scheduled task?

The 10-minute `notion-project-folder-sync` task can stay disabled. The
separate `excel-comps-sync` task is unaffected — it still refreshes
`comps_cache.json` on its own schedule. After a comp cache refresh, rebuild
and redeploy the Cloud Run image so the new cache ships with it:

```bash
cd _automation_n8n/render_service
cp "/path/to/_automation/comps_cache.json" ./
gcloud run deploy kingdom-render --source . --region us-central1
```

(If comp freshness becomes a pain point, next iteration: have the service
pull the cache from a GCS bucket at startup instead of baking it into the
image.)

---

## Troubleshooting quickref

| Symptom                                            | Where to look                                               |
| -------------------------------------------------- | ----------------------------------------------------------- |
| n8n webhook returns 400 "missing page_id"          | Notion automation payload shape — test with curl above       |
| 401 from render service                            | `SHARED_SECRET` mismatch between Cloud Run and n8n credential |
| Template copy fails with 403                       | Graph `Files.ReadWrite.All` not admin-consented              |
| `Folder` never flips to checked in Notion          | Integration not added to Projects DB, or `NOTION_TOKEN` wrong |
| Docx renders but uploads 0 bytes                   | Check Cloud Run logs — likely pipeline exception pre-upload  |
| Duplicate renders                                  | Shouldn't happen (idempotency check), but if it does the job number prefix on uploaded filenames has drifted — verify `project_folder_name()` returns the same `proj_short` on both calls |
