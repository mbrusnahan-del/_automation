# Kingdom Structural — n8n Migration Package

This folder contains everything needed to migrate the six automation jobs
(A0, A02, A, C, D, B) from the PC-based `sweep.py` runner to n8n.

## Architecture

n8n handles the **Notion-side** automation. The PC handles the **filesystem-side**
automation (OneDrive folder creation, Word contract rendering) via a small
HTTP server. The two sides talk to each other over HTTP.

```
   ┌──────────────────────┐
   │       Notion         │
   │  (Projects, Briefs,  │
   │   Automation Log)    │
   └──────────┬───────────┘
              │ polled / webhook
              ▼
   ┌──────────────────────┐                ┌────────────────────┐
   │         n8n          │   POST job     │  webhook_server.py │
   │  (workflows for      │◄──────────────►│   (runs on your    │
   │   Jobs A0, A02,      │   HTTP  result │        PC)         │
   │   A→, C, D, B→)      │                │                    │
   └──────────────────────┘                │  creates folders,  │
                                           │  renders .docx,    │
                                           │  reports back      │
                                           └──────────┬─────────┘
                                                      │
                                                      ▼
                                           ┌────────────────────┐
                                           │  OneDrive folder   │
                                           │   (_Projects/...)  │
                                           └────────────────────┘
```

Jobs running fully in n8n (no PC dependency):
- **A0** — Auto-create Proposal Brief on new Project
- **A02** — Provision Fee Schedule on new Brief
- **C** — Flip `Engineer Notified` when Project.Folder becomes true
- **D** — Flip `Admins Notified` and bump Project.Status when Brief.Status = Ready for Admin Review

Jobs that need the PC (because filesystem / Word rendering):
- **A** — Create OneDrive folder tree and stamp SharePoint URL on Project
- **B** — Render contract `.docx` into the project's `/Contracts/` folder

The hybrid approach means:
- Most of the pipeline runs in the cloud (fast, event-driven, easy to monitor)
- The two pieces that fundamentally need local file access stay local
- Everything is observable in n8n's execution log AND the Notion Automation Log

Later, if you move off OneDrive/local rendering (e.g. Microsoft Graph API
or a cloud doc generation service), Jobs A and B become pure n8n workflows
too and you can shut off `webhook_server.py`.

---

## Files in this folder

| File | Purpose |
|---|---|
| `README.md` | this file |
| `ARCHITECTURE.md` | deeper dive on how each workflow works and why |
| `webhook_server.py` | PC-side HTTP server for folder creation + contract rendering |
| `run_webhook_server.bat` | Windows launcher for the PC server |
| `01_A0_create_brief.json` | n8n workflow: new Project → create Brief with template body |
| `02_A02_provision_fee_schedule.json` | n8n workflow: new Brief → create Fee Schedule child DB + seed rows |
| `03_C_engineer_notified.json` | n8n workflow: Project.Folder = true → flip Engineer Notified |
| `04_D_admins_notified.json` | n8n workflow: Brief ready for review → flip Admins Notified, bump Project.Status |
| `05_A_dispatch_folder_creation.json` | n8n workflow: new Project gated → POST to PC to create OneDrive folder |
| `06_B_dispatch_contract_render.json` | n8n workflow: Brief approved → POST to PC to render contract |

---

## Migration steps

### Prerequisites

1. n8n running — either **n8n Cloud** ($20/mo, easiest), **self-hosted** (Docker on a VPS), or **local** (not recommended for production; requires your laptop to be on).
2. Notion internal integration with the three KS databases shared:
   - 🏗️ Projects
   - 📄 Proposal Brief
   - 🤖 Automation Log
   - (Same integration you set up for `sweep.py` — token goes in n8n instead of `.env`)

### Step 1 — Install Notion credentials in n8n

1. In n8n, open **Credentials** → **Add Credential** → **Notion API**.
2. Paste the internal integration secret from Notion.
3. Name it `KS Notion` (the workflow JSONs reference this credential name).
4. Save.

### Step 2 — Import the workflows

For each of the six `*.json` files:

1. In n8n, click **Workflows** → **Import from File**.
2. Select the JSON file.
3. The workflow appears in your list (initially inactive).
4. Open it and click each Notion or HTTP node once to confirm the credential
   binding is correct (they default to the `KS Notion` credential name).

### Step 3 — Set up the PC-side webhook server

On your PC, in the `_automation` folder:

```powershell
python n8n\webhook_server.py
```

The server listens on `http://localhost:8787` by default. Leave this
window open — closing it stops the server.

To auto-start on login, create a Windows shortcut to `run_webhook_server.bat`
in your Startup folder (`shell:startup` in Win+R).

### Step 4 — Make the PC reachable from n8n

If n8n is running on the same machine as the webhook server, use
`http://localhost:8787` directly — already the default.

If n8n is running in the cloud (n8n Cloud, or a VPS), you need to expose
the PC to the internet. Two clean options:

**Option A — Cloudflare Tunnel (free, recommended):**
```powershell
# one-time install
winget install --id Cloudflare.cloudflared
# persistent tunnel (replace with your chosen hostname)
cloudflared tunnel --url http://localhost:8787
```
Cloudflare prints a public URL like `https://some-words-here.trycloudflare.com`.
Paste that URL into the `KS_WEBHOOK_BASE_URL` env var inside each of the
n8n dispatch workflows (05 and 06).

**Option B — ngrok (free tier OK):**
```powershell
winget install ngrok
ngrok http 8787
```
Same deal — use the printed URL.

### Step 5 — Activate the workflows

In n8n, one by one:
1. Click each workflow → toggle **Active**.
2. Watch the first firing in the execution log.
3. Confirm Notion state updates as expected.

### Step 6 — Shut off the old PC sweep

Once the six n8n workflows are active and firing correctly, turn off the
Windows Task Scheduler entry for `run_sweep.bat` so it doesn't run in
parallel:

- Win+R → `taskschd.msc`
- Find **KS Automation Sweep**
- Right-click → **Disable** (or Delete)

The PC will still run `webhook_server.py` (for Jobs A and B), but nothing
polls Notion anymore — n8n does it.

---

## Upgrade path: from polling to true webhooks

The workflows imported here use n8n's **Notion Trigger** node, which
polls the database at a configurable interval (default: 1 minute).
That's already 10x faster than the current 10-min sweep and near-real-time
in practice.

If you want **true real-time webhooks** (event-driven, zero polling):

1. In Notion, enable **Outgoing Webhooks** on each database. Requires a
   paid Notion plan; configured in Settings → Integrations → Webhooks.
2. In n8n, replace each Notion Trigger node with a standard **Webhook**
   node. n8n gives you a URL; paste it into Notion's webhook config.
3. Notion now POSTs directly to n8n on every change. Latency drops from
   ~60s to ~2s.

Polling is fine for production. Upgrade to webhooks when the polling
latency becomes a real friction point.

---

## Testing the migration

Recommended rollout sequence:

1. **Import all workflows but keep them inactive.** Open each one and
   click through the nodes to make sure the credential bindings are right
   and the database/property IDs match your workspace.

2. **Activate workflow 01 (A0) only.** Create a test Project in Notion.
   Watch n8n's execution log. Confirm a Brief gets created within 60s
   of the Project appearing.

3. **Activate workflow 02 (A02).** Keep the test Project. Confirm the
   Brief has a Fee Schedule child DB appear within 60s of the Brief being
   created.

4. **Activate workflows 03 (C) and 04 (D).** Manually flip
   `Project.Folder` to true on the test Project. Confirm
   `Engineer Notified` flips true within 60s.

5. **Activate workflows 05 (A) and 06 (B).** These require
   `webhook_server.py` running on your PC. Pick one project, clear
   `Folder = false` on the Project, wait for the dispatch to run, confirm
   the folder appears on disk.

6. **Let it run for a day** against production data. Compare the n8n
   execution log against the Automation Log in Notion — the numbers
   should match.

7. **Shut off the old sweep.py Task Scheduler entry.**

---

## Observability

Three places to watch the system operate:

1. **n8n Executions view** — one row per workflow firing, green/red status,
   timing, full input/output JSON for debugging. This is your day-to-day
   monitoring surface.

2. **Notion Automation Log DB** — same rows the old sweep.py was writing,
   still populated by these workflows. Useful for partner-level visibility
   ("did this project get a Brief yet?").

3. **`webhook_server.log`** on the PC — timestamped log of folder creation
   and contract render requests. Check this if Jobs A or B aren't firing.

---

## Reverting

If something breaks and you want to revert to the PC-based sweep:

1. Deactivate all six n8n workflows.
2. Re-enable the `KS Automation Sweep` task in Windows Task Scheduler.
3. The PC sweep is fully idempotent — it'll pick up wherever n8n left off
   without creating duplicates.

No data migration or cleanup required either direction.
