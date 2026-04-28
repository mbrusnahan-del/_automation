# KS Automation — n8n Architecture Deep Dive

This document walks through each of the six n8n workflows, what triggers
them, what they do, and where state flows through the system.

## Shared context

- **Notion databases** live at these data source IDs (hardcoded in each workflow):
  - Projects: `262b73dc-460e-8137-b3bb-000b62103b15`
  - Proposal Brief: `6b64658f-5fb3-4329-b02c-3ac3cb8a0828`
  - Automation Log: `04d7aa20-eb8d-4d51-b302-25fb4eaaf26e`
- **Notion credential** in n8n must be named `KS Notion` (or update each workflow).
- **Brief template** lives at page `34ab73dc-460e-80bf-b97c-da13d53310e8`. Its
  body blocks are fetched and copied into every new Brief.
- **Notion API version** pinned to `2025-09-03` in every HTTP Request node.

## Workflow 01 — A0: Create Proposal Brief on new Project

**Triggered by:** New page in Projects DB (polled every 1 min).

**Gate:**
- Project has no linked Brief (`Proposal Brief.relation.length == 0`)
- Engineering Status is NOT in {Dead, IGNNROE FO RNOW, On Hold}

**Flow:**
1. Fetch the Brief template page's blocks (`GET /blocks/{template}/children`).
2. Clean the blocks: drop unsupported block types, strip null-valued fields
   (`icon: null` etc.) and Notion output-only metadata. Keep: paragraph,
   heading, to_do, callout, divider, etc.
3. Create a new page in Proposal Brief DB with:
   - `Name` = Project.Name
   - `Project` relation → the new Project
   - `Status` = "Not Started"
   - `children` = the cleaned template blocks
4. Patch the Project with `Proposal Brief` relation → new Brief.
5. Write a Success row to Automation Log.

**Failure modes:**
- Notion 429 rate limit → n8n's default retry handles it.
- Malformed block type → the "Clean template blocks" Code node filters it.
- Project is missing Name → the created Brief gets an empty title; partner
  can rename manually.

## Workflow 02 — A02: Provision Fee Schedule on new Brief

**Triggered by:** New page in Proposal Brief DB (polled every 1 min).

**No gate** beyond the trigger — every brand-new Brief gets a Fee Schedule.
The idempotency is implicit: if the Brief was just created by Workflow 01,
it can't have a Fee Schedule yet.

**Flow:**
1. Create child database `💵 Fee Schedule` under the Brief page with
   `initial_data_source.properties` containing Service / Include / Type /
   Amount / Order.
2. Extract `data_source_id` from the response (not `database_id` — the
   2025-09-03 API distinguishes them).
3. Fan out six seed rows (Schematic Design through Engineering Site Visits).
4. Create each seed row via `POST /v1/pages` with `parent.data_source_id`.
5. Write a Success row to Automation Log.

**Failure modes:**
- `database_id` vs `data_source_id` confusion — fixed here by always using
  `data_source_id` for pages.create.
- Seed row creation partial failure → 5 rows might succeed before one fails.
  The retry is manual; check the Automation Log.

## Workflow 03 — C: Engineer Notified on Folder Created

**Triggered by:** Update to any Project page (polled every 1 min).

**Gate:**
- `Folder` checkbox is true
- `Engineer Notified` checkbox is false

This gate fires exactly once per Project, when `Folder` flips from false to
true. After Workflow 03 runs, `Engineer Notified` becomes true and the gate
excludes the Project on subsequent updates.

**Flow:**
1. `PATCH /v1/pages/{project}` setting `Engineer Notified = true`.
2. Write a Success row to Automation Log.

**Downstream:** Your Notion workflow automation watches for
`Engineer Notified` flipping to true and sends the actual notification
(email / @mention / Slack / whatever you wired up).

## Workflow 04 — D: Admins Notified on Brief Ready for Review

**Triggered by:** Update to any Brief page (polled every 1 min).

**Gate:**
- Brief.Status = "Ready for Admin Review"
- Brief.Admins Notified = false

**Flow:**
1. `PATCH` the Brief → `Admins Notified = true`.
2. `PATCH` the linked Project → `Status = "Admin Review"`.
3. Write a Success row to Automation Log.

**Downstream:** Your Notion automation watches for the Project.Status change
and notifies admins.

## Workflow 05 — A: Dispatch OneDrive folder creation

**Triggered by:** Update to any Project page (polled every 1 min).

**Gate (v2.1.2):**
- `Folder` = false
- `City` populated
- `State` populated
- `ENGINEER.people.length > 0`
- `Admin.people.length > 0`

**Flow:**
1. `POST http://<your-pc>/create-folder` with `{ project_id: "..." }` and
   header `X-KS-Token: <shared secret>`.
2. PC-side webhook_server.py:
   - Fetches the Project from Notion
   - Re-checks the gate (in case state changed between n8n firing and
     the request arriving)
   - Parses the 8-digit KS job number from Project Name
   - Creates `_Projects/{YEAR}/{num} - {name} - {City}, {State}/`
   - Clones the template folder tree into it
   - Patches the Project: `Folder = true`, `SharePoint Folder = URL`,
     `Status = "Proposal Brief Pending"`
   - Writes a Success/Failed row to Automation Log (via the shared
     `sweep.log_row` helper)
   - Returns JSON `{ ok, path, url, stats }` or `{ ok: false, error }`
3. n8n branches on `ok`:
   - Success → write an A-Success row to Automation Log
   - Failure → write an A-Failed row with the error text

**Environment variables used by n8n:**
- `KS_WEBHOOK_BASE_URL` — e.g. `http://192.168.1.50:8787` (local LAN)
  or `https://abc-def.trycloudflare.com` (Cloudflare Tunnel)
- `KS_WEBHOOK_TOKEN` — must match the PC's `.env` entry

## Workflow 06 — B: Dispatch Contract Render

**Triggered by:** Update to any Brief page (polled every 1 min).

**Gate:**
- Brief.Status = "Approved"
- Brief.Rendered At is empty

**Flow:**
1. `POST http://<your-pc>/render-contract` with `{ brief_id: "..." }` and
   the shared token header. 60-second timeout because the render pipeline
   involves template cloning and docx manipulation.
2. PC-side webhook_server.py:
   - Fetches the Brief and its linked Project from Notion
   - Resolves Client, Project Contact, Engineer dicts (same helpers as
     sweep.py's Job B)
   - Reads the Fee Schedule child DB into fee_lines
   - Reconstructs the Brief body markdown
   - Calls `render_contract_package(...)`
   - Patches Brief: `Rendered At = now`, `Status = "Rendered"`
   - Patches Project: `Status = "Contract Rendered"`
   - Writes Success/Failed to Automation Log
   - Returns JSON result
3. n8n branches on `ok` and writes its own summary log row.

Both sides log to the Automation Log — slightly redundant but provides
cross-checking if something goes wrong.

## Environment variables you need to configure in n8n

Settings → Variables (or your n8n host's env file):

| Variable | Example | Used by |
|---|---|---|
| `KS_WEBHOOK_BASE_URL` | `https://abc.trycloudflare.com` | Workflows 05, 06 |
| `KS_WEBHOOK_TOKEN`    | (random 32+ char string) | Workflows 05, 06 |

The same `KS_WEBHOOK_TOKEN` must be set in the PC's `.env` file.

## Environment variables PC-side (webhook_server.py reads from .env)

| Variable | Purpose |
|---|---|
| `NOTION_TOKEN` | Same Notion integration secret used by sweep.py |
| `KS_WEBHOOK_TOKEN` | Shared secret that authenticates n8n requests |
| `KS_ONEDRIVE_WIN_ROOT` | (optional) Override OneDrive Windows path |

## Trade-offs of this architecture

**Pros:**
- True event-driven behavior — a new Project gets a folder + Brief in <90s.
- Zero PC dependency for 4 of 6 jobs.
- Clean separation: Notion-bound logic in n8n, filesystem/Word in Python.
- Same idempotency as the PC sweep — can run alongside the old sweep.py
  without duplicate writes.
- Observable in two places: n8n executions + Notion Automation Log.

**Cons:**
- More moving parts than pure PC sweep.
- `webhook_server.py` must be running for Jobs A and B to work.
- If the PC is offline, A and B queue up in Notion (gate stays true) and
  catch up next time the PC is online. Not a data-loss scenario, just a
  delay.
- n8n polls every minute — not truly real-time. Upgrade to true webhooks
  later if needed (see README § Upgrade path).

## Polling cost

n8n calls these Notion endpoints every minute:
- `POST /v1/data_sources/{projects}/query` (used by triggers 01, 03, 05)
- `POST /v1/data_sources/{briefs}/query` (used by triggers 02, 04, 06)

That's ~6 Notion reads per minute, or ~8,640 per day. Well under Notion's
rate limit of 3 requests/second (~260,000/day).
