# Kingdom Structural — n8n Workflows v2.1.8 (Webhook Mode)

Fresh set of 6 workflow JSONs generated from `n8n/build_workflows.py`.
All workflows use **Notion outgoing-webhook triggers** (real-time, ~2s latency)
instead of polling.

## Common pattern

Every workflow follows the same shape:

```
Webhook ──▶ Get [Project|Brief] Page ──▶ Gate ──▶ rest of pipeline
```

The Webhook receives Notion's POST (containing only the page ID), then `Get
... Page` fetches the full page properties via Notion API. Downstream nodes
read `$("Get Project Page").item.json.properties.*` (or `Get Brief Page`).

## Workflow → Notion automation mapping

For each workflow, set up a **separate Notion automation** on the indicated
database with the indicated trigger. Action: "Send webhook" with the Production
URL from the corresponding n8n workflow.

| Workflow | Webhook path slug | Notion DB | Notion automation trigger |
|---|---|---|---|
| 01 — A0 (Create Brief) | `ks-a0-create-brief` | Projects | `Intake Complete` set to checked |
| 02 — A02 (Fee Schedule) | `ks-a02-fee-schedule` | Proposal Brief | `Page added` |
| 03 — C (Engineer Notified) | `ks-c-engineer-notified` | Projects | `Folder` set to checked |
| 04 — D (Admins Notified) | `ks-d-admins-notified` | Proposal Brief | `Status` set to `Ready for Admin Review` |
| 05 — A (Folder Creation) | `ks-a-folder-creation` | Projects | `Intake Complete` set to checked |
| 06 — B (Contract Render) | `ks-b-contract-render` | Proposal Brief | `Status` set to `Approved` |

Notes:
- Workflows 01 and 05 share the **same Notion trigger** (Intake Complete checked).
  That's intentional — both should run when intake completes. They have separate
  webhooks so each can be enabled/disabled independently.
- For each Notion automation, leave **Content → Select all existing properties**
  checked so the page ID is included in the payload.

## Production URLs

After importing each workflow and clicking **Publish**, the Production URL is:

```
https://automation.kingdomstructural.com/webhook/<path-slug>
```

For example, workflow 01 becomes:
```
https://automation.kingdomstructural.com/webhook/ks-a0-create-brief
```

Paste this into the matching Notion automation's URL field.

## Environment variables (for workflows 05 and 06 only)

Workflows 05 (Job A — folder creation) and 06 (Job B — contract render) call
back to the PC's `webhook_server.py`. They reference two env vars:

| Variable | Example | Where it's set |
|---|---|---|
| `KS_WEBHOOK_BASE_URL` | `https://abc.trycloudflare.com` | n8n → Settings → Variables |
| `KS_WEBHOOK_TOKEN` | (random 32+ char string) | n8n env vars AND PC's `.env` (must match) |

## Import procedure

1. **Delete the current consolidated workflow** in n8n (the one that was
   merged from 6 into 1 a few weeks back).
2. For each of the 6 `*.json` files, **Workflows → Import from File**.
3. Open each imported workflow, click each Notion or HTTP-to-Notion node, confirm
   the `Notion account` credential is bound. (Should auto-bind by name.)
4. **Publish** each workflow individually.
5. Copy each Production URL into its matching Notion automation.
6. Test workflow 01 first — create a Project, fill in fields, check Intake
   Complete. A Brief should appear. Then activate 02 (which fires automatically
   on Brief creation), and so on.

## Credential reference

All Notion-touching nodes reference an n8n credential named **`Notion account`**.
If yours is named differently, edit `CREDENTIAL_NAME` in `build_workflows.py`
and regenerate.

## Regenerating

```bash
cd _automation/n8n
python build_workflows.py
```

The builder is idempotent — re-runs overwrite the 6 JSONs. Edit the builder,
not the JSONs, when something needs to change.
