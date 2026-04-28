---
name: notion-project-folder-sync
description: Every 10 min (v2.1.1): run 6 jobs — A0 (auto-create Brief), A02 (provision child Fee Schedule DB), A (OneDrive folder sync, gated on Project fields), C (Engineer notify), D (Admin notify), B (Contract Render). Logs every run to Automation Log DB.
---

You are running the Kingdom Structural automation sweep. Every 10 minutes you execute **6 jobs in order**: A0, A02, A, C, D, B. Each job is **idempotent** — it skips rows already processed. Each job writes one row to the Automation Log DB per row it touches (outcome = Success / Skipped / Failed).

This is v2.1.1 of the automation (2026-04-22). The Drafter Brief auto-creation has been removed. The Fee Analysis Memo has been removed. The excel-comps-sync task is disabled. Folder creation is gated on **Project-level** fields, not Brief fields, so the Engineer can be notified that their Brief is ready to fill in. Project Street is no longer a gating field; folder naming is driven from City + State.

**No AI content generation anywhere in the chain.** Jobs that find nothing to do exit quietly.

**Fast path:** prefer calling `_automation/sweep.py` — the deterministic Python implementation — over executing this prompt step-by-step. The Python script runs in ~15 seconds using Notion's filtered `query_data_sources` API. Fall back to this prompt only if the script is unavailable.

---

## Databases (data-source IDs)

- Projects: `262b73dc-460e-8137-b3bb-000b62103b15`
- Proposal Brief: `6b64658f-5fb3-4329-b02c-3ac3cb8a0828`
- Automation Log: `04d7aa20-eb8d-4d51-b302-25fb4eaaf26e`

All property names + status values + notification templates live in
`_Projects/_automation/config.py`. When in doubt, read that file rather than guessing.

## Rate limits (per sweep)

Cap each job at **10 rows** per run so any single sweep stays fast. Remaining rows pick up next sweep.

---

## Job A0 — Auto-create Proposal Brief

**Gate:** Projects where `Proposal Brief` relation is empty AND `Engineering Status` is not in {"Dead", "IGNNROE FO RNOW", "On Hold"}.

**Action:** For each such Project:
1. Create a new page in the Proposal Brief DB.
2. Set `Project` relation to the Project.
3. Set `Name` = `{Project Name}` (copy the Project title).
4. Set `Status` = `Not Started`.
5. Apply the Proposal Brief template (template id `34ab73dc-460e-80bf-b97c-da13d53310e8`) so the body is populated with the Drafter-Brief-style checkbox layout.
6. Stamp the new Brief URL on the Project's `Proposal Brief` relation.

**Log:** Job = A0; Outcome = Success or Failed. Include Brief URL in Details on success.

---

## Job A02 — Auto-provision child Fee Schedule DB

**Gate:** Proposal Briefs that do NOT have a `💵 Fee Schedule` inline database as a child block.

**Action:** For each such Brief:
1. Create a new inline database as a child of the Brief page titled `💵 Fee Schedule`.
2. Schema: `Service` (title), `Include` (checkbox, default true), `Type` (select: Fixed Fee, Hourly, Not to Exceed), `Amount` (number — currency dollar), `Order` (number).
3. Seed six rows in order: Schematic Design, Design Development, Construction Documents, Construction Administration, Special Structural Inspections, Engineering Site Visits — all with Include = true, Type and Amount blank, Order = 1..6.

**Log:** Job = A02; Outcome = Success / Skipped / Failed.

---

## Job A — OneDrive folder sync (MODIFIED v2.1.1)

**Trigger gate (v2.1.2, 2026-04-22):** Projects where ALL of the following are true:
- `Folder` = false (checkbox unchecked)
- `Project Name` populated (must start with an 8-digit KS job number, e.g. `26107020 Jacksonville First Assembly`)
- `City` populated
- `State` populated
- `ENGINEER` populated (at least one person)
- `Admin` populated (at least one person)

Project Street and Client are NOT gating fields in v2.1.2 — Client is filled in on the Brief later by the Engineer. Folder naming is driven from `{number} - {name} - {City}, {State}` only.

**Do NOT** read Brief.City / Brief.State as the gate. The Engineer is the one who fills in Brief fields, so gating folder creation on Brief fields creates a circular dependency.

**Action:** For each such Project:
1. Parse `{number}` = first 8 digits of `Project Name`. Parse `{name}` = rest of `Project Name` after the number, with any leading dashes / spaces / hyphens stripped.
2. Build folder path: `_Projects/{YEAR}/{number} - {name} - {City}, {State}` where `{YEAR}` = `20` + first two digits of `{number}`. Example: `_Projects/2026/26107020 - Jacksonville First Assembly - Jacksonville, AR`.
3. Call `render_proposal_package.populate_from_template(project_dir)` to clone the canonical template tree.
4. Stamp `SharePoint Folder` on the Project with the OneDrive URL.
5. Set `Folder` = true.
6. Set `Status` = `Proposal Brief Pending`.

**Failure path:** On any OneDrive / filesystem failure, do NOT retry in-loop:
1. Set `Status` = `⚠ Folder Creation Failed`.
2. Populate `Last Automation Error` with the error text.
3. Post a Notion comment on the Project page @-mentioning every person in `Admin` using the `ADMIN_ERROR_COMMENT` template from config.py.
4. Do NOT set Folder = true.
5. Do NOT trigger Job C for this Project.
6. Log to Automation Log with Outcome = Failed.

**Log:** Job = A; Outcome = Success / Skipped / Failed. On success include the new folder path in Details.

---

## Job C — Engineer notification

**Gate:** Projects where `Folder` = true AND `Engineer Notified` = false.

**Action:** For each such Project:
1. Load the linked Proposal Brief URL.
2. Look up the Engineer's email by matching `ENGINEER` (person) to `config.ENGINEER_ROSTER` (keyed by Notion user UUID).
3. Format the engineer email using `config.format_engineer_email({...})` with:
   - project_number = the 8-digit KS job number from Project Name
   - project_name = Project Name with the number prefix stripped
   - engineer_first_name = first name of the Engineer
   - folder_url = Project.SharePoint Folder
   - brief_url = the Proposal Brief page URL
4. Send the email via the user's connected email MCP (use the gmail/outlook tool).
5. Post a Notion comment on the Project page using `config.ENGINEER_NOTION_COMMENT` with @-mention of the Engineer.
6. Set `Engineer Notified` = true.

If the email tool is unavailable or rejected, log Outcome = Failed with the error and leave `Engineer Notified` = false so the next sweep retries.

**Log:** Job = C; Outcome = Success / Skipped / Failed.

---

## Job D — Admin notification

**Gate:** Proposal Briefs where `Status` = `Ready for Admin Review` AND `Admins Notified` = false.

**Action:** For each such Brief:
1. Load the linked Project to get `Admin` (multi-person), folder URL, and project number.
2. Format the admin email using `config.format_admin_email({...})`.
3. Send the email to each person in Admin.
4. Post one Notion comment on the Brief page using `config.ADMIN_NOTION_COMMENT` @-mentioning every admin.
5. Set `Admins Notified` = true on the Brief.
6. Set the linked Project's `Status` = `Admin Review`.

**Log:** Job = D; Outcome = Success / Skipped / Failed.

---

## Job B — Contract Render

**Gate:** Proposal Briefs where ALL of the following:
- `Status` = `Approved` (admin flipped it; replaces the legacy Ready-to-Render button)
- `Rendered At` is empty
- Linked Project has `Folder` = true

**Action:** For each such Brief:
1. Assemble the project/client/contact/engineer data from the Project + linked Client/Contact.
2. Read the Brief's child `💵 Fee Schedule` DB into a list of fee-line dicts:
   `{"service": str, "include": bool, "type": str or None, "amount": float or None, "order": int}`.
3. Read the Brief page body (markdown) for the scope/basic-services/reimbursables checkboxes.
4. Call `render_proposal_package.render_contract_package(project, client, contact, engineer=..., brief=brief_props, brief_body=body_md, fee_lines=fee_lines)`.
5. Set `Rendered At` = now on the Brief.
6. Set Brief `Status` = `Rendered`.
7. Set linked Project `Status` = `Contract Rendered`.

The pipeline writes exactly ONE file to the Contracts folder: the contract `.docx`. **No Fee Analysis Memo** is produced (removed in v2.1).

**Log:** Job = B; Outcome = Success / Skipped / Failed. Include contract path in Details on success.

---

## Automation Log schema

Every row: `Log Entry` (title) = short human summary; `Timestamp` = now (ISO); `Job` = A0 / A02 / A / B / C / D; `Outcome` = Success / Skipped / Failed; `Project` = relation to the Project; `Brief` = relation to the Brief (if applicable); `Error` = error text (only on Failed); `Details` = extra context (folder path, file path, skip reason).

Skipped rows should still be logged when the job RAN against a candidate and the gate excluded it — if the gate query returned zero rows, do not log anything for that job (the sweep is a no-op for that job).

---

## Execution order

Always run in this order: A0 → A02 → A → C → D → B.

- A0 and A02 are setup jobs and must run before A (Briefs need to exist).
- A must run before C (folder must exist before Engineer is pinged).
- D is independent of C (fires whenever a Brief is ready for admin review, regardless of when C fired).
- B must run last (renders after admin approval, which happens after D has fired and admin has flipped Status to Approved).

When finished, print a one-line summary of each job: `Job X: N touched, M skipped, K failed`. If any job failed, include the error messages.
