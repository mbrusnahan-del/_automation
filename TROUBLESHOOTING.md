# KS Automation — Troubleshooting Runbook

When something breaks, start here. Each section is "symptom → root cause → fix"
written for the team running the automation, not for whoever wrote it.

If a problem isn't here, run the preflight check first — it surfaces 90% of
configuration issues without needing to dig into logs:

```powershell
cd "C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\_automation"
& "C:\Users\MichaelBrusnahan\AppData\Local\Programs\Python\Python312\python.exe" preflight.py --deep
```

The deep check walks a real Project → Contact → Contact-page traversal and
fails loud if the integration token can't reach any link in the chain.

---

## Symptom 1 — Contract renders with `<<FILL IN: Contact Name>>` markers

The data is in Notion but the renderer can't see it.

**Most common cause:** the Notion integration token in `.env`
(currently `Excel Power Query`) is not connected to the **Contacts database**.

**How to confirm:** run preflight and look for `[FAIL] Contacts DB reachable`.

**Fix:**
1. Open the Contacts database in Notion (the one with rows like Carlos Dávalos,
   Omar Holguin — NOT the Client Database which has companies).
2. Click `•••` (top-right) → `Connections` → `+ Add connection`.
3. Select the integration named in your `.env` (default: `Excel Power Query`).
4. Re-run preflight to confirm `[OK] Contacts DB reachable`.
5. Re-render the affected contract:
   - Clear `Rendered At` on the Brief in Notion (click the ✕)
   - Run `sweep.py`

If preflight says Contacts is reachable but the contact still renders blank,
the relation cache on a specific Project may be stale. Re-link Project Contact
in Notion: open the Project, click the ✕ next to the contact name in
`Project Contact`, then re-add the same contact. Save.

---

## Symptom 2 — Sweep reports `Job B: 1 touched` but no new contract file appears

**Cause:** the contract file path is locked or unwritable.

**Common reasons (in order of likelihood):**

1. **The previous contract is open in Word.** Word puts an exclusive lock
   on `.docx` files. Close all Word windows. If unsure, open Task Manager
   (Ctrl+Shift+Esc) and end any `WINWORD.EXE` processes.
2. **OneDrive sync was mid-write.** Wait 30 seconds and re-run sweep.
3. **Antivirus is scanning the Contracts folder.** Pause real-time scanning
   briefly, re-run, then re-enable.

**How to confirm:** the sweep output now shows a clear error:
`PermissionError: Could not write contract to '...' after 5 retries.`

If you see that message, follow the steps above.

**Verify a write actually landed before opening:**
```powershell
$f = "C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\2026\<PROJECT FOLDER>\Contracts\<contract filename>.docx"
(Get-Item $f).LastWriteTime
```
Should be within the last minute.

---

## Symptom 3 — `Job B: 0 touched` even though the Brief is approved

**Cause:** Job B's filter requires BOTH `Status = Approved` AND `Rendered At = empty`.

**Fix:** open the Brief in Notion. Clear the `Rendered At` date (click the ✕
inside the property). Save. Re-run sweep.

This is by design — without that filter, Job B would re-render every approved
Brief on every 5-minute sweep cycle.

---

## Symptom 4 — `Job A: 0 touched` for a project that should fire

**Cause:** Job A (and A0, C) require ALL of these on the Project:
- `Engineering Status = Proposal Requested`
- `Intake Complete = ✓` (checkbox ticked)
- `Folder = ✗` (not yet created)
- City, State, ENGINEER, Admin all populated

**Fix:** open the Project page and look at the `Intake Status` formula. It
spells out exactly which field is missing (`⚠ Missing: City`, etc.). Fix the
field, tick `Intake Complete`, save. Job A fires on the next sweep.

---

## Symptom 5 — Contract scope bullets are generic / wrong

The scope library fallback ran instead of the Brief body parser.

**Cause:** the Brief body parser found 0 checked items at render time. Either:
1. The render happened before the engineer filled in the scope checkboxes.
2. The Brief body has unexpected content the parser doesn't recognize.

**Fix:** re-render after the Brief is fully filled:
1. Verify the Brief body has at least one checked scope item (`- [x]`).
2. Clear `Rendered At` on the Brief.
3. Re-run sweep.

If scope STILL renders generic after a full Brief, the body markdown reader
in `sweep.py` (`_fetch_brief_body_markdown`) may be missing a block type the
engineer used. Check the sweep log for parser output.

---

## Symptom 6 — `Fee Schedule` table is missing or shows template placeholders

**Cause:** Fee Schedule data source has zero rows with `Include = ✓` AND a
populated `Amount`.

**Fix:** open the Brief, scroll to the inline `💵 Fee Schedule` database, and
ensure at least one row has both:
- `Include` checkbox ticked
- `Amount` populated with a dollar value
- `Type` set (Fixed Fee / Hourly / Not to Exceed)

Then clear `Rendered At` and re-run sweep.

---

## Symptom 7 — Header on pages 2-4 is missing

**Cause:** the merge-ready template's `header1.xml` doesn't have content
(it was rendered from the older empty-header version).

**Fix:** run the template patch script once:
```powershell
& "C:\Users\MichaelBrusnahan\AppData\Local\Programs\Python\Python312\python.exe" patch_template_header.py
```
This pulls the page-2+ header from `contract template std...docx` and
patches it into `contract_template_merge_ready.docx` with renamed
placeholders. Backup is saved to `contract_template_merge_ready.bak.docx`.

After running, re-render any contract you want the header on.

---

## Symptom 8 — Sweep log shows `fetch_relation_ids(...): unexpected response shape`

**Cause:** Notion's API returned a property-item structure we don't recognize.
Usually means the integration's permissions changed mid-run, OR Notion
deployed an API change.

**Fix:** check Notion's API changelog at
<https://developers.notion.com/page/changelog>. If the response shape
changed, update `fetch_relation_ids()` in `sweep.py` to handle the new
structure. Until then, the function returns `[]` and the affected resolver
falls through to its `<<FILL IN>>` placeholder.

---

## Symptom 9 — Excel sync (`sync_clients_from_excel.py`) creates duplicates

**Cause:** an Excel row's `Number` (col B) doesn't match any Notion Client's
`Number` field. The script can't match by name (intentional — names drift),
so it creates a new Client.

**Fix:** before running the sync live, review the dry-run output:
```powershell
& "C:\Users\MichaelBrusnahan\AppData\Local\Programs\Python\Python312\python.exe" sync_clients_from_excel.py --force
```
Look at the **Pre-flight match report**:
- `MATCH` rows are safe (existing Notion Client, contacts-only sync)
- `CREATE` rows are about to create a new Notion Client. Verify each one
  is genuinely new. If a row should match an existing Client, populate the
  `Number` field on that Client in Notion to match Excel's `Number`.

---

## How to manually re-render one specific project

When you need to rebuild one contract from scratch (e.g., after fixing
data in Notion):

1. In Notion, open the affected Brief page.
2. Set `Status = Approved`.
3. Click the ✕ next to `Rendered At` to clear the date.
4. Save (click outside the field).
5. Make sure no Word windows have the old contract open.
6. Run sweep. The new contract appears in the project's `Contracts` folder.

---

## Two integrations connected — should I clean this up?

You currently have two Notion integrations connected to your databases:
`Excel Power Query` and `KS Automation`. The automation only uses one
(whichever's token is in `.env`, currently Excel Power Query).

**It's safe to leave both** — they're independent. But it's tidier to
consolidate:

1. Pick the integration you want to keep (probably the one in `.env` —
   leaves `.env` alone).
2. Make sure it's connected to all four databases: Projects, Briefs,
   Clients, Contacts.
3. Disconnect the other from those databases.
4. Run preflight to confirm everything still works.

Don't do this on a Friday afternoon.

---

## Where to find the logs

- **Scheduled task runs:** `_automation/sweep.log` (Windows Task Scheduler
  redirects sweep.py output here via `run_sweep.bat`)
- **Manual runs:** stdout of the PowerShell window you ran `sweep.py` from
- **Excel sync:** `_automation/excel_sync.log`
- **Notion-side audit trail:** the `📊 Automation Log` database in your
  workspace — every job logs a row there with outcome, project ref, and
  error detail
