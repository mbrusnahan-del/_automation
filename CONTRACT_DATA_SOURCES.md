# Contract data-source map (v2.1.4, 2026-04-23)

Every field that gets stamped into `contract_template_merge_ready.docx` is
listed here with the **exact Notion source** it reads from. Use this doc as
the source of truth when cleaning up Notion properties — if a field isn't
rendering correctly, walk the row and confirm the upstream source is filled
in.

## What changed in v2.1.4

- **`Intake Complete` checkbox** is now a gating requirement for Jobs A0, A,
  and C. Admin must tick this on the Project before automation acts.
- **`Intake Status` formula** on Project surfaces the first missing onboarding
  field (or "✓ Ready for Proposal Requested" when complete).
- **Contact rollups** on Project pull through `Contact Email`, `Contact Phone`,
  `Contact Mobile`, and `Contact Role` from the linked Contact page. Admins
  can see these at a glance without clicking into the Contact page.
- **Brief body reader** now walks nested blocks and preserves bold markers,
  fixing the Basic Services rendering bug where checkboxes nested inside a
  callout were being missed.
- **Project Type select options** on the Brief are the new 16-item list
  (Residential – Single-Family, Canopy, etc.).

## Legend

- **Project** = a page in the `🏗️ Projects` database
- **Brief** = a page in the `📄 Proposal Brief` database (linked via Project.Proposal Brief)
- **Client** = a page in the `Client Database` (linked via Project.Client)
- **Contact** = a page in the `Contacts` database (linked via Project.Project Contact)
- **Fee Schedule** = child database inside the Brief page
- **Brief body** = the checkbox/callout content on the Brief page (not the properties sidebar)
- **Config** = hardcoded in `_automation/config.py` or `render_from_notion.py`
- **Derived** = computed at render time

---

## 1. Letterhead — top of page 1

### Client block (top-left)

| Contract field | Reads from | Exact property |
|---|---|---|
| Client company (ALL CAPS) | **Client** | `Name` (title) — uppercased |
| Client company (normal case) | **Client** | `Name` (title) |
| Client street | **Client** | `Address` (text) — substring before first comma |
| Client city/state/ZIP | **Client** | `Address` (text) — substring after first comma |
| Client email | **Contact** | `Email` (email property) |
| Client phone | **Contact** | `Phone` (phone), falls back to `Phone (Mobile)` |

### Partner / Engineer block (top-right)

| Contract field | Reads from |
|---|---|
| Partner name + credentials | **Config** → `render_from_notion.PARTNER_BLOCKS` keyed by `Project.ENGINEER[0]` UUID |
| Partner position | **Config** → PARTNER_BLOCKS defaults (JB/CV = Partner, GO = Project Manager) |
| Partner phone | **Config** → `PARTNER_BLOCKS.phone` formatted as `XXX.XXX.XXXX` |
| Partner email | **Config** → `PARTNER_BLOCKS.email` |
| Fallback when Project.ENGINEER is empty | **Client** → `The Guy` rollup → PARTNER_BLOCKS |

### Project block

| Contract field | Reads from |
|---|---|
| KS job number | **Derived** — first 8 digits of `Project.Project Name` (strict) |
| Project name (short) | **Derived** — `Project.Project Name` minus the 8-digit prefix |
| Project name (ALL CAPS) | **Derived** — same, uppercased |
| Project street | **Project** `Project Street` (text) |
| Project city/state | **Project** `{City}, {State}` |
| Fallback when City/State missing | **Derived** — parsed off `Project.SharePoint Folder` URL |

### Date

| Field | Source |
|---|---|
| Contract date | **Derived** — `datetime.today()` at render time |
| Revised block | **Project** `revised_date` (if set; blank on initial renders) |

---

## 2. Salutation

| Contract field | Source |
|---|---|
| "Dear ___," | **Contact** `Contact Name` (title) |

---

## 3. SCOPE OF SERVICES bullets

| Contract field | Source |
|---|---|
| Scope bullets | **Brief body** — every `- [x]` checkbox under `> 📋 SCOPE OF SERVICES` callout. **v2.1.4: reader recurses into nested children so checkboxes stored as children of the callout are now captured.** Only checked items render; unchecked items are ignored. |
| Category headers (CUSTOM RESIDENCE, PEMB, etc.) | **Brief body** — `**CATEGORY**` bold paragraphs. Used for internal classification only; not emitted as bullets. |
| `{{icc_year}}` placeholder in bullet text | **Brief** `ICC Code Year` (select: 2015/2018/2021/2024) |
| `{{jurisdiction}}` placeholder in bullet text | **Brief** `Jurisdiction` (text) |
| Structural data bullet (optional prepend) | **Project** — Gravity System + Lateral System + Foundation Type if all set |
| Fallback when Brief has zero checkboxes | **Config** — `scope_library.py` keyed by Brief `Project Type` |

---

## 4. BASIC SERVICES paragraph

| Contract field | Source |
|---|---|
| "Basic services shall include …" | **Brief body** — `- [x]` checkboxes under `> ✅ BASIC SERVICES` callout, Oxford-comma joined |
| Exclusion tail (RFI / value engineering / etc.) | **Config** — hardcoded in `build_basic_services_paragraph()` |

---

## 5. REIMBURSABLES paragraph

Three variants live in the template; automation keeps one and deletes the others.

| Variant kept when | Trigger |
|---|---|
| "in addition to basic services…" (Standard) | Brief body `- [x] Standard — …` |
| "not applicable. All work shall be carried out digitally" (Digital Only) | Brief body `- [x] Digital Only — …` |
| "not anticipated. Travel to site has been included…" (Included in Fee) | Brief body `- [x] Included in Fee — …` |

---

## 6. SPECIAL STRUCTURAL INSPECTIONS paragraph

| Field | Source |
|---|---|
| SSI paragraph included by default | (implicit) |
| SSI paragraph suppressed | **Brief body** — `- [x] Special structural inspections shall be excluded` under `> 📋 SCOPE OF SERVICES` |

---

## 7. FEE SCHEDULE table

| Contract field | Source |
|---|---|
| Row per service | **Fee Schedule** child DB on the Brief, only `Include = true` rows |
| Service name | Fee Schedule `Service` (title) |
| Fee type | Fee Schedule `Type` (select: Fixed Fee / Hourly / Not to Exceed) |
| Dollar amount | Fee Schedule `Amount` (number, dollar format) |
| Row order | Fee Schedule `Order` (number, ascending) |
| Fee total in words | **Derived** — sum of included rows → `dollars_to_words()` |

---

## 8. Signature block (bottom of page 2)

| Contract field | Source |
|---|---|
| "send signed copy to {email}" | **Config** → `PARTNER_BLOCKS[code].email` |
| Partner name + creds | **Config** → `PARTNER_BLOCKS[code].name_with_creds` |
| Partner position | **Config** → `resolve_engineer()` position logic |
| Partner phone | **Config** → `PARTNER_BLOCKS[code].phone` |
| Partner email | **Config** → `PARTNER_BLOCKS[code].email` |

---

## 9. Highlighting rule (v2.1.2)

After every placeholder is filled in:

1. Run contains a real value (no `<<FILL IN:>>` marker) → highlight and
   shading are stripped. Text appears **clean** on the rendered contract.
2. Run still contains `<<FILL IN: …>>` text → highlight is preserved.
   Gaps stand out visually so the engineer can spot and fix them.

**Highlighted text on the rendered contract = a Notion data gap you need
to fill in.** Unhighlighted = data was pulled successfully.

---

## 10. Automation gates (v2.1.4)

### Trigger status (v2.1.3)

Jobs A0, A, and C only fire when Project `Engineering Status` equals one of:

- `Proposal Requested`

(Configured in `sweep.py` → `PROPOSAL_TRIGGER_STATUSES`. Add more statuses
here to widen the trigger.)

### Intake Complete gate (v2.1.4)

Jobs A0, A, and C additionally require `Intake Complete = true` on the
Project. Admin ticks this once onboarding fields are all populated (the
Intake Status formula confirms when it's safe to tick).

### Visible onboarding status (v2.1.4)

The `Intake Status` formula property on Projects displays the first
missing required field, or `✓ Ready for Proposal Requested` when all of
these are populated:

1. Project Name (with 8-digit prefix)
2. City
3. State
4. ENGINEER (at least one person)
5. Admin (at least one person)
6. Client (relation)
7. Project Contact (relation)
8. Contact Email (rollup from Project Contact → Email)
9. Contact Phone OR Contact Mobile (rollup from Project Contact)

When the formula shows `⚠`, the field name in the message tells the
admin exactly what to fix. When it shows `✓`, they can safely tick
`Intake Complete` and set Engineering Status = `Proposal Requested`.

---

## 11. Onboarding data-hygiene checklist

Before setting `Intake Complete = true` on a Project, ensure:

### On the Project page
- [ ] `Project Name` starts with an 8-digit KS job number
- [ ] `City` and `State` populated
- [ ] `Project Street` (optional but recommended for letterhead)
- [ ] `Client` relation → a Client page (see below)
- [ ] `Project Contact` relation → a Contact page (see below)
- [ ] `ENGINEER` at least one person
- [ ] `Admin` at least one person
- [ ] `Intake Status` formula shows `✓ Ready for Proposal Requested`

### On the Client page
- [ ] `Name` — real company name (not a placeholder)
- [ ] `Address` with a comma between street and city (e.g. `1241 E Sheffield Ct, Gilbert, AZ 85296`)

### On the Contact page
- [ ] `Contact Name` (title)
- [ ] `Email` (email property type, not text)
- [ ] `Phone` or `Phone (Mobile)` (phone property type)

### On the Brief page (after engineer takes over)
- [ ] `ICC Code Year` (2015/2018/2021/2024)
- [ ] `Jurisdiction` (text)
- [ ] `Project Type` (one of the 16 new options)
- [ ] Brief body: scope-of-services checkboxes
- [ ] Brief body: basic-services checkboxes
- [ ] Brief body: exactly one Reimbursables option
- [ ] Fee Schedule: `Type` and `Amount` on each included row

---

## 12. Brief Project Type — new options (v2.1.4)

The 16 categories on the Brief's Project Type dropdown:

| Category | Typical scope type |
|---|---|
| Residential – Remodel / Addition | Existing home alterations / additions |
| Residential – Single-Family | New construction single home |
| Residential – Multi‑Family | Apartments / condos / duplexes |
| Religious | Churches, temples, places of worship |
| Restaurant | Dining TI or new build |
| Storage | Self-storage, warehouse storage |
| Medical | Clinics, dental, surgery centers |
| Office | Office TI or new office builds |
| Education | Schools, daycares, classroom facilities |
| Commercial – TI | Retail / office tenant improvements |
| Industrial | Manufacturing, heavy equipment |
| Canopy | Gas station canopies, drive-thru covers |
| Specialty Structures | Anything non-standard (towers, art, etc.) |
| Existing Building / Assessment | Structural observation, peer review |
| PEMB | Pre-engineered metal buildings |
| Other | Escape hatch for edge cases |

---

**Version history**
- v2.1 (2026-04-22) — Fee Analysis Memo removed, gating moved to Project level
- v2.1.1 (2026-04-22) — Project Street dropped from Job A gate; strict 8-digit lint
- v2.1.2 (2026-04-23) — Client/Contact/Engineer resolved from Project relations; highlight-on-gaps-only
- v2.1.3 (2026-04-23) — Jobs A0/A/C tightened to `Engineering Status = Proposal Requested` only
- v2.1.4 (2026-04-23) — Intake Complete gate; Intake Status formula; Contact rollups; nested-block reader; new Project Type options
