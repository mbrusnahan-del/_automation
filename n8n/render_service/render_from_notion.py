"""
Render a KS Structural contract from a Notion project record.
Pulls: project → client → contact(s), maps to merge fields, fills template.

This is the core function the scheduled task will call. For now it takes
pre-fetched Notion data dicts (since the MCP tool is the integration point);
the scheduled task prompt will handle the Notion fetches and pass data in.
"""
from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy
import shutil
import os
import sys
import re
from datetime import date

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)
TEMPLATE = os.path.join(_SCRIPT_DIR, "contract_template_merge_ready.docx")

from scope_library import build_scope_bullets


# ----------------------------------------------------------------------
# Structural data + scope bullets (Notion Scope Templates DB driven)
# ----------------------------------------------------------------------
# As of 2026-04-21 the canonical source of scope bullets is the Notion
# `Scope Templates` database, related to Projects via the `Scope Items`
# multi-relation. The scheduled task resolves that relation into a list of
# bullet strings and passes them in as project["Scope Bullets"]. When that
# list is empty (older projects, or engineer hasn't filled it in yet), we
# fall back to scope_library.build_scope_bullets() so renders still produce
# a usable contract.
#
# Three new structural-data select fields on Projects feed an additional
# summary bullet that gets prepended to the scope list:
#   - Gravity System  (e.g. "Wood Frame", "Steel Frame", "Concrete")
#   - Lateral System  (e.g. "Shear Walls", "Braced Frames")
#   - Foundation Type (e.g. "Spread Footings", "Mat", "Piles")


def _structural_phrase(g, l, f):
    """
    Build a compact engineer-friendly summary of the three structural systems.
    Skips parts that are blank so we never print 'None' or '' in the contract.
    Returns '' if all three fields are empty.

    Format: "Gravity: Wood Frame  |  Lateral: Shear Walls  |  Foundation: Spread Footings"
    Sized so it reads as a single bullet of structural info, not a sentence.
    """
    g = (g or "").strip()
    l = (l or "").strip()
    f = (f or "").strip()
    if not (g or l or f):
        return ""

    parts = []
    if g:
        parts.append(f"Gravity: {g}")
    if l:
        parts.append(f"Lateral: {l}")
    if f:
        parts.append(f"Foundation: {f}")
    return "  |  ".join(parts)


def build_structural_data_bullet(project):
    """
    Return a single bullet string summarizing the project's structural systems,
    or None if no structural fields are populated.
    Format: 'Structural System — Gravity: Wood Frame  |  Lateral: Shear Walls  |  Foundation: Spread Footings'
    """
    phrase = _structural_phrase(
        project.get("Gravity System"),
        project.get("Lateral System"),
        project.get("Foundation Type"),
    )
    if not phrase:
        return None
    return f"Structural System — {phrase}"


BRIEF_FIELDS = (
    # Proposal-specific fields that moved to the Brief in 2026-04-21 rework
    "ICC Code Year",
    "Jurisdiction",
    "Reimbursables Treatment",
    "Approx. Structural SF",
    "Scope Description",
    # Structural data fields (no longer rendered, kept for Brief overlay)
    "Gravity System",
    "Lateral System",
    "Foundation Type",
    # Location + type fields moved off Project in 2026-04-22 cleanup.
    # Brief is now the canonical source; Project no longer stores these.
    "City",
    "State",
    "Project Street",
    "Project Type",
)


def merge_brief_into_project(project, brief):
    """
    Brief-first, Project-fallback: for each proposal-only field that lives on
    the Proposal Brief (see BRIEF_FIELDS + Scope Items), overlay the Brief's
    value on the Project dict. The Brief wins when it has a non-empty value;
    the Project property is used as fallback for legacy projects without a
    Brief yet.

    `brief` may be None (no Proposal Brief linked yet). In that case the
    project dict is returned unchanged. Never mutates either input dict.

    Also normalizes the scope bullets field:
      - If brief has "Scope Bullets" list (resolved from Scope Items relation
        by the scheduled task) it takes precedence over project's.
      - The renderer downstream uses project_data["Scope Bullets"].
    """
    merged = dict(project)
    if not brief:
        return merged

    for key in BRIEF_FIELDS:
        bv = brief.get(key)
        # Treat None, "", and 0 (only for SF) carefully: only overlay when the
        # brief value is truthy OR an explicitly-set number (including 0).
        if bv is None:
            continue
        if isinstance(bv, str) and not bv.strip():
            continue
        merged[key] = bv

    # Scope bullets: if the brief resolved any bullets (from its Scope Items
    # relation), prefer them over whatever the project had.
    brief_bullets = brief.get("Scope Bullets")
    if brief_bullets:
        merged["Scope Bullets"] = brief_bullets

    return merged


def resolve_scope_bullets(project):
    """
    Build the final SCOPE OF SERVICES bullet list for a project.

    Priority for bullet body:
      1. Engineer-selected scope items from Notion (project["Scope Bullets"]).
         The scheduled task resolves the `Scope Items` relation into a list
         of strings before calling the renderer. The list may come from the
         Proposal Brief (preferred) or the legacy Project property.
      2. Fallback to the hardcoded scope_library by Project Type.

    Then prepend the Structural System bullet (if structural fields are set)
    so the contract opens with a one-line description of what's being built.
    """
    notion_bullets = project.get("Scope Bullets") or []
    # Trim and dedupe while preserving order
    seen = set()
    cleaned = []
    for b in notion_bullets:
        if not b:
            continue
        s = b.strip()
        if not s:
            continue
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(s)

    if cleaned:
        bullets = cleaned
    else:
        bullets = build_scope_bullets(project)

    structural = build_structural_data_bullet(project)
    if structural:
        bullets = [structural] + bullets
    return bullets


def split_address(addr):
    """
    Split an address string like
      '4310 N 75th St Suite 145, Scottsdale, AZ 85251'
    into (line1, line2) = ('4310 N 75th St Suite 145', 'Scottsdale, AZ 85251').

    Heuristic: split on the first comma. Works for typical US addresses.
    Returns (addr, '') if no comma present.
    """
    if not addr:
        return "", ""
    parts = addr.split(",", 1)
    if len(parts) == 1:
        return parts[0].strip(), ""
    return parts[0].strip(), parts[1].strip()


def parse_project_number(project_name, city=None, state=None):
    """
    Notion project names vary in format:
      '26102144 - General Metal Construction PEMB'  → canonical
      '26103061 Quarry Pines Driving Range Tucson, AZ'  → no dash, location suffix
      '26103061 Quarry Pines Driving Range'          → no dash, no suffix
    Returns (job_number, project_name_short) where project_name_short
    has any leading number and trailing " City, State" location stripped.
    """
    s = project_name.strip()
    m = re.match(r'^(\d{8})[\s\-]+(.+?)\s*$', s)
    if m:
        job, rest = m.group(1), m.group(2).strip()
    else:
        return "", s

    # Strip trailing " City, State" or " City,State" if it matches known location
    if city and state:
        for pat in (fr'\s+{re.escape(city)},\s*{re.escape(state)}\s*$',
                    fr'\s+{re.escape(city)}\s*,\s*{re.escape(state)}\s*$'):
            rest = re.sub(pat, '', rest, flags=re.IGNORECASE).strip()
    elif city:
        rest = re.sub(fr'\s+{re.escape(city)}\s*,?\s*$', '', rest, flags=re.IGNORECASE).strip()
    # Strip trailing punctuation that might remain
    rest = rest.rstrip(' ,-')
    return job, rest


def derive_location_from_sharepoint_url(url):
    """
    SharePoint folder URLs like
      file:///C:/.../_Projects/2026/26102144 General Metal Construction PEMB - Phoenix, AZ
    encode the project location in the folder name after the last ' - '.
    Return that suffix (e.g. 'Phoenix, AZ') or empty string.
    """
    if not url:
        return ""
    # Take the last path segment
    from urllib.parse import unquote
    segment = unquote(url.rstrip('/').split('/')[-1])
    if ' - ' in segment:
        return segment.rsplit(' - ', 1)[-1].strip()
    return ""


# ----------------------------------------------------------------------
# Engineer (partner) resolution
# ----------------------------------------------------------------------
# Source of truth is the Notion People database
# (collection://72ab13bc-f91f-47c4-9759-e392e0f79e1e). The scheduled task
# fetches it and passes an `engineer` dict into build_merge_data().
# Shape: {"code": "GO", "name": "Gabriel O'Reilly",
#         "email": "GOreilly@kingdomstructural.com", "phone": "5148054150"}
#
# Credentials (", SE, PE" / ", PE") aren't stored on the People record yet,
# so we keep a tiny code→creds map here; partners update as needed.
_CREDS_BY_CODE = {
    "JB": ", SE, PE",
    "CV": ", PE",
    "GO": ", PE",
}

# Hard-coded fallback used ONLY if the scheduled task couldn't load People
# (network/permissions issue). Kept in sync with People DB snapshot 2026-04-21.
PARTNER_BLOCKS = {
    "JB": {"code": "JB", "name": "Jonathan Brusnahan",
           "phone": "4802745493", "email": "Jbrusnahan@kingdomstructural.com"},
    "CV": {"code": "CV", "name": "Chris Valdez",
           "phone": "6233266119", "email": "CValdez@kingdomstructural.com"},
    "GO": {"code": "GO", "name": "Gabriel O'Reilly",
           "phone": "5148054150", "email": "GOreilly@kingdomstructural.com"},
}


def _format_phone(raw):
    """Format a 10-digit US phone as 'XXX.XXX.XXXX'. Pass through other shapes."""
    if not raw:
        return ""
    digits = re.sub(r'\D', '', raw)
    if len(digits) == 10:
        return f"{digits[0:3]}.{digits[3:6]}.{digits[6:10]}"
    if len(digits) == 11 and digits.startswith("1"):
        return f"{digits[1:4]}.{digits[4:7]}.{digits[7:11]}"
    return raw


def resolve_engineer(project, client, engineer=None):
    """
    Pick the engineer (partner) record for this project. Priority:
      1. Explicit `engineer` dict passed in by the scheduled task (freshest —
         resolved from Project.ENGINEER → People DB lookup). Expected to include
         at minimum: name, email, phone, code. May also include 'roles' (list
         from People DB Roles multi-select) or 'position' directly.
      2. Client.The Guy code → PARTNER_BLOCKS fallback snapshot (position = "Partner").
      3. None → caller fills <<FILL IN>> placeholders.
    Returns: {code, name, email, phone, name_with_creds, position}.

    Position is derived from People.Roles (first entry). JB/CV fall back to
    "Partner" via PARTNER_BLOCKS; GO is "Project Manager" per Roles.
    """
    if engineer and (engineer.get("name") or engineer.get("email")):
        rec = dict(engineer)
    else:
        code = (client.get("The Guy") or "").strip().upper()
        rec = dict(PARTNER_BLOCKS.get(code) or {})

    if not rec:
        return {
            "code": "",
            "name": "<<FILL IN: Partner name>>",
            "email": "<<FILL IN: Partner email>>",
            "phone": "<<FILL IN: Partner phone>>",
            "name_with_creds": "<<FILL IN: Partner name>>",
            "position": "<<FILL IN: Partner position>>",
        }

    code = (rec.get("code") or "").upper()
    creds = _CREDS_BY_CODE.get(code, "")
    name = rec.get("name") or ""

    # Position: prefer explicit 'position' key, then first of 'roles' list,
    # then fall back to PARTNER_BLOCKS defaults (all three partners = Partner).
    position = rec.get("position") or ""
    if not position:
        roles = rec.get("roles") or []
        if isinstance(roles, list) and roles:
            position = roles[0]
    if not position:
        # Fallback by code
        position = "Partner" if code in ("JB", "CV") else ""

    return {
        "code": code,
        "name": name,
        "email": rec.get("email") or "",
        "phone": _format_phone(rec.get("phone") or ""),
        "name_with_creds": f"{name}{creds}" if name else "",
        "position": position,
    }


def salutation_from_contact(contact_name):
    """
    Build salutation from a contact name. Uses 'Mr. {LastName}' when the
    contact name looks like a personal name (two+ words, not a company).
    Falls back to the full contact name, or 'Sir or Madam' when empty.
    Defaults to 'Mr.' since title/gender isn't captured in the Contacts DB;
    partner can adjust to Ms. manually if needed.
    """
    if not contact_name or not contact_name.strip():
        return "Sir or Madam"
    name = contact_name.strip()
    # Skip obvious company names (contain LLC, Inc, Corp, Studio, Architects, etc.)
    company_markers = ("LLC", "Inc", "Corp", "Studio", "Architects", "Company",
                       "Associates", "Partners", "Group")
    if any(m in name for m in company_markers):
        return name
    parts = name.split()
    if len(parts) >= 2:
        return f"Mr. {parts[-1]}"  # partner can swap to Ms. if needed
    return name


def build_short_scope_summary(project):
    """
    Build the 1-sentence scope summary that goes in the contract opening paragraph.
    Format: "a <type> project, <name>, located at <city>, <state>, and general
            services to be provided including engineering, drafting, and
            construction administration for the structural scope of work"
    This is INTENTIONALLY short — the detailed Scope Description goes elsewhere
    (SCOPE OF SERVICES section, via the scope generator).
    """
    ptype = (project.get("Project Type") or "").strip().lower()
    city = (project.get("City") or "").strip()
    state = (project.get("State") or "").strip()

    # Short name without number/location suffix
    _, short = parse_project_number(project.get("Project Name", ""), city=city, state=state)

    # Type phrasing
    type_phrase_map = {
        "pemb": "a pre-engineered metal building (PEMB)",
        "tenant improvement": "a tenant improvement",
        "commercial new": "a new commercial",
        "commercial remodel": "a commercial remodel",
        "new residence": "a new residence",
        "residential remodel": "a residential remodel",
        "multifamily": "a multifamily",
        "structural observation": "a structural observation",
        "feasibility": "a structural feasibility",
        "peer review": "a structural peer review",
        "retrofit": "a structural retrofit",
    }
    type_phrase = type_phrase_map.get(ptype, "a")
    loc_phrase = f" in {city}, {state}" if city and state else (f" in {city}" if city else "")
    return (
        f"{type_phrase} project, {short}{loc_phrase}, "
        f"and general services to be provided including engineering, drafting, "
        f"and construction administration for the structural scope of work"
    )


def build_merge_data(project, client, contact, engineer=None, today=None, brief=None):
    """
    Given dict forms of {project, client, contact}, produce the merge-field
    dict for the contract template. Missing fields return '<<FILL IN: ...>>'
    placeholders the engineer can fix before sending.

    `engineer` is an optional dict pulled from the Notion People database
    (keyed by the Project.ENGINEER Notion user). When omitted, falls back
    to Client.The Guy → PARTNER_BLOCKS snapshot.

    `brief` is an optional dict pulled from the linked Proposal Brief record
    (via Project.Proposal Brief relation). When provided, its proposal-only
    fields (ICC Code Year, Jurisdiction, Reimbursables Treatment, Approx.
    Structural SF, Gravity/Lateral/Foundation System, Scope Description,
    Scope Bullets) overlay the Project dict. Legacy projects without a Brief
    render from their original Project properties. See merge_brief_into_project().
    """
    # Brief-first, Project-fallback: overlay Brief values onto Project before
    # anything else touches the dict.
    project = merge_brief_into_project(project, brief)
    today = today or date.today()
    date_str = today.strftime("%B %d, %Y")

    # --- Project info
    proj_city = project.get("City") or ""
    proj_state = project.get("State") or ""
    project_number, project_short_name = parse_project_number(
        project.get("Project Name", ""), city=proj_city, state=proj_state
    )
    project_location = derive_location_from_sharepoint_url(project.get("SharePoint Folder", ""))
    project_line1 = project.get("Project Street") or ""
    project_line2 = ""
    if proj_city and proj_state:
        project_line2 = f"{proj_city}, {proj_state}"
    elif proj_city or proj_state:
        project_line2 = (proj_city or proj_state).strip()
    # Fallback: derive from SharePoint folder if city/state not set
    if not project_line2:
        project_line2 = project_location

    # --- Client info
    client_address_line1, client_address_line2 = split_address(client.get("Address", ""))
    client_company = client.get("Name", "") or "<<FILL IN: Client Name>>"

    # --- Contact info
    contact_name = contact.get("Contact Name", "") if contact else ""
    email = (contact or {}).get("Email", "")
    phone = (contact or {}).get("Phone", "") or (contact or {}).get("Phone (Mobile)", "")

    def fill(val, label):
        return val if val else f"<<FILL IN: {label}>>"

    # --- Engineer (partner) info — from Notion People DB via Project.ENGINEER,
    # falling back to Client.The Guy → PARTNER_BLOCKS snapshot.
    partner = resolve_engineer(project, client, engineer=engineer)
    partner_name_with_creds = partner["name_with_creds"]

    # --- Revised date is conditional — blank on initial renders, set on re-renders
    # Caller can set project.get("revised_date") explicitly to show the revised line
    revised_date_val = project.get("revised_date") or ""
    revised_block = f"Revised {revised_date_val}" if revised_date_val else ""

    return {
        "date": date_str,
        "revised_block": revised_block,

        # Client block — both normal-case (opening paragraph) and upper-case
        # (letterhead + signature) variants for the v2 template
        "client_company_name":  client_company,
        "client_company_upper": client_company.upper() if client_company else "",
        "contact_name": fill(contact_name, "Contact Name — no Contact linked on Client"),
        "client_address_line1": fill(client_address_line1, "Client street"),
        "client_address_line2": fill(client_address_line2, "Client city/state/ZIP"),
        "client_email": fill(email, "Contact Email"),
        "client_phone": fill(phone, "Contact Phone"),

        # Project block — normal-case (opening paragraph) + upper-case (letterhead)
        "project_name":       project_short_name,
        "project_name_upper": project_short_name.upper() if project_short_name else "",
        "project_address_line1": fill(project_line1, "Project street (not in Notion)"),
        "project_address_line2": fill(project_line2, "Project city/state"),
        "ks_job_number":         fill(project_number, "KS Job Number"),

        # Code + jurisdiction fill into the submittals scope bullet
        "icc_year":     fill(project.get("ICC Code Year", ""), "ICC Code Year not set on Brief"),
        "jurisdiction": fill(project.get("Jurisdiction", ""),  "Jurisdiction not set on Brief"),

        # Partner (letterhead + signature block)
        "partner_name_with_creds": partner_name_with_creds,
        "partner_position":        partner["position"] or "<<FILL IN: Partner position>>",
        "partner_phone":           partner["phone"],
        "partner_email":           partner["email"],

        # Legacy merge fields — kept for backward compat with older templates
        # still in the wild. The v2 template doesn't reference them.
        "salutation":       "",
        "scope_summary":    "",
        "gravity_system":   project.get("Gravity System") or "",
        "lateral_system":   project.get("Lateral System") or "",
        "foundation_type":  project.get("Foundation Type") or "",
        "fee_schematic":    fill(project.get("fee_schematic", ""), "see Fee Memo"),
        "fee_dev":          fill(project.get("fee_dev", ""),       "see Fee Memo"),
        "fee_cd":        fill(project.get("fee_cd", ""),        "see Fee Memo"),
        "fee_ca":        fill(project.get("fee_ca", ""),        "see Fee Memo"),
        "fee_ssi":       fill(project.get("fee_ssi", ""),       "see Fee Memo"),
        "fee_visits":    fill(project.get("fee_visits", ""),    "see Fee Memo"),
        "fee_total":     fill(project.get("fee_total", ""),     "see Fee Memo"),
    }


def fill_placeholders(doc, data):
    pattern = re.compile(r'\{\{([a-z_0-9]+)\}\}')

    def substitute(text):
        return pattern.sub(
            lambda m: str(data.get(m.group(1), f"<<FILL IN: {m.group(1)}>>")),
            text
        )

    def process_paragraph(paragraph):
        for run in paragraph.runs:
            if '{{' in run.text:
                run.text = substitute(run.text)
        # python-docx's paragraph.runs doesn't include runs inside <w:hyperlink>
        # elements. Walk the raw XML to catch those.
        from docx.oxml.ns import qn
        for hyperlink in paragraph._element.findall(qn('w:hyperlink')):
            for r_elem in hyperlink.findall(qn('w:r')):
                for t_elem in r_elem.findall(qn('w:t')):
                    if t_elem.text and '{{' in t_elem.text:
                        t_elem.text = substitute(t_elem.text)

    for para in doc.paragraphs:
        process_paragraph(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process_paragraph(para)
    # Headers and footers live in doc.sections[*].header / .footer
    for section in doc.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            if header is None:
                continue
            try:
                for para in header.paragraphs:
                    process_paragraph(para)
            except Exception:
                pass
        for footer in (section.footer, section.first_page_footer, section.even_page_footer):
            if footer is None:
                continue
            try:
                for para in footer.paragraphs:
                    process_paragraph(para)
            except Exception:
                pass


def cleanup_whitespace_artifacts(doc):
    """
    After template edits (like removing instructional text runs), stray
    whitespace can end up adjacent to punctuation: "work ." / "intervals ." /
    "email ..". Clean these by walking every paragraph's joined text.
    """
    import re as _re
    patterns = [
        (_re.compile(r'\s+\.\.'), '.'),   # " .." -> "."
        (_re.compile(r'\s+\.'),   '.'),   # " ."  -> "."
        (_re.compile(r'\s+,'),    ','),   # " ,"  -> ","
        (_re.compile(r'  +'),     ' '),   # collapse multiple spaces
    ]

    def clean_paragraph(p):
        # Work at the run level to preserve formatting. Walk runs and trim
        # trailing whitespace when the next run starts with punctuation.
        runs = p.runs
        for i in range(len(runs) - 1):
            cur = runs[i].text
            nxt = runs[i + 1].text
            if cur and nxt and cur.endswith((' ', '\t')) and nxt[:1] in '.,;:':
                runs[i].text = cur.rstrip()
        # Then run a text-level regex on each run for patterns contained within one run
        for r in runs:
            if r.text:
                for pat, repl in patterns:
                    r.text = pat.sub(repl, r.text)

    for p in doc.paragraphs:
        clean_paragraph(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    clean_paragraph(p)


def apply_reimbursables_treatment(doc, treatment):
    """
    The template includes three 'REIMBURSABLE EXPENSES...' paragraphs covering
    Standard, Digital Only, and Included in Fee.  Based on the project's
    Reimbursables Treatment, keep the matching one and delete the other two.
    """
    treatment = (treatment or "Standard").strip().lower()

    # Paragraph-starting text to identifier mapping
    paragraph_signatures = {
        "standard":        "REIMBURSABLE EXPENSES are in addition",
        "digital only":    "REIMBURSABLE EXPENSES are not applicable",
        "included in fee": "REIMBURSABLE EXPENSES are not anticipated",
    }
    keep_sig = paragraph_signatures.get(treatment, paragraph_signatures["standard"])

    to_remove = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t.startswith("REIMBURSABLE EXPENSES") and not t.startswith(keep_sig):
            # It's one of the two we don't want
            to_remove.append(p)
    for p in to_remove:
        p._element.getparent().remove(p._element)


def expand_scope_bullets(doc, bullets):
    """
    Find the paragraph containing {{SCOPE_BULLET_ANCHOR}} and expand it into
    len(bullets) paragraphs, one per bullet, preserving the anchor's list-item
    styling.  The original anchor paragraph is removed.
    """
    anchor_para = None
    for p in doc.paragraphs:
        if 'SCOPE_BULLET_ANCHOR' in p.text:
            anchor_para = p
            break
    if anchor_para is None:
        return

    anchor_elem = anchor_para._element
    prev = anchor_elem

    for bullet_text in bullets:
        new_elem = deepcopy(anchor_elem)
        # Find the first <w:t> in the cloned paragraph and set its text
        # (the anchor has a single <w:t>{{SCOPE_BULLET_ANCHOR}}</w:t>)
        for t in new_elem.iter(qn('w:t')):
            t.text = bullet_text
            break
        prev.addnext(new_elem)
        prev = new_elem

    # Remove the original anchor paragraph
    anchor_elem.getparent().remove(anchor_elem)


def render_contract(project, client, contact, out_dir):
    data = build_merge_data(project, client, contact)
    project_number = data["ks_job_number"]
    project_name_short = data["project_name"][:80]
    client_short = data["client_company_name"][:40]
    filename = f"{project_number} KS Structural Contract - {client_short} {project_name_short} 26.04.20.docx"
    # sanitize filename
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    out = os.path.join(out_dir, filename)

    os.makedirs(out_dir, exist_ok=True)
    shutil.copyfile(TEMPLATE, out)
    doc = Document(out)
    fill_placeholders(doc, data)
    doc.save(out)
    return out, data


# ===================================================================
# V2 template helpers (2026-04-22)
#   - Brief body parser (checkbox lists)
#   - Basic Services paragraph composer
#   - Reimbursables + SSI conditional paragraphs
#   - Fee Schedule table filler with number→words
# ===================================================================

def parse_brief_body(body):
    """
    Parse a Brief page body (markdown) and return:
      {
        "basic_services":    [...],             # checked item texts
        "scope_of_services": [{category, text}, ...],
        "reimbursables":     "Standard"|"Digital Only"|"Included in Fee"|None,
      }

    Sections are identified by callout headers ('> ✅ BASIC SERVICES',
    '> 📋 SCOPE OF SERVICES', '> 💵 REIMBURSABLES'). Under SCOPE OF SERVICES,
    bold text on its own line (`**CATEGORY**`) sets the current category for
    subsequent checked items.
    """
    result = {
        "basic_services": [],
        "scope_of_services": [],
        "reimbursables": None,
    }
    section = None
    category = None
    for raw_line in (body or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith(">"):
            up = line.upper()
            if "BASIC SERVICES" in up:
                section = "basic"; category = None
            elif "SCOPE OF SERVICES" in up:
                section = "scope"; category = None
            elif "REIMBURSABLES" in up:
                section = "reimb"; category = None
            elif "NOTES TO ENGINEER" in up or "FEE NOTES" in up or "FEE SCHEDULE" in up:
                section = None; category = None
            continue
        if section == "scope":
            m = re.match(r'^\*\*([A-Z ]{2,40})\*\*\s*$', line)
            if m:
                category = m.group(1).strip()
                continue
        m = re.match(r'^- \[(x| )\]\s+(.+?)\s*$', line, re.IGNORECASE)
        if not m:
            continue
        if m.group(1).lower() != "x":
            continue
        text = m.group(2).strip()
        if section == "basic":
            result["basic_services"].append(text)
        elif section == "scope":
            result["scope_of_services"].append({
                "category": category or "Uncategorized",
                "text": text,
            })
        elif section == "reimb":
            # Label before the em-dash
            label = re.split(r'\s*[—-]\s*', text, maxsplit=1)[0].strip()
            result["reimbursables"] = label
    return result


def build_basic_services_paragraph(parsed):
    """
    Compose the 'Basic services shall include …' paragraph from the list of
    checked BASIC SERVICES items. Oxford-comma joined, with the standard
    'not included in basic services' exclusions tail.
    """
    items = parsed.get("basic_services") or []
    if not items:
        return ("<<FILL IN: no Basic Services items were checked on the Brief; "
                "check at least one before rendering>>")
    if len(items) == 1:
        joined = items[0]
    elif len(items) == 2:
        joined = f"{items[0]} and {items[1]}"
    else:
        joined = ", ".join(items[:-1]) + f", and {items[-1]}"
    tail = (". Construction administration services include shop drawing review "
            "and RFI responses for clarification of structural designs. Value "
            "engineering, contractor design alterations, construction "
            "administration tasks such as shop drawings review, RFI response, "
            "site visits, and field repairs are not included in the scope of "
            "basic services.")
    return f"Basic services shall include {joined}{tail}"


def _replace_paragraph_text(p, new_text):
    """Replace a paragraph's run text: set first run to new_text, blank the rest."""
    if not p.runs:
        p.add_run(new_text)
        return
    p.runs[0].text = new_text
    for r in p.runs[1:]:
        r.text = ""


def handle_basic_services_paragraph(doc, paragraph_text):
    """
    Find the paragraph that starts with 'Basic services shall include' and
    wraps its body in {{ }} (v2 template pattern). Replace the whole paragraph
    with the composed text.
    """
    for p in doc.paragraphs:
        t = p.text
        if 'Basic services shall include' in t and '{{' in t:
            _replace_paragraph_text(p, paragraph_text)
            return True
    return False


def apply_reimbursables_v2(doc, treatment):
    """
    v2 template wraps the 3 REIMBURSABLE paragraphs in {{ }}. Keep only the
    one matching `treatment`; delete the other two; strip the {{ }} from the
    survivor.
      - Standard        -> "in addition to basic services"
      - Digital Only    -> "not applicable"
      - Included in Fee -> "not anticipated"
    """
    sig_map = {
        "Standard":        "are in addition",
        "Digital Only":    "are not applicable",
        "Included in Fee": "are not anticipated",
    }
    want_sig = sig_map.get((treatment or "Standard"), sig_map["Standard"])
    remove = []
    for p in doc.paragraphs:
        t = p.text
        if t.lstrip().startswith("{{REIMBURSABLE EXPENSES"):
            if want_sig in t:
                cleaned = t.replace("{{", "", 1)
                cleaned = re.sub(r'\}\}\s*$', '', cleaned)
                _replace_paragraph_text(p, cleaned)
            else:
                remove.append(p)
    for p in remove:
        p._element.getparent().remove(p._element)


def handle_ssi_paragraph(doc, include=True):
    """
    v2 template has the SPECIAL STRUCTURAL INSPECTIONS paragraph wrapped in
    {{ }}. Keep + strip wrapper if include=True; delete entirely if False.
    """
    for p in doc.paragraphs:
        t = p.text
        if t.lstrip().startswith("{{SPECIAL STRUCTURAL INSPECTIONS"):
            if include:
                cleaned = t.replace("{{", "", 1)
                cleaned = re.sub(r'\}\}\s*$', '', cleaned)
                _replace_paragraph_text(p, cleaned)
            else:
                p._element.getparent().remove(p._element)
            return


# --- Number → words (for fee schedule) ---

_NUM_SMALL = ("zero","one","two","three","four","five","six","seven","eight",
              "nine","ten","eleven","twelve","thirteen","fourteen","fifteen",
              "sixteen","seventeen","eighteen","nineteen")
_NUM_TENS = ("","","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety")


def _num_words_under_1000(n):
    if n < 20:
        return _NUM_SMALL[n]
    if n < 100:
        return _NUM_TENS[n // 10] + (("-" + _NUM_SMALL[n % 10]) if n % 10 else "")
    return _NUM_SMALL[n // 100] + " hundred" + (
        " " + _num_words_under_1000(n % 100) if n % 100 else ""
    )


def dollars_to_words(amount):
    """Convert a dollar amount to 'Eight Thousand Dollars' style (title cased)."""
    if amount is None:
        return ""
    amt = int(round(amount))
    if amt == 0:
        return "Zero Dollars"
    parts = []
    if amt >= 1_000_000:
        parts.append(_num_words_under_1000(amt // 1_000_000) + " million")
        amt %= 1_000_000
    if amt >= 1000:
        parts.append(_num_words_under_1000(amt // 1000) + " thousand")
        amt %= 1000
    if amt > 0:
        parts.append(_num_words_under_1000(amt))
    return " ".join(parts).strip().title() + " Dollars"


def format_fee_cell(fee_line):
    """
    Render the FEE cell text for one included service line. Uses the
    'Words Dollars ($X,XXX.XX)' amount format. Phrasing differs for SSI
    and Site Visits (both use 'not to exceed' semantics).

    Caller should only pass included rows; unchecked rows are removed from
    the table by render_fee_schedule().
    """
    ftype = (fee_line.get("type") or "").strip()
    amount = fee_line.get("amount")
    service = fee_line.get("service", "")

    def amt_str(a):
        return f"{dollars_to_words(a)} (${a:,.2f})"

    if service == "Special Structural Inspections":
        if ftype == "Fixed Fee" and amount is not None:
            return f"Estimated Fixed Fee of {amt_str(amount)}"
        if amount is not None:
            return f"Estimated hourly not to exceed {amt_str(amount)}"
        return "Estimated hourly, billed at actual time expended"
    if service == "Engineering Site Visits":
        if ftype == "Fixed Fee" and amount is not None:
            return f"Fixed Fee of {amt_str(amount)}"
        if amount is not None:
            return f"Hourly not to exceed {amt_str(amount)}"
        return "Hourly, billed at actual time expended"

    # Standard rows (SD / DD / CD / CA)
    if ftype == "Fixed Fee" and amount is not None:
        return f"Fixed Fee of {amt_str(amount)}"
    if ftype == "Hourly":
        if amount is not None:
            return f"Hourly at {amt_str(amount)} per hour"
        return "Hourly, billed at actual time expended"
    return "Hourly"


def render_fee_schedule(doc, fee_lines):
    """
    Fill the FEE SCHEDULE table from fee_lines.
      - Rows whose Service has Include=False are REMOVED from the table.
      - Included rows get 'Words Dollars ($X,XXX.XX)' formatted cell text.
      - TOTAL row sums Fixed-Fee + 'not to exceed' amounts of included rows.
    """
    by_service = {fl["service"]: fl for fl in (fee_lines or [])}

    def set_cell_text(cell, text):
        first = True
        for p in cell.paragraphs:
            if first:
                if p.runs:
                    p.runs[0].text = text
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.add_run(text)
                first = False
            else:
                for r in p.runs:
                    r.text = ""

    for table in doc.tables:
        table_text = " ".join(cell.text for row in table.rows for cell in row.cells)
        if "Schematic Design" not in table_text or "FEE SCHEDULE" not in table_text:
            continue

        rows_to_remove = []
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            service_cell = row.cells[0]
            fee_cell     = row.cells[1]
            service_name = service_cell.text.strip()

            if service_name in by_service:
                fl = by_service[service_name]
                if not fl.get("include"):
                    rows_to_remove.append(row)
                else:
                    set_cell_text(fee_cell, format_fee_cell(fl))
            elif service_name.upper() == "TOTAL":
                total = 0.0
                for fl in (fee_lines or []):
                    if not fl.get("include"):
                        continue
                    amt = fl.get("amount") or 0
                    ftype = (fl.get("type") or "").lower()
                    if ftype == "fixed fee" or "not to exceed" in ftype:
                        total += amt
                if total > 0:
                    new_text = f"{dollars_to_words(total)} (${total:,.2f}) + Reimbursables"
                else:
                    new_text = "TBD + Reimbursables"
                set_cell_text(fee_cell, new_text)

        for row in rows_to_remove:
            row._element.getparent().remove(row._element)
        break


# ===================================================================
# Run against real Trevor Pan project (Notion-pulled data)
# ===================================================================
if __name__ == "__main__":
    # --- Pulled from Notion 2026-04-20 ---
    project_general_metal = {
        "Project Name": "26102144 - General Metal Construction PEMB",
        "State": "Arizona",
        "City": "Phoenix",
        "SharePoint Folder": "file:///C:/Users/MichaelBrusnahan/OneDrive%20-%20Kingdom%20Structural%20LLC/_Projects/2026/26102144%20General%20Metal%20Construction%20PEMB%20-%20Phoenix%2C%20AZ",
        "Engineering Status": "Proposal Sent",
        # NEW FIELDS (sample values — partner fills these on each project)
        "Project Type": "PEMB",
        "ICC Code Year": "2021",
        "Jurisdiction": "Phoenix",
        "Project Street": "<<FILL IN at intake: project street>>",
        "Scope Description": (
            "a pre-engineered metal building (PEMB) project, General Metal Construction PEMB, "
            "and general services to be provided including structural engineering, drafting, and "
            "construction administration for the primary steel frame, secondary members, and "
            "foundation design"
        ),
        "Reimbursables Treatment": "Standard",
        "Approx. Structural SF": 12000,
    }
    client_trevor_pan = {
        "Name": "Trevor Pan Architects",
        "Address": "4310 N 75th St Suite 145, Scottsdale, AZ 85251",
        "The Guy": "CV",
    }
    contact_trevor = {
        "Contact Name": "Trevor Pan",
        "Email": "trevor.pan@trevorpan.com",
        "Phone": "",
        "Phone (Mobile)": "480.277.3499",
        "Role": "",
    }

    out_dir = "/sessions/admiring-friendly-brown/mnt/outputs/26102144 General Metal Construction PEMB"
    out_path, merge_data = render_contract(
        project_general_metal, client_trevor_pan, contact_trevor, out_dir
    )
    print(f"Rendered: {out_path}")
    print(f"Size: {os.path.getsize(out_path)} bytes")
    print("\nMerge values used:")
    for k, v in merge_data.items():
        marker = "✓" if not str(v).startswith("<<FILL IN") else "⚠ "
        print(f"  {marker} {k:30s} = {v[:70]}")
