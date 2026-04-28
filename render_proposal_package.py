"""
Render the Kingdom Structural contract .docx for a project and drop it into
the project's OneDrive Contracts folder.

Per v2.1 Change Order (2026-04-22): Fee Analysis Memo rendering has been
removed from this pipeline. Only the contract .docx is produced. Fee memo
code is preserved under _archive/ for future revival.

This is the function the "Contract Render" scheduled task (Job B) calls
after the Admin approves a Brief (Brief.Status = "Approved" AND
Brief.Ready to Render = true).

Folder/naming conventions match Kingdom Structural's existing OneDrive layout:
    _Projects/
        {year}/
            {KS_Job_Number} {Project Name} - {City, State}/
                Contracts/
                    {KS_Job_Number} KS Structural Contract - {Client} - {Project} - {YY.MM.DD}.docx
                    NOT USED/           (legacy — partners put abandoned drafts here)
                    Signed Contract/    (legacy — signed contracts land here)
"""
import os
import sys
import re
import time
from datetime import date

# Pull in the existing render functions (scripts live alongside this file)
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)
from render_from_notion import (
    build_merge_data,
    fill_placeholders,
    expand_scope_bullets,
    apply_reimbursables_treatment,   # v1 (legacy template) — kept for fallback
    cleanup_whitespace_artifacts,
    clear_highlight_on_resolved_runs,  # v2.1.2: strip highlight on resolved fields
    resolve_scope_bullets,
    merge_brief_into_project,
    # v2 template pipeline (2026-04-22)
    parse_brief_body,
    build_basic_services_paragraph,
    handle_basic_services_paragraph,
    apply_reimbursables_v2,
    handle_ssi_paragraph,
    render_fee_schedule,
    TEMPLATE,
)
from docx import Document
import shutil


# _automation/ is a direct child of _Projects/, so root is the parent
ONEDRIVE_ROOT = os.path.dirname(_SCRIPT_DIR)
# Canonical project folder template — cloned into every new project folder.
# Partners keep this folder as the source of truth for standard subfolders,
# template files, calc cover sheets, CAD title blocks, contract variants, etc.
TEMPLATE_PROJECT_FOLDER = os.path.join(
    ONEDRIVE_ROOT, "2026", "_26001001 New 2026 Job Name - City, STATE"
)


def sanitize(s):
    """Strip characters not safe for filenames on Windows."""
    return re.sub(r'[<>:"/\\|?*]', '', s).strip()


# ----------------------------------------------------------------------
# Path-length safety (Windows MAX_PATH = 260, 259 usable)
# ----------------------------------------------------------------------
# On partner machines the OneDrive root resolves to something like:
#   C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_Projects\
# We size folder + filename budgets so the deepest file in Contracts/ stays
# under 259 chars even on the longest partner path we know about.
_MAX_FULL_PATH = 255  # 4-char safety buffer under 259
_ONEDRIVE_ROOT_WIN = "C:\\Users\\MichaelBrusnahan\\OneDrive - Kingdom Structural LLC\\_Projects\\"


def _norm_for_path(name):
    """Make a string safe to embed in a Windows path. Does NOT truncate."""
    s = name.replace('/', ' - ')
    s = re.sub(r'[<>:"\\|?*]', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def _truncate_at_word(s, max_len):
    """Trim s to <= max_len, preferring the previous word boundary."""
    if len(s) <= max_len:
        return s
    cut = s[:max_len]
    sp = cut.rfind(' ')
    if sp >= max_len - 10:
        cut = cut[:sp]
    return cut.rstrip(' -_,.')


def project_folder_name(job_number, project_name_raw, city, state, client_name=""):
    """
    Build a Windows-safe project folder name AND return the truncated project
    name that should be used in filenames. Both results are tuned so the
    deepest contract/memo file still fits under Windows' 259-char path limit.

    Returns: (folder_name, project_name_short)
    """
    proj = _norm_for_path(project_name_raw)
    # Strip any leading "26XXXXXX " or "26XXXXXX - " job-number prefix so
    # we don't double up when we prepend the job number below.
    m = re.match(r'^\d{8}[\s\-]+(.+)$', proj)
    if m:
        proj = m.group(1).strip()

    loc_parts = [p for p in [city, state] if p]
    loc = ", ".join(loc_parts)
    loc_suffix = f" - {loc}" if loc else ""

    # Budget the project-name portion so the full path fits under _MAX_FULL_PATH.
    # Structure:
    #   {ONEDRIVE_ROOT}{YEAR}\{FOLDER}\Contracts\{FILE}
    # The contract filename is the longest-template case:
    #   "{JOB} KS Structural Contract - {CLIENT} - {PROJ} - YY.MM.DD.docx"
    prefix_len = len(_ONEDRIVE_ROOT_WIN) + len("2026\\")
    contracts_mid = len("\\Contracts\\")
    max_client = 30  # cap client name portion of the filename
    file_overhead = (
        len(f"{job_number} KS Structural Contract - ")
        + max_client
        + len(" - ")
        + len(" - YY.MM.DD.docx")
    )
    folder_overhead = len(f"{job_number} ") + len(loc_suffix)
    fixed_total = prefix_len + contracts_mid + file_overhead + folder_overhead
    # Split remaining budget between the folder's project portion and the
    # filename's project portion. They're embedded in different places so
    # each needs its own allowance.
    available_for_proj = max(20, (_MAX_FULL_PATH - fixed_total) // 2)

    proj_trunc = _truncate_at_word(proj, available_for_proj)
    folder_name = f"{job_number} {proj_trunc}{loc_suffix}"
    return folder_name, proj_trunc


def short_client_name(client_name, max_len=30):
    """Cap the client name we embed in filenames."""
    s = _norm_for_path(client_name or "")
    return _truncate_at_word(s, max_len)


def find_project_folder(ks_job_number, project_name_short):
    """
    Find the on-disk project folder matching the KS job number.
    Returns the full path or None.

    Year search order: the year implied by the number prefix first ("26…" →
    2026), then every existing sibling year under _Projects/ as a fallback
    in case a folder was filed under an unexpected year directory.
    """
    # Primary: use the 2-digit prefix (handles 2028, 2029, etc., not just 24-27)
    years_to_check: list[str] = []
    if ks_job_number[:2].isdigit():
        years_to_check.append("20" + ks_job_number[:2])
    # Fallback: any other year dirs that actually exist under _Projects/
    try:
        for entry in sorted(os.listdir(ONEDRIVE_ROOT)):
            if entry.isdigit() and len(entry) == 4 and entry not in years_to_check:
                years_to_check.append(entry)
    except OSError:
        pass

    for year in years_to_check:
        year_dir = os.path.join(ONEDRIVE_ROOT, year)
        if not os.path.isdir(year_dir):
            continue
        for name in os.listdir(year_dir):
            # Match on full 8-digit job number at the start of folder name.
            # Accept both "{num} Name - City, ST" (legacy) and "{num} - Name - City, ST" (v2.1.2).
            if name.startswith(ks_job_number + " ") or name.startswith(ks_job_number + "-"):
                return os.path.join(year_dir, name)
    return None


def populate_from_template(project_dir):
    """
    Copy the canonical template tree (subfolders + template files) INTO
    project_dir, preserving any files/folders that already exist. Never
    deletes existing content. Safe to re-run on established projects.

    Files that are OneDrive cloud-only (not yet downloaded locally) will
    error on copy; we log and skip them — partners can fetch them manually.
    """
    if not os.path.isdir(TEMPLATE_PROJECT_FOLDER):
        return {"error": f"Template folder not found: {TEMPLATE_PROJECT_FOLDER}"}

    stats = {"folders_created": 0, "files_copied": 0, "files_skipped_cloudonly": 0, "files_already_present": 0}

    for root, dirs, files in os.walk(TEMPLATE_PROJECT_FOLDER):
        rel = os.path.relpath(root, TEMPLATE_PROJECT_FOLDER)
        target_root = project_dir if rel == "." else os.path.join(project_dir, rel)
        if not os.path.isdir(target_root):
            os.makedirs(target_root, exist_ok=True)
            stats["folders_created"] += 1
        for f in files:
            src = os.path.join(root, f)
            dst = os.path.join(target_root, f)
            if os.path.exists(dst):
                stats["files_already_present"] += 1
                continue  # NEVER overwrite partner's work
            try:
                shutil.copy2(src, dst)
                stats["files_copied"] += 1
            except OSError as e:
                # OneDrive cloud-only files raise OSError (errno 22). Skip.
                stats["files_skipped_cloudonly"] += 1
    return stats


def render_contract(project_data, client_data, contact_data, contracts_dir,
                    today=None, engineer=None, brief=None,
                    brief_body=None, fee_lines=None):
    """Render the filled contract .docx into the Contracts folder.

    v2 pipeline (2026-04-22):
      - `brief`       — Proposal Brief properties overlay dict (City, State,
                        Project Type, Project Street, ICC, Jurisdiction, etc.)
      - `brief_body`  — raw markdown of the Brief page body (for checkbox parsing)
      - `fee_lines`   — list of {service, include, type, amount, order} from
                        the Contract Fee Lines DB filtered to this Brief

    Flow:
      1. Overlay Brief properties onto Project dict.
      2. Parse Brief body → Basic Services / Scope of Services / Reimbursables.
      3. Build scope bullets (Brief body preferred; scope_library fallback).
      4. Apply reimbursables variant + conditional SSI paragraph.
      5. Fill fee schedule from fee_lines.
      6. Compose basic services paragraph from checked items.
      7. Substitute simple {{placeholder}} merge fields.
    """
    project_data = merge_brief_into_project(project_data, brief)
    merge = build_merge_data(
        project_data, client_data, contact_data,
        engineer=engineer, today=today, brief=None,  # already merged above
    )

    today = today or date.today()
    date_short = today.strftime("%y.%m.%d")
    job = merge["ks_job_number"]
    client_full = merge["client_company_name"]
    proj_full = merge["project_name"]

    _, proj_short = project_folder_name(
        job, proj_full,
        project_data.get("City", ""),
        project_data.get("State", ""),
        client_name=client_full,
    )
    client_short = short_client_name(client_full)
    filename = f"{job} KS Structural Contract - {client_short} - {proj_short} - {date_short}.docx"
    filename = sanitize(filename)
    out_path = os.path.join(contracts_dir, filename)

    # Retry the template copy when the destination is locked. Causes we've
    # actually seen: (a) Word has the previous render open with an exclusive
    # lock, (b) OneDrive sync handler is mid-write to that path, (c) the
    # Contracts folder is being moved/restored by OneDrive Files-on-Demand.
    # Total back-off ~16s (1+2+3+4+5s) is generous enough to outlast a
    # transient OneDrive lock without dragging out a real Word-is-open case.
    # If it ultimately fails, raise with a clear, actionable error.
    last_err = None
    for attempt in range(5):
        try:
            shutil.copyfile(TEMPLATE, out_path)
            break
        except (PermissionError, OSError) as e:
            last_err = e
            time.sleep(attempt + 1)   # 1s, 2s, 3s, 4s, 5s
    else:
        raise PermissionError(
            f"Could not write contract to {out_path!r} after 5 retries. "
            f"Most common cause: the file is open in Word — close it and re-run. "
            f"Last OS error: {last_err}"
        )
    doc = Document(out_path)

    # Parse Brief body for body-driven content. Safe when brief_body=None
    # (returns empty lists; renderer falls back gracefully).
    parsed = parse_brief_body(brief_body or "")

    # 1. Reimbursables variant — body checkbox wins; Project.Reimbursables
    #    Treatment is legacy fallback.
    treatment = parsed.get("reimbursables") or project_data.get("Reimbursables Treatment")
    apply_reimbursables_v2(doc, treatment or "Standard")

    # 2. SSI paragraph — include by default. Body-checkbox "Special structural
    #    inspections shall be excluded" suppresses it. The scope bullet itself
    #    still appears (separate from the SSI billing paragraph).
    ssi_excluded = any(
        "special structural inspections shall be excluded" in i.get("text", "").lower()
        for i in parsed.get("scope_of_services", [])
    )
    handle_ssi_paragraph(doc, include=not ssi_excluded)

    # 3. Scope of Services bullets from Brief body (preferred) or fallback.
    scope_items = parsed.get("scope_of_services") or []
    if scope_items:
        bullets = [i["text"] for i in scope_items]
    else:
        bullets = resolve_scope_bullets(project_data)
    expand_scope_bullets(doc, bullets)

    # 4. Fee Schedule — remove unchecked rows, format included rows, sum total.
    if fee_lines:
        render_fee_schedule(doc, fee_lines)

    # 5. Basic Services paragraph — comma-join checked items.
    basic_para = build_basic_services_paragraph(parsed)
    handle_basic_services_paragraph(doc, basic_para)

    # 6. Simple placeholder substitution (must be last so runs are stable
    #    after any prior manipulations).
    fill_placeholders(doc, merge)

    # 7. Cleanup stray whitespace around punctuation.
    cleanup_whitespace_artifacts(doc)

    # 8. Clear highlighting on every run that holds a resolved value. Leave
    #    highlight intact on any run whose text still contains '<<FILL IN:'
    #    so engineers visually spot the data gaps before sending the contract.
    #    (v2.1.2 spec)
    clear_highlight_on_resolved_runs(doc)

    doc.save(out_path)
    return out_path, merge


def render_contract_package(project_data, client_data, contact_data,
                             today=None, engineer=None,
                             brief=None, brief_body=None, fee_lines=None):
    """
    Full pipeline: find project folder, ensure Contracts subfolder exists,
    render the contract, return a dict of paths.

    v2.1 (2026-04-22): Fee memo rendering removed from this pipeline per
    the v2.1 Change Order. Only the contract .docx is produced.

    Parameters:
      - `brief`      — dict of Brief properties (overlays Project)
      - `brief_body` — raw markdown string of the Brief page body
      - `fee_lines`  — list of fee-line dicts from the Brief's child
                       Fee Schedule DB
    """
    ks_job = project_data.get("Project Name", "")
    # Extract the 8-digit job number from project name ("26102144 - General Metal Construction PEMB")
    m = re.match(r'^(\d+)', ks_job)
    if not m:
        return {"error": "Could not extract KS Job Number from project name"}
    ks_job_number = m.group(1)

    folder = find_project_folder(ks_job_number, ks_job)
    if not folder:
        return {"error": f"Project folder for {ks_job_number} not found under {ONEDRIVE_ROOT}"}

    # Clone the canonical template tree into this project folder. This is
    # additive — never removes partner's existing files; only fills in what's
    # missing. Safe to run on every render cycle.
    populate_stats = populate_from_template(folder)

    contracts_dir = os.path.join(folder, "Contracts")
    os.makedirs(contracts_dir, exist_ok=True)

    contract_path, merge_data = render_contract(
        project_data, client_data, contact_data, contracts_dir,
        today=today, engineer=engineer, brief=brief,
        brief_body=brief_body, fee_lines=fee_lines,
    )

    return {
        "project_folder": folder,
        "contracts_dir": contracts_dir,
        "contract": contract_path,
        "merge_data": merge_data,
        "template_populate": populate_stats,
    }


# Backwards-compat alias. Old callers that still say render_proposal_package()
# will keep working. New code should use render_contract_package().
render_proposal_package = render_contract_package


# ===================================================================
# Test: run against real Trevor Pan PEMB project (contract-only, v2.1)
# ===================================================================
if __name__ == "__main__":
    project_general_metal = {
        "Project Name": "26102144 - General Metal Construction PEMB",
        "State": "Arizona",
        "City": "Phoenix",
        "SharePoint Folder": "file:///C:/Users/MichaelBrusnahan/OneDrive%20-%20Kingdom%20Structural%20LLC/_Projects/2026/26102144%20General%20Metal%20Construction%20PEMB%20-%20Phoenix%2C%20AZ",
        "Engineering Status": "Proposal Sent",
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

    result = render_contract_package(
        project_general_metal, client_trevor_pan, contact_trevor
    )
    if "error" in result:
        print(f"ERROR: {result['error']}")
    else:
        print(f"Project folder: {result['project_folder']}")
        print(f"Contracts dir:  {result['contracts_dir']}")
        print(f"Contract:       {result['contract']}")
