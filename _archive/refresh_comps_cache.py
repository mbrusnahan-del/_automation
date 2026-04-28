"""
Refresh the comp cache from live Notion data.

STANDALONE Python scripts can't call Notion MCP tools directly — only a
Claude session can. So this module provides the NORMALIZER (convert raw
Notion records into the COMPS dict shape) and expects the caller (a Claude
session in the scheduled task) to have already fetched the records.

Usage from a Claude session:

    import sys
    sys.path.insert(0, "/mnt/_Projects/_automation")  # adjust path as needed
    from refresh_comps_cache import normalize_jobs, commit_cache

    # 1. Claude session queries Notion:
    #    - All Jobs 26KS records
    #    - For each job's Project relation, fetch the Project page too
    jobs_records = [...]   # list of dicts with the Notion property values
    project_lookup = {url: project_dict, ...}  # keyed by project URL

    # 2. Normalize into COMPS format
    comps = normalize_jobs(jobs_records, project_lookup)

    # 3. Write the cache
    commit_cache(comps, source="scheduled-task-refresh-2026-04-21")
"""
import sys
import os

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)
from fee_memo import write_comps_cache


def _first_url(val):
    """A Notion relation property is a JSON array of page URLs (stringified)."""
    if not val:
        return None
    if isinstance(val, str):
        import json
        try:
            parsed = json.loads(val)
        except Exception:
            parsed = [val]
        val = parsed
    if isinstance(val, list) and val:
        return val[0]
    return None


def normalize_job(job, project_lookup):
    """
    Convert ONE Jobs 26KS raw record dict + the linked Project dict into a
    comp entry. Returns None if the job doesn't have enough data to be useful.

    Expected keys on `job` (as they appear in Notion query results):
      - Name (title)
      - Job (text)
      - Design, CD, CA (dollar numbers)
      - TOTAL SIGNED (dollar number)
      - Y (checkbox — "__YES__" or "__NO__")
      - Project (relation — JSON array of page URLs)
      - Approx. Construction Cost (dollar number, optional)
      - Complexity (select, optional)
      - Lost Reason (select, optional)
      - Project Type (rollup) — often comes through as the raw select name
      - Approx. Structural SF (rollup) — number

    Expected keys on the linked project dict (when no rollup available):
      - Project Type (select)
      - Approx. Structural SF (number)
    """
    name = (job.get("Name") or "").strip()
    if not name:
        return None

    project_url = _first_url(job.get("Project"))
    project = project_lookup.get(project_url) if project_url else None

    # Project Type — prefer rollup, fall back to direct join
    ptype = job.get("Project Type (rollup)") or (project or {}).get("Project Type") or "Other"
    # Approx Structural SF — prefer rollup, fall back to direct join
    sf = job.get("Approx. Structural SF (rollup)")
    if not sf and project:
        sf = project.get("Approx. Structural SF")

    won = (job.get("Y") or "").upper() in ("__YES__", "YES", "TRUE")

    design = job.get("Design")
    cd = job.get("CD")
    ca = job.get("CA")
    total_signed = job.get("TOTAL SIGNED")

    # Proposed total for lost rows: sum of design+cd+ca components
    total_proposed = None
    if not won:
        parts = [v for v in (design, cd, ca) if v and v > 0]
        total_proposed = sum(parts) if parts else None

    # Skip records with no fee information at all
    if won and not total_signed:
        return None
    if not won and not total_proposed:
        return None

    return {
        "name": name[:80],
        "job": (job.get("Job") or "").strip(),
        "type": ptype,
        "design": design or 0,
        "cd": cd or 0,
        "ca": ca or 0,
        "total_signed": total_signed if won else None,
        "total_proposed": total_proposed if not won else None,
        "won": won,
        "partner": None,  # deferred — could derive from partner split fields
        "approx_sf": int(sf) if sf else None,
        "construction_cost": job.get("Approx. Construction Cost") or None,
        "complexity": job.get("Complexity") or None,
        "lost_reason": job.get("Lost Reason") or None,
    }


def normalize_jobs(jobs_records, project_lookup):
    """Apply normalize_job to a list of records, dropping any that return None."""
    comps = []
    for j in jobs_records:
        c = normalize_job(j, project_lookup)
        if c:
            comps.append(c)
    return comps


def commit_cache(comps, source="unspecified"):
    """Write the normalized comps to the cache file, ready for next render."""
    return write_comps_cache(comps, source=source)
