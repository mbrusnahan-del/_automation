"""
Fee Analysis Memo generator.
Takes a project intake (dict) and produces a .docx memo with:
  - Project summary
  - Closest comps from the comp database (with win/loss status)
  - Fee distribution stats
  - Close-rate context and pricing tier recommendations
  - External reference points

Uses real Jobs 26KS data pulled from Notion.
"""
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import date, datetime, timezone
import os
import json
import statistics as stats

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_COMPS_CACHE_PATH = os.path.join(_SCRIPT_DIR, "comps_cache.json")


def get_comps():
    """
    Return the active comp list. Preference order:
      1. comps_cache.json — refreshed from live Notion by scheduled task
      2. hardcoded COMPS fallback below

    The cache is a JSON object: {"generated_at": ISO8601, "comps": [...]}
    where each comp dict matches the same shape as COMPS entries.
    """
    if os.path.exists(_COMPS_CACHE_PATH):
        try:
            with open(_COMPS_CACHE_PATH, "r", encoding="utf-8") as f:
                payload = json.load(f)
            comps = payload.get("comps")
            if isinstance(comps, list) and comps:
                return comps
        except Exception:
            pass
    return COMPS


def write_comps_cache(comps, source="manual"):
    """
    Write a fresh comp list to the cache. Called by the Claude scheduled task
    after querying Jobs 26KS + Projects. `source` is just a label for logs.
    """
    payload = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "source": source,
        "count": len(comps),
        "comps": comps,
    }
    with open(_COMPS_CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
    return _COMPS_CACHE_PATH


# =====================================================================
# Comp Database — real Jobs 26KS records (pulled from Notion 2026-04-20)
# Project Type inferred from project name where Notion field not set.
# =====================================================================
COMPS = [
    # approx_sf values are rough estimates for MVP; these should be replaced
    # with actual values from the new "Approx. Structural SF" Notion field
    # as teams populate it. None = unknown (shown as "—" in memo).
    # --- PEMB / Commercial / Industrial ---
    {"name": "Faith PEMB Final Revisions",       "job": "25052001.1", "type": "PEMB",
     "design": 23600, "cd": 1100, "ca": 2300, "total_signed": 23600, "won": True,
     "partner": "JB", "approx_sf": 15000},
    {"name": "Emerald Industrial Building Suite 5", "job": "25101049", "type": "Commercial Remodel",
     "design": 1200, "cd": 1200, "ca": 300, "total_signed": 1200, "won": True,
     "partner": "JB", "approx_sf": 2500},
    {"name": "Emerald Industrial Building Suite 2", "job": "26042049", "type": "Commercial Remodel",
     "design": 1500, "cd": 1500, "ca": 400, "total_signed": 1500, "won": True,
     "partner": "CV", "approx_sf": 3000},
    {"name": "91st and McDowell Commercial Pad",  "job": "26054049", "type": "Commercial New",
     "design": 19800, "cd": 19800, "ca": 3700, "total_proposed": 23500, "won": False,
     "partner": None, "approx_sf": 10000},
    {"name": "Sunstate Companies PEMB Foundation","job": "25192033", "type": "PEMB",
     "design": 7400, "cd": 0, "ca": 850, "total_proposed": 8250, "won": False,
     "partner": None, "approx_sf": 8000},
    {"name": "Flynn Family Farm PEMB",            "job": "26073136", "type": "PEMB",
     "design": None, "cd": None, "ca": None, "total_proposed": None, "won": False,
     "partner": None, "approx_sf": None},
    {"name": "Azteca Milling Pump Building",      "job": "26007057", "type": "Commercial New",
     "design": 3800, "cd": 3800, "ca": 0, "total_proposed": 3800, "won": False,
     "partner": None, "approx_sf": 1500},
    {"name": "Apache Building Apartments",        "job": "25207019", "type": "Multifamily",
     "design": None, "cd": None, "ca": None, "total_proposed": None, "won": False,
     "partner": None, "approx_sf": None},

    # --- Residential ---
    {"name": "Coder Residence",                   "job": "25055054", "type": "New Residence",
     "design": 11600, "cd": 11600, "ca": 0, "total_signed": 11600, "won": True,
     "partner": "CV", "approx_sf": 3500},
    {"name": "Rios Remodel",                      "job": "25097072", "type": "Residential Remodel",
     "design": 3222, "cd": 3200, "ca": 140, "total_signed": 3222, "won": True,
     "partner": "CV", "approx_sf": 2000},
    {"name": "Ayala Residence Remodel",           "job": "26009051", "type": "Residential Remodel",
     "design": 1000, "cd": 1000, "ca": 0, "total_signed": 1000, "won": True,
     "partner": "CV", "approx_sf": 1500},
    {"name": "Salas Residence New ADU",           "job": "26008094", "type": "New Residence",
     "design": 4500, "cd": 0, "ca": 0, "total_proposed": 4500, "won": False,
     "partner": None, "approx_sf": 800},

    # --- Observation / Feasibility ---
    {"name": "Legacy Hotel Structural Observation","job": "25202004", "type": "Structural Observation",
     "design": 3000, "cd": 1500, "ca": 0, "total_signed": 3000, "won": True,
     "partner": "CV", "approx_sf": None},  # Observation — SF not meaningful

    # ====================================================================
    # Expansion batch — commercial / hotel / TI / retail comps pulled from
    # Jobs 26KS on 2026-04-21.  Rough SF estimates — refine as teams update.
    # ====================================================================

    # --- Tenant Improvement / Retail / Restaurant TI ---
    {"name": "PHNX Retail at Miller and Southern", "job": "240040", "type": "Tenant Improvement",
     "design": 6950, "cd": 6500, "ca": 400, "total_signed": 6950, "won": True,
     "partner": "JB", "approx_sf": 3500},
    {"name": "Roma Pizzeria Remodel", "job": "25133033", "type": "Tenant Improvement",
     "design": 1994, "cd": 1950, "ca": 0, "total_signed": 1994, "won": True,
     "partner": "CV", "approx_sf": 2000},
    {"name": "Dental Office Roanoke TX", "job": "26046049", "type": "Tenant Improvement",
     "design": 4250, "cd": 4250, "ca": 0, "total_proposed": 4250, "won": False,
     "partner": None, "approx_sf": 2500},
    {"name": "Overton Office", "job": "26026061", "type": "Tenant Improvement",
     "design": 11900, "cd": 11000, "ca": 1100, "total_proposed": 11900, "won": False,
     "partner": None, "approx_sf": 5000},

    # --- Commercial Remodel / Hotel ---
    {"name": "Wailea Hotel Remodel", "job": "240044", "type": "Commercial Remodel",
     "design": 8480, "cd": 4240, "ca": 2120, "total_signed": 8480, "won": True,
     "partner": "JB", "approx_sf": 5000},
    {"name": "Cambria Hotel VE Revisions", "job": "240038.2", "type": "Commercial Remodel",
     "design": 2400, "cd": 2400, "ca": 1225, "total_signed": 2400, "won": True,
     "partner": "CV", "approx_sf": 2000},
    {"name": "Williams Four Points Hotel Conversion", "job": "25007029", "type": "Commercial Remodel",
     "design": 14650, "cd": 10000, "ca": 150, "total_signed": 14650, "won": True,
     "partner": "JB", "approx_sf": 30000},
    {"name": "Townplace Suites Hotel Phoenix AZ", "job": "25193107", "type": "Commercial Remodel",
     "design": 47600, "cd": 45800, "ca": 4600, "total_proposed": 47600, "won": False,
     "partner": None, "approx_sf": 60000},
    {"name": "United Cambria Hotel Mesa", "job": "26018004", "type": "Commercial Remodel",
     "design": 1000, "cd": 1000, "ca": 0, "total_signed": 1000, "won": True,
     "partner": "CV", "approx_sf": 500},
]


# Which types are "close" to each other for fallback matching
TYPE_GROUPS = {
    "PEMB":                   ["PEMB", "Commercial New"],
    "Commercial New":         ["Commercial New", "PEMB"],
    "Commercial Remodel":     ["Commercial Remodel", "Commercial New"],
    "Multifamily":            ["Multifamily", "Commercial New"],
    "New Residence":          ["New Residence", "Residential Remodel"],
    "Residential Remodel":    ["Residential Remodel", "New Residence"],
    "Structural Observation": ["Structural Observation", "Peer Review", "Feasibility"],
    "Feasibility":            ["Feasibility", "Structural Observation"],
    "Peer Review":            ["Peer Review", "Structural Observation"],
    "Retrofit":               ["Retrofit", "Commercial Remodel"],
    "Tenant Improvement":     ["Tenant Improvement", "Commercial Remodel", "Commercial New"],
}


def find_comps(project_type, comps=None, min_with_fee=4, subject_sf=None):
    if comps is None:
        comps = get_comps()
    """
    Return (primary, related) comps.

    Primary = exact type match.
    Related = adjacent types added when any of these is true:
      1. Primary doesn't have at least min_with_fee comps with fee data
      2. Subject project is much larger than any primary comp (subject_sf >
         2x the max primary SF) — small primary comps aren't useful for a
         much bigger project in the same category (e.g. 24K SF driving range
         classified as TI vs. 2-5K SF pizzeria TIs).

    This keeps the comp set rich enough for reasonable statistics AND keeps
    the engine honest about when it needs broader reference points.
    """
    primary_types = TYPE_GROUPS.get(project_type, [project_type])
    primary = [c for c in comps if c["type"] == project_type]
    primary_with_fee = sum(1 for c in primary if comp_fee(c) is not None)

    # Scale mismatch check: is the subject project dramatically bigger than
    # any primary comp in its category?
    scale_mismatch = False
    if subject_sf:
        primary_sfs = [c.get("approx_sf") for c in primary if c.get("approx_sf")]
        if primary_sfs and subject_sf > max(primary_sfs) * 2:
            scale_mismatch = True

    if primary_with_fee >= min_with_fee and not scale_mismatch:
        return primary, []
    related = [c for c in comps if c["type"] in primary_types and c not in primary]
    return primary, related


def comp_fee(c):
    """Return the relevant fee for a comp (signed if won, proposed if lost)."""
    if c["won"]:
        return c.get("total_signed")
    return c.get("total_proposed")


def analyze_comps(comps):
    """Compute statistics on won fees and win rate."""
    won = [c for c in comps if c["won"] and comp_fee(c) is not None]
    lost = [c for c in comps if not c["won"] and comp_fee(c) is not None]
    won_fees = [comp_fee(c) for c in won]

    stats_d = {}
    if won_fees:
        won_fees_sorted = sorted(won_fees)
        stats_d["won_count"] = len(won_fees)
        stats_d["lost_count"] = len(lost)
        stats_d["median"] = stats.median(won_fees_sorted)
        stats_d["min"] = min(won_fees_sorted)
        stats_d["max"] = max(won_fees_sorted)
        n = len(won_fees_sorted)
        if n >= 4:
            stats_d["p25"] = won_fees_sorted[int(n * 0.25)]
            stats_d["p75"] = won_fees_sorted[int(n * 0.75)]
        else:
            stats_d["p25"] = stats_d["min"]
            stats_d["p75"] = stats_d["max"]
        # p90 — used as the stretch anchor (top 10% of historical wins).
        # For sparse comp sets (n<10), p90 falls back to max.
        if n >= 10:
            stats_d["p90"] = won_fees_sorted[int(n * 0.90)]
        else:
            stats_d["p90"] = stats_d["max"]
    else:
        stats_d = {"won_count": 0, "lost_count": len(lost), "median": None,
                   "min": None, "max": None, "p25": None, "p75": None, "p90": None}

    # Win rate in category
    total = len(won) + len(lost)
    stats_d["win_rate"] = (len(won) / total * 100) if total else None
    stats_d["total_comps"] = total
    return stats_d


def analyze_phase_distribution(comps):
    """Compute typical Design vs CA fee split from won comps that have both
    fields populated. Returns averaged ratios across comps where the split
    is computable (design > 0 AND ca >= 0 AND at least one is non-zero).

    Returns None if fewer than 3 comps qualify (sample too small to mean
    anything statistically — engineers should fall back to standard heuristics).
    """
    ratios = []
    for c in comps:
        if not c.get("won"):
            continue
        design = c.get("design")
        ca = c.get("ca")
        if design is None or ca is None:
            continue
        if design <= 0 and ca <= 0:
            continue
        total = design + ca
        if total <= 0:
            continue
        ratios.append((design / total, ca / total))

    if len(ratios) < 3:
        return None

    avg_design_pct = sum(r[0] for r in ratios) / len(ratios)
    avg_ca_pct = sum(r[1] for r in ratios) / len(ratios)

    # Count how many comps had non-zero CA — that's a useful signal.
    ca_inclusion_rate = sum(1 for c in comps
                             if c.get("won") and (c.get("ca") or 0) > 0) / max(
                                 1, sum(1 for c in comps if c.get("won")))

    return {
        "design_pct": avg_design_pct,
        "ca_pct": avg_ca_pct,
        "sample_size": len(ratios),
        "ca_inclusion_rate": ca_inclusion_rate,
    }


def recommend_tiers(stats_d, project_type):
    """Compute Conservative / Market / Stretch fee tiers from stats.

    Tier definitions (v2.1.11):
      Conservative = median of historical wins. The fee that closes most often.
      Market       = max(median * 1.25, P75 of wins). Above-typical pricing
                     that still wins at a meaningful rate.
      Stretch      = market * 1.50. Premium quote — 50% above market. Per
                     KS partner experience, clients regularly sign same-day
                     at this tier when scope and relationship support it.

    History: prior formulas anchored to max_won*1.30 or P90 produced numbers
    that felt either too conservative or too aggressive. market*1.50 reflects
    the actual pricing posture that's been working in practice.
    """
    if not stats_d["median"]:
        return None

    median = stats_d["median"]
    p75 = stats_d["p75"]

    conservative = int(round(median / 100) * 100)
    market_raw = max(median * 1.25, p75)
    market = int(round(market_raw / 100) * 100)
    stretch_raw = market_raw * 1.50
    stretch = int(round(stretch_raw / 100) * 100)

    return {
        "conservative": conservative,
        "market": market,
        "stretch": stretch,
    }


# =====================================================================
# Document rendering helpers
# =====================================================================
def _set_cell_bg(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)


def _set_cell_borders(cell, color='888888'):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), color)
        tc_borders.append(b)
    tc_pr.append(tc_borders)


def _fill_row(table, row_idx, values, bold=False, bg=None, font_size=10, white_text=False):
    for i, v in enumerate(values):
        cell = table.rows[row_idx].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        run = p.add_run(str(v))
        run.font.size = Pt(font_size)
        run.font.bold = bold
        if white_text:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        if bg:
            _set_cell_bg(cell, bg)
        _set_cell_borders(cell)


def _add_heading(doc, text, size=9, color=None, space_before=2, space_after=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.bold = True
    if color:
        run.font.color.rgb = color


def _add_para(doc, text, size=11, italic=False, bold=False, space_after=6, space_before=0, indent=0):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    if space_before:
        p.paragraph_format.space_before = Pt(space_before)
    if indent:
        p.paragraph_format.left_indent = Inches(indent)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.italic = italic
    run.font.bold = bold


def _add_bullet(doc, text, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.25)
    p.paragraph_format.first_line_indent = Inches(-0.15)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run('•  ' + text)
    r.font.size = Pt(size)


def _fmt(v):
    if v is None:
        return "—"
    return f"${v:,}"


def _fmt_psf(fee, sf):
    """Format $/SF, return "—" if either is missing."""
    if not fee or not sf:
        return "—"
    return f"${fee/sf:.2f}/SF"


def _fmt_sf(sf):
    if not sf:
        return "—"
    return f"{int(sf):,} SF"


# =====================================================================
# Main memo render
# =====================================================================
def render_fee_memo(project, out_dir, overall_close_rate=70, filename=None):
    """
    project dict keys:
        project_name (short, e.g., 'General Metal Construction PEMB')
        ks_job_number
        project_type (must match keys in TYPE_GROUPS)
        approx_sf (number, optional)
        jurisdiction
        location (e.g., 'Phoenix, AZ')
        scope_description
        partner (JB/CV/GO)
        client_company_name
        engineering_status
    """
    doc = Document()

    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(9)

    KS_BLUE = RGBColor(0x1F, 0x3A, 0x5F)

    # -------- Title (compact) --------
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run('FEE ANALYSIS MEMO  —  Internal Working Document, Kingdom Structural LLC')
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = KS_BLUE

    # -------- Metadata (compact single line under title) --------
    partner_map = {"JB": "Jonathan Brusnahan, SE, PE", "CV": "Chris Valdez, PE", "GO": "Griffin O'Reilly"}
    partner_name = partner_map.get(project.get('partner', ''), project.get('partner', ''))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(
        f"{project.get('project_name', '')}  |  Job # {project.get('ks_job_number', '')}  |  "
        f"{project.get('project_type', '')}  |  Partner: {partner_name}  |  "
        f"Prepared {date.today().strftime('%b %d, %Y')}"
    )
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    # -------- 1. Project Summary --------
    _add_heading(doc, '1. Project Summary', color=KS_BLUE, space_before=4)
    subject_sf = project.get('approx_sf')
    sf_note = f" Approx. {subject_sf:,} SF." if subject_sf else ""
    _add_para(doc,
        project.get('scope_description', 'Scope not yet set in Notion.') +
        sf_note +
        f" {project.get('location', 'TBD')} | Jurisdiction: {project.get('jurisdiction', 'TBD')}.",
        size=8, space_after=2
    )

    # Gather comps for all subsequent sections
    ptype = project.get('project_type', '')
    primary, related = find_comps(ptype, subject_sf=subject_sf)
    all_comps = primary + related
    lost_with_fee = [c for c in all_comps if not c["won"] and comp_fee(c) is not None]
    won_with_fee = [c for c in all_comps if c["won"] and comp_fee(c) is not None]

    # -------- 2. LOSS SIGNALS (prominent early section) --------
    _add_heading(doc, '2. Loss Signals — Where Clients Walked', color=RGBColor(0xB0, 0x3A, 0x2E),
                 space_before=14)
    if lost_with_fee:
        # Cap loss-signals display at 5 to keep memo on one page; all losses still inform stats.
        MAX_LOSS_DISPLAY = 5
        displayed_losses = sorted(lost_with_fee, key=lambda c: -(comp_fee(c) or 0))[:MAX_LOSS_DISPLAY]

        lead = f'{len(lost_with_fee)} lost proposal(s) — ceiling references where clients walked.'
        if len(lost_with_fee) > MAX_LOSS_DISPLAY:
            lead += f' Top {MAX_LOSS_DISPLAY} shown by proposed fee.'
        _add_para(doc, lead, size=8, italic=True, space_after=2)

        loss_table = doc.add_table(rows=len(displayed_losses) + 1, cols=5)
        widths = [Inches(2.3), Inches(1.2), Inches(1.0), Inches(1.0), Inches(1.0)]
        for i, w in enumerate(widths):
            loss_table.columns[i].width = w
        _fill_row(loss_table, 0, ['Project', 'Type', 'Proposed Fee', 'SF', '$/SF'],
                  bold=True, bg='B03A2E', white_text=True, font_size=9)
        for i, c in enumerate(displayed_losses, start=1):
            fee = comp_fee(c)
            sf = c.get("approx_sf")
            _fill_row(loss_table, i, [
                c['name'][:40],
                c['type'],
                _fmt(fee),
                _fmt_sf(sf),
                _fmt_psf(fee, sf),
            ], bg='FFF2F2', font_size=7)

        # Highlight dangerous proximity: lost bids near or below where we'd recommend pricing
        if won_with_fee:
            lowest_lost = min(comp_fee(c) for c in lost_with_fee)
            highest_won = max(comp_fee(c) for c in won_with_fee)
            if lowest_lost <= highest_won:
                _add_para(doc,
                    f'⚠ Notable: lowest lost bid ({_fmt(lowest_lost)}) sits at or below the highest won '
                    f'({_fmt(highest_won)}) in this category. Price alone does not determine close — '
                    f'scope fit, client relationship, and timing matter. Treat these as directional signals, '
                    f'not hard ceilings.',
                    size=10, bold=True, space_before=6, space_after=4
                )
    else:
        _add_para(doc,
            'No lost proposals recorded in this category yet. Pricing ceiling is unknown — '
            'recommend treating Market tier as aspirational and watching closely for the first loss.',
            size=10, italic=True)

    # -------- 3. Comparable Projects --------
    _add_heading(doc, '3. Comparable Projects from Jobs 26KS', color=KS_BLUE)

    # Select up to 10 most-relevant comps for display (stats still use ALL comps)
    # Ranking: exact type match > related type, comps with fees > without,
    # closer SF to subject SF preferred.
    def _relevance_key(c):
        exact_type = 0 if c["type"] == ptype else 1
        has_fee = 0 if comp_fee(c) is not None else 1
        sf_distance = 0
        if subject_sf and c.get("approx_sf"):
            sf_distance = abs(c["approx_sf"] - subject_sf) / max(subject_sf, 1)
        else:
            sf_distance = 999  # unknown SF = deprioritized for display
        return (exact_type, has_fee, sf_distance)

    MAX_DISPLAY = 5
    ranked_comps = sorted(all_comps, key=_relevance_key)
    display_comps = ranked_comps[:MAX_DISPLAY]

    lead_text = f'{len(primary)} exact "{ptype}" match(es)'
    if related:
        lead_text += f' + {len(related)} related-type'
    lead_text += f'. Top {len(display_comps)} shown by relevance; all {len(all_comps)} feed the statistics.'
    _add_para(doc, lead_text, size=8, italic=True, space_after=2)

    comp_table = doc.add_table(rows=len(display_comps) + 1, cols=6)
    widths = [Inches(2.0), Inches(1.2), Inches(0.9), Inches(0.9), Inches(0.9), Inches(0.7)]
    for i, w in enumerate(widths):
        comp_table.columns[i].width = w
    _fill_row(comp_table, 0, ['Project', 'Type', 'Total Fee', 'SF', '$/SF', 'Outcome'],
              bold=True, bg='1F3A5F', white_text=True, font_size=9)
    for i, c in enumerate(display_comps, start=1):
        outcome = 'WON' if c['won'] else 'LOST'
        fee = comp_fee(c)
        sf = c.get("approx_sf")
        bg = None
        if outcome == 'LOST':
            bg = 'FFF2F2'
        elif i % 2 == 0:
            bg = 'F5F5F5'
        _fill_row(comp_table, i, [
            c['name'][:36],
            c['type'],
            _fmt(fee),
            _fmt_sf(sf),
            _fmt_psf(fee, sf),
            outcome,
        ], bg=bg, font_size=7)

    # SF coverage compact note
    sf_coverage = sum(1 for c in all_comps if c.get("approx_sf"))
    _add_para(doc,
        f'SF coverage: {sf_coverage} of {len(all_comps)} comps. Older SFs are partner estimates.',
        size=8, italic=True, space_before=1, space_after=0
    )

    # -------- 4. Fee Distribution + Win Rate Context --------
    _add_heading(doc, '4. Fee Distribution and Win Rate', color=KS_BLUE)
    stats_d = analyze_comps(all_comps)

    if stats_d["won_count"] > 0:
        dist_table = doc.add_table(rows=2, cols=5)
        for i, w in enumerate([Inches(1.3)] * 5):
            dist_table.columns[i].width = w
        _fill_row(dist_table, 0, ['Won Count', 'Lost Count', 'Median', 'P25', 'P75 / Max'],
                  bold=True, bg='1F3A5F', white_text=True, font_size=9)
        _fill_row(dist_table, 1, [
            stats_d["won_count"], stats_d["lost_count"],
            _fmt(stats_d["median"]), _fmt(stats_d["p25"]),
            f'{_fmt(stats_d["p75"])} / {_fmt(stats_d["max"])}'
        ], font_size=9)

        cat_win_rate = f'{stats_d["win_rate"]:.0f}%' if stats_d["win_rate"] else 'N/A'
        _add_para(doc,
            f'Category win rate: {cat_win_rate} ({stats_d["won_count"]} won / '
            f'{stats_d["lost_count"]} lost). Overall KS: {overall_close_rate}%. Industry: 30–50%.',
            size=8, space_before=2, space_after=0
        )
    else:
        _add_para(doc,
            'No won comps in category yet — pricing tiers below are based on related-type projects '
            'and industry heuristics only. Recommend broader comp matching once more category data is available.',
            size=10, italic=True)

    # -------- 5. Recommended Fee Tiers --------
    _add_heading(doc, '5. Recommended Fee Tiers', color=KS_BLUE)
    tiers = recommend_tiers(stats_d, ptype)

    if tiers:
        tier_table = doc.add_table(rows=4, cols=4)
        for i, w in enumerate([Inches(1.1), Inches(1.1), Inches(0.9), Inches(3.3)]):
            tier_table.columns[i].width = w
        _fill_row(tier_table, 0, ['Tier', 'Total Fee', '$/SF', 'Basis'],
                  bold=True, bg='1F3A5F', white_text=True, font_size=9)

        median = stats_d["median"] or 0

        tier_rows = [
            ['Conservative', _fmt(tiers["conservative"]),
             _fmt_psf(tiers["conservative"], subject_sf),
             f'Matches category median (${median:,}). Safe — preserves current win behavior.'],
            ['Market  ★', _fmt(tiers["market"]),
             _fmt_psf(tiers["market"], subject_sf),
             f'+25% over median OR at P75 of wins. Tests upward price discipline.'],
            ['Stretch', _fmt(tiers["stretch"]),
             _fmt_psf(tiers["stretch"], subject_sf),
             f'50% above market. Premium quote — defensible when scope, complexity, or relationship supports it.'],
        ]
        for i, row in enumerate(tier_rows, start=1):
            bg = 'E8F0F8' if i == 2 else ('F5F5F5' if i % 2 == 0 else None)
            _fill_row(tier_table, i, row, bg=bg, font_size=7)

        # Danger-zone check: any lost bids within ±20% of each tier's total fee?
        # That's the pricing band where losses actually occurred.
        if lost_with_fee:
            for tier_name, fee_val in [('Conservative', tiers['conservative']),
                                        ('Market', tiers['market']),
                                        ('Stretch', tiers['stretch'])]:
                nearby = [c for c in lost_with_fee
                          if 0.80 * fee_val <= comp_fee(c) <= 1.20 * fee_val]
                if nearby:
                    examples = '; '.join(
                        f'{c["name"][:30]} lost at {_fmt(comp_fee(c))}'
                        for c in nearby[:2]
                    )
                    _add_para(doc,
                        f'⚠ {tier_name} ({_fmt(fee_val)}) within 20% of lost bids: {examples}.',
                        size=8, bold=True, space_before=2, space_after=0
                    )

        _add_para(doc,
            '★ Market = default posture. Stretch reflects pricing that has been signing same-day in practice — partner judgment.',
            size=8, italic=True, space_before=2)
    else:
        _add_para(doc, 'Unable to compute tiers — no won comps in category.',
                  size=8, italic=True)

    # -------- 6. Phase Distribution --------
    _add_heading(doc, '6. Phase Distribution', color=KS_BLUE)
    phase = analyze_phase_distribution(all_comps)
    if phase and tiers:
        design_pct = phase["design_pct"]
        ca_pct = phase["ca_pct"]
        sample = phase["sample_size"]
        ca_rate = phase["ca_inclusion_rate"]

        _add_para(doc,
            f'Typical fee split (n={sample}): {design_pct*100:.0f}% Design / {ca_pct*100:.0f}% CA. '
            f'CA included on {ca_rate*100:.0f}% of comparable wins.',
            size=8, space_after=2
        )

        phase_table = doc.add_table(rows=4, cols=4)
        for i, w in enumerate([Inches(1.4), Inches(1.4), Inches(1.4), Inches(1.4)]):
            phase_table.columns[i].width = w
        _fill_row(phase_table, 0, ['Tier', 'Total Fee', f'Design ({design_pct*100:.0f}%)', f'CA ({ca_pct*100:.0f}%)'],
                  bold=True, bg='1F3A5F', white_text=True, font_size=9)

        for i, tier_name in enumerate(['Conservative', 'Market', 'Stretch'], start=1):
            tier_key = tier_name.lower()
            total = tiers[tier_key]
            design_amt = int(round(total * design_pct / 100) * 100)
            ca_amt = int(round(total * ca_pct / 100) * 100)
            bg = 'E8F0F8' if tier_name == 'Market' else ('F5F5F5' if i % 2 == 0 else None)
            _fill_row(phase_table, i,
                      [tier_name, _fmt(total), _fmt(design_amt), _fmt(ca_amt)],
                      bg=bg, font_size=7)

        _add_para(doc,
            'Design phase subdivides ~15–20% SD / ~25–30% DD / ~50–60% CD (industry convention). Use as Fee Schedule starting points.',
            size=8, italic=True, space_before=2
        )
    else:
        _add_para(doc,
            'Insufficient comp data with phase breakdown to compute distribution. '
            'Fall back to standard heuristic: ~80% Design phase / ~20% CA for typical projects.',
            size=10, italic=True
        )

    # -------- 7. External Reference Points --------
    _add_heading(doc, '7. External Reference Points', color=KS_BLUE)
    _add_para(doc,
        f'Structural fees typically run 1.5–3.0% of structural construction cost. '
        f'Phoenix demand elevated 2025–2026 (supports upward pressure). '
        f'Industry close-rate 30–50% (KS overall 70% is an outlier). '
        f'Comp DB: {len(get_comps())} projects total.',
        size=8, italic=True, space_after=2
    )

    # -------- 8. Partner Decision --------
    _add_heading(doc, '8. Partner Decision', color=KS_BLUE)
    _add_para(doc, 'Circle selected tier and record final fee below. Contract fee schedule will be populated with this value.',
              size=10, italic=True, space_after=8)

    if tiers:
        decision = doc.add_table(rows=1, cols=4)
        for i, w in enumerate([Inches(1.6), Inches(1.6), Inches(1.6), Inches(1.7)]):
            decision.columns[i].width = w
        _fill_row(decision, 0, [
            f'☐ Conservative  {_fmt(tiers["conservative"])}',
            f'☐ Market  {_fmt(tiers["market"])}',
            f'☐ Stretch  {_fmt(tiers["stretch"])}',
            'Final Fee: $________'
        ], bold=True, font_size=9)


    # -------- Footer --------
    p = doc.add_paragraph()
    r = p.add_run('_' * 80)
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(
        f'Generated from Jobs 26KS comp database ({len(get_comps())} projects) on '
        f'{date.today().isoformat()}. Recommendations are model-generated and subject to partner override.'
    )
    r.font.size = Pt(8)
    r.font.italic = True

    os.makedirs(out_dir, exist_ok=True)
    if filename is None:
        filename = f'{project.get("ks_job_number","")} Fee Analysis Memo - {project.get("project_name","")}.docx'
    filename = ''.join(c for c in filename if c not in '<>:"/\\|?*')
    out_path = os.path.join(out_dir, filename)
    doc.save(out_path)
    return out_path, stats_d, tiers


# =====================================================================
# Run against real Trevor Pan project
# =====================================================================
if __name__ == "__main__":
    trevor_pan_pemb = {
        "project_name": "General Metal Construction PEMB",
        "ks_job_number": "26102144",
        "project_type": "PEMB",
        "approx_sf": 12000,
        "jurisdiction": "Phoenix",
        "location": "Phoenix, AZ",
        "scope_description": (
            "Pre-engineered metal building (PEMB) project including structural engineering, "
            "drafting, and construction administration for the primary steel frame, secondary "
            "members, and foundation design."
        ),
        "partner": "CV",
        "client_company_name": "Trevor Pan Architects",
        "engineering_status": "Proposal Sent",
    }
    out_dir = "/tmp/fee_memo_test_output"
    os.makedirs(out_dir, exist_ok=True)
    out_path, stats_d, tiers = render_fee_memo(trevor_pan_pemb, out_dir)
    print(f"Rendered: {out_path}")
    print(f"\nStats: {stats_d}")
    print(f"\nTiers: {tiers}")
