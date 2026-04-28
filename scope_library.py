"""
Scope library — canonical SCOPE OF SERVICES bullets per Project Type.

The render script merges three sources to build the SCOPE OF SERVICES section:
  1. Standard bullets from this library (keyed by Project Type)
  2. Project-specific bullets parsed from Notion's Scope Description field
  3. The catch-all "Tasks not listed..." bullet (always appended last)

Partners: Edit this file directly to improve bullet language.  Add/remove lines
freely. The render picks up the latest version at runtime. Keep bullets short
and specific — long prose should go in Scope Description on the Notion project.

Each list is ordered from most-central (top) to more-specific (bottom).
"""

# =============================================================================
# STANDARD BULLETS BY PROJECT TYPE
# =============================================================================

STANDARD_BULLETS = {
    "PEMB": [
        "Primary steel frame design (columns, rafters, bracing) coordinated with PEMB manufacturer-supplied loads",
        "Secondary member design (girts, purlins, bridging) where not provided by PEMB manufacturer",
        "Foundation design including spread footings, slab-on-grade, and anchor bolt layout",
        "Lateral analysis for wind and seismic loads per governing code",
        "Review of and coordination with PEMB manufacturer shop drawings",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Site-specific geotechnical report is NOT included — Client to provide if required by jurisdiction",
    ],
    "Commercial New": [
        "Structural design of foundation, slab-on-grade, and retaining walls within the building footprint",
        "Structural design of the primary vertical and lateral load-resisting system (framing, bearing walls, braced frames, shear walls)",
        "Roof framing design including gravity and diaphragm systems",
        "Lateral analysis for wind and seismic loads per governing code",
        "Coordination with architectural, MEP, and civil disciplines through one full design cycle",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Site retaining walls outside the building footprint are NOT included unless specifically listed in the fee schedule",
    ],
    "Commercial Remodel": [
        "Evaluation of existing structure based on available as-built drawings and site observation",
        "Structural design of modifications to existing framing (new openings, added loads, reinforcing)",
        "Foundation review for added loads or new equipment where applicable",
        "Lateral compliance review per current code for the modified portions",
        "Coordination with architectural, MEP, and civil disciplines for the modified scope",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Full building seismic retrofit is NOT included unless specifically listed in the fee schedule",
    ],
    "Tenant Improvement": [
        "Evaluation of existing structure relative to proposed interior modifications",
        "Structural design of new framing for added mezzanines, equipment platforms, or interior walls",
        "Review of existing roof or floor framing capacity for new mechanical, electrical, or suspended loads",
        "Design of framing for new openings, stair wells, or penetrations in existing structure",
        "Coordination with architectural and MEP disciplines for the TI scope",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Evaluation of the base building's lateral system capacity is NOT included unless specifically listed",
    ],
    "New Residence": [
        "Structural design of foundation (spread footings, slab-on-grade, or post-tension slab as required)",
        "Structural design of wood-framed floor and roof framing",
        "Lateral design including braced walls and/or shear walls per governing code",
        "Design of structural steel components where architecturally required (lintels, moment frames, stair stringers)",
        "Coordination with architectural drawings through one full design cycle",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Site retaining walls outside the building footprint are NOT included unless specifically listed in the fee schedule",
    ],
    "Residential Remodel": [
        "Evaluation of existing structure based on available as-built drawings and site observation",
        "Structural design of modifications to existing framing (new openings, load path revisions, reinforcing)",
        "Foundation review and design for additions where applicable",
        "Lateral compliance review for the modified portions per current code",
        "Coordination with architectural drawings through one full design cycle",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Full-building seismic retrofit of unmodified portions is NOT included unless specifically listed",
    ],
    "Multifamily": [
        "Structural design of foundation and slab-on-grade for the building footprint",
        "Structural design of the primary vertical and lateral load-resisting system (framing, shear walls, braced frames)",
        "Roof and floor framing design including gravity and diaphragm systems",
        "Lateral analysis for wind and seismic loads per governing code",
        "Coordination with architectural, MEP, and civil disciplines through one full design cycle",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Site retaining walls outside the building footprint are NOT included unless specifically listed in the fee schedule",
    ],
    "Retrofit": [
        "Evaluation of existing structure based on available as-built drawings and site observation",
        "Structural analysis of existing vertical and lateral systems per current code",
        "Design of retrofit elements (new braced frames, shear walls, chord/drag members, foundation reinforcing)",
        "Coordination with architectural and MEP disciplines for the retrofit scope",
        "Response to one round of plan-check comments from the jurisdiction",
        "Sealed structural drawings, calculations, and submittal package",
        "Demolition means and methods are NOT included — Client's contractor to develop",
        "Historic preservation documentation is NOT included unless specifically listed",
    ],
    "Structural Observation": [
        "Visual site observation of the structural system at the stage(s) requested by Client",
        "Written summary report documenting observed conditions and any concerns noted",
        "Digital photographs of observed conditions included in the summary report",
        "One round of clarification correspondence following report delivery",
        "Structural calculations are NOT included unless specifically listed in the fee schedule",
        "Destructive investigation and/or material testing is NOT included",
        "Sealed drawings are NOT included under this scope",
    ],
    "Peer Review": [
        "Independent structural peer review of drawings and calculations provided by Client",
        "Written summary report identifying code-compliance concerns, design gaps, or recommended revisions",
        "One round of review comments, provided in writing",
        "One follow-up conference call or meeting to discuss comments",
        "Original structural design work is NOT included — review only",
        "Response to subsequent revisions beyond the first submittal is NOT included unless specifically listed",
    ],
    "Feasibility": [
        "Preliminary structural feasibility study based on proposed concept documents",
        "Evaluation of structural options and constructibility considerations",
        "Written feasibility report summarizing findings and recommendations",
        "One conference call or meeting to discuss findings",
        "Sealed drawings and final calculations are NOT included under this scope",
        "Plan-check submittal is NOT included",
    ],
    "Other": [
        "Structural engineering services as described in the Scope Description provided by Client",
        "Response to one round of plan-check comments from the jurisdiction (where applicable)",
        "Sealed structural deliverables as required by the project scope",
    ],
}


# =============================================================================
# CONTEXT-DRIVEN MODIFIERS
# =============================================================================
# Bullets added conditionally based on project attributes. Each entry is a
# (condition_fn, bullet_text) tuple. condition_fn takes the project dict and
# returns True if the bullet should be appended.

CONDITIONAL_BULLETS = [
    # If Reimbursables Treatment = Digital Only, flag it
    (lambda p: (p.get("Reimbursables Treatment") or "").lower() == "digital only",
     "All deliverables to be transmitted digitally — no printed documents provided"),
    # Hint for projects with large SF — multi-phase submittal
    (lambda p: (p.get("Approx. Structural SF") or 0) >= 20000,
     "One submittal package is anticipated; additional phasing submittals are NOT included unless specifically listed"),
]


# =============================================================================
# CATCH-ALL — ALWAYS APPENDED LAST
# =============================================================================
# Retained per client requirement: if it's not specified, it's not in scope at
# the quoted fee. Belt-and-suspenders protection alongside specific exclusions.

CATCH_ALL_BULLET = "Tasks not listed in this scope of work are excluded"


# =============================================================================
# CODE/JURISDICTION BULLET (kept from original template)
# =============================================================================
# This one is always first in the scope list because it grounds the submittal.

def code_jurisdiction_bullet(icc_year, jurisdiction):
    return (
        f"One submittal package is anticipated designed under the {icc_year} ICC Codes "
        f"with City of {jurisdiction} amendments"
    )


def build_scope_bullets(project):
    """
    Build the full ordered list of scope bullets for a project dict.
    Returns list of strings.
    """
    ptype = (project.get("Project Type") or "").strip() or "Other"
    icc = project.get("ICC Code Year") or "current"
    juris = project.get("Jurisdiction") or "the governing jurisdiction"

    bullets = []

    # 1. Code/jurisdiction anchor
    bullets.append(code_jurisdiction_bullet(icc, juris))

    # 2. Standard bullets for the project type
    bullets.extend(STANDARD_BULLETS.get(ptype, STANDARD_BULLETS["Other"]))

    # 3. Project-specific bullets parsed from Scope Description
    #    Break Scope Description into sentences/phrases that look scope-like
    extra = parse_scope_description(project.get("Scope Description") or "")
    bullets.extend(extra)

    # 4. Conditional bullets
    for cond, text in CONDITIONAL_BULLETS:
        if cond(project):
            bullets.append(text)

    # 5. Catch-all (always last)
    bullets.append(CATCH_ALL_BULLET)

    # Deduplicate while preserving order
    seen = set()
    unique = []
    for b in bullets:
        key = b.lower().strip()
        if key not in seen:
            seen.add(key)
            unique.append(b)
    return unique


def parse_scope_description(desc):
    """
    Very light heuristic: split description into sentences, filter to ones that
    look like scope items (contain nouns like 'design', 'structure', 'frame',
    dimensions, etc.), strip numbering, and return as bullets.

    Intentionally conservative — partners can always add scope in Notion with
    clear, bulletable phrasing.
    """
    if not desc or len(desc) < 30:
        return []
    import re
    # Split on sentence boundaries
    sentences = re.split(r'(?<=[\.\?\!])\s+(?=[A-Z])', desc.strip())
    bullets = []
    for s in sentences:
        s = s.strip()
        # Strip numbered list prefixes
        s = re.sub(r'^\s*\d+\.\s*', '', s)
        # Skip chitchat / questions
        if len(s) < 25 or s.endswith('?'):
            continue
        if any(w in s.lower() for w in (
            "thanks", "let me know", "questions?", "fee in email", "so what",
            "hi jonathan", "hi chris", "hi griffin",
        )):
            continue
        # Cap length per bullet
        if len(s) > 200:
            s = s[:200].rstrip() + "…"
        bullets.append(s)
    # Cap to a reasonable count to avoid scope bloat
    return bullets[:6]
