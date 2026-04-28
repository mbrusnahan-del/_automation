"""
One-shot patch script: copy the page-2+ header from the 'typical' template
into contract_template_merge_ready.docx, normalizing the legacy uppercase
placeholders ({{KS PROJECT NUMBER}}, {{PROJECT NAME-CAPITALIZED}}) to the
merge-system-friendly snake_case names ({{ks_job_number}}, {{project_name_upper}})
that fill_placeholders() in render_from_notion.py recognizes.

Run once. Backs up the original template to contract_template_merge_ready.bak.docx.
"""
from __future__ import annotations
import os
import re
import shutil
import sys
import zipfile

_HERE = os.path.dirname(os.path.abspath(__file__))

MERGE_READY = os.path.join(_HERE, "contract_template_merge_ready.docx")
TYPICAL_DEFAULT = os.path.join(
    _HERE, "..", "..", "_Projects", "2026",
    "26113149 - Blue Wash Road Residential Remodel - Cave Creek, Arizona",
    "Contracts", "NOT USED",
    "contract template std - 26xxxx KS Structural Contract - CLIENT NAME - JOB NAME - 26.xx.xx.docx",
)

# Allow override on the command line if user wants to point at a different
# typical template
TYPICAL_OVERRIDE = sys.argv[1] if len(sys.argv) > 1 else None


def normalize_header_xml(xml_bytes: bytes) -> bytes:
    """Take the typical template's header1.xml and fix its placeholders.

    Two transformations:

    1. {{KS PROJECT NUMBER}} (single run) → {{ks_job_number}}
       Simple string substitution.

    2. {{</w:t>...PROJECT NAME...-CAPITALIZED}} (split across 3 runs) →
       collapse to a single run holding {{project_name_upper}}.
       We use a regex over the raw XML to spot the broken trio and
       rewrite as a single <w:t> body.
    """
    text = xml_bytes.decode("utf-8")

    # Fix 1 — simple single-run placeholder
    text = text.replace("{{KS PROJECT NUMBER}}", "{{ks_job_number}}")

    # Fix 2 — collapse the 3-run {{...PROJECT NAME...-CAPITALIZED}} pattern.
    # The source XML splits the placeholder across three sibling runs:
    #   <w:r>...<w:t>{{</w:t></w:r>
    #   <w:r>...<w:t>PROJECT NAME</w:t></w:r>
    #   <w:r>...<w:t>-CAPITALIZED}}</w:t></w:r>
    # We match from `{{` through the closing `</w:t></w:r>` of the third run
    # and rewrite as if the whole token lived in the first run's <w:t>.
    pattern = re.compile(
        r"\{\{</w:t></w:r>"
        r"<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>PROJECT NAME</w:t></w:r>"
        r"<w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>-CAPITALIZED\}\}</w:t></w:r>",
        re.DOTALL,
    )
    # Replacement keeps the first run intact: closes its <w:t> and </w:r>
    # so the XML stays balanced after collapse.
    text, n_collapsed = pattern.subn("{{project_name_upper}}</w:t></w:r>", text)
    print(f"  Collapsed {n_collapsed} split PROJECT NAME-CAPITALIZED placeholder(s)")

    # Drop any remaining "yellow" highlights from the header — those came
    # from the source template marking the placeholders. After substitution
    # we want the rendered text to look clean (the renderer's
    # clear_highlight_on_resolved_runs() only walks the body, so this strips
    # them at template level).
    text = re.sub(r'<w:highlight w:val="yellow"/>', '', text)

    return text.encode("utf-8")


def main() -> int:
    typical = TYPICAL_OVERRIDE or TYPICAL_DEFAULT
    if not os.path.isfile(typical):
        # Fallback: probe a couple of likely OneDrive locations
        candidates = [
            os.path.join(_HERE, "contract template std - 26xxxx KS Structural Contract - CLIENT NAME - JOB NAME - 26.xx.xx.docx"),
            "/sessions/nice-compassionate-pasteur/mnt/uploads/contract template std - 26xxxx KS Structural Contract - CLIENT NAME - JOB NAME - 26.xx.xx.docx",
        ]
        for c in candidates:
            if os.path.isfile(c):
                typical = c
                break
        else:
            print(f"ERROR: typical template not found. Looked at:\n  {typical}", file=sys.stderr)
            for c in candidates:
                print(f"  {c}", file=sys.stderr)
            return 1
    if not os.path.isfile(MERGE_READY):
        print(f"ERROR: merge-ready template not found: {MERGE_READY}", file=sys.stderr)
        return 1

    print(f"Typical template:    {typical}")
    print(f"Merge-ready target:  {MERGE_READY}")

    # Backup the merge-ready template
    backup = MERGE_READY.replace(".docx", ".bak.docx")
    if not os.path.isfile(backup):
        shutil.copy2(MERGE_READY, backup)
        print(f"Backed up to:        {backup}")
    else:
        print(f"(Backup already exists: {backup})")

    # Read typical's header1.xml
    with zipfile.ZipFile(typical, "r") as z:
        if "word/header1.xml" not in z.namelist():
            print("ERROR: typical template has no word/header1.xml", file=sys.stderr)
            return 1
        header_xml = z.read("word/header1.xml")

    print(f"\nOriginal typical header1.xml: {len(header_xml)} bytes")

    # Normalize placeholders
    cleaned = normalize_header_xml(header_xml)
    print(f"Normalized header1.xml:        {len(cleaned)} bytes")

    # Write into merge-ready (recreating the zip — the only way to update
    # an entry inside an Office Open XML package without corrupting it)
    tmp_out = MERGE_READY + ".new"
    with zipfile.ZipFile(MERGE_READY, "r") as zin:
        with zipfile.ZipFile(tmp_out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/header1.xml":
                    zout.writestr(item, cleaned)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(tmp_out, MERGE_READY)
    print(f"\n✓ Patched: {MERGE_READY}")

    # Verify by reading back
    with zipfile.ZipFile(MERGE_READY, "r") as z:
        verify = z.read("word/header1.xml").decode("utf-8")
    placeholders = re.findall(r"\{\{[a-z_]+\}\}", verify)
    print(f"\nFinal header1.xml placeholders: {placeholders}")
    if "{{ks_job_number}}" in placeholders and "{{project_name_upper}}" in placeholders:
        print("✓ Both expected placeholders present and in single-run form.")
        return 0
    else:
        print("⚠ Missing one or both expected placeholders. Inspect manually.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
