#!/usr/bin/env python
"""
Template-fidelity spike (workstream 5.1).

Exercises the same code paths the MCP tools use (utils/*) against a supplied
.pptx template, then checks whether the template's theme, slide masters, and
layouts survive the round-trip unchanged.

Usage:
    .venv/bin/python spike/fidelity_spike.py <template.pptx> [output.pptx]

The output deck must additionally be opened in PowerPoint for visual
inspection — the automated checks catch structural drift (lost/changed theme
or master/layout XML), not rendering subtleties.
"""
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import utils as ppt_utils


def part_names(pptx_path, prefix):
    """Real parts under prefix — zip directory entries (trailing /) excluded;
    some authoring tools write them, python-pptx legitimately does not."""
    with zipfile.ZipFile(pptx_path) as z:
        return sorted(n for n in z.namelist()
                      if n.startswith(prefix) and not n.endswith("/"))


def read_part(pptx_path, name):
    with zipfile.ZipFile(pptx_path) as z:
        return z.read(name)


def parts_equal(template, output, name):
    """Semantic comparison: XML parts are compared canonically with
    inter-element whitespace stripped (python-pptx re-serializes XML, so
    byte equality is too strict for real-world templates); other parts
    byte-for-byte."""
    t_raw, o_raw = read_part(template, name), read_part(output, name)
    if not name.endswith((".xml", ".rels")):
        return t_raw == o_raw
    from lxml import etree
    parser = etree.XMLParser(remove_blank_text=True)

    def norm(raw):
        return etree.canonicalize(
            etree.tostring(etree.fromstring(raw, parser)).decode())

    try:
        return norm(t_raw) == norm(o_raw)
    except etree.XMLSyntaxError:
        return t_raw == o_raw


def compare_parts(template, output, prefix, label):
    """Compare a family of OOXML parts between the two files."""
    t_names, o_names = part_names(template, prefix), part_names(output, prefix)
    issues = []
    missing = set(t_names) - set(o_names)
    if missing:
        issues.append(f"{label}: MISSING in output: {sorted(missing)}")
    changed = [
        n for n in t_names
        if n in o_names and not parts_equal(template, output, n)
    ]
    if changed:
        issues.append(f"{label}: CHANGED (beyond whitespace): {changed}")
    ok = f"{label}: {len(t_names)} part(s) preserved" if not issues else None
    return issues, ok


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    template = Path(sys.argv[1])
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else template.with_name(
        template.stem + "_spike_output.pptx")

    # 1. Create from template — same util the create_presentation_from_template tool calls
    pres = ppt_utils.create_presentation_from_template(str(template))
    layouts = ppt_utils.get_slide_layouts(pres)
    print(f"Template: {template.name} — {len(pres.slides)} slide(s), "
          f"{len(layouts)} layout(s) available")

    # 2. Add slides using a handful of existing tool code paths
    slide, _ = ppt_utils.add_slide(pres, layout_index=0)      # title layout
    ppt_utils.set_title(slide, "Fidelity Spike — Title Slide")

    body_layout = 1 if len(layouts) > 1 else 0
    slide2, layout2 = ppt_utils.add_slide(pres, layout_index=body_layout)
    ppt_utils.set_title(slide2, "Bullets")
    for ph in slide2.placeholders:
        if ph.placeholder_format.idx != 0 and ph.has_text_frame:
            ppt_utils.add_bullet_points(ph, ["Point one", "Point two", "Point three"])
            break

    slide3, _ = ppt_utils.add_slide(pres, layout_index=body_layout)
    ppt_utils.set_title(slide3, "Table and textbox")
    ppt_utils.add_table(slide3, rows=3, cols=3, left=1.0, top=2.0, width=6.0, height=3.0)
    ppt_utils.add_textbox(slide3, left=1.0, top=5.5, width=6.0, height=1.0,
                          text="Textbox added by spike")

    slide4, _ = ppt_utils.add_slide(pres, layout_index=body_layout)
    ppt_utils.set_title(slide4, "Chart")
    ppt_utils.add_chart(slide4, "column", left=1.0, top=2.0, width=7.0, height=4.0,
                        categories=["Q1", "Q2", "Q3"],
                        series_names=["Actual", "Plan"],
                        series_values=[[10, 20, 30], [12, 18, 33]])

    # 3. Save
    ppt_utils.save_presentation(pres, str(output))
    print(f"Saved: {output}")

    # 4. Automated structural fidelity checks
    print("\n--- Structural fidelity checks (template vs output) ---")
    all_issues = []
    for prefix, label in [
        ("ppt/theme/", "Theme XML"),
        ("ppt/slideMasters/", "Slide masters"),
        ("ppt/slideLayouts/", "Slide layouts"),
        ("ppt/media/", "Media (logos/branding images)"),
    ]:
        issues, ok = compare_parts(template, output, prefix, label)
        for i in issues:
            print(f"  DRIFT  {i}")
        if ok:
            print(f"  OK     {ok}")
        all_issues.extend(issues)

    print("\nRESULT:", "STRUCTURAL DRIFT DETECTED — see above" if all_issues
          else "no structural drift — now open the output in PowerPoint for visual check")
    return 1 if all_issues else 0


if __name__ == "__main__":
    sys.exit(main())
