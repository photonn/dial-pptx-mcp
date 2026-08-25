"""
Structural validation tool: is the deck well-formed, and are its slides sane?

This is the axis visual QA cannot see. A deck with a dangling relationship or a
chart with no series renders in LibreOffice and opens in python-pptx — the two
things the visual pass relies on — and still arrives broken on the user's
machine. The checks live in deck_validation.py; this module is the MCP surface.
"""
from typing import Dict

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.validation")

UNKNOWN_ID = (
    "Unknown or expired presentation_id. Pass the presentation_id returned "
    "by create_presentation, create_presentation_from_template, or "
    "open_presentation"
)

SEVERITIES = ("error", "warning", "info")


def register_validation_tools(app: FastMCP, presentations):
    """Register the structural validation tool."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Validate Presentation Structure",
            readOnlyHint=True,
        ),
    )
    def validate_presentation(presentation_id: str,
                              min_severity: str = "warning") -> Dict:
        """Check the deck for structural faults, without rendering it.

        Run this before export. It is fast — no rendering, no model call — and
        it catches the class of defect visual inspection cannot: relationships
        that do not resolve, charts with no data, shapes off the canvas,
        pictures stretched out of proportion, leftover template placeholder
        text, notes shared between two slides.

        It does NOT judge appearance. Overflowing text, overlaps and spacing
        are visual_inspect_slides' job; the two are complementary and a
        finished deck deserves both.

        min_severity: "error" reports only faults that can break the file in
          PowerPoint; "warning" (default) adds the defects a user would
          notice; "info" adds advisories, notably fonts this server's renderer
          substitutes with different metrics — which is what makes a
          text-fit finding from visual QA approximate.

        Returns "ok" (no errors), per-severity "counts", and a "problems" list
        where each entry names the slide, the shape, what is wrong and the
        tool that fixes it.
        """
        import deck_validation

        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}
        if min_severity not in SEVERITIES:
            return {"error": f"Invalid min_severity: {min_severity}. Must be "
                             f"one of {', '.join(SEVERITIES)}."}

        pres = presentations[presentation_id]
        try:
            report = deck_validation.validate_presentation(pres)
        except Exception as e:
            logger.error("validation_failed presentation_id=%s error=%s",
                         short_id(presentation_id), e)
            return {"error": f"Failed to validate the presentation: {e}"}

        cutoff = SEVERITIES.index(min_severity)
        report["problems"] = [p for p in report["problems"]
                              if SEVERITIES.index(p["severity"]) <= cutoff]
        report["min_severity"] = min_severity

        counts = report["counts"]
        logger.info("validation_done presentation_id=%s slides=%d errors=%d "
                    "warnings=%d", short_id(presentation_id), report["slides"],
                    counts["error"], counts["warning"])

        if not report["problems"]:
            report["message"] = (
                f"No problems at severity '{min_severity}' or above. The deck "
                "is structurally sound; appearance is a separate question — "
                "use visual_inspect_slides for that."
            )
        elif counts["error"]:
            report["message"] = (
                f"{counts['error']} error(s) and {counts['warning']} "
                "warning(s). The errors are faults PowerPoint may refuse to "
                "open — fix them before exporting, following each problem's "
                "\"fix\"."
            )
        else:
            report["message"] = (
                f"{counts['warning']} warning(s) and no errors. The file is "
                "valid; each warning names a defect the user would notice and "
                "the tool that resolves it."
            )
        return report
