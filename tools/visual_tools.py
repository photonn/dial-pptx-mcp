"""
Visual QA tools for the PowerPoint MCP Server.

Visual QA is orchestrator-driven: the calling agent inspects and repairs
slides whenever it wants — typically right after building each slide, not
only at the end — instead of the server running one hidden deck-wide pass
inside export.

Two tools, both scoped to an optional slide selection:
- visual_inspect_slides  render + vision review, report issues (read-only)
- visual_repair_slides   review, then let the server fix the reported issues
                         with whitelisted python-pptx operations and re-review

Both are registered whenever the vision LLM is configured
(visual_qa.enforcement_enabled). Operators who additionally want export to
refuse an unverified deck set VISUAL_QA_EXPORT_GATE=true; see
tools/presentation_tools.py.

Where a deck has brand rules attached (attach_brand_profile), both tools carry
the profile's `review_notes` into every review as standing focus, and its
reference deck is shown to the reviewer as the template to match. That is where
the brand rules no measurement can express live — "a headline must state a
message", "no slide may be plain text on white" — and they belong here rather
than only on the export path, because this is where the agent can still fix the
slide it just built.
"""
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.visual")

UNKNOWN_ID = (
    "Unknown or expired presentation_id. Pass the presentation_id returned "
    "by create_presentation, create_presentation_from_template, or "
    "open_presentation"
)


def _scope(slides):
    return "deck" if not slides else ",".join(str(n) for n in slides)


def _brand_context(presentations, presentation_id, focus):
    """-> (focus, reference_pres, note) from the deck's attached brand rules.

    The caller's own instruction comes first: it is about the slide in front
    of it, the profile's notes are standing rules.

    A deck with no rules attached is reviewed without them and told so, rather
    than refused: the findings are worth having either way, and a build that
    never called attach_brand_profile should not lose its QA pass over it.
    """
    import brand_validation

    context = presentations.brand_for(presentation_id) or {}
    profile = context.get("profile")
    if not profile:
        if not brand_validation.enabled():
            return focus, None, None
        from tools.brand_tools import not_attached_error

        return focus, None, not_attached_error(
            brand_validation.profile_file_name()) + (
            " The slides were reviewed without the brand rules.")

    notes = brand_validation.review_focus(profile)
    if notes:
        focus = f"{focus.strip()} {notes}" if focus and focus.strip() else notes
    return focus, context.get("reference"), None


def register_visual_tools(app: FastMCP, presentations):
    """Register the visual QA tools, if the vision LLM is configured."""
    import visual_qa

    if not visual_qa.enforcement_enabled():
        logger.info("visual_qa_tools_hidden reason=vision_llm_not_configured")
        return
    logger.info("visual_qa_tools_registered export_gate=%s",
                visual_qa.export_gate_enabled())

    def _resolve(presentation_id, slides):
        """-> (pres, normalized_slides, error_dict)."""
        if presentation_id not in presentations:
            return None, None, {"error": UNKNOWN_ID}
        pres = presentations[presentation_id]
        try:
            return pres, visual_qa.normalize_slides(pres, slides), None
        except ValueError as e:
            return None, None, {"error": str(e)}

    @app.tool(
        annotations=ToolAnnotations(
            title="Visually Inspect Slides",
            readOnlyHint=True,
        ),
    )
    def visual_inspect_slides(
        presentation_id: str,
        slides: Optional[List[int]] = None,
        focus: Optional[str] = None,
        reference_presentation_id: Optional[str] = None,
    ) -> Dict:
        """Render slides and have a vision model review them for template
        fidelity and visible errors (overflowing or clipped text, overlaps,
        elements off the slide, unfilled placeholders, broken charts).

        slides: 1-based slide numbers to review, e.g. [3] or [1,2,3]. Omit to
        review the whole deck. Inspecting the one slide you just built is
        cheap and precise; inspect as often as you like.
        focus: extra instruction for the reviewer, e.g. "check the chart
        labels are legible".
        reference_presentation_id: a template deck to compare against for
        brand fidelity (whole-deck review only). If this server has a brand
        reference deck configured, it is used automatically and you do not
        need to supply one.

        Returns {"passed": bool, "issues": [{"slide", "severity",
        "description", "suggested_fix"}]}. Slide numbers in issues are
        absolute deck positions. This tool only reports — use
        visual_repair_slides, or your own editing tools, to fix what it finds.
        """
        pres, numbers, error = _resolve(presentation_id, slides)
        if error:
            return error
        focus, reference, brand_note = _brand_context(
            presentations, presentation_id, focus)
        if reference_presentation_id is not None:
            if reference_presentation_id not in presentations:
                return {"error": "Unknown or expired reference_presentation_id."}
            reference = presentations[reference_presentation_id]

        try:
            verdict = visual_qa.inspect_presentation(pres, reference, focus,
                                                     numbers)
        except visual_qa.VisualQAError as e:
            logger.error("visual_inspect_failed presentation_id=%s scope=%s "
                         "reason=qa_error error=%s", short_id(presentation_id),
                         _scope(numbers), e)
            return {"error": str(e)}
        except Exception as e:
            logger.error("visual_inspect_failed presentation_id=%s scope=%s "
                         "reason=%s error=%s", short_id(presentation_id),
                         _scope(numbers), type(e).__name__, e)
            return {"error": f"Visual inspection failed: {str(e)}"}

        # Only a clean review of the *whole* deck clears it for export.
        if verdict.get("passed") is True and not numbers:
            presentations.clear_dirty(presentation_id)
        verdict["scope"] = numbers or "deck"
        if brand_note:
            verdict["brand_note"] = brand_note
        return verdict

    @app.tool(
        annotations=ToolAnnotations(
            title="Visually Repair Slides",
        ),
    )
    def visual_repair_slides(
        presentation_id: str,
        slides: Optional[List[int]] = None,
        focus: Optional[str] = None,
        max_iterations: Optional[int] = None,
    ) -> Dict:
        """Inspect the selected slides and fix what the review finds, then
        re-inspect — repeating until they pass or the iteration budget runs
        out. Repairs are made server-side with a fixed set of validated
        operations (move, resize, font size, set text, word wrap, delete) and
        never touch slides outside `slides`.

        slides: 1-based slide numbers to repair; omit for the whole deck.
        focus: extra instruction for the reviewer.
        max_iterations: inspect/repair rounds for this call (default
        VISUAL_QA_MAX_ITERATIONS, normally 10). Lower it for a quick pass on
        a single slide.

        Returns {"passed", "iterations", "repair_rounds", "issues"} — with
        "issues" listing what could not be resolved. A false "passed" is a
        report, not a request to retry the same call: either edit the slide
        content yourself and inspect again, or tell the user what remains.
        """
        pres, numbers, error = _resolve(presentation_id, slides)
        if error:
            return error
        focus, reference, brand_note = _brand_context(
            presentations, presentation_id, focus)

        logger.info("visual_repair_start presentation_id=%s scope=%s",
                    short_id(presentation_id), _scope(numbers))
        try:
            outcome = visual_qa.inspect_and_repair(
                pres, numbers, focus, max_iterations,
                reference_pres=reference)
        except visual_qa.VisualQAError as e:
            logger.error("visual_repair_failed presentation_id=%s scope=%s "
                         "reason=qa_error error=%s", short_id(presentation_id),
                         _scope(numbers), e)
            return {"error": str(e)}
        except Exception as e:
            logger.error("visual_repair_failed presentation_id=%s scope=%s "
                         "reason=%s error=%s", short_id(presentation_id),
                         _scope(numbers), type(e).__name__, e)
            return {"error": f"Visual repair failed: {str(e)}"}

        if outcome.get("passed") and not numbers:
            presentations.clear_dirty(presentation_id)
        outcome["scope"] = numbers or "deck"
        if brand_note:
            outcome["brand_note"] = brand_note
        return outcome
