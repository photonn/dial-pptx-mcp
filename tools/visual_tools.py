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
        """Report what is wrong with slides, without changing anything.
        Prefer visual_repair_slides, which runs the same review and then
        fixes what it can.

        The review checklist is complete and fixed server-side: blank or
        misplaced content, hidden text, overflow and overlap (text, charts,
        tables, diagrams), legibility, broken characters, charts, pictures,
        alignment, brand fidelity — and, for the whole deck, the deck's
        story: agenda vs. actual sections, content on the wrong slide,
        contradictory figures, order.

        slides: 1-based slide numbers to review, e.g. [3] or [1,2,3]. Omit to
        review the whole deck (the only scope that runs the story review).
        focus: optional extra instruction, added on top of the checklist —
        leave it out unless the user asked for something specific.
        reference_presentation_id: a template deck to compare against for
        brand fidelity.

        Returns {"passed", "issues": [{"slide", "severity", "category",
        "element", "description", "suggested_fix"}], "slides_reviewed",
        "checks"}. Slide numbers are absolute deck positions.
        """
        pres, numbers, error = _resolve(presentation_id, slides)
        if error:
            return error
        reference = None
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
        if verdict.get("passed") is True and not numbers \
                and not verdict.get("slides_not_reviewed"):
            presentations.clear_dirty(presentation_id)
        verdict["scope"] = numbers or "deck"
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
        """Review slides and fix what the review finds, then re-review —
        repeating until they pass or the iteration budget runs out.

        The review checklist is complete and fixed server-side (see
        visual_inspect_slides); you do not need to say what to look for.
        Called without `slides` it also reviews the deck's story — agenda
        vs. actual sections, content built on the wrong slide, hidden text,
        contradictory figures — and can fix those too: repairs move shapes
        between slides, restack, delete, clear or rewrite text, resize,
        refit, recolour, fix tables and charts, and reorder slides, all
        through validated operations. It never invents content.

        slides: 1-based slide numbers to repair; omit for the whole deck
        (recommended once the deck is built). A scoped call never touches
        slides outside `slides`.
        focus: optional extra instruction, added on top of the checklist.
        max_iterations: reviews for this call, including the first (default
        VISUAL_QA_MAX_ITERATIONS, normally 3); values below 2 are raised to 2.

        A round that makes a slide worse is undone on that slide (the round
        then lists it under "reverted_slides"), so the deck you get back is
        the best version of each slide, not the last attempt.

        Returns {"passed", "iterations", "repair_rounds" (each with the
        "changes" it made), "issues" (blocking issues left), "minor_issues",
        "action_required" (what only you can fix — usually missing content:
        build it with the editing tools, then call this again on those
        slides), "slides_reviewed", "checks"}. A false "passed" is a report,
        not a request to retry the same call.
        """
        pres, numbers, error = _resolve(presentation_id, slides)
        if error:
            return error

        logger.info("visual_repair_start presentation_id=%s scope=%s",
                    short_id(presentation_id), _scope(numbers))
        try:
            outcome = visual_qa.inspect_and_repair(pres, numbers, focus,
                                                   max_iterations)
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

        if outcome.get("passed") and not numbers \
                and not outcome.get("slides_not_reviewed"):
            presentations.clear_dirty(presentation_id)
        outcome["scope"] = numbers or "deck"
        return outcome
