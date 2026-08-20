"""
Visual inspection tool registration for PowerPoint MCP Server.

Visual QA normally runs automatically as an export gate (see
tools/presentation_tools.py): when the vision LLM is configured, any deck
edited since its last passed inspection is inspected inside
export_presentation/save_presentation, and a failing deck is refused with
the issue list until the agent fixes it.

The standalone visual_inspect_presentation tool is therefore NOT exposed by
default — set VISUAL_QA_EXPOSE_TOOL=true to add it to the tool list (useful
for debugging or for orchestrators that want mid-build checks).
"""
import os
from typing import Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.visual")


def register_visual_tools(app: FastMCP, presentations):
    """Register visual inspection tools with the FastMCP app (opt-in)."""
    if os.environ.get("VISUAL_QA_EXPOSE_TOOL", "false").lower() != "true":
        logger.debug("visual_inspect_tool_hidden reason=VISUAL_QA_EXPOSE_TOOL_unset")
        return
    logger.info("visual_inspect_tool_exposed reason=VISUAL_QA_EXPOSE_TOOL=true")

    @app.tool(
        annotations=ToolAnnotations(
            title="Visually Inspect Presentation",
            readOnlyHint=True,
        ),
    )
    def visual_inspect_presentation(
        presentation_id: str,
        reference_presentation_id: Optional[str] = None,
        focus: Optional[str] = None,
    ) -> Dict:
        """Render the presentation and have a vision model review every slide
        for template/brand fidelity and visible errors. Note: exports run
        this check automatically; use this tool only for mid-build checks.
        """
        from visual_qa import VisualQAError, inspect_presentation

        if presentation_id not in presentations:
            return {
                "error": "Unknown or expired presentation_id. Pass the presentation_id returned by create_presentation, create_presentation_from_template, or open_presentation"
            }
        reference = None
        if reference_presentation_id is not None:
            if reference_presentation_id not in presentations:
                return {"error": "Unknown or expired reference_presentation_id."}
            reference = presentations[reference_presentation_id]

        try:
            verdict = inspect_presentation(presentations[presentation_id],
                                           reference, focus)
        except VisualQAError as e:
            logger.error("visual_inspect_failed presentation_id=%s reason=qa_error "
                         "error=%s", short_id(presentation_id), e)
            return {"error": str(e)}
        except Exception as e:
            logger.error("visual_inspect_failed presentation_id=%s reason=%s "
                         "error=%s", short_id(presentation_id),
                         type(e).__name__, e)
            return {"error": f"Visual inspection failed: {str(e)}"}

        if verdict.get("passed") is True:
            presentations.clear_dirty(presentation_id)
        return verdict
