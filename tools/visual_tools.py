"""
Visual inspection tools for PowerPoint MCP Server.
Renders the deck and has a configured external vision LLM review it for
template fidelity and visible errors, so the calling agent can iterate.
"""
import io
import os
from typing import Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations


def register_visual_tools(app: FastMCP, presentations):
    """Register visual inspection tools with the FastMCP app"""

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
        for template/brand fidelity and visible errors (overflowing or
        clipped text, overlaps, empty placeholders, broken charts, etc.).

        Use this after building or editing a deck. If the result has
        "passed": false, fix the reported issues with the editing tools and
        call this again — LOOP until "passed" is true before exporting.

        reference_presentation_id: optional second presentation (e.g. the
        original template opened as its own presentation) whose rendered
        slides are shown to the reviewer as the fidelity reference.
        focus: optional extra instruction for the reviewer.
        """
        from visual_qa import (
            VisionLLM, VisualQAError, render_pptx_bytes_to_pngs, review_prompt,
        )

        if presentation_id not in presentations:
            return {
                "error": "Unknown or expired presentation_id. Pass the presentation_id returned by create_presentation, create_presentation_from_template, or open_presentation"
            }
        try:
            llm = VisionLLM()
        except VisualQAError as e:
            return {"error": str(e)}

        max_slides = int(os.environ.get("VISION_LLM_MAX_SLIDES", "15"))

        def render(pres_id):
            buf = io.BytesIO()
            presentations[pres_id].save(buf)
            return render_pptx_bytes_to_pngs(buf.getvalue(), max_slides=max_slides)

        try:
            deck_images = render(presentation_id)
            ref_images = []
            if reference_presentation_id is not None:
                if reference_presentation_id not in presentations:
                    return {"error": "Unknown or expired reference_presentation_id."}
                ref_images = render(reference_presentation_id)

            prompt = review_prompt(bool(ref_images), focus)
            if ref_images:
                prompt += (
                    f"\nImage order: images 1-{len(ref_images)} are the "
                    f"reference template; images {len(ref_images) + 1}-"
                    f"{len(ref_images) + len(deck_images)} are the deck under "
                    "review. Report issue slide numbers relative to the deck "
                    "under review."
                )
            verdict = llm.review(ref_images + deck_images, prompt)
        except VisualQAError as e:
            return {"error": str(e)}
        except Exception as e:
            return {"error": f"Visual inspection failed: {str(e)}"}

        result = {
            "passed": verdict.get("passed"),
            "issues": verdict.get("issues", []),
            "slides_reviewed": len(deck_images),
        }
        for key in ("raw_review", "note"):
            if key in verdict:
                result[key] = verdict[key]
        if result["passed"] is False:
            result["message"] = (
                "Visual inspection found issues. Fix them with the editing "
                "tools, then call visual_inspect_presentation again; repeat "
                "until passed is true before exporting."
            )
        elif result["passed"] is True:
            result["message"] = "Visual inspection passed."
        return result
