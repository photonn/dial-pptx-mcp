"""
Slide previews: see the template before building on it.

The metadata tools report layout names and indices, which is not enough to
choose between the eight near-identical layouts a corporate template ships.
This tool renders the deck into labelled contact sheets and, when a vision model
is configured, describes each slide's structure back to the caller — the agent
cannot look at an image, so the description is the part it can act on, and the
uploaded sheet is for the person.

Rendering needs LibreOffice, the same dependency visual QA has, so the tool is
registered only when the renderer is present.
"""
from typing import Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.preview")

UNKNOWN_ID = (
    "Unknown or expired presentation_id. Pass the presentation_id returned "
    "by create_presentation, create_presentation_from_template, or "
    "open_presentation"
)

MAX_PREVIEW_SLIDES = 24


def register_preview_tools(app: FastMCP, presentations):
    """Register the slide-preview tool, if slides can be rendered at all."""
    import visual_qa

    try:
        visual_qa._soffice_binary()
    except visual_qa.VisualQAError:
        logger.info("preview_tools_hidden reason=libreoffice_not_found")
        return
    logger.info("preview_tools_registered describe=%s",
                visual_qa.enforcement_enabled())

    @app.tool(
        annotations=ToolAnnotations(
            title="Render Slide Previews",
            readOnlyHint=True,
        ),
    )
    def render_slide_previews(
        presentation_id: str,
        slides: Optional[List[int]] = None,
        describe: bool = True,
        columns: int = 4,
    ) -> Dict:
        """Render the deck's slides as a labelled grid, to see what you have.

        Use this right after opening a corporate template, before planning:
        it tells you which of the template's slides exist and what each one is
        structurally for, so you can map your content onto real slides and
        duplicate_slide the right ones. Layout names alone do not tell you
        where a template puts its content.

        slides: 1-based slide numbers to preview; omit for the whole deck
          (capped at 24 — preview a range if the deck is longer).
        describe: ask the vision model what each slide is structurally suited
          to, returned as "slide_descriptions". This is the part you can act
          on; without it you get only the image URL, which you cannot see.
        columns: cells per row in the grid image.

        Returns "contact_sheet_urls" (DIAL file URLs — offer them to the user
        if they want to look at the template) and, with describe, one entry
        per slide naming what it holds and what to reuse it for.
        """
        import previews
        from dial_client import DialFileClient

        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}
        pres = presentations[presentation_id]

        try:
            wanted = visual_qa.normalize_slides(pres, slides)
        except ValueError as e:
            return {"error": str(e)}
        if not len(pres.slides):
            return {"error": "This presentation has no slides to preview."}
        if wanted is None and len(pres.slides) > MAX_PREVIEW_SLIDES:
            wanted = list(range(1, MAX_PREVIEW_SLIDES + 1))
            truncated = len(pres.slides)
        else:
            truncated = None
        if columns < 1 or columns > 6:
            return {"error": "columns must be between 1 and 6"}

        try:
            sheets, images, numbers = previews.render_contact_sheets(
                pres, wanted, columns)
        except visual_qa.VisualQAError as e:
            logger.warning("preview_render_failed presentation_id=%s error=%s",
                           short_id(presentation_id), e)
            return {"error": f"Could not render the slides: {e}"}

        result = {
            "slides_previewed": numbers,
            "slide_count": len(pres.slides),
        }
        if truncated:
            result["note"] = (
                f"This deck has {truncated} slides; only the first "
                f"{MAX_PREVIEW_SLIDES} were previewed. Pass `slides` to see "
                f"a specific range."
            )

        # The sheet is for the user, so a failure to upload must not lose the
        # descriptions — which are what the agent came for.
        try:
            client = DialFileClient()
            result["contact_sheet_urls"] = [
                client.upload(sheet, f"slide-previews-{index + 1}.jpg",
                              content_type=previews.JPEG_MIME)
                for index, sheet in enumerate(sheets)
            ]
        except Exception as e:
            logger.warning("preview_upload_failed presentation_id=%s error=%s",
                           short_id(presentation_id), e)
            result["contact_sheet_note"] = (
                f"The preview images could not be stored in DIAL ({e}), so "
                "there is no URL to show the user. The descriptions below are "
                "unaffected." if describe else
                f"The preview images could not be stored in DIAL ({e})."
            )

        if describe:
            if not visual_qa.enforcement_enabled():
                result["describe_note"] = (
                    "No vision model is configured on this server, so the "
                    "slides could not be described. Choose a slide from the "
                    "contact sheet with the user, or from "
                    "get_presentation_info."
                )
            else:
                try:
                    described = previews.describe_slides(images, numbers)
                except Exception as e:
                    logger.warning("preview_describe_failed presentation_id=%s "
                                   "error=%s", short_id(presentation_id), e)
                    described = None
                if described:
                    result["slide_descriptions"] = described
                else:
                    result["describe_note"] = (
                        "The vision model did not return usable descriptions. "
                        "Retry, or work from get_presentation_info."
                    )

        logger.info("previews_rendered presentation_id=%s slides=%d sheets=%d "
                    "described=%s", short_id(presentation_id), len(numbers),
                    len(sheets), "slide_descriptions" in result)
        return result
