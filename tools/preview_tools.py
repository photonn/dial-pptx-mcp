"""
Slide previews: see the template before building on it.

The metadata tools report layout names and indices, which is not enough to
choose between the eight near-identical layouts a corporate template ships.
This tool renders the deck into labelled contact sheets and, when a vision model
is configured, describes each slide's structure back to the caller — the agent
cannot look at an image, so the description is the part it can act on, and the
uploaded sheet is for the person.

`render_deck_summary_card` is the other end of the same pipeline: one image of
the finished deck, for showing the user what was built without making them open
the file.

Rendering needs LibreOffice, the same dependency visual QA has, so the tools are
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
DEFAULT_CARD_NAME = "deck-summary.jpg"


def register_preview_tools(app: FastMCP, presentations):
    """Register the preview tools, if slides can be rendered at all."""
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


    @app.tool(
        annotations=ToolAnnotations(
            title="Render Deck Summary Card",
            readOnlyHint=True,
        ),
    )
    def render_deck_summary_card(
        presentation_id: str,
        title: Optional[str] = None,
        columns: Optional[int] = None,
        filename: str = DEFAULT_CARD_NAME,
    ) -> Dict:
        """Render the whole deck as ONE image, to show the user what you built.

        Call this at the end of the job, after export_presentation, and put
        the returned image_url in your final answer as an attachment
        alongside the .pptx. It tiles every slide into a single labelled
        card, so the user can see the finished deck at a glance without
        opening the file.

        This is for the user, not for you: it is one image and you cannot
        see it. To check your own work use visual_inspect_slides, and to
        choose a template slide use render_slide_previews.

        title: caption for the card's header bar — use the deck's title.
          Omit for no header.
        columns: cells per row; omit to fit the grid to the deck's length.
        filename: name of the stored image.

        Returns "image_url" (a DIAL file URL), its mime_type and pixel size.
        """
        import previews
        from dial_client import DialFileClient

        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}
        pres = presentations[presentation_id]

        if not len(pres.slides):
            return {"error": "This presentation has no slides to summarize."}
        if columns is not None and not 1 <= columns <= 8:
            return {"error": "columns must be between 1 and 8, or omitted to "
                             "fit the grid to the deck."}
        if not filename.strip():
            return {"error": "filename must not be empty."}

        try:
            card, size, used_columns, count = previews.render_summary_card(
                pres, title, columns)
        except visual_qa.VisualQAError as e:
            logger.warning("summary_card_render_failed presentation_id=%s "
                           "error=%s", short_id(presentation_id), e)
            return {"error": f"Could not render the slides: {e}"}

        # Unlike render_slide_previews, the stored image is the entire
        # deliverable here — there is nothing left to return without it.
        try:
            client = DialFileClient()
            url = client.upload(card, filename,
                                content_type=previews.JPEG_MIME)
        except Exception as e:
            logger.warning("summary_card_upload_failed presentation_id=%s "
                           "error=%s", short_id(presentation_id), e)
            return {"error": f"The summary card was rendered but could not be "
                             f"stored in DIAL ({e}), so there is no image to "
                             f"show the user. Deliver the exported deck "
                             f"without it."}

        logger.info("summary_card_rendered presentation_id=%s slides=%d "
                    "columns=%d size=%dx%d bytes=%d",
                    short_id(presentation_id), count, used_columns,
                    size[0], size[1], len(card))
        return {
            "message": f"Deck summary card stored at {url}. Include this "
                       "image URL in your final answer as an attachment so "
                       "the user can see the whole deck at a glance.",
            "image_url": url,
            "mime_type": previews.JPEG_MIME,
            "size_bytes": len(card),
            "image_size": {"width": size[0], "height": size[1]},
            "slide_count": count,
            "columns": used_columns,
        }
