"""
Slide structure tools: delete, reorder, duplicate, copy between decks, and
speaker notes.

python-pptx offers none of these, so the deck-building loop upstream supports
is append-only: a slide can be added at the end and filled in, never removed,
moved, or repeated. That is the wrong shape for the corporate-template work
this server exists for, where the template arrives with its designed slides
already in place and the job is to repeat the ones that fit the content and
drop the ones that don't. The package-level work is in utils/slide_utils.py.
"""
from typing import Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

import utils as ppt_utils
from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.slide")

UNKNOWN_ID = (
    "Unknown or expired presentation_id. Pass the presentation_id returned "
    "by create_presentation, create_presentation_from_template, or "
    "open_presentation"
)

NOTES_OPERATIONS = ("get", "set", "clear")


def register_slide_tools(app: FastMCP, presentations):
    """Register the slide-structure and speaker-notes tools."""

    def _resolve(presentation_id, slide_index=None, what="slide_index"):
        """Return (pres, None) or (None, error_dict)."""
        if presentation_id not in presentations:
            return None, {"error": UNKNOWN_ID}
        pres = presentations[presentation_id]
        if slide_index is not None:
            problem = ppt_utils.check_index(pres, slide_index, what)
            if problem:
                return None, {"error": problem}
        return pres, None

    @app.tool(
        annotations=ToolAnnotations(
            title="Delete Slide",
        ),
    )
    def delete_slide(presentation_id: str, slide_index: int) -> Dict:
        """Remove a slide from the presentation.

        Use this to drop template slides whose content you don't need — a
        template's unused sections, or the fourth column of a three-item
        layout. Deleting is preferable to blanking a slide out: an emptied
        slide still shows the template's decoration and reads as a mistake.

        slide_index: 0-based. Every later slide shifts down by one, so when
        removing several, delete from the highest index downwards or re-read
        get_presentation_info between calls.
        """
        pres, error = _resolve(presentation_id, slide_index)
        if error:
            return error

        try:
            ppt_utils.delete_slide(pres, slide_index)
        except Exception as e:
            logger.error("delete_slide_failed presentation_id=%s index=%d "
                         "error=%s", short_id(presentation_id), slide_index, e)
            return {"error": f"Failed to delete slide {slide_index}: {e}"}

        remaining = len(pres.slides)
        logger.info("slide_deleted presentation_id=%s index=%d remaining=%d",
                    short_id(presentation_id), slide_index, remaining)
        return {
            "message": f"Deleted slide {slide_index}. The presentation now has "
                       f"{remaining} slide(s), indexed 0-{remaining - 1}."
                       if remaining else
                       f"Deleted slide {slide_index}. The presentation is now empty.",
            "deleted_index": slide_index,
            "slide_count": remaining,
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Move Slide",
        ),
    )
    def move_slide(presentation_id: str, slide_index: int,
                   new_index: int) -> Dict:
        """Move a slide to a different position in the deck.

        slide_index and new_index are both 0-based, and new_index is the
        position the slide ends up at once it has been lifted out — moving
        slide 0 to index 3 in a five-slide deck leaves it fourth.
        """
        pres, error = _resolve(presentation_id, slide_index)
        if error:
            return error

        total = len(pres.slides)
        if not isinstance(new_index, int) or isinstance(new_index, bool):
            return {"error": "new_index must be an integer"}
        if new_index < 0 or new_index >= total:
            return {"error": f"Invalid new_index: {new_index}. This "
                             f"presentation has {total} slide(s), indexed "
                             f"0-{total - 1}"}

        try:
            ppt_utils.move_slide(pres, slide_index, new_index)
        except Exception as e:
            logger.error("move_slide_failed presentation_id=%s from=%d to=%d "
                         "error=%s", short_id(presentation_id), slide_index,
                         new_index, e)
            return {"error": f"Failed to move slide {slide_index}: {e}"}

        logger.info("slide_moved presentation_id=%s from=%d to=%d",
                    short_id(presentation_id), slide_index, new_index)
        return {
            "message": f"Moved slide {slide_index} to position {new_index}.",
            "slide_index": new_index,
            "slide_count": total,
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Duplicate Slide",
        ),
    )
    def duplicate_slide(presentation_id: str, slide_index: int,
                        insert_after: Optional[int] = None,
                        count: int = 1) -> Dict:
        """Copy an existing slide, keeping all of its content and formatting.

        This is the way to reuse a corporate template's designed slides:
        duplicate the slide whose layout suits your next section, then edit
        the copy's text with manage_text / populate_placeholder. add_slide
        builds a bare slide from a layout and gets you none of the template's
        artwork, so prefer duplicating when the template already contains a
        slide that looks right.

        slide_index: 0-based index of the slide to copy.
        insert_after: 0-based index the copy is placed after; omit to append
          the copy (or copies) to the end of the deck.
        count: how many copies to make, for a section that repeats — the
          copies land consecutively.

        Pictures are shared with the original (same bytes, no size cost);
        charts, SmartArt and embedded objects are cloned, so editing the
        copy's chart does not change the original's. Speaker notes are copied.
        Returns "new_slide_indexes" — read them before editing, since
        inserting in the middle shifts every later slide.
        """
        pres, error = _resolve(presentation_id, slide_index)
        if error:
            return error

        if not isinstance(count, int) or isinstance(count, bool) or count < 1:
            return {"error": "count must be a positive integer"}
        if count > 50:
            return {"error": f"count of {count} is too large; duplicate at "
                             "most 50 slides in one call"}
        if insert_after is not None:
            problem = ppt_utils.check_index(pres, insert_after, "insert_after")
            if problem:
                return {"error": problem}

        created = []
        try:
            after = insert_after
            for _ in range(count):
                new_index = ppt_utils.duplicate_slide(pres, slide_index, after)
                created.append(new_index)
                if after is not None:
                    # Keep copies in order, and follow the source if inserting
                    # ahead of it pushed it down the deck.
                    after = new_index
                    if new_index <= slide_index:
                        slide_index += 1
        except Exception as e:
            logger.error("duplicate_slide_failed presentation_id=%s index=%d "
                         "error=%s", short_id(presentation_id), slide_index, e)
            return {"error": f"Failed to duplicate slide {slide_index}: {e}"}

        logger.info("slide_duplicated presentation_id=%s source=%d count=%d "
                    "new=%s", short_id(presentation_id), slide_index, count,
                    ",".join(str(i) for i in created))
        return {
            "message": f"Duplicated slide {slide_index} {count} time(s). New "
                       f"slide index(es): {created}.",
            "new_slide_indexes": created,
            "slide_count": len(pres.slides),
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Copy Slide Between Presentations",
        ),
    )
    def copy_slide_between_presentations(
        source_presentation_id: str,
        slide_index: int,
        target_presentation_id: str,
        insert_after: Optional[int] = None,
        layout_index: Optional[int] = None,
    ) -> Dict:
        """Copy one slide from one open presentation into another.

        Use this to merge decks, or to lift a single slide out of a reference
        deck into the one you are building. Both presentations must be open on
        this server (you hold both handles).

        The slide keeps its own shapes, text and direct formatting, and its
        images and charts are copied into the target deck. Anything it
        inherited from its old theme — placeholder fonts, scheme colours — is
        re-resolved against the TARGET deck's master, so the copy can look
        different from the original. Inspect the result with
        visual_inspect_slides.

        layout_index: force a particular layout in the target deck. By default
        the layout is matched by name, then by placeholder structure; the
        response reports whether a real match was found in "layout_matched".
        """
        if source_presentation_id not in presentations:
            return {"error": "Unknown or expired source_presentation_id. " + UNKNOWN_ID}
        if target_presentation_id not in presentations:
            return {"error": "Unknown or expired target_presentation_id. " + UNKNOWN_ID}
        if source_presentation_id == target_presentation_id:
            return {"error": "source_presentation_id and "
                             "target_presentation_id are the same "
                             "presentation; use duplicate_slide instead"}

        source = presentations[source_presentation_id]
        target = presentations[target_presentation_id]

        problem = ppt_utils.check_index(source, slide_index)
        if problem:
            return {"error": f"In the source presentation: {problem}"}
        if insert_after is not None:
            problem = ppt_utils.check_index(target, insert_after, "insert_after")
            if problem:
                return {"error": f"In the target presentation: {problem}"}
        if layout_index is not None:
            layouts = len(target.slide_layouts)
            if layout_index < 0 or layout_index >= layouts:
                return {"error": f"Invalid layout_index: {layout_index}. The "
                                 f"target presentation has {layouts} layout(s), "
                                 f"indexed 0-{layouts - 1}"}

        try:
            new_index, layout_name, matched = ppt_utils.copy_slide_to_presentation(
                source, slide_index, target, insert_after, layout_index)
        except Exception as e:
            logger.error("copy_slide_failed source=%s target=%s index=%d "
                         "error=%s", short_id(source_presentation_id),
                         short_id(target_presentation_id), slide_index, e)
            return {"error": f"Failed to copy slide {slide_index} between "
                             f"presentations: {e}"}

        logger.info("slide_copied source=%s target=%s index=%d new=%d "
                    "layout=%s matched=%s", short_id(source_presentation_id),
                    short_id(target_presentation_id), slide_index, new_index,
                    layout_name, matched)
        result = {
            "message": f"Copied slide {slide_index} into the target "
                       f"presentation at index {new_index}.",
            "slide_index": new_index,
            "layout_name": layout_name,
            "layout_matched": matched,
            "slide_count": len(target.slides),
        }
        if not matched:
            result["note"] = (
                "The target deck has no layout matching the source slide's, "
                "so a fallback layout was used. The copied shapes are intact, "
                "but placeholder styling comes from the target template — "
                "check this slide visually."
            )
        return result

    @app.tool(
        annotations=ToolAnnotations(
            title="Manage Speaker Notes",
        ),
    )
    def manage_speaker_notes(
        presentation_id: str,
        operation: str,
        slide_index: Optional[int] = None,
        text: Optional[str] = None,
    ) -> Dict:
        """Read, write or clear a slide's speaker notes.

        Notes belong in the notes pane, never in a text box on the slide
        itself — a "notes" textbox is visible to the audience during the
        presentation.

        operation:
          "get"   return the notes for slide_index, or for every slide when
                  slide_index is omitted;
          "set"   replace the notes on slide_index with `text`;
          "clear" remove the notes from slide_index.
        """
        if presentation_id not in presentations:
            return {"error": UNKNOWN_ID}
        pres = presentations[presentation_id]

        if operation not in NOTES_OPERATIONS:
            return {"error": f"Invalid operation: {operation}. Must be one of "
                             f"{', '.join(NOTES_OPERATIONS)}."}

        if operation == "get" and slide_index is None:
            notes = [{"slide_index": i,
                      "notes": ppt_utils.get_speaker_notes(slide)}
                     for i, slide in enumerate(pres.slides)]
            return {"notes": notes,
                    "slides_with_notes": sum(1 for n in notes if n["notes"])}

        if slide_index is None:
            return {"error": f"slide_index is required for operation "
                             f"'{operation}'"}
        problem = ppt_utils.check_index(pres, slide_index)
        if problem:
            return {"error": problem}
        slide = pres.slides[slide_index]

        if operation == "get":
            return {"slide_index": slide_index,
                    "notes": ppt_utils.get_speaker_notes(slide)}

        if operation == "set" and text is None:
            return {"error": "text is required for operation 'set'. Use "
                             "operation 'clear' to remove notes."}

        try:
            ppt_utils.set_speaker_notes(slide, "" if operation == "clear" else text)
        except Exception as e:
            logger.error("speaker_notes_failed presentation_id=%s index=%d "
                         "operation=%s error=%s", short_id(presentation_id),
                         slide_index, operation, e)
            return {"error": f"Failed to {operation} speaker notes on slide "
                             f"{slide_index}: {e}"}

        verb = "Cleared" if operation == "clear" else "Set"
        return {
            "message": f"{verb} speaker notes on slide {slide_index}.",
            "slide_index": slide_index,
            "characters": 0 if operation == "clear" else len(text),
        }
