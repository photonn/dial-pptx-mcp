"""
Custom icons drawn as SVG by the orchestrator, rendered and vetted here.

The icon problem a deck runs into is not "which icon", it is "which icon on
which colour": a stock PNG carries an opaque background, so the moment a card
is tinted or a panel is brand blue, a library icon shows as a white square.
Drawing the icon per slide solves that — an SVG with no background renders to a
transparent PNG — and it also covers the concepts no library has.

The split is the same one `add_image_from_dial_url` makes: the orchestrator
composes (it has the style guide, served by `get_icon_guidance`), the server
renders. What is added here is the check in between. The agent cannot see the
icon it just drew, and once a bad one is on a slide the visual-repair
whitelist can only move, resize or delete it — so the icon is reviewed on its
own, before placement, while regenerating it is one cheap call.

The result lands in DIAL file storage like any other image, which is what makes
the last step uniform: an icon and a generated illustration reach a slide
through the same `add_image_from_dial_url`.
"""
import re
from pathlib import Path
from typing import Dict, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from logging_utils import get_logger

logger = get_logger("tools.icon")

GUIDANCE_PATH = (Path(__file__).resolve().parent.parent / "docs"
                 / "ICON_GUIDANCE.md")

DEFAULT_FILENAME = "icon.png"

_cache = {}


def _guidance():
    """The icon style guide, cached on the file's mtime (as
    get_design_guidance does) so an edit needs no restart."""
    try:
        stamp = GUIDANCE_PATH.stat().st_mtime
    except OSError as e:
        logger.warning("icon_guidance_unreadable path=%s error=%s",
                       GUIDANCE_PATH, e)
        return None
    if _cache.get("stamp") != stamp:
        _cache.update(stamp=stamp,
                      text=GUIDANCE_PATH.read_text(encoding="utf-8"))
    return _cache["text"]


def _slug(text):
    return re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")[:40]


def register_icon_tools(app: FastMCP):
    """Register the SVG icon tools. No renderer dependency: PyMuPDF rasterizes
    the SVG, so these work even where LibreOffice (and therefore visual QA) is
    absent — the review step is the only part that needs a vision model."""
    import svg_icons
    import visual_qa

    logger.info("icon_tools_registered review=%s", visual_qa.enforcement_enabled())

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Icon Guidance",
            readOnlyHint=True,
        ),
    )
    def get_icon_guidance() -> Dict:
        """The style guide for drawing an icon as SVG, and how to use it.

        Read this before writing SVG for render_svg_icon: it holds the two
        icon templates (line art for a light surface, white-on-filled for a
        coloured one), the rules that keep a hand-drawn icon looking like the
        rest of the set, worked path examples you can adapt, and the decision
        of when a custom icon is the right answer at all.

        Returns the document as "guidance".
        """
        text = _guidance()
        if text is None:
            return {"error": f"The icon guidance document is missing from this "
                             f"server ({GUIDANCE_PATH.name}). Draw icons as "
                             f"simple line art on a 200x200 viewBox: one "
                             f"stroke weight, fill=\"none\", rounded caps and "
                             f"joins, no text."}
        return {"guidance": text, "characters": len(text)}

    @app.tool(
        annotations=ToolAnnotations(
            title="Render SVG Icon",
            readOnlyHint=True,
        ),
    )
    def render_svg_icon(
        svg: str,
        concept: Optional[str] = None,
        size: int = svg_icons.DEFAULT_PX,
        background: str = "transparent",
        slide_background: str = "#FFFFFF",
        review: bool = True,
        filename: Optional[str] = None,
    ) -> Dict:
        """Render an SVG icon you have written to a PNG, check it, and store it.

        Use this when a slide needs an icon: call get_icon_guidance, write the
        SVG yourself following it, pass it here, and then place the returned
        image_url with add_image_from_dial_url. PowerPoint cannot embed SVG,
        which is why the rasterizing happens here rather than in your answer.

        The returned icon has been looked at. You cannot see what you drew, so
        the vision model is asked whether the render is free of artifacts and
        still readable at slide size. **If "passed" is false, do not place the
        icon**: fix the SVG as the issues say and call this again. An
        image_url is returned only for an icon that is fit to use.

        svg: the SVG source, self-contained line art — <path>, <circle>,
          <rect>, <line>, <polyline>, <polygon>. No text elements, no embedded
          images, no external references; a text label belongs in a textbox on
          the slide, not in the icon.
        concept: what the icon depicts ("supply chain", "regulatory
          approval"). Passed to the reviewer, which is the difference between
          "this is a clean pictogram" and "this reads as the thing you meant".
        size: longest side of the PNG in pixels (default 800). An icon is
          displayed at about an inch, so this is deliberate headroom; there is
          no reason to raise it.
        background: the icon's own background. Leave "transparent" — that is
          the whole advantage over a stock icon, and it is what lets the same
          icon sit on a white slide and on a brand-coloured panel. A hex
          colour bakes a solid rectangle in behind it.
        slide_background: hex colour of the surface the icon will sit on. The
          review judges legibility and contrast against it, so set it to the
          panel colour for an icon going on a coloured card.
        review: set false only to skip the vision check (faster, unverified).
        filename: name for the stored PNG; defaults to a name from `concept`.

        Returns "image_url" (pass it straight to add_image_from_dial_url),
        the pixel size, and "review" with the verdict and any issues.
        """
        from dial_client import DialFileClient

        if size < svg_icons.MIN_PX or size > svg_icons.MAX_PX:
            return {"error": f"size must be between {svg_icons.MIN_PX} and "
                             f"{svg_icons.MAX_PX} pixels."}
        try:
            canvas = svg_icons.parse_color(background, "background")
            surface = svg_icons.parse_color(slide_background,
                                            "slide_background")
        except svg_icons.SvgIconError as e:
            return {"error": str(e)}

        try:
            png, info = svg_icons.render_svg(svg, size, canvas)
        except svg_icons.SvgIconError as e:
            logger.info("icon_render_rejected concept=%s reason=%s",
                        (concept or "-")[:40], e)
            return {"error": str(e)}

        result = {
            "image_size": {"width": info["width"], "height": info["height"]},
            "size_bytes": len(png),
            "mime_type": svg_icons.PNG_MIME,
            "ink_coverage": info["ink_coverage"],
        }
        # Named before any model call: a blank or solid-black render has a
        # specific cause in the SVG, and saying so beats a vision verdict.
        coverage_note = svg_icons.coverage_note(info["ink_coverage"])
        if coverage_note:
            result["render_note"] = coverage_note

        verdict = None
        if review:
            if not visual_qa.enforcement_enabled():
                result["review_note"] = (
                    "No vision model is configured on this server, so the "
                    "icon could not be checked. Place it if you are confident "
                    "in the SVG, and check the deck with the user."
                )
            else:
                try:
                    verdict = svg_icons.review_icon(png, concept, surface)
                except visual_qa.VisualQAError as e:
                    logger.warning("icon_review_failed concept=%s error=%s",
                                   (concept or "-")[:40], e)
                    result["review_note"] = (
                        f"The icon was rendered but could not be reviewed "
                        f"({e}). It is stored below unchecked."
                    )
                else:
                    result["review"] = verdict

        if verdict is not None and verdict.get("passed") is False:
            # Nothing to place, so nothing is stored: the next call carries a
            # corrected SVG, not this URL.
            logger.info("icon_rejected concept=%s issues=%d",
                        (concept or "-")[:40], len(verdict.get("issues", [])))
            result["message"] = (
                "This icon did not pass review, so it was not stored and "
                "there is no image_url to place. Rewrite the SVG following "
                "the issues above and call render_svg_icon again. If two "
                "attempts do not fix it, simplify the drawing — fewer, larger "
                "shapes — or leave the icon out rather than placing a broken "
                "one."
            )
            return result

        try:
            url = DialFileClient().upload(
                png, filename or f"{_slug(concept) or 'icon'}.png",
                content_type=svg_icons.PNG_MIME)
        except Exception as e:
            logger.warning("icon_upload_failed concept=%s error=%s",
                           (concept or "-")[:40], e)
            return {"error": f"The icon rendered but could not be stored in "
                             f"DIAL file storage ({e}), so there is no URL to "
                             f"place. Retry, or build the slide without the "
                             f"icon."}

        result["image_url"] = url
        result["message"] = (
            f"Icon stored at {url}. Place it with add_image_from_dial_url — "
            "give it a square box (width == height) and leave fit at "
            "\"contain\" so it is not distorted."
        )
        logger.info("icon_ready concept=%s bytes=%d reviewed=%s",
                    (concept or "-")[:40], len(png),
                    bool(verdict) and verdict.get("passed"))
        return result
