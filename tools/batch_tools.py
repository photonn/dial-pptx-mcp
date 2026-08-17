"""
Batch deck-building tool for PowerPoint MCP Server.

Orchestrators (DIAL Quick Apps) cap the number of LLM operations per user
request — 15 by default — so building a deck one tool call per slide runs
out of iterations before the deck is finished ("Agent stopped due to max
iterations"). build_presentation does the whole job in a SINGLE call:
template in, every slide's content, and (optionally) the DIAL export.

The granular tools remain available for follow-up edits.
"""
import base64
import binascii
import io
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

import utils as ppt_utils


def _decode_template(template_content: str) -> bytes:
    payload = template_content.strip()
    if payload.startswith("data:"):
        _, _, payload = payload.partition(",")
    raw = base64.b64decode(payload, validate=True)
    if not raw.startswith(b"PK"):
        raise ValueError("decoded template_content is not a .pptx/.potx file")
    return raw


def _body_placeholder(slide):
    """First non-title placeholder with a text frame, if any."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx != 0 and ph.has_text_frame:
            return ph
    return None


def _fill_slide(pres, slide, spec: Dict[str, Any]) -> List[str]:
    """Apply one slide spec. Returns a list of non-fatal notes."""
    warnings = []
    title = spec.get("title")
    if title:
        if slide.shapes.title is not None:
            ppt_utils.set_title(slide, str(title))
        else:
            # set_title silently no-ops without a title placeholder, which
            # would drop the text: add a heading textbox instead.
            width_in = pres.slide_width / 914400
            height_in = pres.slide_height / 914400
            ppt_utils.add_textbox(
                slide,
                left=width_in * 0.08, top=height_in * 0.08,
                width=width_in * 0.84, height=height_in * 0.15,
                text=str(title), font_size=32, bold=True,
            )
            warnings.append(
                "layout has no title placeholder; used a heading textbox")

    bullets = spec.get("bullets")
    body_text = spec.get("body_text")
    if bullets or body_text:
        placeholder = _body_placeholder(slide)
        if placeholder is not None:
            if bullets:
                ppt_utils.add_bullet_points(placeholder, [str(b) for b in bullets])
            else:
                placeholder.text_frame.text = str(body_text)
        else:
            # Layout has no body placeholder: fall back to a textbox inside
            # the slide, sized from the deck's own dimensions.
            width_in = pres.slide_width / 914400
            height_in = pres.slide_height / 914400
            text = ("\n".join(f"• {b}" for b in bullets) if bullets
                    else str(body_text))
            ppt_utils.add_textbox(
                slide,
                left=width_in * 0.08, top=height_in * 0.30,
                width=width_in * 0.84, height=height_in * 0.55,
                text=text,
            )
            warnings.append("layout has no body placeholder; used a textbox")

    notes = spec.get("notes")
    if notes:
        try:
            slide.notes_slide.notes_text_frame.text = str(notes)
        except Exception:
            warnings.append("could not set speaker notes")
    return warnings


def register_batch_tools(app: FastMCP, presentations):
    """Register the batch deck-building tool with the FastMCP app"""

    @app.tool(
        annotations=ToolAnnotations(
            title="Build Presentation (single call)",
        ),
    )
    def build_presentation(
        slides: List[Dict[str, Any]],
        template_content: Optional[str] = None,
        filename: str = "presentation.pptx",
        export: bool = True,
        keep_template_slides: bool = False,
        title: Optional[str] = None,
        subject: Optional[str] = None,
        author: Optional[str] = None,
    ) -> Dict:
        """Build a complete PowerPoint presentation in ONE call — the
        preferred way to create a deck. Do not add slides one at a time.

        slides: list of slide specs, in order. Each spec supports:
          - title: slide title text
          - bullets: list of bullet strings (body content)
          - body_text: plain body text (alternative to bullets)
          - notes: speaker notes
          - layout_index: template layout to use (default 1; use 0 for the
            opening title slide). Call get_presentation_info/
            get_template_file_info first if you need the layout list.
        template_content: optional .pptx/.potx template as a data: URI or
          base64 string (in DIAL Quick Apps pass the template file as
          file:data::files/{bucket}/{path}); its theme, layouts, masters and
          branding are preserved. Omit for a blank default deck.
        export: when true (default), the finished deck is exported to DIAL
          file storage and file_url is returned — include that URL in your
          final answer as an attachment.
        keep_template_slides: templates often ship example slides; by default
          they are removed so the deck contains only your slides (theme,
          layouts, masters and branding are always preserved). Set true to
          append your slides after the template's own.
        title/subject/author: optional document properties.

        Returns presentation_id (for any follow-up edits with the granular
        tools), slide_count, file_url when exported, and any warnings.
        """
        if not isinstance(slides, list) or not slides:
            return {"error": "slides must be a non-empty list of slide specs."}
        if not all(isinstance(s, dict) for s in slides):
            return {"error": "each entry in slides must be an object."}

        # 1. Create the deck (from template when supplied)
        try:
            if template_content:
                pres = ppt_utils.open_presentation_bytes(
                    _decode_template(template_content))
            else:
                pres = ppt_utils.create_presentation()
        except (binascii.Error, ValueError) as e:
            return {
                "error": f"Invalid template_content: {e}. Pass the template as "
                         "a data: URI or base64 string (in DIAL Quick Apps: "
                         "file:data::<dial file path>)."
            }
        except Exception as e:
            return {"error": f"Failed to open template: {e}"}

        layout_count = len(pres.slide_layouts)
        if layout_count == 0:
            return {"error": "The template provides no slide layouts."}

        warnings = []
        if template_content and not keep_template_slides:
            removed = ppt_utils.remove_all_slides(pres)
            if removed:
                warnings.append(
                    f"removed {removed} example slide(s) from the template "
                    "(pass keep_template_slides=true to keep them)")

        # 2. Add every slide
        for position, spec in enumerate(slides, start=1):
            requested = spec.get("layout_index", 1)
            try:
                layout_index = int(requested)
            except (TypeError, ValueError):
                layout_index = 1
            if not 0 <= layout_index < layout_count:
                warnings.append(
                    f"slide {position}: layout_index {requested} out of range "
                    f"(0-{layout_count - 1}); used "
                    f"{min(max(layout_index, 0), layout_count - 1)}")
                layout_index = min(max(layout_index, 0), layout_count - 1)
            try:
                slide, _ = ppt_utils.add_slide(pres, layout_index)
            except Exception as e:
                return {"error": f"Failed to add slide {position}: {e}"}
            for note in _fill_slide(pres, slide, spec):
                warnings.append(f"slide {position}: {note}")

        # 3. Document properties
        if any((title, subject, author)):
            try:
                ppt_utils.set_core_properties(
                    pres, title=title, subject=subject, author=author)
            except Exception:
                warnings.append("could not set document properties")

        # 4. Register the deck so follow-up tools can edit it
        pres_id = presentations.new_id()
        presentations[pres_id] = pres

        result = {
            "presentation_id": pres_id,
            "slide_count": len(pres.slides),
            "message": f"Built presentation with {len(pres.slides)} slide(s).",
        }
        if warnings:
            result["warnings"] = warnings
        if not export:
            return result

        # 5. Export through the same path as export_presentation (visual QA
        # gate included when enabled)
        from tools.presentation_tools import export_via_dial
        export_result = export_via_dial(presentations, pres_id, filename)
        if "error" in export_result:
            # Keep the handle so the agent can fix and re-export
            export_result["presentation_id"] = pres_id
            export_result["slide_count"] = len(pres.slides)
            if warnings:
                export_result["warnings"] = warnings
            return export_result
        result.update(export_result)
        result["message"] = (
            f"Built presentation with {len(pres.slides)} slide(s) and exported "
            "it. Include the file_url in your final answer as an attachment."
        )
        return result
