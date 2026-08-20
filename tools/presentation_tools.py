"""
Presentation management tools for PowerPoint MCP Server.
Handles presentation creation, opening, saving, and core properties.
"""
from typing import Dict, List, Optional, Any
import os
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
import utils as ppt_utils
from logging_utils import get_logger
from state import short_id

logger = get_logger("tools.presentation")


def _visual_qa_gate(presentations, pres_id):
    """Automatic visual QA before a deck leaves the server (export or save).

    Fully internal: inspects the deck and, when issues are found, repairs it
    in place and re-inspects (visual_qa.inspect_and_repair), so callers
    receive a finished, verified presentation — never a fix-and-retry
    request. Runs only when the vision LLM is configured and the deck was
    edited since it last passed. Returns None to let the export proceed, or
    an error dict when QA could not verify the deck.
    """
    import visual_qa

    if not visual_qa.enforcement_enabled():
        logger.debug("qa_gate_skipped presentation_id=%s reason=not_configured",
                     short_id(pres_id))
        return None
    if not presentations.is_dirty(pres_id):
        logger.debug("qa_gate_skipped presentation_id=%s reason=already_passed",
                     short_id(pres_id))
        return None
    logger.info("qa_gate_start presentation_id=%s", short_id(pres_id))
    try:
        outcome = visual_qa.inspect_and_repair(presentations[pres_id])
    except visual_qa.VisualQAError as e:
        if visual_qa.fail_open_on_error():
            logger.warning("qa_gate_failed_open presentation_id=%s error=%s",
                           short_id(pres_id), e)
            return None
        logger.error("qa_gate_blocked presentation_id=%s reason=inspection_error "
                     "error=%s", short_id(pres_id), e)
        return {
            "error": f"Automatic visual QA could not run ({e}). "
                     "The export was blocked; fix the inspection setup, or "
                     "set VISUAL_QA_ON_ERROR=allow to export uninspected."
        }
    if outcome["passed"]:
        logger.info("qa_gate_passed presentation_id=%s iterations=%d",
                    short_id(pres_id), outcome["iterations"])
        presentations.clear_dirty(pres_id)
        return None
    if visual_qa.unresolved_policy() == "export_as_is":
        logger.warning("qa_gate_override presentation_id=%s policy=export_as_is "
                       "unresolved_issues=%d", short_id(pres_id),
                       len(outcome.get("issues", [])))
        return None  # operator chose to ship best-effort decks
    logger.error("qa_gate_blocked presentation_id=%s reason=unresolved "
                 "iterations=%d unresolved_issues=%d", short_id(pres_id),
                 outcome["iterations"], len(outcome.get("issues", [])))
    refusal = {
        "error": "Automatic visual QA could not bring the presentation to a "
                 f"passing state after {outcome['iterations']} internal "
                 "inspection/repair round(s). The presentation was not "
                 "exported.",
        "unresolved_issues": outcome.get("issues", []),
        "repair_rounds": outcome.get("repair_rounds", []),
        "message": "This is a server-side quality failure, not a request to "
                   "retry. Report the unresolved issues to the user, or "
                   "adjust the deck content and export again.",
    }
    for key in ("raw_review", "note"):
        if key in outcome:
            refusal[key] = outcome[key]
    return refusal


def register_presentation_tools(app: FastMCP, presentations: Dict, get_current_presentation_id, get_template_search_directories):
    """Register presentation management tools with the FastMCP app"""
    
    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation",
        ),
    )
    def create_presentation() -> Dict:
        """Create a new PowerPoint presentation. Returns a presentation_id
        that must be passed to all subsequent tool calls for this deck."""
        # Create a new presentation
        pres = ppt_utils.create_presentation()

        # Server-generated unguessable handle (multi-tenant safety):
        # client-chosen IDs are not accepted.
        id = presentations.new_id()
        presentations[id] = pres

        logger.info("presentation_created presentation_id=%s source=blank",
                    short_id(id))
        return {
            "presentation_id": id,
            "message": f"Created new presentation with ID: {id}",
            "slide_count": len(pres.slides)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation from Template",
        ),
    )
    def create_presentation_from_template(template_path: str) -> Dict:
        """Create a new PowerPoint presentation from a template file. Returns
        a presentation_id that must be passed to all subsequent tool calls."""
        # Check if template file exists
        if not os.path.exists(template_path):
            # Try to find the template by searching in configured directories
            search_dirs = get_template_search_directories()
            template_name = os.path.basename(template_path)
            
            for directory in search_dirs:
                potential_path = os.path.join(directory, template_name)
                if os.path.exists(potential_path):
                    template_path = potential_path
                    break
            else:
                env_path_info = f" (PPT_TEMPLATE_PATH: {os.environ.get('PPT_TEMPLATE_PATH', 'not set')})" if os.environ.get('PPT_TEMPLATE_PATH') else ""
                return {
                    "error": f"Template file not found: {template_path}. Searched in {', '.join(search_dirs)}{env_path_info}"
                }
        
        # Create presentation from template
        try:
            pres = ppt_utils.create_presentation_from_template(template_path)
        except Exception as e:
            logger.error("template_load_failed source=path path=%s error=%s",
                         template_path, e)
            return {
                "error": f"Failed to create presentation from template: {str(e)}"
            }
        
        # Server-generated unguessable handle (multi-tenant safety)
        id = presentations.new_id()
        presentations[id] = pres

        logger.info("template_loaded source=path path=%s presentation_id=%s "
                    "slides=%d layouts=%d", template_path, short_id(id),
                    len(pres.slides), len(pres.slide_layouts))
        return {
            "presentation_id": id,
            "message": f"Created new presentation from template '{template_path}' with ID: {id}",
            "template_path": template_path,
            "slide_count": len(pres.slides),
            "layout_count": len(pres.slide_layouts)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Open Presentation",
            readOnlyHint=True,
        ),
    )
    def open_presentation(file_path: str) -> Dict:
        """Open an existing PowerPoint presentation from a file. Returns a
        presentation_id that must be passed to all subsequent tool calls."""
        # Check if file exists
        if not os.path.exists(file_path):
            return {
                "error": f"File not found: {file_path}"
            }
        
        # Open the presentation
        try:
            pres = ppt_utils.open_presentation(file_path)
        except Exception as e:
            logger.error("open_failed path=%s error=%s", file_path, e)
            return {
                "error": f"Failed to open presentation: {str(e)}"
            }
        
        # Server-generated unguessable handle (multi-tenant safety)
        id = presentations.new_id()
        presentations[id] = pres

        logger.info("presentation_opened path=%s presentation_id=%s slides=%d",
                    file_path, short_id(id), len(pres.slides))
        return {
            "presentation_id": id,
            "message": f"Opened presentation from {file_path} with ID: {id}",
            "slide_count": len(pres.slides)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Save Presentation",
            destructiveHint=True,
        ),
    )
    def save_presentation(file_path: str, presentation_id: Optional[str] = None) -> Dict:
        """Save a presentation to a file."""
        # Use the specified presentation or the current one
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "Unknown or expired presentation_id. Pass the presentation_id returned by create_presentation, create_presentation_from_template, or open_presentation"
            }

        refusal = _visual_qa_gate(presentations, pres_id)
        if refusal is not None:
            return refusal

        # Save the presentation
        try:
            saved_path = ppt_utils.save_presentation(presentations[pres_id], file_path)
            logger.info("presentation_saved presentation_id=%s path=%s",
                        short_id(pres_id), saved_path)
            return {
                "message": f"Presentation saved to {saved_path}",
                "file_path": saved_path
            }
        except Exception as e:
            logger.error("save_failed presentation_id=%s path=%s error=%s",
                         short_id(pres_id), file_path, e)
            return {
                "error": f"Failed to save presentation: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Presentation Info",
            readOnlyHint=True,
        ),
    )
    def get_presentation_info(presentation_id: Optional[str] = None) -> Dict:
        """Get information about a presentation."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "Unknown or expired presentation_id. Pass the presentation_id returned by create_presentation, create_presentation_from_template, or open_presentation"
            }
        
        pres = presentations[pres_id]
        
        try:
            info = ppt_utils.get_presentation_info(pres)
            info["presentation_id"] = pres_id
            return info
        except Exception as e:
            return {
                "error": f"Failed to get presentation info: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Template File Info",
            readOnlyHint=True,
        ),
    )
    def get_template_file_info(template_path: str) -> Dict:
        """Get information about a template file including layouts and properties."""
        # Check if template file exists
        if not os.path.exists(template_path):
            # Try to find the template by searching in configured directories
            search_dirs = get_template_search_directories()
            template_name = os.path.basename(template_path)
            
            for directory in search_dirs:
                potential_path = os.path.join(directory, template_name)
                if os.path.exists(potential_path):
                    template_path = potential_path
                    break
            else:
                return {
                    "error": f"Template file not found: {template_path}. Searched in {', '.join(search_dirs)}"
                }
        
        try:
            return ppt_utils.get_template_info(template_path)
        except Exception as e:
            return {
                "error": f"Failed to get template info: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation from Template Content",
        ),
    )
    def create_presentation_from_template_content(template_content: str) -> Dict:
        """Create a new presentation from an uploaded .pptx template file.

        template_content: the .pptx template file content, as a data: URI or a
        base64-encoded string (in DIAL Quick Apps, pass the template file as
        file:data::files/{bucket}/{path} and it is resolved automatically).

        The template's theme, layouts, masters and branding are preserved.
        Returns a presentation_id that must be passed to all subsequent tool
        calls for this deck.
        """
        import base64
        import binascii
        import io

        payload = template_content.strip()
        if payload.startswith("data:"):
            # RFC 2397: data:<mime>;base64,<payload>
            _, _, payload = payload.partition(",")
        try:
            raw = base64.b64decode(payload, validate=True)
        except (binascii.Error, ValueError):
            logger.warning("template_decode_failed reason=invalid_base64 chars=%d",
                           len(payload))
            return {
                "error": "template_content is not valid base64. Pass the "
                         "template as a data: URI or base64 string (in DIAL "
                         "Quick Apps: file:data::<dial file path>)."
            }
        if not raw.startswith(b"PK"):
            logger.warning("template_decode_failed reason=not_ooxml bytes=%d",
                           len(raw))
            return {
                "error": "Decoded template_content is not a .pptx/.potx file "
                         "(expected a ZIP/OOXML container)."
            }

        try:
            # Handles .pptx directly and .potx via content-type coercion
            pres = ppt_utils.open_presentation_bytes(raw)
        except Exception as e:
            logger.error("template_load_failed source=upload bytes=%d error=%s",
                         len(raw), e)
            return {"error": f"Failed to open template: {str(e)}"}

        # Server-generated unguessable handle (multi-tenant safety)
        id = presentations.new_id()
        presentations[id] = pres

        logger.info("template_loaded source=upload bytes=%d presentation_id=%s "
                    "slides=%d layouts=%d", len(raw), short_id(id),
                    len(pres.slides), len(pres.slide_layouts))
        return {
            "presentation_id": id,
            "message": f"Created new presentation from uploaded template with ID: {id}",
            "slide_count": len(pres.slides),
            "layout_count": len(pres.slide_layouts)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Export Presentation to DIAL Files",
        ),
    )
    def export_presentation(presentation_id: str, filename: str = "presentation.pptx") -> Dict:
        """Export the presentation to DIAL file storage and return its file URL.

        Use this (not save_presentation) to deliver the finished deck to the
        user. When visual QA is enabled on the server, the deck is
        automatically inspected first and a failing deck is refused with the
        issue list — fix the issues and retry. ALWAYS include the returned
        file_url in your final answer as an attachment so the user can
        download the presentation.
        """
        from dial_client import DialFileClient, DialConfigError, PPTX_MIME

        if presentation_id not in presentations:
            return {
                "error": "Unknown or expired presentation_id. Pass the presentation_id returned by create_presentation, create_presentation_from_template, or open_presentation"
            }

        refusal = _visual_qa_gate(presentations, presentation_id)
        if refusal is not None:
            return refusal

        pres = presentations[presentation_id]

        if not filename.lower().endswith(".pptx"):
            filename += ".pptx"

        try:
            import io
            buf = io.BytesIO()
            pres.save(buf)
            client = DialFileClient()
            file_url = client.upload(buf.getvalue(), filename)
            logger.info("export_ok presentation_id=%s filename=%s slides=%d "
                        "bytes=%d", short_id(presentation_id), filename,
                        len(pres.slides), buf.getbuffer().nbytes)
            return {
                "message": f"Presentation exported to DIAL file storage: {file_url}. "
                           "Include this file URL in your final answer.",
                "file_url": file_url,
                "mime_type": PPTX_MIME,
                "size_bytes": buf.getbuffer().nbytes
            }
        except DialConfigError as e:
            logger.error("export_failed presentation_id=%s reason=dial_config "
                         "error=%s", short_id(presentation_id), e)
            return {"error": str(e)}
        except Exception as e:
            logger.error("export_failed presentation_id=%s reason=%s error=%s",
                         short_id(presentation_id), type(e).__name__, e)
            return {"error": f"Failed to export presentation to DIAL: {str(e)}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Set Core Properties",
        ),
    )
    def set_core_properties(
        title: Optional[str] = None,
        subject: Optional[str] = None,
        author: Optional[str] = None,
        keywords: Optional[str] = None,
        comments: Optional[str] = None,
        presentation_id: Optional[str] = None
    ) -> Dict:
        """Set core document properties."""
        pres_id = presentation_id if presentation_id is not None else get_current_presentation_id()
        
        if pres_id is None or pres_id not in presentations:
            return {
                "error": "Unknown or expired presentation_id. Pass the presentation_id returned by create_presentation, create_presentation_from_template, or open_presentation"
            }
        
        pres = presentations[pres_id]
        
        try:
            ppt_utils.set_core_properties(
                pres,
                title=title,
                subject=subject,
                author=author,
                keywords=keywords,
                comments=comments
            )
            
            return {
                "message": "Core properties updated successfully"
            }
        except Exception as e:
            return {
                "error": f"Failed to set core properties: {str(e)}"
            }