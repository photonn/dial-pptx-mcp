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

# Compound File Binary header — the container PowerPoint 97-2003 (.ppt) uses.
# python-pptx reads OOXML only, so these are converted on the way in.
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _visual_qa_gate(presentations, pres_id):
    """Optional export gate (VISUAL_QA_EXPORT_GATE=true).

    Visual QA is normally orchestrator-driven — the agent calls
    visual_inspect_slides / visual_repair_slides while it builds (see
    tools/visual_tools.py) — and export does not inspect anything. Operators
    who want the guarantee that no unverified deck ever leaves the server
    turn this gate on: it runs the inspect-repair loop over the whole deck
    when the deck was edited since it last passed, and refuses the export if
    QA cannot verify it.

    Returns None to let the export proceed, or an error dict.
    """
    import visual_qa

    if not visual_qa.export_gate_enabled():
        logger.debug("qa_gate_skipped presentation_id=%s reason=gate_disabled",
                     short_id(pres_id))
        return None
    if not presentations.is_dirty(pres_id):
        logger.debug("qa_gate_skipped presentation_id=%s reason=already_passed",
                     short_id(pres_id))
        return None
    logger.info("qa_gate_start presentation_id=%s", short_id(pres_id))
    try:
        # The brand profile's review_notes are the rules no measurement can
        # express (a headline must carry a message; a slide must not be plain
        # text on white), and the reference deck lets the reviewer compare
        # against a real branded deck instead of inferring the brand from the
        # slides in front of it. Both come from the deck's attached brand
        # context and are simply absent when nothing was attached: the gate
        # exists to catch bad slides, not to hold a deck hostage to a build
        # that never called attach_brand_profile.
        focus, reference = _brand_review_context(presentations, pres_id)
        outcome = visual_qa.inspect_and_repair(
            presentations[pres_id], focus=focus, reference_pres=reference)
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


def _brand_review_context(presentations, pres_id):
    """-> (reviewer focus, reference deck) from the deck's brand rules, if any."""
    import brand_validation

    context = presentations.brand_for(pres_id) or {}
    return (brand_validation.review_focus(context.get("profile")),
            context.get("reference"))


def _export_pdf(client, pres, blob, stem):
    """Render the deck to PDF and upload it; -> (file_entry, error_message).

    Never raises: a PDF is a convenience next to the .pptx, and losing the
    renderer must not cost the user the deck itself. The caller decides
    whether the failure is fatal (format="pdf") or a note (format="both").
    """
    from dial_client import PDF_MIME

    try:
        import visual_qa
        pdf = visual_qa.render_pptx_bytes_to_pdf(blob)
    except Exception as e:
        logger.warning("export_pdf_failed slides=%d error=%s",
                       len(pres.slides), e)
        return None, (
            f"The deck could not be converted to PDF ({e}). PDF export needs "
            "LibreOffice on the server; the .pptx is unaffected."
        )
    try:
        return {
            "file_url": client.upload(pdf, f"{stem}.pdf",
                                      content_type=PDF_MIME),
            "mime_type": PDF_MIME,
            "size_bytes": len(pdf),
            "format": "pdf",
        }, None
    except Exception as e:
        logger.warning("export_pdf_upload_failed bytes=%d error=%s",
                       len(pdf), e)
        return None, f"The PDF was rendered but could not be stored ({e})."


def _structure_summary(pres, blob):
    """Cheap structural check folded into every export.

    Export is the last moment anything can be caught, and the check needs no
    renderer and no model call, so it runs unconditionally — but it never
    blocks: refusing to deliver a finished deck over a warning costs the user
    more than the warning does. A failure of the checker itself is reported as
    "unavailable" rather than raised, for the same reason.
    """
    try:
        import deck_validation
        report = deck_validation.validate_presentation(pres, blob)
        return {"validated": True,
                "errors": report["counts"]["error"],
                "warnings": report["counts"]["warning"]}
    except Exception as e:
        logger.warning("export_structure_check_failed error=%s", e)
        return {"validated": False, "note": "structural check unavailable"}


def _brand_summary(presentations, pres_id, pres):
    """Brand check folded into every export, when the deck has rules attached.

    Reported, never blocking — same reasoning as _structure_summary, and more
    so here: a brand rule is a house style, and refusing to deliver a finished
    deck over one costs the user far more than the finding does. On a server
    with no brand profile the key is left out entirely, so an agent is never
    told about rules it cannot see; on one that has a profile the deck never
    attached, it is told exactly that, since a deck delivered unchecked should
    not look like a deck that passed.
    """
    import brand_validation

    if not brand_validation.enabled():
        return None
    context = presentations.brand_for(pres_id) or {}
    profile = context.get("profile")
    if not profile:
        from tools.brand_tools import not_attached_error

        return {"validated": False,
                "note": not_attached_error(
                    brand_validation.profile_file_name())}
    try:
        report = brand_validation.validate_brand(pres, profile)
    except Exception as e:
        logger.warning("export_brand_check_failed error=%s", e)
        return {"validated": False, "note": "brand check unavailable"}
    return {"validated": True,
            "brand": report["brand"],
            "errors": report["counts"]["error"],
            "warnings": report["counts"]["warning"]}


def _visual_qa_status(presentations, pres_id):
    """Advisory QA state of the deck being exported, for the agent's benefit.

    "passed" the deck passed a whole-deck inspection and was not edited
             since; "unverified" it has slides that no inspection has
             cleared; "unavailable" no vision LLM is configured.
    """
    import visual_qa

    if not visual_qa.enforcement_enabled():
        return "unavailable"
    return "unverified" if presentations.is_dirty(pres_id) else "passed"


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

        A PowerPoint 97-2003 (.ppt) file is accepted too and converted to
        .pptx on the way in, where the server has LibreOffice.

        template_content: the .pptx/.potx/.ppt template file content, as a
        data: URI or a base64-encoded string (in DIAL Quick Apps, pass the template file as
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
        converted_from = None
        if raw.startswith(OLE_MAGIC):
            # PowerPoint 97-2003. python-pptx reads OOXML only, so a legacy
            # deck has to be converted before anything else can touch it.
            try:
                import visual_qa
                raw = visual_qa.convert_legacy_ppt(raw)
                converted_from = "ppt"
            except Exception as e:
                logger.warning("template_convert_failed reason=legacy_ppt "
                               "error=%s", e)
                return {
                    "error": "template_content is a PowerPoint 97-2003 (.ppt) "
                             f"file, which this server could not convert ({e}). "
                             "Converting it needs LibreOffice on the server. "
                             "Re-save the file as .pptx and try again."
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
        result = {
            "presentation_id": id,
            "message": f"Created new presentation from uploaded template with ID: {id}",
            "slide_count": len(pres.slides),
            "layout_count": len(pres.slide_layouts)
        }
        if converted_from:
            result["converted_from"] = converted_from
            result["conversion_note"] = (
                "This was a PowerPoint 97-2003 (.ppt) file, converted to .pptx "
                "on the way in. Conversion is approximate — check the layouts "
                "with render_slide_previews before building on them."
            )
        return result

    @app.tool(
        annotations=ToolAnnotations(
            title="Get DIAL Storage Info",
            readOnlyHint=True,
        ),
    )
    def get_dial_storage_info() -> Dict:
        """Report which DIAL file storage this server can actually read and
        write on this request: the identity in use, the bucket it owns, and
        its appdata path.

        Use it when a file URL is refused: DIAL storage is per-user, so a
        file whose URL names a different bucket than the one reported here
        cannot be downloaded, however valid the URL looks.
        """
        from dial_client import DialFileClient, DialConfigError, identity_label

        try:
            info = DialFileClient().get_bucket_info()
        except DialConfigError as e:
            return {"error": str(e)}
        except Exception as e:
            logger.error("dial_storage_info_failed reason=%s error=%s",
                         type(e).__name__, e)
            return {"error": f"Could not reach DIAL file storage: {str(e)}"}
        return {
            "identity": identity_label(),
            "bucket": info.get("bucket"),
            "appdata": info.get("appdata"),
            "note": "add_image_from_dial_url can only read files under this "
                    "bucket (and export writes under it). A URL naming any "
                    "other bucket will be refused by DIAL Core.",
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Export Presentation to DIAL Files",
        ),
    )
    def export_presentation(presentation_id: str,
                            filename: str = "presentation.pptx",
                            format: str = "pptx") -> Dict:
        """Export the presentation to DIAL file storage and return its file URL.

        Use this (not save_presentation) to deliver the finished deck to the
        user.

        format: "pptx" (default) for the editable deck; "pdf" for a read-only
        copy; "both" to deliver the pair, which is what a user who asked to
        "share" or "send" a deck usually wants. PDF is produced by rendering
        the deck through LibreOffice, so it needs the renderer and reflects
        that renderer's fonts — see get_design_guidance("type"). Export does NOT run visual QA for you: inspect the deck with
        visual_inspect_slides / visual_repair_slides while you build it, or
        at least once before exporting. The response carries "visual_qa":
        "passed" | "unverified" | "unavailable" so you can tell whether the
        deck was ever checked. ALWAYS include the returned file_url in your
        final answer as an attachment so the user can download the
        presentation.
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

        if format not in ("pptx", "pdf", "both"):
            return {"error": f"Invalid format: {format}. Must be 'pptx', "
                             f"'pdf' or 'both'."}

        stem = filename[:-5] if filename.lower().endswith(".pptx") else filename
        if stem.lower().endswith(".pdf"):
            stem = stem[:-4]
        if not stem:
            return {"error": "filename must not be empty."}

        try:
            import io
            buf = io.BytesIO()
            pres.save(buf)
            blob = buf.getvalue()
            structure = _structure_summary(pres, blob)
            client = DialFileClient()

            files = []
            pdf_error = None
            if format in ("pptx", "both"):
                files.append({
                    "file_url": client.upload(blob, f"{stem}.pptx"),
                    "mime_type": PPTX_MIME,
                    "size_bytes": len(blob),
                    "format": "pptx",
                })
            if format in ("pdf", "both"):
                pdf_file, pdf_error = _export_pdf(client, pres, blob, stem)
                if pdf_file:
                    files.append(pdf_file)
                elif format == "pdf":
                    return {"error": pdf_error}

            qa_status = _visual_qa_status(presentations, presentation_id)
            primary = files[0]
            urls = ", ".join(f["file_url"] for f in files)
            logger.info("export_ok presentation_id=%s filename=%s format=%s "
                        "slides=%d bytes=%d visual_qa=%s",
                        short_id(presentation_id), stem, format,
                        len(pres.slides), len(blob), qa_status)
            result = {
                "message": f"Presentation exported to DIAL file storage: {urls}. "
                           "Include the file URL(s) in your final answer.",
                "file_url": primary["file_url"],
                "mime_type": primary["mime_type"],
                "size_bytes": primary["size_bytes"],
                "files": files,
                "visual_qa": qa_status,
                "structure": structure,
            }
            if format == "both" and len(files) == 1:
                result["pdf_note"] = pdf_error
            if structure.get("errors"):
                result["structure_note"] = (
                    f"The exported deck has {structure['errors']} structural "
                    "error(s) that may stop it opening in PowerPoint. The "
                    "file was uploaded regardless; call "
                    "validate_presentation for the details and fixes."
                )
            brand = _brand_summary(presentations, presentation_id, pres)
            if brand:
                result["brand"] = brand
                if brand.get("validated") is False:
                    result["brand_note"] = (
                        f"{brand['note']} The file was delivered regardless, "
                        "but it has not been checked against the brand rules."
                    )
                elif brand.get("errors") or brand.get("warnings"):
                    result["brand_note"] = (
                        f"The deck breaks {brand.get('errors', 0)} rule(s) at "
                        f"error and {brand.get('warnings', 0)} at warning "
                        f"severity in the {brand.get('brand')} brand profile. "
                        "The file was delivered regardless; call "
                        "validate_brand_profile for the details and fixes."
                    )
            if qa_status == "unverified":
                result["visual_qa_note"] = (
                    "This deck was never visually inspected, or was edited "
                    "since its last passing inspection. The export succeeded "
                    "regardless; run visual_repair_slides if quality matters "
                    "before you hand the file to the user."
                )
            return result
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