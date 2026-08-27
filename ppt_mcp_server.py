#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
Fork for EPAM AI DIAL: remote transport, multi-tenant state, DIAL file I/O
and agent-driven visual QA on top of the upstream content tools.
"""
import os
import argparse
import logging
import re
from pathlib import Path
from typing import Dict, Any

from dotenv import load_dotenv
# .env next to this file, if present; real env vars take precedence
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

# Logging is configured before FastMCP is imported/instantiated: FastMCP's
# own logging.basicConfig (RichHandler, multi-line output) is a no-op once
# the root logger already has a handler. See logging_utils.
from logging_utils import configure_logging, get_logger

LOG_LEVEL = configure_logging()
logger = get_logger("server")

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

# import utils  # Currently unused
from tools import (
    register_presentation_tools,
    register_content_tools,
    register_structural_tools,
    register_professional_tools,
    register_template_tools,
    register_hyperlink_tools,
    register_chart_tools,
    register_connector_tools,
    register_master_tools,
    register_transition_tools,
    register_visual_tools,
    register_image_tools,
    register_slide_tools,
    register_slide_number_tools,
    register_validation_tools,
    register_brand_tools,
    register_guidance_tools,
    register_preview_tools
)

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-mcp-server",
    # Surfaced to the client on initialize. Kept to a pointer rather than the
    # advice itself: the full guidance is a few pages, and get_design_guidance
    # serves it (whole or by section) only to a caller that is actually
    # building a deck.
    instructions=(
        "Builds PowerPoint decks, usually from a corporate template. Every "
        "tool takes an explicit presentation_id — there is no current "
        "presentation.\n"
        "Before planning a deck, call get_design_guidance: it covers deck "
        "structure, layout, type and colour, and how to work with a supplied "
        "template rather than over it.\n"
        "If attach_brand_profile is in your tool list, this deployment holds "
        "decks to a brand profile: call it right after creating the "
        "presentation and build to the rules it returns — they are cheaper to "
        "follow than to retrofit.\n"
        "While building, inspect each slide you finish with "
        "visual_inspect_slides(slides=[n]). Before exporting, run "
        "validate_presentation for structural faults and visual_repair_slides "
        "for appearance; they check different things. Deliver the deck with "
        "export_presentation and give the user the file URL it returns, then "
        "call render_deck_summary_card and attach its image so the user can "
        "see the finished deck without opening the file."
    ),
)

# Presentation state: concurrency-safe, UUID-handle store (see state.py).
# Dict-like, so existing tool call sites work unchanged.
from state import PresentationStore, serialize_per_presentation

presentations = PresentationStore()

# Template configuration
def get_template_search_directories():
    """
    Get list of directories to search for templates.
    Uses environment variable PPT_TEMPLATE_PATH if set, otherwise uses default directories.
    
    Returns:
        List of directories to search for templates
    """
    template_env_path = os.environ.get('PPT_TEMPLATE_PATH')
    
    if template_env_path:
        # If environment variable is set, use it as the primary template directory
        # Support multiple paths separated by colon (Unix) or semicolon (Windows)
        import platform
        separator = ';' if platform.system() == "Windows" else ':'
        env_dirs = [path.strip() for path in template_env_path.split(separator) if path.strip()]
        
        # Verify that the directories exist
        valid_env_dirs = []
        for dir_path in env_dirs:
            expanded_path = os.path.expanduser(dir_path)
            if os.path.exists(expanded_path) and os.path.isdir(expanded_path):
                valid_env_dirs.append(expanded_path)
        
        if valid_env_dirs:
            # Add default fallback directories
            logger.debug("template_search_dirs source=PPT_TEMPLATE_PATH dirs=%s",
                         ",".join(valid_env_dirs))
            return valid_env_dirs + ['.', './templates', './assets', './resources']
        else:
            logger.warning("template_path_missing PPT_TEMPLATE_PATH=%s "
                           "falling_back_to=defaults", template_env_path)
    
    # Default search directories when no environment variable or invalid paths
    return ['.', './templates', './assets', './resources']

# ---- Helper Functions ----

def get_current_presentation_id():
    """There is no 'current presentation' on a shared multi-tenant server:
    every tool call must pass the presentation_id handle explicitly.
    (Upstream's implicit-current mechanism was already non-functional —
    nothing ever set the global.)"""
    return None

def validate_parameters(params):
    """
    Validate parameters against constraints.
    
    Args:
        params: Dictionary of parameter name: (value, constraints) pairs
        
    Returns:
        (True, None) if all valid, or (False, error_message) if invalid
    """
    for param_name, (value, constraints) in params.items():
        for constraint_func, error_msg in constraints:
            if not constraint_func(value):
                return False, f"Parameter '{param_name}': {error_msg}"
    return True, None

def is_positive(value):
    """Check if a value is positive."""
    return value > 0

def is_non_negative(value):
    """Check if a value is non-negative."""
    return value >= 0

def is_in_range(min_val, max_val):
    """Create a function that checks if a value is in a range."""
    return lambda x: min_val <= x <= max_val

def is_in_list(valid_list):
    """Create a function that checks if a value is in a list."""
    return lambda x: x in valid_list

def is_valid_rgb(color_list):
    """Check if a color list is a valid RGB tuple."""
    if not isinstance(color_list, list) or len(color_list) != 3:
        return False
    return all(isinstance(c, int) and 0 <= c <= 255 for c in color_list)

def add_shape_direct(slide, shape_type: str, left: float, top: float, width: float, height: float) -> Any:
    """
    Add an auto shape to a slide using direct integer values instead of enum objects.
    
    This implementation provides a reliable alternative that bypasses potential 
    enum-related issues in the python-pptx library.
    
    Args:
        slide: The slide object
        shape_type: Shape type string (e.g., 'rectangle', 'oval', 'triangle')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches
        
    Returns:
        The created shape
    """
    from pptx.util import Inches
    
    # Direct mapping of shape types to their integer values
    # Values from MSO_AUTO_SHAPE_TYPE enum: https://github.com/scanny/python-pptx/blob/master/src/pptx/enum/shapes.py
    shape_type_map = {
        'rectangle': 1,              # RECTANGLE
        'rounded_rectangle': 5,      # ROUNDED_RECTANGLE
        'oval': 9,                   # OVAL
        'diamond': 4,                # DIAMOND
        'triangle': 7,               # ISOSCELES_TRIANGLE
        'right_triangle': 8,         # RIGHT_TRIANGLE
        'pentagon': 51,              # PENTAGON
        'hexagon': 10,               # HEXAGON
        'heptagon': 145,             # HEPTAGON
        'octagon': 6,                # OCTAGON
        'star': 92,                  # STAR_5_POINT
        'arrow': 33,                 # RIGHT_ARROW
        'cloud': 179,                # CLOUD
        'heart': 21,                 # HEART
        'lightning_bolt': 22,        # LIGHTNING_BOLT
        'sun': 23,                   # SUN
        'moon': 24,                  # MOON
        'smiley_face': 17,           # SMILEY_FACE
        'no_symbol': 19,             # NO_SYMBOL
        'flowchart_process': 61,     # FLOWCHART_PROCESS
        'flowchart_decision': 63,    # FLOWCHART_DECISION
        'flowchart_data': 64,        # FLOWCHART_DATA
        'flowchart_document': 67     # FLOWCHART_DOCUMENT
    }
    
    # Check if shape type is valid before trying to use it
    shape_type_lower = str(shape_type).lower()
    if shape_type_lower not in shape_type_map:
        available_shapes = ', '.join(sorted(shape_type_map.keys()))
        raise ValueError(f"Unsupported shape type: '{shape_type}'. Available shape types: {available_shapes}")
    
    # Get the integer value for the shape type
    shape_value = shape_type_map[shape_type_lower]
    
    # Create the shape using the direct integer value
    try:
        # The integer value is passed directly to add_shape
        shape = slide.shapes.add_shape(
            shape_value, Inches(left), Inches(top), Inches(width), Inches(height)
        )
        return shape
    except Exception as e:
        raise ValueError(f"Failed to create '{shape_type}' shape using direct value {shape_value}: {str(e)}")

# ---- Register Tools ----
register_presentation_tools(
    app, 
    presentations, 
    get_current_presentation_id, 
    get_template_search_directories
)

register_content_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb
)

register_structural_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb,
    add_shape_direct
)

register_professional_tools(
    app,
    presentations,
    get_current_presentation_id
)

register_template_tools(
    app,
    presentations,
    get_current_presentation_id
)

register_hyperlink_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb
)

register_chart_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb
)


register_connector_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb
)

register_master_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb
)

register_transition_tools(
    app,
    presentations,
    get_current_presentation_id,
    validate_parameters,
    is_positive,
    is_non_negative,
    is_in_range,
    is_valid_rgb
)

register_visual_tools(
    app,
    presentations
)

register_image_tools(
    app,
    presentations
)

register_slide_tools(
    app,
    presentations
)

register_slide_number_tools(
    app,
    presentations
)

register_validation_tools(
    app,
    presentations
)

register_brand_tools(
    app,
    presentations
)

register_guidance_tools(
    app,
    presentations
)

register_preview_tools(
    app,
    presentations
)


# ---- Additional Utility Tools ----
# Note: upstream's list_presentations and switch_presentation tools are
# removed on purpose: on a shared multi-tenant server the former leaked every
# tenant's presentation handles and the latter only mutated the (removed)
# global "current presentation" pointer.

@app.tool(
    annotations=ToolAnnotations(
        title="Get Server Info",
        readOnlyHint=True,
    ),
)
def get_server_info() -> Dict:
    """Report this server's version, how many tools it is actually exposing,
    and whether the optional DIAL and visual QA integrations are configured.

    Tool registration is dynamic — the visual QA tools appear only when a
    vision LLM is configured — so "tools" is counted at call time rather
    than hardcoded."""
    return {
        "name": "dial-pptx-mcp (PowerPoint MCP Server for EPAM AI DIAL)",
        "version": _project_version(),
        "upstream": "GongRzhe/Office-PowerPoint-MCP-Server",
        "tools": _registered_tool_count(),
        "loaded_presentations": len(presentations),
        "visual_qa": _visual_qa_status(),
        "dial_file_storage": ("configured" if os.environ.get("DIAL_CORE_URL")
                              else "unset"),
        "brand_profile": _brand_profile_status(),
    }


def _brand_profile_status():
    """The brand files this deployment expects, by name.

    Names, not contents: the files live in DIAL storage and are fetched from a
    URL the orchestrator resolves per deck (attach_brand_profile), so there is
    nothing here to read and nothing to report as reachable. Worth reporting
    at all because "the brand tools are missing" is otherwise
    indistinguishable from a bug."""
    import brand_validation

    if not brand_validation.enabled():
        return "unset"
    reference = brand_validation.reference_deck_file_name()
    return {"profile_file": brand_validation.profile_file_name(),
            "reference_deck_file": reference or None}


def _transport_security_for(host: str):
    """Host-header (DNS-rebinding) protection for the HTTP transports.

    FastMCP auto-enables a localhost-only allowlist when constructed, which
    421-rejects any remote Host header (e.g. Kubernetes service DNS names)
    once the server binds a non-loopback address.

    - PPT_MCP_ALLOWED_HOSTS (comma-separated hostnames; default
      dial-pptx-mcp.dial.svc.cluster.local — the standard in-cluster service
      name): protection is ON and exactly those hosts (any port) plus
      localhost are accepted.
    - PPT_MCP_ALLOWED_HOSTS=*: protection is DISABLED (any Host accepted) —
      for deployments fronted by ingress/service routing under other names.
    - Loopback bind with the value unset: keep FastMCP's localhost-only
      default (None).
    """
    from mcp.server.transport_security import TransportSecuritySettings

    default = ("" if host in ("127.0.0.1", "localhost", "::1")
               else "dial-pptx-mcp.dial.svc.cluster.local")
    allowed = os.environ.get("PPT_MCP_ALLOWED_HOSTS", default).strip()
    if allowed == "*":
        return TransportSecuritySettings(enable_dns_rebinding_protection=False)
    if allowed:
        hosts = []
        for h in allowed.split(","):
            h = h.strip()
            if h:
                hosts += [h, f"{h}:*"]
        hosts += ["127.0.0.1:*", "localhost:*", "[::1]:*"]
        return TransportSecuritySettings(
            enable_dns_rebinding_protection=True, allowed_hosts=hosts)
    return None


# ---- Main Function ----

def _project_version():
    """Version from pyproject.toml next to this file — the one place it is
    maintained. Parsed by hand: tomllib is 3.11+ and CI still runs 3.10."""
    try:
        text = (Path(__file__).resolve().parent / "pyproject.toml").read_text()
        match = re.search(r'^version\s*=\s*"([^"]+)"', text, re.MULTILINE)
        if match:
            return match.group(1)
    except OSError:
        pass
    logger.debug("project_version_unavailable")
    return "unknown"


def _registered_tool_count():
    """Number of tools the SDK ended up exposing (registration is dynamic —
    e.g. the visual QA tools appear only when a vision LLM is configured)."""
    try:
        return len(app._tool_manager._tools)
    except AttributeError:
        return -1


def _visual_qa_status():
    """Whether the QA tools are available, and whether the optional export
    gate is armed, for the startup line."""
    try:
        import visual_qa
        if not visual_qa.enforcement_enabled():
            return "disabled"
        return "tools+export_gate" if visual_qa.export_gate_enabled() else "tools"
    except Exception:
        return "unavailable"


def main(transport: str = "stdio", port: int = 8000, host: str = "127.0.0.1"):
    # Serialize tool calls that target the same presentation (python-pptx is
    # not thread-safe); calls on different presentations run concurrently.
    serialize_per_presentation(app, presentations)
    # Keep uvicorn's own verbosity in step with LOG_LEVEL.
    app.settings.log_level = LOG_LEVEL
    logger.info(
        "server_starting transport=%s host=%s port=%s log_level=%s tools=%d "
        "visual_qa=%s dial_core=%s",
        transport, host, port, LOG_LEVEL, _registered_tool_count(),
        _visual_qa_status(), "configured" if os.environ.get("DIAL_CORE_URL")
        else "unset",
    )
    if transport == "http":
        import asyncio
        # Set the host/port for HTTP transport (host must be 0.0.0.0 in containers)
        app.settings.host = host
        app.settings.port = port
        security = _transport_security_for(host)
        if security is not None:
            app.settings.transport_security = security
        # Start the FastMCP server with HTTP transport
        try:
            app.run(transport='streamable-http')
        except asyncio.exceptions.CancelledError:
            logger.info("server_stopped reason=cancelled")
        except KeyboardInterrupt:
            logger.info("server_stopped reason=keyboard_interrupt")
        except Exception as e:
            logger.error("server_start_failed transport=http error=%s", e,
                         exc_info=logger.isEnabledFor(logging.DEBUG))

    elif transport == "sse":
        # Run the FastMCP server in SSE (Server Side Events) mode
        app.settings.host = host
        app.settings.port = port
        security = _transport_security_for(host)
        if security is not None:
            app.settings.transport_security = security
        try:
            app.run(transport='sse')
        except KeyboardInterrupt:
            logger.info("server_stopped reason=keyboard_interrupt")

    else:
        # Run the FastMCP server
        try:
            app.run(transport='stdio')
        except KeyboardInterrupt:
            logger.info("server_stopped reason=keyboard_interrupt")

if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(description="MCP Server for PowerPoint manipulation using python-pptx")

    parser.add_argument(
        "-t",
        "--transport",
        type=str,
        default=os.environ.get("PPT_MCP_TRANSPORT", "stdio"),
        choices=["stdio", "http", "sse"],
        help="Transport method for the MCP server (default: stdio; env: PPT_MCP_TRANSPORT)"
    )

    parser.add_argument(
        "-p",
        "--port",
        type=int,
        default=int(os.environ.get("PPT_MCP_PORT", "8000")),
        help="Port to run the MCP server on (default: 8000; env: PPT_MCP_PORT)"
    )

    parser.add_argument(
        "--host",
        type=str,
        default=os.environ.get("PPT_MCP_HOST", "127.0.0.1"),
        help="Host to bind for http/sse transports (default: 127.0.0.1; "
             "set to 0.0.0.0 in containers; env: PPT_MCP_HOST)"
    )
    args = parser.parse_args()
    main(args.transport, args.port, args.host)
