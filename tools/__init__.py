"""
Tools package for PowerPoint MCP Server.
Organizes tools into logical modules for better maintainability.
"""

from .presentation_tools import register_presentation_tools
from .content_tools import register_content_tools
from .structural_tools import register_structural_tools
from .professional_tools import register_professional_tools
from .template_tools import register_template_tools
from .hyperlink_tools import register_hyperlink_tools
from .chart_tools import register_chart_tools
from .connector_tools import register_connector_tools
from .master_tools import register_master_tools
from .transition_tools import register_transition_tools
from .visual_tools import register_visual_tools
from .image_tools import register_image_tools
from .slide_tools import register_slide_tools
from .slide_number_tools import register_slide_number_tools
from .validation_tools import register_validation_tools
from .guidance_tools import register_guidance_tools
from .preview_tools import register_preview_tools

__all__ = [
    "register_visual_tools",
    "register_presentation_tools",
    "register_content_tools", 
    "register_structural_tools",
    "register_professional_tools",
    "register_template_tools",
    "register_hyperlink_tools",
    "register_chart_tools",
    "register_connector_tools",
    "register_master_tools",
    "register_transition_tools",
    "register_image_tools",
    "register_slide_tools",
    "register_slide_number_tools",
    "register_validation_tools",
    "register_guidance_tools",
    "register_preview_tools"
]