"""
PowerPoint utilities package.
Organized utility functions for PowerPoint manipulation.
"""

from .core_utils import *
from .presentation_utils import *
from .content_utils import *
from .design_utils import *
from .validation_utils import *
from .slide_utils import *
from .combo_chart_utils import *

__all__ = [
    # Combo chart utilities
    "COMBO_SERIES_TYPES",
    "apply_combo_layout",
    "value_axis_with_id",
    "set_series_color",
    "set_series_data_labels",
    "add_series_trendline",
    "TRENDLINE_TYPES",

    # Slide structure utilities
    "delete_slide",
    "move_slide",
    "duplicate_slide",
    "copy_slide_to_presentation",
    "get_speaker_notes",
    "set_speaker_notes",
    "check_index",

    # Core utilities
    "safe_operation",
    "try_multiple_approaches",
    
    # Presentation utilities
    "create_presentation",
    "open_presentation", 
    "save_presentation",
    "create_presentation_from_template",
    "get_presentation_info",
    "get_template_info",
    "set_core_properties",
    "get_core_properties",
    
    # Content utilities
    "add_slide",
    "get_slide_info",
    "set_title",
    "populate_placeholder",
    "add_bullet_points",
    "add_textbox",
    "format_text",
    "format_text_advanced",
    "add_image",
    "add_table",
    "format_table_cell",
    "add_chart",
    "format_chart",
    
    # Design utilities
    "get_professional_color",
    "get_professional_font", 
    "get_color_schemes",
    "add_professional_slide",
    "apply_professional_theme",
    "enhance_existing_slide",
    "apply_professional_image_enhancement",
    "enhance_image_with_pillow",
    "set_slide_gradient_background",
    "create_professional_gradient_background",
    "format_shape",
    "apply_picture_shadow",
    "apply_picture_reflection",
    "apply_picture_glow",
    "apply_picture_soft_edges",
    "apply_picture_rotation",
    "apply_picture_transparency",
    "apply_picture_bevel",
    "apply_picture_filter",
    "analyze_font_file",
    "optimize_font_for_presentation",
    "get_font_recommendations",
    
    # Validation utilities
    "validate_text_fit",
    "validate_and_fix_slide"
]