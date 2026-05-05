#!/usr/bin/env python3
"""
PowerPoint Editor MCP Server Main Program
Provides MCP server functionality, calls PowerPointEditor class from tool.py

Every tool requires file_path as the first parameter. The presentation is loaded
from file_path if it exists, otherwise a new one is created. After applying the
requested changes, the presentation is automatically saved back to file_path.
"""

import asyncio
import json
import logging

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from tool import PowerPointEditor

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

ppt_editor = PowerPointEditor()

server = Server("powerpoint-editor")

FILE_PATH_PARAM = {
    "type": "string",
    "description": "Full path to the PowerPoint file (.pptx). If the file exists it will be loaded; if not, a new presentation will be created and saved there.",
}


@server.list_tools()
async def handle_list_tools():
    return [
        Tool(
            name="add_slide",
            description="Add a new slide to a PowerPoint presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "layout_index": {
                        "type": "integer",
                        "description": "Slide layout index (0=title slide, 1=title and content, default=1)",
                        "default": 1,
                    },
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="add_text_box",
            description="Add a text box to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "text": {"type": "string", "description": "Text content to add"},
                    "left": {
                        "type": "number",
                        "description": "Text box left margin (inches)",
                        "default": 1,
                    },
                    "top": {
                        "type": "number",
                        "description": "Text box top margin (inches)",
                        "default": 1,
                    },
                    "width": {
                        "type": "number",
                        "description": "Text box width (inches)",
                        "default": 8,
                    },
                    "height": {
                        "type": "number",
                        "description": "Text box height (inches)",
                        "default": 1,
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "Font size",
                        "default": 18,
                    },
                    "font_color": {
                        "type": "string",
                        "description": "Font color (hexadecimal, e.g. 000000)",
                        "default": "000000",
                    },
                },
                "required": ["file_path", "slide_index", "text"],
            },
        ),
        Tool(
            name="add_title_slide",
            description="Add a title slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "title": {"type": "string", "description": "Slide title"},
                    "subtitle": {
                        "type": "string",
                        "description": "Slide subtitle (optional)",
                        "default": "",
                    },
                },
                "required": ["file_path", "title"],
            },
        ),
        Tool(
            name="add_bullet_points",
            description="Add bullet points content to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "title": {"type": "string", "description": "Slide title"},
                    "bullet_points": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Bullet points list",
                    },
                },
                "required": ["file_path", "slide_index", "title", "bullet_points"],
            },
        ),
        Tool(
            name="add_image",
            description="Add an image to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "image_path": {"type": "string", "description": "Image file path"},
                    "left": {
                        "type": "number",
                        "description": "Image left margin (inches)",
                        "default": 1,
                    },
                    "top": {
                        "type": "number",
                        "description": "Image top margin (inches)",
                        "default": 2,
                    },
                    "width": {
                        "type": "number",
                        "description": "Image width (inches, optional)",
                    },
                    "height": {
                        "type": "number",
                        "description": "Image height (inches, optional)",
                    },
                },
                "required": ["file_path", "slide_index", "image_path"],
            },
        ),
        Tool(
            name="add_shape",
            description="Add a shape to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_type": {
                        "type": "string",
                        "description": "Shape type (rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow)",
                    },
                    "left": {
                        "type": "number",
                        "description": "Shape left margin (inches)",
                        "default": 1,
                    },
                    "top": {
                        "type": "number",
                        "description": "Shape top margin (inches)",
                        "default": 1,
                    },
                    "width": {
                        "type": "number",
                        "description": "Shape width (inches)",
                        "default": 2,
                    },
                    "height": {
                        "type": "number",
                        "description": "Shape height (inches)",
                        "default": 1,
                    },
                    "fill_color": {
                        "type": "string",
                        "description": "Fill color (hexadecimal, e.g. 0066CC)",
                        "default": "0066CC",
                    },
                },
                "required": ["file_path", "slide_index", "shape_type"],
            },
        ),
        Tool(
            name="get_presentation_info",
            description="Get presentation information (slide count, titles, etc.)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="delete_slide",
            description="Delete a specified slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of slide to delete (starting from 0)",
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="duplicate_slide",
            description="Duplicate a specified slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Index of slide to duplicate (starting from 0)",
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="move_slide",
            description="Move slide position",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "from_index": {
                        "type": "integer",
                        "description": "Source position index (starting from 0)",
                    },
                    "to_index": {
                        "type": "integer",
                        "description": "Target position index (starting from 0)",
                    },
                },
                "required": ["file_path", "from_index", "to_index"],
            },
        ),
        Tool(
            name="add_table",
            description="Add a table to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "rows": {"type": "integer", "description": "Number of table rows"},
                    "cols": {
                        "type": "integer",
                        "description": "Number of table columns",
                    },
                    "left": {
                        "type": "number",
                        "description": "Table left margin (inches)",
                        "default": 1,
                    },
                    "top": {
                        "type": "number",
                        "description": "Table top margin (inches)",
                        "default": 2,
                    },
                    "width": {
                        "type": "number",
                        "description": "Table width (inches)",
                        "default": 8,
                    },
                    "height": {
                        "type": "number",
                        "description": "Table height (inches)",
                        "default": 4,
                    },
                },
                "required": ["file_path", "slide_index", "rows", "cols"],
            },
        ),
        Tool(
            name="set_table_cell_text",
            description="Set table cell text",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "table_index": {
                        "type": "integer",
                        "description": "Table index (starting from 0)",
                    },
                    "row": {
                        "type": "integer",
                        "description": "Row index (starting from 0)",
                    },
                    "col": {
                        "type": "integer",
                        "description": "Column index (starting from 0)",
                    },
                    "text": {"type": "string", "description": "Text content to set"},
                },
                "required": [
                    "file_path",
                    "slide_index",
                    "table_index",
                    "row",
                    "col",
                    "text",
                ],
            },
        ),
        Tool(
            name="set_slide_background_color",
            description="Set slide background color",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "color": {
                        "type": "string",
                        "description": "Background color (hexadecimal, e.g. FF0000)",
                    },
                },
                "required": ["file_path", "slide_index", "color"],
            },
        ),
        Tool(
            name="add_hyperlink",
            description="Add hyperlink to a shape",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index (starting from 0)",
                    },
                    "url": {"type": "string", "description": "Hyperlink URL"},
                    "display_text": {
                        "type": "string",
                        "description": "Display text (optional)",
                    },
                },
                "required": ["file_path", "slide_index", "shape_index", "url"],
            },
        ),
        Tool(
            name="set_text_formatting",
            description="Set text formatting for a shape",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index (starting from 0)",
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font name (optional)",
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "Font size (optional)",
                    },
                    "font_color": {
                        "type": "string",
                        "description": "Font color (hexadecimal, optional)",
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Whether to bold (optional)",
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Whether to italicize (optional)",
                    },
                    "underline": {
                        "type": "boolean",
                        "description": "Whether to underline (optional)",
                    },
                },
                "required": ["file_path", "slide_index", "shape_index"],
            },
        ),
        Tool(
            name="get_slide_shapes_info",
            description="Get information of all shapes on a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="add_slide_animation",
            description="Add animation transition effects to a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "animation_style": {
                        "type": "string",
                        "description": "Animation style: fade (recommended), push, wipe, zoom, split, blinds, dissolve, none",
                        "default": "fade",
                    },
                    "speed": {
                        "type": "string",
                        "description": "Animation speed: fast, medium, slow",
                        "default": "medium",
                    },
                    "auto_advance": {
                        "type": "boolean",
                        "description": "Whether to automatically advance to next slide",
                        "default": False,
                    },
                    "auto_advance_seconds": {
                        "type": "number",
                        "description": "Auto advance delay in seconds (only effective when auto_advance is true)",
                        "default": 3.0,
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="make_presentation_dynamic",
            description="Add uniform animation effects to the entire presentation so all slides have smooth transition animations",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "animation_style": {
                        "type": "string",
                        "description": "Uniform animation style: fade (recommended), push, wipe, zoom",
                        "default": "fade",
                    },
                    "speed": {
                        "type": "string",
                        "description": "Animation speed: fast, medium, slow",
                        "default": "medium",
                    },
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="get_animation_options",
            description="View all available slide animation effect options",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="make_professional_presentation",
            description="One-click to make presentation professional - automatically add elegant fade transition effects to all slides",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="add_smooth_transitions",
            description="Add smooth transition animations to presentation to make slide transitions more natural",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="add_dynamic_effects",
            description="Add dynamic transition effects to presentation to make presentations more lively",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="set_slide_transition",
            description="Set slide transition effect (advanced)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "transition_type": {
                        "type": "string",
                        "description": "Transition type: fade, push, wipe, zoom, split, blinds, dissolve, none",
                        "default": "fade",
                    },
                    "duration": {
                        "type": "number",
                        "description": "Transition duration in seconds (default=1.0)",
                        "default": 1.0,
                    },
                    "advance_on_click": {
                        "type": "boolean",
                        "description": "Whether to advance on click (default=true)",
                        "default": True,
                    },
                    "advance_after_time": {
                        "type": "number",
                        "description": "Auto advance after this many seconds (optional)",
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="get_available_transitions",
            description="Get list of available transition effects",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="generate_outline",
            description="Generate a structured JSON outline based on a topic for subsequent PPT creation",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "topic": {"type": "string", "description": "Presentation topic"},
                },
                "required": ["file_path", "topic"],
            },
        ),
        Tool(
            name="set_core_properties",
            description="Set core document properties (title, author, subject, keywords, comments)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "title": {"type": "string", "description": "Document title"},
                    "subject": {"type": "string", "description": "Document subject"},
                    "author": {"type": "string", "description": "Document author"},
                    "keywords": {"type": "string", "description": "Document keywords"},
                    "comments": {"type": "string", "description": "Document comments"},
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="get_core_properties",
            description="Get core document properties (title, author, subject, etc.)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="get_slide_info",
            description="Get detailed information about a specific slide including placeholders and shapes",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="extract_slide_text",
            description="Extract all text content from a specific slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
        Tool(
            name="extract_presentation_text",
            description="Extract all text content from all slides in the presentation",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="populate_placeholder",
            description="Populate a placeholder with text by placeholder index",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "placeholder_idx": {
                        "type": "integer",
                        "description": "Placeholder index",
                    },
                    "text": {
                        "type": "string",
                        "description": "Text content to populate",
                    },
                },
                "required": ["file_path", "slide_index", "placeholder_idx", "text"],
            },
        ),
        Tool(
            name="add_text_box_advanced",
            description="Add a text box with advanced formatting options (font name, bold, italic, underline, alignment, background color)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "text": {"type": "string", "description": "Text content"},
                    "left": {
                        "type": "number",
                        "description": "Left margin (inches)",
                        "default": 1.0,
                    },
                    "top": {
                        "type": "number",
                        "description": "Top margin (inches)",
                        "default": 1.0,
                    },
                    "width": {
                        "type": "number",
                        "description": "Width (inches)",
                        "default": 4.0,
                    },
                    "height": {
                        "type": "number",
                        "description": "Height (inches)",
                        "default": 2.0,
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "Font size in points",
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font name (e.g. Arial, Segoe UI)",
                    },
                    "bold": {"type": "boolean", "description": "Bold text"},
                    "italic": {"type": "boolean", "description": "Italic text"},
                    "underline": {"type": "boolean", "description": "Underline text"},
                    "font_color": {
                        "type": "string",
                        "description": "Font color (hex, e.g. 000000)",
                    },
                    "bg_color": {
                        "type": "string",
                        "description": "Background color (hex, e.g. FFFFFF)",
                    },
                    "alignment": {
                        "type": "string",
                        "description": "Text alignment: left, center, right, justify",
                    },
                },
                "required": ["file_path", "slide_index", "text"],
            },
        ),
        Tool(
            name="format_text_runs",
            description="Format multiple text runs with different formatting in a shape",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index (starting from 0)",
                    },
                    "text_runs": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "text": {
                                    "type": "string",
                                    "description": "Text content for this run",
                                },
                                "bold": {"type": "boolean", "description": "Bold"},
                                "italic": {"type": "boolean", "description": "Italic"},
                                "underline": {
                                    "type": "boolean",
                                    "description": "Underline",
                                },
                                "font_size": {
                                    "type": "integer",
                                    "description": "Font size",
                                },
                                "font_name": {
                                    "type": "string",
                                    "description": "Font name",
                                },
                                "color": {
                                    "type": "string",
                                    "description": "Font color (hex or RGB list)",
                                },
                                "hyperlink": {
                                    "type": "string",
                                    "description": "Hyperlink URL",
                                },
                            },
                            "required": ["text"],
                        },
                        "description": "List of text runs with formatting",
                    },
                },
                "required": ["file_path", "slide_index", "shape_index", "text_runs"],
            },
        ),
        Tool(
            name="format_table_cell",
            description="Format a specific table cell (font, color, background, alignment)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index of the table (starting from 0)",
                    },
                    "row": {
                        "type": "integer",
                        "description": "Row index (starting from 0)",
                    },
                    "col": {
                        "type": "integer",
                        "description": "Column index (starting from 0)",
                    },
                    "font_size": {
                        "type": "integer",
                        "description": "Font size in points",
                    },
                    "font_name": {"type": "string", "description": "Font name"},
                    "bold": {"type": "boolean", "description": "Bold text"},
                    "italic": {"type": "boolean", "description": "Italic text"},
                    "font_color": {"type": "string", "description": "Font color (hex)"},
                    "bg_color": {
                        "type": "string",
                        "description": "Background color (hex)",
                    },
                    "alignment": {
                        "type": "string",
                        "description": "Text alignment: left, center, right, justify",
                    },
                },
                "required": ["file_path", "slide_index", "shape_index", "row", "col"],
            },
        ),
        Tool(
            name="add_chart",
            description="Add a chart to a slide (column, bar, line, pie, doughnut, area, scatter, radar)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "chart_type": {
                        "type": "string",
                        "description": "Chart type: column, stacked_column, bar, stacked_bar, line, line_markers, pie, doughnut, area, stacked_area, scatter, radar, radar_markers",
                    },
                    "left": {"type": "number", "description": "Left margin (inches)"},
                    "top": {"type": "number", "description": "Top margin (inches)"},
                    "width": {"type": "number", "description": "Chart width (inches)"},
                    "height": {
                        "type": "number",
                        "description": "Chart height (inches)",
                    },
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Category names",
                    },
                    "series_names": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Series names",
                    },
                    "series_values": {
                        "type": "array",
                        "items": {"type": "array", "items": {"type": "number"}},
                        "description": "Series values (list of number lists)",
                    },
                    "has_legend": {
                        "type": "boolean",
                        "description": "Show chart legend",
                        "default": True,
                    },
                    "title": {"type": "string", "description": "Chart title"},
                },
                "required": [
                    "file_path",
                    "slide_index",
                    "chart_type",
                    "left",
                    "top",
                    "width",
                    "height",
                    "categories",
                    "series_names",
                    "series_values",
                ],
            },
        ),
        Tool(
            name="update_chart_data",
            description="Replace existing chart data with new categories and series",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index of the chart (starting from 0)",
                    },
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "New category names",
                    },
                    "series_data": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": {
                                    "type": "string",
                                    "description": "Series name",
                                },
                                "values": {
                                    "type": "array",
                                    "items": {"type": "number"},
                                    "description": "Series values",
                                },
                            },
                            "required": ["name", "values"],
                        },
                        "description": "New series data",
                    },
                },
                "required": [
                    "file_path",
                    "slide_index",
                    "shape_index",
                    "categories",
                    "series_data",
                ],
            },
        ),
        Tool(
            name="add_connector",
            description="Add a connector line between two points on a slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "connector_type": {
                        "type": "string",
                        "description": "Connector type: straight, elbow, curved",
                    },
                    "start_x": {
                        "type": "number",
                        "description": "Starting X coordinate (inches)",
                    },
                    "start_y": {
                        "type": "number",
                        "description": "Starting Y coordinate (inches)",
                    },
                    "end_x": {
                        "type": "number",
                        "description": "Ending X coordinate (inches)",
                    },
                    "end_y": {
                        "type": "number",
                        "description": "Ending Y coordinate (inches)",
                    },
                    "line_width": {
                        "type": "number",
                        "description": "Line width in points",
                        "default": 1.0,
                    },
                    "color": {
                        "type": "string",
                        "description": "Line color (hex, e.g. 000000)",
                    },
                },
                "required": [
                    "file_path",
                    "slide_index",
                    "connector_type",
                    "start_x",
                    "start_y",
                    "end_x",
                    "end_y",
                ],
            },
        ),
        Tool(
            name="manage_hyperlinks",
            description="Manage hyperlinks in text shapes (add, remove, list, update)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "operation": {
                        "type": "string",
                        "description": "Operation: add, remove, list, update",
                    },
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index (starting from 0)",
                    },
                    "text": {
                        "type": "string",
                        "description": "Text for hyperlink (for add operation)",
                    },
                    "url": {"type": "string", "description": "URL for hyperlink"},
                    "run_index": {
                        "type": "integer",
                        "description": "Run index within shape (for update/remove)",
                        "default": 0,
                    },
                },
                "required": ["file_path", "operation", "slide_index"],
            },
        ),
        Tool(
            name="manage_slide_masters",
            description="Access and manage slide masters and layouts (list, get_layouts, get_info)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "operation": {
                        "type": "string",
                        "description": "Operation: list, get_layouts, get_info",
                    },
                    "master_index": {
                        "type": "integer",
                        "description": "Slide master index (default 0)",
                        "default": 0,
                    },
                    "layout_index": {
                        "type": "integer",
                        "description": "Layout index within master (for get_info)",
                    },
                },
                "required": ["file_path", "operation"],
            },
        ),
        Tool(
            name="add_shape_enhanced",
            description="Add a shape with enhanced options (more shape types, line color, line width, text inside shape)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_type": {
                        "type": "string",
                        "description": "Shape type: rectangle, rounded_rectangle, oval, diamond, triangle, right_triangle, pentagon, hexagon, heptagon, octagon, star, arrow, cloud, heart, lightning_bolt, sun, moon, smiley_face, no_symbol, flowchart_process, flowchart_decision, flowchart_data, flowchart_document",
                    },
                    "left": {"type": "number", "description": "Left margin (inches)"},
                    "top": {"type": "number", "description": "Top margin (inches)"},
                    "width": {"type": "number", "description": "Width (inches)"},
                    "height": {"type": "number", "description": "Height (inches)"},
                    "fill_color": {
                        "type": "string",
                        "description": "Fill color (hex, e.g. 0066CC)",
                    },
                    "line_color": {
                        "type": "string",
                        "description": "Line/border color (hex)",
                    },
                    "line_width": {
                        "type": "number",
                        "description": "Line width in points",
                    },
                    "text": {"type": "string", "description": "Text inside the shape"},
                    "font_size": {
                        "type": "integer",
                        "description": "Font size for shape text",
                    },
                    "font_color": {
                        "type": "string",
                        "description": "Font color for shape text (hex)",
                    },
                },
                "required": [
                    "file_path",
                    "slide_index",
                    "shape_type",
                    "left",
                    "top",
                    "width",
                    "height",
                ],
            },
        ),
        Tool(
            name="set_slide_gradient_background",
            description="Set a gradient background for a slide (requires Pillow)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "start_color": {
                        "type": "string",
                        "description": "Start color (hex, e.g. FFFFFF)",
                    },
                    "end_color": {
                        "type": "string",
                        "description": "End color (hex, e.g. 0000FF)",
                    },
                    "direction": {
                        "type": "string",
                        "description": "Gradient direction: horizontal, vertical, diagonal",
                        "default": "horizontal",
                    },
                },
                "required": ["file_path", "slide_index", "start_color", "end_color"],
            },
        ),
        Tool(
            name="apply_shape_effects",
            description="Apply visual effects to a shape (rotation, shadow, glow, etc.)",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "shape_index": {
                        "type": "integer",
                        "description": "Shape index (starting from 0)",
                    },
                    "effects": {
                        "type": "object",
                        "description": "Effects to apply. Keys are effect types (rotation, shadow, reflection, glow, soft_edges, transparency, bevel, filter), values are effect parameters.",
                        "properties": {
                            "rotation": {
                                "type": "object",
                                "properties": {
                                    "rotation": {
                                        "type": "number",
                                        "description": "Rotation angle in degrees",
                                    }
                                },
                            },
                            "shadow": {"type": "object"},
                            "reflection": {"type": "object"},
                            "glow": {"type": "object"},
                            "soft_edges": {"type": "object"},
                            "transparency": {"type": "object"},
                            "bevel": {"type": "object"},
                            "filter": {"type": "object"},
                        },
                    },
                },
                "required": ["file_path", "slide_index", "shape_index", "effects"],
            },
        ),
        Tool(
            name="list_slide_templates",
            description="List all available slide layout templates",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                },
                "required": ["file_path"],
            },
        ),
        Tool(
            name="get_template_info",
            description="Get detailed information about a specific slide template",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "template_id": {
                        "type": "string",
                        "description": "Template ID (e.g. title_slide, text_with_image)",
                    },
                },
                "required": ["file_path", "template_id"],
            },
        ),
        Tool(
            name="apply_slide_template",
            description="Apply a layout template to an existing slide",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "template_id": {
                        "type": "string",
                        "description": "Template ID (e.g. title_slide, text_with_image, two_column_text)",
                    },
                    "color_scheme": {
                        "type": "string",
                        "description": "Color scheme: modern_blue, corporate_gray, elegant_green, warm_red",
                        "default": "modern_blue",
                    },
                    "content_mapping": {
                        "type": "object",
                        "description": "Dictionary mapping element roles to content",
                    },
                    "image_paths": {
                        "type": "object",
                        "description": "Dictionary mapping image element roles to file paths",
                    },
                },
                "required": ["file_path", "slide_index", "template_id"],
            },
        ),
        Tool(
            name="create_slide_from_template",
            description="Create a new slide using a layout template",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "template_id": {
                        "type": "string",
                        "description": "Template ID (e.g. title_slide, text_with_image, two_column_text)",
                    },
                    "color_scheme": {
                        "type": "string",
                        "description": "Color scheme: modern_blue, corporate_gray, elegant_green, warm_red",
                        "default": "modern_blue",
                    },
                    "content_mapping": {
                        "type": "object",
                        "description": "Dictionary mapping element roles to content",
                    },
                    "image_paths": {
                        "type": "object",
                        "description": "Dictionary mapping image element roles to file paths",
                    },
                    "layout_index": {
                        "type": "integer",
                        "description": "PowerPoint layout index (default 1)",
                        "default": 1,
                    },
                },
                "required": ["file_path", "template_id"],
            },
        ),
        Tool(
            name="auto_generate_presentation",
            description="Automatically generate a presentation based on topic and preferences",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "topic": {
                        "type": "string",
                        "description": "Main topic/theme for the presentation",
                    },
                    "slide_count": {
                        "type": "integer",
                        "description": "Number of slides to generate (3-20)",
                        "default": 5,
                    },
                    "presentation_type": {
                        "type": "string",
                        "description": "Type: business, academic, creative",
                        "default": "business",
                    },
                    "color_scheme": {
                        "type": "string",
                        "description": "Color scheme: modern_blue, corporate_gray, elegant_green, warm_red",
                        "default": "modern_blue",
                    },
                },
                "required": ["file_path", "topic"],
            },
        ),
        Tool(
            name="save_slide_as_png",
            description="Save a slide from a PowerPoint presentation as a PNG image file",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "output_path": {
                        "type": "string",
                        "description": "Full path for the output PNG file",
                    },
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index to export (starting from 0, default=0)",
                        "default": 0,
                    },
                    "width": {
                        "type": "integer",
                        "description": "Output image width in pixels (default=1920)",
                        "default": 1920,
                    },
                    "height": {
                        "type": "integer",
                        "description": "Output image height in pixels (default=1080)",
                        "default": 1080,
                    },
                },
                "required": ["file_path", "output_path"],
            },
        ),
        Tool(
            name="optimize_slide_text",
            description="Optimize text elements on a slide for better readability",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": FILE_PATH_PARAM,
                    "slide_index": {
                        "type": "integer",
                        "description": "Slide index (starting from 0)",
                    },
                    "auto_resize": {
                        "type": "boolean",
                        "description": "Auto-resize fonts to fit containers",
                        "default": True,
                    },
                    "min_font_size": {
                        "type": "integer",
                        "description": "Minimum font size",
                        "default": 8,
                    },
                    "max_font_size": {
                        "type": "integer",
                        "description": "Maximum font size",
                        "default": 36,
                    },
                },
                "required": ["file_path", "slide_index"],
            },
        ),
    ]


def _get_file_path(arguments: dict) -> str:
    file_path = arguments.get("file_path", "")
    if not file_path:
        return ""
    return file_path


@server.call_tool()
async def handle_call_tool(name: str, arguments: dict):
    try:
        file_path = _get_file_path(arguments)
        if not file_path:
            return [
                TextContent(
                    type="text",
                    text=json.dumps(
                        {
                            "success": False,
                            "error": "Missing required parameter: file_path",
                        },
                        ensure_ascii=False,
                        indent=2,
                    ),
                )
            ]

        if name == "add_slide":
            layout_index = arguments.get("layout_index", 1)
            result = ppt_editor.add_slide(file_path, layout_index)

        elif name == "add_text_box":
            slide_index = arguments.get("slide_index")
            text = arguments.get("text")
            if slide_index is None or text is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or text",
                }
            else:
                result = ppt_editor.add_text_box(
                    file_path,
                    slide_index,
                    text,
                    arguments.get("left", 1),
                    arguments.get("top", 1),
                    arguments.get("width", 8),
                    arguments.get("height", 1),
                    arguments.get("font_size", 18),
                    arguments.get("font_color", "000000"),
                )

        elif name == "add_title_slide":
            title = arguments.get("title")
            if not title:
                result = {
                    "success": False,
                    "error": "Missing required parameter: title",
                }
            else:
                subtitle = arguments.get("subtitle", "")
                result = ppt_editor.add_title_slide(file_path, title, subtitle)

        elif name == "add_bullet_points":
            slide_index = arguments.get("slide_index")
            title = arguments.get("title")
            bullet_points = arguments.get("bullet_points")
            if slide_index is None or not title or not bullet_points:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, title or bullet_points",
                }
            else:
                result = ppt_editor.add_bullet_points(
                    file_path, slide_index, title, bullet_points
                )

        elif name == "add_image":
            slide_index = arguments.get("slide_index")
            image_path = arguments.get("image_path")
            if slide_index is None or not image_path:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or image_path",
                }
            else:
                result = ppt_editor.add_image(
                    file_path,
                    slide_index,
                    image_path,
                    arguments.get("left", 1),
                    arguments.get("top", 2),
                    arguments.get("width"),
                    arguments.get("height"),
                )

        elif name == "add_shape":
            slide_index = arguments.get("slide_index")
            shape_type = arguments.get("shape_type")
            if slide_index is None or not shape_type:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or shape_type",
                }
            else:
                result = ppt_editor.add_shape(
                    file_path,
                    slide_index,
                    shape_type,
                    arguments.get("left", 1),
                    arguments.get("top", 1),
                    arguments.get("width", 2),
                    arguments.get("height", 1),
                    arguments.get("fill_color", "0066CC"),
                )

        elif name == "get_presentation_info":
            result = ppt_editor.get_presentation_info(file_path)

        elif name == "delete_slide":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                result = ppt_editor.delete_slide(file_path, slide_index)

        elif name == "duplicate_slide":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                result = ppt_editor.duplicate_slide(file_path, slide_index)

        elif name == "move_slide":
            from_index = arguments.get("from_index")
            to_index = arguments.get("to_index")
            if from_index is None or to_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: from_index or to_index",
                }
            else:
                result = ppt_editor.move_slide(file_path, from_index, to_index)

        elif name == "add_table":
            slide_index = arguments.get("slide_index")
            rows = arguments.get("rows")
            cols = arguments.get("cols")
            if slide_index is None or rows is None or cols is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, rows or cols",
                }
            else:
                result = ppt_editor.add_table(
                    file_path,
                    slide_index,
                    rows,
                    cols,
                    arguments.get("left", 1),
                    arguments.get("top", 2),
                    arguments.get("width", 8),
                    arguments.get("height", 4),
                )

        elif name == "set_table_cell_text":
            slide_index = arguments.get("slide_index")
            table_index = arguments.get("table_index")
            row = arguments.get("row")
            col = arguments.get("col")
            text = arguments.get("text")
            required_params = {
                "slide_index": slide_index,
                "table_index": table_index,
                "row": row,
                "col": col,
                "text": text,
            }
            if any(v is None for v in required_params.values()):
                missing = [k for k, v in required_params.items() if v is None]
                result = {
                    "success": False,
                    "error": f"Missing required parameter: {', '.join(missing)}",
                }
            else:
                result = ppt_editor.set_table_cell_text(
                    file_path, slide_index, table_index, row, col, text
                )

        elif name == "set_slide_background_color":
            slide_index = arguments.get("slide_index")
            color = arguments.get("color")
            if slide_index is None or not color:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or color",
                }
            else:
                result = ppt_editor.set_slide_background_color(
                    file_path, slide_index, color
                )

        elif name == "add_hyperlink":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            url = arguments.get("url")
            if slide_index is None or shape_index is None or not url:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, shape_index or url",
                }
            else:
                display_text = arguments.get("display_text")
                result = ppt_editor.add_hyperlink(
                    file_path, slide_index, shape_index, url, display_text
                )

        elif name == "set_text_formatting":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            if slide_index is None or shape_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or shape_index",
                }
            else:
                result = ppt_editor.set_text_formatting(
                    file_path,
                    slide_index,
                    shape_index,
                    arguments.get("font_name"),
                    arguments.get("font_size"),
                    arguments.get("font_color"),
                    arguments.get("bold"),
                    arguments.get("italic"),
                    arguments.get("underline"),
                )

        elif name == "get_slide_shapes_info":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                result = ppt_editor.get_slide_shapes_info(file_path, slide_index)

        elif name == "add_slide_animation":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                animation_style = arguments.get("animation_style", "fade")
                speed = arguments.get("speed", "medium")
                auto_advance = arguments.get("auto_advance", False)
                auto_advance_seconds = arguments.get("auto_advance_seconds", 3.0)
                speed_mapping = {"fast": 0.5, "medium": 1.0, "slow": 2.0}
                duration = speed_mapping.get(speed, 1.0)
                advance_after_time = auto_advance_seconds if auto_advance else None
                result = ppt_editor.set_slide_transition(
                    file_path,
                    slide_index,
                    animation_style,
                    duration,
                    True,
                    advance_after_time,
                )

        elif name == "make_presentation_dynamic":
            animation_style = arguments.get("animation_style", "fade")
            speed = arguments.get("speed", "medium")
            speed_mapping = {"fast": 0.5, "medium": 1.0, "slow": 2.0}
            duration = speed_mapping.get(speed, 1.0)
            result = ppt_editor.apply_transition_to_all_slides(
                file_path, animation_style, duration
            )

        elif name == "get_animation_options":
            result = ppt_editor.get_available_transitions(file_path)

        elif name == "make_professional_presentation":
            result = ppt_editor.make_presentation_professional(file_path)

        elif name == "add_smooth_transitions":
            result = ppt_editor.add_smooth_transitions(file_path)

        elif name == "add_dynamic_effects":
            result = ppt_editor.add_dynamic_effects(file_path)

        elif name == "set_slide_transition":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                transition_type = arguments.get("transition_type", "fade")
                duration = arguments.get("duration", 1.0)
                advance_on_click = arguments.get("advance_on_click", True)
                advance_after_time = arguments.get("advance_after_time")
                result = ppt_editor.set_slide_transition(
                    file_path,
                    slide_index,
                    transition_type,
                    duration,
                    advance_on_click,
                    advance_after_time,
                )

        elif name == "get_available_transitions":
            result = ppt_editor.get_available_transitions(file_path)

        elif name == "generate_outline":
            topic = arguments.get("topic")
            if not topic:
                result = {
                    "success": False,
                    "error": "Missing required parameter: topic",
                }
            else:
                result = ppt_editor.generate_outline_for_topic(file_path, topic)

        elif name == "set_core_properties":
            result = ppt_editor.set_core_properties(
                file_path,
                title=arguments.get("title"),
                subject=arguments.get("subject"),
                author=arguments.get("author"),
                keywords=arguments.get("keywords"),
                comments=arguments.get("comments"),
            )

        elif name == "get_core_properties":
            result = ppt_editor.get_core_properties(file_path)

        elif name == "get_slide_info":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                result = ppt_editor.get_slide_info(file_path, slide_index)

        elif name == "extract_slide_text":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                result = ppt_editor.extract_slide_text(file_path, slide_index)

        elif name == "extract_presentation_text":
            result = ppt_editor.extract_presentation_text(file_path)

        elif name == "populate_placeholder":
            slide_index = arguments.get("slide_index")
            placeholder_idx = arguments.get("placeholder_idx")
            text = arguments.get("text")
            if slide_index is None or placeholder_idx is None or text is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, placeholder_idx, or text",
                }
            else:
                result = ppt_editor.populate_placeholder(
                    file_path, slide_index, placeholder_idx, text
                )

        elif name == "add_text_box_advanced":
            slide_index = arguments.get("slide_index")
            text = arguments.get("text")
            if slide_index is None or text is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or text",
                }
            else:
                result = ppt_editor.add_text_box_advanced(
                    file_path,
                    slide_index,
                    text,
                    left=arguments.get("left", 1.0),
                    top=arguments.get("top", 1.0),
                    width=arguments.get("width", 4.0),
                    height=arguments.get("height", 2.0),
                    font_size=arguments.get("font_size"),
                    font_name=arguments.get("font_name"),
                    bold=arguments.get("bold"),
                    italic=arguments.get("italic"),
                    underline=arguments.get("underline"),
                    font_color=arguments.get("font_color"),
                    bg_color=arguments.get("bg_color"),
                    alignment=arguments.get("alignment"),
                )

        elif name == "format_text_runs":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            text_runs = arguments.get("text_runs")
            if slide_index is None or shape_index is None or not text_runs:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, shape_index, or text_runs",
                }
            else:
                result = ppt_editor.format_text_runs(
                    file_path, slide_index, shape_index, text_runs
                )

        elif name == "format_table_cell":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            row = arguments.get("row")
            col = arguments.get("col")
            if slide_index is None or shape_index is None or row is None or col is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, shape_index, row, or col",
                }
            else:
                result = ppt_editor.format_table_cell(
                    file_path,
                    slide_index,
                    shape_index,
                    row,
                    col,
                    font_size=arguments.get("font_size"),
                    font_name=arguments.get("font_name"),
                    bold=arguments.get("bold"),
                    italic=arguments.get("italic"),
                    font_color=arguments.get("font_color"),
                    bg_color=arguments.get("bg_color"),
                    alignment=arguments.get("alignment"),
                )

        elif name == "add_chart":
            slide_index = arguments.get("slide_index")
            chart_type = arguments.get("chart_type")
            categories = arguments.get("categories", [])
            series_names = arguments.get("series_names", [])
            series_values = arguments.get("series_values", [])
            if slide_index is None or not chart_type:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or chart_type",
                }
            else:
                result = ppt_editor.add_chart(
                    file_path,
                    slide_index,
                    chart_type,
                    left=arguments.get("left", 1.0),
                    top=arguments.get("top", 2.0),
                    width=arguments.get("width", 8.0),
                    height=arguments.get("height", 4.5),
                    categories=categories,
                    series_names=series_names,
                    series_values=series_values,
                    has_legend=arguments.get("has_legend", True),
                    title=arguments.get("title"),
                )

        elif name == "update_chart_data":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            categories = arguments.get("categories", [])
            series_data = arguments.get("series_data", [])
            if slide_index is None or shape_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or shape_index",
                }
            else:
                result = ppt_editor.update_chart_data(
                    file_path, slide_index, shape_index, categories, series_data
                )

        elif name == "add_connector":
            slide_index = arguments.get("slide_index")
            connector_type = arguments.get("connector_type")
            start_x = arguments.get("start_x")
            start_y = arguments.get("start_y")
            end_x = arguments.get("end_x")
            end_y = arguments.get("end_y")
            if (
                slide_index is None
                or not connector_type
                or start_x is None
                or start_y is None
                or end_x is None
                or end_y is None
            ):
                result = {
                    "success": False,
                    "error": "Missing required parameters for add_connector",
                }
            else:
                result = ppt_editor.add_connector(
                    file_path,
                    slide_index,
                    connector_type,
                    start_x,
                    start_y,
                    end_x,
                    end_y,
                    line_width=arguments.get("line_width", 1.0),
                    color=arguments.get("color"),
                )

        elif name == "manage_hyperlinks":
            operation = arguments.get("operation")
            slide_index = arguments.get("slide_index")
            if not operation or slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: operation or slide_index",
                }
            else:
                result = ppt_editor.manage_hyperlinks(
                    file_path,
                    operation,
                    slide_index,
                    shape_index=arguments.get("shape_index"),
                    text=arguments.get("text"),
                    url=arguments.get("url"),
                    run_index=arguments.get("run_index", 0),
                )

        elif name == "manage_slide_masters":
            operation = arguments.get("operation")
            if not operation:
                result = {
                    "success": False,
                    "error": "Missing required parameter: operation",
                }
            else:
                result = ppt_editor.manage_slide_masters(
                    file_path,
                    operation,
                    master_index=arguments.get("master_index", 0),
                    layout_index=arguments.get("layout_index"),
                )

        elif name == "add_shape_enhanced":
            slide_index = arguments.get("slide_index")
            shape_type = arguments.get("shape_type")
            left = arguments.get("left")
            top = arguments.get("top")
            width = arguments.get("width")
            height = arguments.get("height")
            if (
                slide_index is None
                or not shape_type
                or left is None
                or top is None
                or width is None
                or height is None
            ):
                result = {
                    "success": False,
                    "error": "Missing required parameters for add_shape_enhanced",
                }
            else:
                result = ppt_editor.add_shape_enhanced(
                    file_path,
                    slide_index,
                    shape_type,
                    left,
                    top,
                    width,
                    height,
                    fill_color=arguments.get("fill_color"),
                    line_color=arguments.get("line_color"),
                    line_width=arguments.get("line_width"),
                    text=arguments.get("text"),
                    font_size=arguments.get("font_size"),
                    font_color=arguments.get("font_color"),
                )

        elif name == "set_slide_gradient_background":
            slide_index = arguments.get("slide_index")
            start_color = arguments.get("start_color")
            end_color = arguments.get("end_color")
            if slide_index is None or not start_color or not end_color:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index, start_color, or end_color",
                }
            else:
                result = ppt_editor.set_slide_gradient_background(
                    file_path,
                    slide_index,
                    start_color,
                    end_color,
                    direction=arguments.get("direction", "horizontal"),
                )

        elif name == "apply_shape_effects":
            slide_index = arguments.get("slide_index")
            shape_index = arguments.get("shape_index")
            effects = arguments.get("effects", {})
            if slide_index is None or shape_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or shape_index",
                }
            else:
                result = ppt_editor.apply_shape_effects(
                    file_path, slide_index, shape_index, effects
                )

        elif name == "list_slide_templates":
            result = ppt_editor.list_slide_templates()

        elif name == "get_template_info":
            template_id = arguments.get("template_id")
            if not template_id:
                result = {
                    "success": False,
                    "error": "Missing required parameter: template_id",
                }
            else:
                result = ppt_editor.get_template_info(template_id)

        elif name == "apply_slide_template":
            slide_index = arguments.get("slide_index")
            template_id = arguments.get("template_id")
            if slide_index is None or not template_id:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index or template_id",
                }
            else:
                result = ppt_editor.apply_slide_template(
                    file_path,
                    slide_index,
                    template_id,
                    color_scheme=arguments.get("color_scheme", "modern_blue"),
                    content_mapping=arguments.get("content_mapping"),
                    image_paths=arguments.get("image_paths"),
                )

        elif name == "create_slide_from_template":
            template_id = arguments.get("template_id")
            if not template_id:
                result = {
                    "success": False,
                    "error": "Missing required parameter: template_id",
                }
            else:
                result = ppt_editor.create_slide_from_template(
                    file_path,
                    template_id,
                    color_scheme=arguments.get("color_scheme", "modern_blue"),
                    content_mapping=arguments.get("content_mapping"),
                    image_paths=arguments.get("image_paths"),
                    layout_index=arguments.get("layout_index", 1),
                )

        elif name == "auto_generate_presentation":
            topic = arguments.get("topic")
            if not topic:
                result = {
                    "success": False,
                    "error": "Missing required parameter: topic",
                }
            else:
                result = ppt_editor.auto_generate_presentation(
                    file_path,
                    topic,
                    slide_count=arguments.get("slide_count", 5),
                    presentation_type=arguments.get("presentation_type", "business"),
                    color_scheme=arguments.get("color_scheme", "modern_blue"),
                )

        elif name == "save_slide_as_png":
            output_path = arguments.get("output_path")
            if not output_path:
                result = {
                    "success": False,
                    "error": "Missing required parameter: output_path",
                }
            else:
                result = ppt_editor.save_slide_as_png(
                    file_path,
                    output_path,
                    slide_index=arguments.get("slide_index", 0),
                    width=arguments.get("width", 1920),
                    height=arguments.get("height", 1080),
                )

        elif name == "optimize_slide_text":
            slide_index = arguments.get("slide_index")
            if slide_index is None:
                result = {
                    "success": False,
                    "error": "Missing required parameter: slide_index",
                }
            else:
                result = ppt_editor.optimize_slide_text(
                    file_path,
                    slide_index,
                    auto_resize=arguments.get("auto_resize", True),
                    min_font_size=arguments.get("min_font_size", 8),
                    max_font_size=arguments.get("max_font_size", 36),
                )

        else:
            result = {"success": False, "error": f"Unknown tool: {name}"}

        return [
            TextContent(
                type="text", text=json.dumps(result, ensure_ascii=False, indent=2)
            )
        ]

    except Exception as e:
        logger.error(f"Tool call error: {e}")
        error_result = {"success": False, "error": str(e)}
        return [
            TextContent(
                type="text", text=json.dumps(error_result, ensure_ascii=False, indent=2)
            )
        ]


async def main():
    from contextlib import AsyncExitStack

    async with AsyncExitStack() as stack:
        streams = await stack.enter_async_context(stdio_server())
        read_stream, write_stream = streams
        await server.run(
            read_stream, write_stream, server.create_initialization_options()
        )


if __name__ == "__main__":
    asyncio.run(main())
