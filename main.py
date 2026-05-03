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
