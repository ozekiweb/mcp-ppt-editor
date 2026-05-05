# PowerPoint Editing MCP Server

This is a PowerPoint editing server based on MCP (Model Context Protocol) that provides comprehensive functionality for creating and editing PowerPoint presentations, including content editing, formatting, charts, connectors, templates, and professional animation effects.

## Key Design

Every tool call requires a `file_path` parameter as the first argument — the full path to the `.pptx` file. The server automatically:

1. **Loads** the presentation if the file exists
2. **Creates** a new presentation if the file does not exist (parent directories are created as needed)
3. **Saves** the presentation back to `file_path` after each operation

This means there are no separate `create_presentation`, `open_presentation`, or `save_presentation` tools — every tool handles loading and saving automatically.

## Project Structure

- `main.py` - MCP server main program, handles MCP protocol communication
- `tool.py` - PowerPoint editor tool class, contains all PPT editing functionality
- `slide_layout_templates.json` - Template definitions for professional slide layouts
- `example.py` - Usage examples
- `requirements.txt` - Project dependencies
- `mcp_config.json` - MCP client configuration file

## Features

### Core Design
- File-based workflow: every operation targets a specific `.pptx` file path
- Auto-load existing files or create new ones
- Auto-save after every mutation

### Slide Operations
- Add, delete, duplicate, and move slides
- Get detailed slide info (placeholders, shapes, layout)
- Extract text from individual slides or entire presentations
- Populate slide placeholders

### Content Editing
- Add text boxes (basic and advanced with full formatting)
- Add title slides and bulleted content
- Insert images
- Add shapes (basic and enhanced with many shape types)
- Add tables and format table cells
- Add charts (column, bar, line, pie, scatter, etc.)
- Add connector lines (straight, elbow, curved)
- Format text runs with per-run styling

### Formatting Features
- Set text formatting (font, size, color, bold, italic, underline)
- Set slide background colors and gradient backgrounds
- Apply shape effects (rotation, shadow, etc.)
- Add and manage hyperlinks (add, update, remove, list)

### Presentation Metadata
- Set and get core properties (title, author, keywords, etc.)

### Template System
- List available slide templates
- Get detailed template info
- Apply templates to existing slides
- Create new slides from templates
- Auto-generate complete presentations from a topic

### Text Optimization
- Auto-optimize slide text (font sizing, line spacing)

### Professional Animations and Transitions
- One-click professionalization — add fade transitions to all slides
- Multiple animation styles: fade, push, wipe, split, zoom, blinds, dissolve
- Smart speed control: fast, medium, slow
- Auto-advance support
- Batch application to all slides
- Preset convenience functions for smooth transitions and dynamic effects

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Running as an MCP Server

```bash
python main.py
```

### Using the PowerPointEditor Class Directly

```python
from tool import PowerPointEditor

editor = PowerPointEditor()
file_path = "my_presentation.pptx"

# Add a title slide (creates new file since it doesn't exist)
editor.add_title_slide(file_path, "My Presentation", "Subtitle")

# Add a content slide (loads existing file, adds slide, saves)
editor.add_slide(file_path, 1)

# Add bullet points
editor.add_bullet_points(file_path, 1, "Main Content", ["Point 1", "Point 2"])

# All changes are automatically saved to file_path after each call
```

### Running Examples

```bash
python example.py
```

## Available Tools

All tools require `file_path` as the first parameter (unless noted). The file is loaded if it exists, or a new presentation is created if it doesn't. After each operation, the file is automatically saved.

### Common Parameter

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `file_path` | string | Yes | Full path to the PowerPoint file (.pptx). Loaded if exists, created if not. |

---

### Slide Operations

#### add_slide
Add a new slide to the presentation.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `layout_index` | integer | No | 1 | Slide layout index (0 = title slide, 1 = title and content) |

#### add_title_slide
Add a title slide to the presentation.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `title` | string | Yes | — | Slide title |
| `subtitle` | string | No | "" | Slide subtitle |

#### delete_slide
Delete a slide from the presentation.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Index of slide to delete (0-based) |

#### duplicate_slide
Duplicate a slide in the presentation.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Index of slide to duplicate (0-based) |

#### move_slide
Move a slide to a new position.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `from_index` | integer | Yes | — | Source position index (0-based) |
| `to_index` | integer | Yes | — | Target position index (0-based) |

#### get_slide_info
Get detailed information about a specific slide including layout, placeholders, and shapes.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |

#### extract_slide_text
Extract all text content from a specific slide, including title, placeholders, text shapes, and table text.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |

#### extract_presentation_text
Extract all text content from every slide in the presentation.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

#### populate_placeholder
Populate a placeholder on a slide with text.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `placeholder_idx` | integer | Yes | — | Placeholder index |
| `text` | string | Yes | — | Text to populate |

---

### Content Editing

#### add_text_box
Add a text box to a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `text` | string | Yes | — | Text content |
| `left` | number | No | 1 | Left margin (inches) |
| `top` | number | No | 1 | Top margin (inches) |
| `width` | number | No | 8 | Width (inches) |
| `height` | number | No | 1 | Height (inches) |
| `font_size` | integer | No | 18 | Font size |
| `font_color` | string | No | "000000" | Font color (hex, e.g. FF0000) |

#### add_text_box_advanced
Add a text box with advanced formatting options (font name, bold, italic, underline, background color, alignment).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `text` | string | Yes | — | Text content |
| `left` | number | No | 1.0 | Left margin (inches) |
| `top` | number | No | 1.0 | Top margin (inches) |
| `width` | number | No | 4.0 | Width (inches) |
| `height` | number | No | 2.0 | Height (inches) |
| `font_size` | integer | No | — | Font size |
| `font_name` | string | No | — | Font name (e.g. Arial, Segoe UI) |
| `bold` | boolean | No | — | Bold text |
| `italic` | boolean | No | — | Italic text |
| `underline` | boolean | No | — | Underline text |
| `font_color` | string | No | — | Font color (hex) |
| `bg_color` | string | No | — | Background color (hex) |
| `alignment` | string | No | — | Text alignment: left, center, right, justify |

#### add_bullet_points
Add bulleted content to a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `title` | string | Yes | — | Slide title |
| `bullet_points` | array of strings | Yes | — | List of bullet points |

#### add_image
Add an image to a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `image_path` | string | Yes | — | Image file path |
| `left` | number | No | 1 | Left margin (inches) |
| `top` | number | No | 2 | Top margin (inches) |
| `width` | number | No | — | Width (inches, optional) |
| `height` | number | No | — | Height (inches, optional) |

#### add_shape
Add a shape to a slide (basic shapes).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_type` | string | Yes | — | Shape type: rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow |
| `left` | number | No | 1 | Left margin (inches) |
| `top` | number | No | 1 | Top margin (inches) |
| `width` | number | No | 2 | Width (inches) |
| `height` | number | No | 1 | Height (inches) |
| `fill_color` | string | No | "0066CC" | Fill color (hex) |

#### add_shape_enhanced
Add a shape with enhanced options including line color, line width, text, and font styling. Supports many more shape types.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_type` | string | Yes | — | Shape type: rectangle, rounded_rectangle, oval, diamond, triangle, right_triangle, pentagon, hexagon, heptagon, octagon, star, arrow, cloud, heart, lightning_bolt, sun, moon, smiley_face, no_symbol, flowchart_process, flowchart_decision, flowchart_data, flowchart_document |
| `left` | number | Yes | — | Left position (inches) |
| `top` | number | Yes | — | Top position (inches) |
| `width` | number | Yes | — | Width (inches) |
| `height` | number | Yes | — | Height (inches) |
| `fill_color` | string | No | — | Fill color (hex) |
| `line_color` | string | No | — | Line/border color (hex) |
| `line_width` | number | No | — | Line/border width (points) |
| `text` | string | No | — | Text inside the shape |
| `font_size` | integer | No | — | Font size for text |
| `font_color` | string | No | — | Font color for text (hex) |

#### add_table
Add a table to a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `rows` | integer | Yes | — | Number of rows |
| `cols` | integer | Yes | — | Number of columns |
| `left` | number | No | 1 | Left margin (inches) |
| `top` | number | No | 2 | Top margin (inches) |
| `width` | number | No | 8 | Width (inches) |
| `height` | number | No | 4 | Height (inches) |

#### set_table_cell_text
Set the text of a table cell.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `table_index` | integer | Yes | — | Table index on the slide (0-based) |
| `row` | integer | Yes | — | Row index (0-based) |
| `col` | integer | Yes | — | Column index (0-based) |
| `text` | string | Yes | — | Text content |

#### format_table_cell
Format a table cell with font styling and background color.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | Yes | — | Shape index of the table (0-based) |
| `row` | integer | Yes | — | Row index (0-based) |
| `col` | integer | Yes | — | Column index (0-based) |
| `font_size` | integer | No | — | Font size |
| `font_name` | string | No | — | Font name (e.g. Arial) |
| `bold` | boolean | No | — | Bold text |
| `italic` | boolean | No | — | Italic text |
| `font_color` | string | No | — | Font color (hex) |
| `bg_color` | string | No | — | Cell background color (hex) |
| `alignment` | string | No | — | Text alignment: left, center, right, justify |

#### add_chart
Add a chart to a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `chart_type` | string | Yes | — | Chart type: column, stacked_column, bar, stacked_bar, line, line_markers, pie, doughnut, area, stacked_area, scatter, radar, radar_markers |
| `left` | number | Yes | — | Left position (inches) |
| `top` | number | Yes | — | Top position (inches) |
| `width` | number | Yes | — | Width (inches) |
| `height` | number | Yes | — | Height (inches) |
| `categories` | array of strings | Yes | — | Category labels |
| `series_names` | array of strings | Yes | — | Names for each data series |
| `series_values` | array of arrays of numbers | Yes | — | Data values for each series (must match series_names length) |
| `has_legend` | boolean | No | true | Show chart legend |
| `title` | string | No | — | Chart title |

#### update_chart_data
Update the data of an existing chart.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | Yes | — | Shape index of the chart (0-based) |
| `categories` | array of strings | Yes | — | New category labels |
| `series_data` | array of objects | Yes | — | Each object must have `name` (string) and `values` (array of numbers) |

#### add_connector
Add a connector line between two points on a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `connector_type` | string | Yes | — | Connector type: straight, elbow, curved |
| `start_x` | number | Yes | — | Start X position (inches) |
| `start_y` | number | Yes | — | Start Y position (inches) |
| `end_x` | number | Yes | — | End X position (inches) |
| `end_y` | number | Yes | — | End Y position (inches) |
| `line_width` | number | No | 1.0 | Line width (points) |
| `color` | string | No | — | Line color (hex) |

---

### Formatting and Style

#### set_slide_background_color
Set the background color of a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `color` | string | Yes | — | Background color (hex, e.g. E6F3FF) |

#### set_slide_gradient_background
Set a gradient background on a slide (requires Pillow).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `start_color` | string | Yes | — | Gradient start color (hex) |
| `end_color` | string | Yes | — | Gradient end color (hex) |
| `direction` | string | No | "horizontal" | Gradient direction: horizontal, vertical, diagonal |

#### add_hyperlink
Add a hyperlink to a shape.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | Yes | — | Shape index (0-based) |
| `url` | string | Yes | — | Hyperlink URL |
| `display_text` | string | No | — | Display text for the hyperlink |

#### manage_hyperlinks
Add, update, remove, or list hyperlinks on a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `operation` | string | Yes | — | Operation: add, update, remove, list |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | No | — | Shape index (0-based, required for add/update/remove) |
| `text` | string | No | — | Display text (required for add) |
| `url` | string | No | — | URL (required for add and update) |
| `run_index` | integer | No | 0 | Run index within the shape's text (for update/remove) |

#### set_text_formatting
Set text formatting for a shape.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | Yes | — | Shape index (0-based) |
| `font_name` | string | No | — | Font name (e.g. Arial) |
| `font_size` | integer | No | — | Font size |
| `font_color` | string | No | — | Font color (hex) |
| `bold` | boolean | No | — | Bold text |
| `italic` | boolean | No | — | Italic text |
| `underline` | boolean | No | — | Underline text |

#### format_text_runs
Replace all text in a shape with individually formatted text runs. Each run can have its own font styling and optional hyperlink.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | Yes | — | Shape index (0-based) |
| `text_runs` | array of objects | Yes | — | Each object: `text` (string, required), plus optional `bold`, `italic`, `underline` (boolean), `font_size` (integer), `font_name` (string), `color` (hex string or [R,G,B] array), `hyperlink` (string URL) |

#### apply_shape_effects
Apply visual effects to a shape (currently supports rotation; other effects are placeholders).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `shape_index` | integer | Yes | — | Shape index (0-based) |
| `effects` | object | Yes | — | Effects to apply. Keys are effect types (e.g. `rotation`, `shadow`, `reflection`), values are parameter objects. Supported: `rotation` with `{"rotation": float}` |

---

### Presentation Metadata

#### set_core_properties
Set core document properties (title, author, etc.).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `title` | string | No | — | Document title |
| `subject` | string | No | — | Document subject |
| `author` | string | No | — | Document author |
| `keywords` | string | No | — | Document keywords |
| `comments` | string | No | — | Document comments |

#### get_core_properties
Get core document properties (title, author, created/modified dates, etc.).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

---

### Slide Masters

#### manage_slide_masters
List, inspect, and get information about slide masters and their layouts.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `operation` | string | Yes | — | Operation: list, get_layouts, get_info |
| `master_index` | integer | No | 0 | Slide master index |
| `layout_index` | integer | No | — | Layout index (required for get_info to get specific layout details) |

---

### Template System

#### list_slide_templates
List all available slide templates. Does not require `file_path`.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| *(no parameters)* | | | | |

#### get_template_info
Get detailed information about a specific template.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `template_id` | string | Yes | — | Template ID (e.g. title_slide, agenda_slide, two_column_text) |

#### apply_slide_template
Apply a template to an existing slide with optional color scheme and content mapping.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `template_id` | string | Yes | — | Template ID |
| `color_scheme` | string | No | "modern_blue" | Color scheme name |
| `content_mapping` | object | No | — | Map of element roles to custom text content |
| `image_paths` | object | No | — | Map of element roles to image file paths |

#### create_slide_from_template
Create a new slide using a template.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `template_id` | string | Yes | — | Template ID |
| `color_scheme` | string | No | "modern_blue" | Color scheme name |
| `content_mapping` | object | No | — | Map of element roles to custom text content |
| `image_paths` | object | No | — | Map of element roles to image file paths |
| `layout_index` | integer | No | 1 | Slide layout index |

#### auto_generate_presentation
Auto-generate a complete presentation from a topic using templates.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `topic` | string | Yes | — | Presentation topic |
| `slide_count` | integer | No | 5 | Number of slides (3–20) |
| `presentation_type` | string | No | "business" | Type: business, academic, creative |
| `color_scheme` | string | No | "modern_blue" | Color scheme name |

---

### Text Optimization

#### optimize_slide_text
Auto-optimize text on a slide (adjust font sizes and line spacing).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `auto_resize` | boolean | No | true | Automatically resize fonts to fit |
| `min_font_size` | integer | No | 8 | Minimum font size (points) |
| `max_font_size` | integer | No | 36 | Maximum font size (points) |

---

### Information Tools

#### get_presentation_info
Get presentation information (slide count, titles, etc.).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

#### get_slide_shapes_info
Get information about all shapes on a slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |

---

### Animation and Transition Tools

#### add_slide_animation
Add animation transition effects to a single slide.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `animation_style` | string | No | "fade" | Animation style: fade, push, wipe, zoom, split, blinds, dissolve, none |
| `speed` | string | No | "medium" | Animation speed: fast, medium, slow |
| `auto_advance` | boolean | No | false | Auto-advance to next slide |
| `auto_advance_seconds` | number | No | 3.0 | Auto-advance delay (seconds) |

#### make_presentation_dynamic
Add uniform animation effects to all slides.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `animation_style` | string | No | "fade" | Animation style: fade, push, wipe, zoom |
| `speed` | string | No | "medium" | Animation speed: fast, medium, slow |

#### make_professional_presentation
One-click professionalization — add elegant fade transition effects to all slides.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

#### add_smooth_transitions
Add smooth fade transitions (0.8s) to all slides.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

#### add_dynamic_effects
Add dynamic push transitions (1.2s) to all slides.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

#### get_animation_options
View all available slide animation effect options.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

#### set_slide_transition
Set slide transition effect (advanced).
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `slide_index` | integer | Yes | — | Slide index (0-based) |
| `transition_type` | string | No | "fade" | Transition type: none, fade, push, wipe, split, zoom, blinds, dissolve |
| `duration` | number | No | 1.0 | Transition duration (seconds) |
| `advance_on_click` | boolean | No | true | Advance on click |
| `advance_after_time` | number | No | — | Auto-advance after this many seconds |

#### get_available_transitions
Get a list of available transition effects.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |

---

### Outline Generation

#### generate_outline
Generate a structured JSON outline based on a topic for subsequent PPT creation.
| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `file_path` | string | Yes | — | Full path to the .pptx file |
| `topic` | string | Yes | — | Presentation topic |

---

## Recommended Workflow

1. **Create structure** — Use `auto_generate_presentation` or `create_slide_from_template` for quick setup, or add slides manually
2. **Add content** — Call tools with `file_path` to add text, shapes, charts, connectors, etc.
3. **Format and style** — Use `add_text_box_advanced`, `format_text_runs`, `format_table_cell`, `set_slide_gradient_background`, `apply_shape_effects`
4. **One-click professionalization** — Use `make_professional_presentation` to add transitions
5. **Personalize** — Set different animation effects for specific slides as needed
6. **Review** — Use `get_presentation_info`, `get_slide_info`, `extract_presentation_text` to inspect
7. The file at `file_path` is always up-to-date after each tool call

## Important Notes

1. Every tool requires `file_path` as the first parameter — there are no separate create/open/save tools
2. If `file_path` exists, the file is loaded; if not, a new presentation is created
3. Parent directories are created automatically if they don't exist
4. The presentation is saved to `file_path` after every mutating operation
5. Make sure all required dependencies are installed (especially `lxml` for animation features and `Pillow` for gradient backgrounds)
6. Image file paths must exist and be accessible
7. Slide indices start from 0
8. Colors use hexadecimal format without `#` (e.g., `000000` for black, `FF0000` for red)
9. Position and size units are in inches
10. Animation effects require opening the file in PowerPoint to see the full effect
11. `list_slide_templates` and `get_template_info` do not require `file_path` — they read from the built-in template file

## Error Handling

All operations return responses in this format:
```json
{
  "success": true/false,
  "message": "Operation result message",
  "error": "Error message (if any)"
}
```

## License

MIT License