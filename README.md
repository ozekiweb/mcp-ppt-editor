# PowerPoint Editing MCP Server

This is a PowerPoint editing server based on MCP (Model Context Protocol) that provides comprehensive functionality for creating and editing PowerPoint presentations, including content editing, formatting, and professional animation effects.

## ✨ Latest Updates

- 🎬 **New Animation System** - Added multiple professional transition animation effects
- 🚀 **One-Click Professionalization** - Quickly make presentations look professional
- 🎯 **Smart Recommendations** - Optimized tool descriptions for better AI model usage
- 🛠️ **Convenience Functions** - Simplified complex operations with intuitive parameter interfaces

## Project Structure

- `main.py` - MCP server main program, handles MCP protocol communication
- `tool.py` - PowerPoint editor tool class, contains all PPT editing functionality
- `example.py` - Usage examples
- `test_transitions.py` - Transition animation feature tests
- `transition_improvements_guide.md` - Animation feature improvement guide
- `requirements.txt` - Project dependencies
- `mcp_config.json` - MCP client configuration file

## Features

### Basic Features
- Create new PowerPoint presentations
- Open existing PowerPoint files
- Save presentations
- Get presentation information

### Slide Operations
- Add new slides (supports different layouts)
- Delete slides
- Duplicate slides
- Move slide positions
- Set slide background colors

### Content Editing
- Add text boxes and text content
- Add title slides
- Add bulleted content
- Insert images
- Add various shapes (rectangle, oval, triangle, etc.)
- Add tables
- Set table cell text

### Formatting Features
- Set text formatting (font, size, color, bold, italic, underline)
- Add hyperlinks to shapes
- Get detailed information about all shapes on a slide

### 🎬 Professional Animations and Transitions
- **One-Click Professionalization** - Quickly add professional transitions to the entire presentation
- **Multiple Animation Styles** - Fade, push, wipe, split, zoom, blinds, dissolve, and 8 other effects
- **Smart Speed Control** - Fast, medium, and slow speed options
- **Auto-Advance Support** - Supports both auto-advance and click-to-advance
- **Batch Application** - Apply uniform animations to all slides at once
- **Convenience Functions** - Preset options for smooth transitions and dynamic effects

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

# Create an editor instance
editor = PowerPointEditor()

# Create a new presentation
editor.create_presentation()

# Add a title slide
editor.add_title_slide("My Presentation", "Subtitle")

# Save the file
editor.save_presentation("my_presentation.pptx")
```

### Running Examples

```bash
python example.py
```

### Testing Animation Features

```bash
python test_transitions.py
```

## 🎬 Animation Quick Start

```python
from tool import PowerPointEditor

editor = PowerPointEditor()
editor.create_presentation()

# Add a few slides
editor.add_title_slide("Welcome", "My Presentation")
editor.add_title_slide("Content", "Main Content")
editor.add_title_slide("End", "Thank You")

# One-click professionalization - add fade effect to all slides
editor.make_presentation_professional()

# Or add dynamic effects
# editor.add_dynamic_effects()

# Save the file
editor.save_presentation("professional_presentation.pptx")
```

## 🛠️ Available Tools

### 🎬 Animation and Transition Tools (New)

#### add_slide_animation
Add animation transition effects to a single slide to make presentations more engaging
- `slide_index`: Slide index
- `animation_style`: Animation style (fade, push, wipe, zoom, split, blinds, dissolve, none)
- `speed`: Animation speed (fast, medium, slow)
- `auto_advance`: Whether to automatically advance to the next slide
- `auto_advance_seconds`: Auto-advance delay time

#### make_presentation_dynamic
Add uniform animation effects to the entire presentation — an important step in creating professional presentations
- `animation_style`: Unified animation style (default: fade)
- `speed`: Animation speed (default: medium)

#### make_professional_presentation ⭐
One-click professionalization! Automatically adds elegant fade transition effects to all slides
- No parameters required

#### add_smooth_transitions
Add smooth transition animations to the presentation for more natural slide switching
- No parameters required, uses preset smooth effects

#### add_dynamic_effects
Add dynamic transition effects to the presentation for more energetic delivery
- No parameters required, uses preset dynamic effects

#### get_animation_options
View all available slide animation effect options
- No parameters required

### 📄 Basic Tools

#### 1. create_presentation
Create a new PowerPoint presentation

#### 2. open_presentation
Open an existing PowerPoint file
- `file_path`: File path

#### 3. save_presentation
Save the presentation
- `file_path`: Save path (optional)

### 📝 Content Editing Tools

#### 4. add_slide
Add a new slide
- `layout_index`: Layout index (0 = title slide, 1 = title and content)

#### 5. add_text_box
Add a text box
- `slide_index`: Slide index
- `text`: Text content
- `left`, `top`, `width`, `height`: Position and size (in inches)
- `font_size`: Font size
- `font_color`: Font color (hexadecimal)

#### 6. add_title_slide
Add a title slide
- `title`: Title
- `subtitle`: Subtitle (optional)

#### 7. add_bullet_points
Add bulleted content
- `slide_index`: Slide index
- `title`: Title
- `bullet_points`: List of bullet points

#### 8. add_image
Add an image
- `slide_index`: Slide index
- `image_path`: Image file path
- `left`, `top`: Position (in inches)
- `width`, `height`: Size (in inches, optional)

#### 9. add_shape
Add a shape
- `slide_index`: Slide index
- `shape_type`: Shape type (rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow)
- `left`, `top`, `width`, `height`: Position and size (in inches)
- `fill_color`: Fill color (hexadecimal)

#### 10. add_table
Add a table
- `slide_index`: Slide index
- `rows`: Number of rows
- `cols`: Number of columns
- `left`, `top`, `width`, `height`: Position and size (in inches)

#### 11. set_table_cell_text
Set table cell text
- `slide_index`: Slide index
- `table_index`: Table index
- `row`: Row index
- `col`: Column index
- `text`: Text content

### 🎨 Formatting and Style Tools

#### 12. set_slide_background_color
Set slide background color
- `slide_index`: Slide index
- `color`: Background color (hexadecimal)

#### 13. add_hyperlink
Add a hyperlink to a shape
- `slide_index`: Slide index
- `shape_index`: Shape index
- `url`: Hyperlink URL
- `display_text`: Display text (optional)

#### 14. set_text_formatting
Set text formatting
- `slide_index`: Slide index
- `shape_index`: Shape index
- `font_name`: Font name (optional)
- `font_size`: Font size (optional)
- `font_color`: Font color (optional)
- `bold`: Whether to bold (optional)
- `italic`: Whether to italicize (optional)
- `underline`: Whether to underline (optional)

### 🔧 Management Tools

#### 15. get_presentation_info
Get presentation information

#### 16. delete_slide
Delete a slide
- `slide_index`: Index of the slide to delete

#### 17. duplicate_slide
Duplicate a slide
- `slide_index`: Index of the slide to duplicate

#### 18. move_slide
Move a slide to a new position
- `from_index`: Source position index
- `to_index`: Target position index

#### 19. get_slide_shapes_info
Get information about all shapes on a slide
- `slide_index`: Slide index

### 🎬 Legacy Animation Tools (Backward Compatible)

#### 20. set_slide_transition
Set slide transition effects (recommended to use the new animation tools instead)
- `slide_index`: Slide index
- `transition_type`: Transition type (none, fade, push, wipe, split, zoom, blinds, dissolve)
- `duration`: Transition duration (seconds)
- `advance_on_click`: Whether to advance on click
- `advance_after_time`: Auto-advance time (seconds, optional)

#### 21. get_available_transitions
Get a list of available transition effects
- No parameters required

## 💡 Usage Tips

### Helping AI Use Animation Features More Effectively

To encourage AI models to use animation features more proactively, use these keywords in your prompts:

- **"Make the presentation more professional"** → AI will call `make_professional_presentation`
- **"Add animation effects"** → AI will use `add_slide_animation` or `make_presentation_dynamic`
- **"Make slide transitions smoother"** → AI will call `add_smooth_transitions`
- **"Make the presentation more dynamic"** → AI will use `add_dynamic_effects`

### Recommended Workflow

1. **Create content** - Add all slides and content first
2. **One-click professionalization** - Use `make_presentation_professional()` to quickly add transitions
3. **Personalize** - Set different animation effects for specific slides as needed
4. **Preview and save** - Save the file and preview the effects in PowerPoint

## ⚠️ Important Notes

1. Make sure all required dependencies are installed (especially `lxml` for animation features)
2. Image file paths must exist and be accessible
3. Slide indices start from 0
4. Colors use hexadecimal format (e.g., 000000 for black, FF0000 for red)
5. Position and size units are in inches
6. Animation effects require opening the file in PowerPoint to see the full effect

## Error Handling

All operations include error handling and return responses in this format:
```json
{
  "success": true/false,
  "message": "Operation result message",
  "error": "Error message (if any)"
}
```

## 🤝 Contributing

Issues and Pull Requests are welcome to help improve this project!

## 📄 License

MIT License
