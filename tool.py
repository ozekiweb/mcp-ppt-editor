#!/usr/bin/env python3
"""
PowerPoint Editor Tool Class
Provides basic PPT editing functionality including adding text, images, shapes, etc.

Every method takes file_path as its first parameter. The presentation is loaded
from file_path if it exists, otherwise a new one is created. After applying the
requested changes, the presentation is automatically saved back to file_path.
"""

import logging
import os
from typing import Any, Dict, List, Optional, TYPE_CHECKING
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
except ImportError:
    raise ImportError("Please install python-pptx library: pip install python-pptx")

if TYPE_CHECKING:
    from pptx.presentation import Presentation as PresentationType

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class PowerPointEditor:
    def __init__(self):
        self.current_presentation: Optional["PresentationType"] = None
        self.current_file_path: Optional[str] = None

    def _load_or_create(self, file_path: str) -> Dict[str, Any]:
        """Load an existing presentation from file_path, or create a new one if it doesn't exist."""
        try:
            if not file_path:
                return {"success": False, "error": "file_path is required"}

            expanded = os.path.expanduser(file_path)
            if Path(expanded).exists():
                self.current_presentation = Presentation(expanded)
                self.current_file_path = expanded
                return {
                    "success": True,
                    "message": f"Loaded existing presentation: {expanded}",
                    "created": False,
                }
            else:
                parent = Path(expanded).parent
                if parent and not parent.exists():
                    parent.mkdir(parents=True, exist_ok=True)
                self.current_presentation = Presentation()
                self.current_file_path = expanded
                return {
                    "success": True,
                    "message": f"Created new presentation: {expanded}",
                    "created": True,
                }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _save_current(self) -> Dict[str, Any]:
        """Save the current presentation to current_file_path."""
        try:
            if not self.current_presentation:
                return {"success": False, "error": "No presentation is open"}
            if not self.current_file_path:
                return {"success": False, "error": "No file path set"}
            self.current_presentation.save(self.current_file_path)
            return {
                "success": True,
                "message": f"Saved presentation to {self.current_file_path}",
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _require_presentation(self) -> Optional[Dict[str, Any]]:
        """Return an error dict if no presentation is loaded, or None if OK."""
        if not self.current_presentation:
            return {"success": False, "error": "No presentation is open"}
        return None

    def _validate_slide_index(self, slide_index: int) -> Optional[Dict[str, Any]]:
        """Return an error dict if slide_index is out of range, or None if OK."""
        err = self._require_presentation()
        if err:
            return err
        if slide_index >= len(self.current_presentation.slides):
            return {
                "success": False,
                "error": f"Slide index out of range: {slide_index}",
            }
        return None

    def add_slide(self, file_path: str, layout_index: int = 1) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            slide_layouts = self.current_presentation.slide_layouts
            if layout_index >= len(slide_layouts):
                layout_index = 1
            layout = slide_layouts[layout_index]
            slide = self.current_presentation.slides.add_slide(layout)
            result = {
                "success": True,
                "message": "Successfully added new slide",
                "slide_index": len(self.current_presentation.slides) - 1,
                "total_slides": len(self.current_presentation.slides),
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_text_box(
        self,
        file_path: str,
        slide_index: int,
        text: str,
        left: float = 1,
        top: float = 1,
        width: float = 8,
        height: float = 1,
        font_size: int = 18,
        font_color: str = "000000",
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            textbox = slide.shapes.add_textbox(
                Inches(left), Inches(top), Inches(width), Inches(height)
            )
            text_frame = textbox.text_frame
            text_frame.text = text
            paragraph = text_frame.paragraphs[0]
            font = paragraph.font
            font.size = Pt(font_size)
            try:
                rgb_color = RGBColor.from_string(font_color)
                font.color.rgb = rgb_color
            except:
                pass
            result = {
                "success": True,
                "message": f"Successfully added text box to slide {slide_index}",
                "text": text,
                "position": {
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                },
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_title_slide(
        self, file_path: str, title: str, subtitle: str = ""
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            title_slide_layout = self.current_presentation.slide_layouts[0]
            slide = self.current_presentation.slides.add_slide(title_slide_layout)
            title_shape = slide.shapes.title
            if title_shape:
                title_shape.text = title
            if subtitle and len(slide.placeholders) > 1:
                try:
                    subtitle_shape = slide.placeholders[1]
                    subtitle_shape.text_frame.text = subtitle
                except (AttributeError, TypeError):
                    pass
            result = {
                "success": True,
                "message": "Successfully added title slide",
                "slide_index": len(self.current_presentation.slides) - 1,
                "title": title,
                "subtitle": subtitle,
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_bullet_points(
        self, file_path: str, slide_index: int, title: str, bullet_points: List[str]
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            if slide.shapes.title:
                slide.shapes.title.text = title
            content_placeholder = None
            for shape in slide.placeholders:
                if shape.placeholder_format.idx == 1:
                    content_placeholder = shape
                    break
            if content_placeholder:
                try:
                    text_frame = content_placeholder.text_frame
                    text_frame.clear()
                    for i, point in enumerate(bullet_points):
                        if i == 0:
                            p = text_frame.paragraphs[0]
                        else:
                            p = text_frame.add_paragraph()
                        p.text = point
                        p.level = 0
                except (AttributeError, TypeError):
                    pass
            result = {
                "success": True,
                "message": f"Successfully added bullet points to slide {slide_index}",
                "title": title,
                "bullet_points": bullet_points,
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_image(
        self,
        file_path: str,
        slide_index: int,
        image_path: str,
        left: float = 1,
        top: float = 2,
        width: Optional[float] = None,
        height: Optional[float] = None,
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            if not Path(image_path).exists():
                return {
                    "success": False,
                    "error": f"Image file does not exist: {image_path}",
                }
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            left_inches = Inches(left)
            top_inches = Inches(top)
            if width and height:
                pic = slide.shapes.add_picture(
                    image_path, left_inches, top_inches, Inches(width), Inches(height)
                )
            else:
                pic = slide.shapes.add_picture(image_path, left_inches, top_inches)
            result = {
                "success": True,
                "message": f"Successfully added image to slide {slide_index}",
                "image_path": image_path,
                "position": {
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                },
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_shape(
        self,
        file_path: str,
        slide_index: int,
        shape_type: str,
        left: float = 1,
        top: float = 1,
        width: float = 2,
        height: float = 1,
        fill_color: str = "0066CC",
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            shape_map = {
                "rectangle": MSO_SHAPE.RECTANGLE,
                "oval": MSO_SHAPE.OVAL,
                "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
                "diamond": MSO_SHAPE.DIAMOND,
                "pentagon": MSO_SHAPE.REGULAR_PENTAGON,
                "hexagon": MSO_SHAPE.HEXAGON,
                "star": MSO_SHAPE.STAR_5_POINT,
                "arrow": MSO_SHAPE.BLOCK_ARC,
            }
            if shape_type.lower() not in shape_map:
                return {
                    "success": False,
                    "error": f"Unsupported shape type: {shape_type}",
                }
            shape = slide.shapes.add_shape(
                shape_map[shape_type.lower()],
                Inches(left),
                Inches(top),
                Inches(width),
                Inches(height),
            )
            try:
                rgb_color = RGBColor.from_string(fill_color)
                shape.fill.solid()
                shape.fill.fore_color.rgb = rgb_color
            except:
                pass
            result = {
                "success": True,
                "message": f"Successfully added {shape_type} shape to slide {slide_index}",
                "shape_type": shape_type,
                "position": {
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                },
                "fill_color": fill_color,
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_presentation_info(self, file_path: str) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            slides_info = []
            for i, slide in enumerate(self.current_presentation.slides):
                slide_info = {
                    "index": i,
                    "shapes_count": len(slide.shapes),
                    "has_title": bool(slide.shapes.title and slide.shapes.title.text),
                    "title": slide.shapes.title.text if slide.shapes.title else "",
                }
                slides_info.append(slide_info)
            return {
                "success": True,
                "file_path": self.current_file_path,
                "slides_count": len(self.current_presentation.slides),
                "slides": slides_info,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def delete_slide(self, file_path: str, slide_index: int) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            xml_slides = self.current_presentation.part._element.sldIdLst
            xml_slides.remove(xml_slides[slide_index])
            slides = self.current_presentation.slides
            slides._sldIdLst.remove(slides._sldIdLst[slide_index])
            result = {
                "success": True,
                "message": f"Successfully deleted slide {slide_index}",
                "remaining_slides": len(self.current_presentation.slides),
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def duplicate_slide(self, file_path: str, slide_index: int) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            source_slide = self.current_presentation.slides[slide_index]
            slide_layout = source_slide.slide_layout
            new_slide = self.current_presentation.slides.add_slide(slide_layout)
            for shape in source_slide.shapes:
                if not shape.is_placeholder:
                    self._copy_shape(shape, new_slide)
            result = {
                "success": True,
                "message": f"Successfully duplicated slide {slide_index}",
                "new_slide_index": len(self.current_presentation.slides) - 1,
                "total_slides": len(self.current_presentation.slides),
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def move_slide(
        self, file_path: str, from_index: int, to_index: int
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            slides = self.current_presentation.slides
            if from_index >= len(slides) or to_index >= len(slides):
                return {"success": False, "error": "Slide index out of range"}
            if from_index == to_index:
                return {"success": True, "message": "Slide position unchanged"}
            source_slide = slides[from_index]
            slide_layout = source_slide.slide_layout
            new_slide = slides.add_slide(slide_layout)
            for shape in source_slide.shapes:
                if not shape.is_placeholder:
                    self._copy_shape(shape, new_slide)
            if from_index < len(slides) - 1:
                actual_from_index = from_index if to_index > from_index else from_index
                xml_slides = self.current_presentation.part._element.sldIdLst
                xml_slides.remove(xml_slides[actual_from_index])
                slides._sldIdLst.remove(slides._sldIdLst[actual_from_index])
            result = {
                "success": True,
                "message": f"Successfully moved slide from position {from_index} to position {to_index}",
                "total_slides": len(slides),
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_table(
        self,
        file_path: str,
        slide_index: int,
        rows: int,
        cols: int,
        left: float = 1,
        top: float = 2,
        width: float = 8,
        height: float = 4,
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            table = slide.shapes.add_table(
                rows, cols, Inches(left), Inches(top), Inches(width), Inches(height)
            )
            result = {
                "success": True,
                "message": f"Successfully added {rows}x{cols} table to slide {slide_index}",
                "rows": rows,
                "cols": cols,
                "position": {
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                },
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_table_cell_text(
        self,
        file_path: str,
        slide_index: int,
        table_index: int,
        row: int,
        col: int,
        text: str,
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            tables = []
            for shape in slide.shapes:
                try:
                    if hasattr(shape, "table"):
                        table_obj = getattr(shape, "table", None)
                        if table_obj is not None:
                            tables.append(shape)
                except:
                    continue
            if table_index >= len(tables):
                return {
                    "success": False,
                    "error": f"Table index out of range: {table_index}",
                }
            table = getattr(tables[table_index], "table")
            if row >= len(table.rows) or col >= len(table.columns):
                return {"success": False, "error": "Cell position out of table range"}
            table.cell(row, col).text = text
            result = {
                "success": True,
                "message": f"Successfully set text for table cell ({row}, {col})",
                "text": text,
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_slide_background_color(
        self, file_path: str, slide_index: int, color: str
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            background = slide.background
            fill = background.fill
            fill.solid()
            try:
                rgb_color = RGBColor.from_string(color)
                fill.fore_color.rgb = rgb_color
            except:
                return {"success": False, "error": f"Invalid color format: {color}"}
            result = {
                "success": True,
                "message": f"Successfully set background color for slide {slide_index}",
                "color": color,
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def add_hyperlink(
        self,
        file_path: str,
        slide_index: int,
        shape_index: int,
        url: str,
        display_text: Optional[str] = None,
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            if shape_index >= len(slide.shapes):
                return {
                    "success": False,
                    "error": f"Shape index out of range: {shape_index}",
                }
            shape = slide.shapes[shape_index]
            try:
                if hasattr(shape, "text_frame"):
                    text_frame = getattr(shape, "text_frame")
                    if text_frame is not None:
                        if display_text:
                            text_frame.text = display_text
                        if not text_frame.paragraphs:
                            paragraph = text_frame.add_paragraph()
                        else:
                            paragraph = text_frame.paragraphs[0]
                        if not paragraph.runs:
                            run = paragraph.add_run()
                        else:
                            run = paragraph.runs[0]
                        run.hyperlink.address = url
                    else:
                        if hasattr(shape, "click_action"):
                            shape.click_action.hyperlink.address = url
                        else:
                            return {
                                "success": False,
                                "error": "Shape does not support hyperlinks",
                            }
                else:
                    if hasattr(shape, "click_action"):
                        shape.click_action.hyperlink.address = url
                    else:
                        return {
                            "success": False,
                            "error": "Shape does not support hyperlinks",
                        }
            except Exception as e:
                return {"success": False, "error": f"Failed to add hyperlink: {str(e)}"}
            result = {
                "success": True,
                "message": f"Successfully added hyperlink to shape {shape_index} on slide {slide_index}",
                "url": url,
                "display_text": display_text,
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def set_text_formatting(
        self,
        file_path: str,
        slide_index: int,
        shape_index: int,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        font_color: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            if shape_index >= len(slide.shapes):
                return {
                    "success": False,
                    "error": f"Shape index out of range: {shape_index}",
                }
            shape = slide.shapes[shape_index]
            try:
                if not hasattr(shape, "text_frame"):
                    return {
                        "success": False,
                        "error": "Shape does not support text boxes",
                    }
                text_frame = getattr(shape, "text_frame")
                if text_frame is None:
                    return {"success": False, "error": "Text frame not available"}
                if not hasattr(text_frame, "paragraphs"):
                    return {
                        "success": False,
                        "error": "Text frame has no paragraphs attribute",
                    }
                paragraphs = text_frame.paragraphs
                if not paragraphs or len(paragraphs) == 0:
                    return {"success": False, "error": "No available text paragraphs"}
                paragraph = paragraphs[0]
                if not hasattr(paragraph, "font"):
                    return {
                        "success": False,
                        "error": "Paragraph has no font attribute",
                    }
                font = paragraph.font
                if font is None:
                    return {"success": False, "error": "Unable to get font object"}
                if font_name:
                    font.name = font_name
                if font_size:
                    font.size = Pt(font_size)
                if font_color:
                    try:
                        rgb_color = RGBColor.from_string(font_color)
                        font.color.rgb = rgb_color
                    except:
                        return {
                            "success": False,
                            "error": f"Invalid color format: {font_color}",
                        }
                if bold is not None:
                    font.bold = bold
                if italic is not None:
                    font.italic = italic
                if underline is not None:
                    font.underline = underline
            except Exception as e:
                return {"success": False, "error": f"Font setting failed: {str(e)}"}
            result = {
                "success": True,
                "message": f"Successfully set text formatting for shape {shape_index} on slide {slide_index}",
                "formatting": {
                    "font_name": font_name,
                    "font_size": font_size,
                    "font_color": font_color,
                    "bold": bold,
                    "italic": italic,
                    "underline": underline,
                },
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_slide_shapes_info(self, file_path: str, slide_index: int) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            shapes_info = []
            for i, shape in enumerate(slide.shapes):
                shape_info = {
                    "index": i,
                    "shape_type": str(shape.shape_type),
                    "name": shape.name,
                    "left": shape.left.inches if hasattr(shape.left, "inches") else 0,
                    "top": shape.top.inches if hasattr(shape.top, "inches") else 0,
                    "width": shape.width.inches
                    if hasattr(shape.width, "inches")
                    else 0,
                    "height": shape.height.inches
                    if hasattr(shape.height, "inches")
                    else 0,
                    "has_text": hasattr(shape, "text_frame")
                    and getattr(shape, "text_frame", None) is not None,
                    "text": getattr(shape, "text", "")
                    if hasattr(shape, "text")
                    else "",
                }
                shapes_info.append(shape_info)
            return {
                "success": True,
                "slide_index": slide_index,
                "shapes_count": len(shapes_info),
                "shapes": shapes_info,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _copy_shape(self, source_shape, target_slide):
        try:
            if hasattr(source_shape, "text_frame") and source_shape.text_frame:
                textbox = target_slide.shapes.add_textbox(
                    source_shape.left,
                    source_shape.top,
                    source_shape.width,
                    source_shape.height,
                )
                textbox.text_frame.text = source_shape.text_frame.text
        except:
            pass

    def set_slide_transition(
        self,
        file_path: str,
        slide_index: int,
        transition_type: str = "fade",
        duration: float = 1.0,
        advance_on_click: bool = True,
        advance_after_time: Optional[float] = None,
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            err = self._validate_slide_index(slide_index)
            if err:
                return err
            slide = self.current_presentation.slides[slide_index]
            supported_transitions = [
                "none",
                "fade",
                "push",
                "wipe",
                "split",
                "zoom",
                "blinds",
                "dissolve",
            ]
            if transition_type.lower() not in supported_transitions:
                return {
                    "success": False,
                    "error": f"Unsupported transition type: {transition_type}. Supported types: {', '.join(supported_transitions)}",
                }
            slide_element = slide._element
            namespaces = {
                "p": "http://schemas.openxmlformats.org/presentationml/2006/main"
            }
            existing_transition = slide_element.find(".//p:transition", namespaces)
            if existing_transition is not None:
                slide_element.remove(existing_transition)
            if transition_type.lower() != "none":
                try:
                    from lxml import etree
                except ImportError:
                    return {
                        "success": False,
                        "error": "Please install lxml library: pip install lxml",
                    }
                transition_xml = self._create_transition_xml(
                    transition_type, duration, advance_on_click, advance_after_time
                )
                parser = etree.XMLParser(ns_clean=True, recover=True)
                transition_elem = etree.fromstring(
                    transition_xml.encode("utf-8"), parser
                )
                color_map_override = slide_element.find(".//p:clrMapOvr", namespaces)
                if color_map_override is not None:
                    color_map_override.addprevious(transition_elem)
                else:
                    slide_element.append(transition_elem)
                verification_elem = slide_element.find(".//p:transition", namespaces)
                if verification_elem is None:
                    return {
                        "success": False,
                        "error": "Transition effect XML insertion failed",
                    }
            result = {
                "success": True,
                "message": f"Successfully set transition effect for slide {slide_index}",
                "transition_type": transition_type,
                "duration": duration,
                "advance_on_click": advance_on_click,
                "advance_after_time": advance_after_time,
                "verification": "Transition effect verified successfully"
                if transition_type.lower() != "none"
                else "Transition effect removed",
            }
            save_result = self._save_current()
            if not save_result.get("success"):
                return save_result
            return result
        except Exception as e:
            return {"success": False, "error": str(e)}

    def _create_transition_xml(
        self,
        transition_type: str,
        duration: float,
        advance_on_click: bool,
        advance_after_time: Optional[float],
    ) -> str:
        if duration <= 0.5:
            speed = "fast"
        elif duration <= 2.0:
            speed = "med"
        else:
            speed = "slow"
        transition_attrs = f'spd="{speed}"'
        if advance_on_click:
            transition_attrs += ' advClick="1"'
        else:
            transition_attrs += ' advClick="0"'
        if advance_after_time is not None:
            advance_time_ms = int(advance_after_time * 1000)
            transition_attrs += f' advTm="{advance_time_ms}"'
        transition_content = ""
        if transition_type.lower() == "fade":
            transition_content = "<p:fade/>"
        elif transition_type.lower() == "push":
            transition_content = '<p:push dir="l"/>'
        elif transition_type.lower() == "wipe":
            transition_content = '<p:wipe dir="l"/>'
        elif transition_type.lower() == "zoom":
            transition_content = "<p:zoom/>"
        elif transition_type.lower() == "split":
            transition_content = '<p:split orient="horz" dir="out"/>'
        elif transition_type.lower() == "blinds":
            transition_content = '<p:blinds dir="horz"/>'
        elif transition_type.lower() == "dissolve":
            transition_content = "<p:dissolve/>"
        else:
            transition_content = "<p:fade/>"
        xml_string = f"""<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" {transition_attrs}>
    {transition_content}
</p:transition>"""
        return xml_string

    def apply_transition_to_all_slides(
        self, file_path: str, transition_type: str = "fade", duration: float = 1.0
    ) -> Dict[str, Any]:
        load_result = self._load_or_create(file_path)
        if not load_result.get("success"):
            return load_result
        try:
            slides = self.current_presentation.slides
            if len(slides) == 0:
                return {"success": False, "error": "No slides in presentation"}
            success_count = 0
            failed_slides = []
            for i in range(len(slides)):
                result = self.set_slide_transition(
                    file_path, i, transition_type, duration, True, None
                )
                if result.get("success"):
                    success_count += 1
                else:
                    failed_slides.append(i)
            if success_count == len(slides):
                return {
                    "success": True,
                    "message": f"Successfully set '{transition_type}' transition effect for all {len(slides)} slides",
                    "transition_type": transition_type,
                    "duration": duration,
                    "slides_processed": len(slides),
                }
            else:
                return {
                    "success": True,
                    "message": f"Set transition effect for {success_count}/{len(slides)} slides",
                    "transition_type": transition_type,
                    "duration": duration,
                    "slides_processed": success_count,
                    "failed_slides": failed_slides,
                    "warning": f"{len(failed_slides)} slides failed to set",
                }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_available_transitions(self, file_path: str) -> Dict[str, Any]:
        transitions = [
            {"name": "none", "description": "No transition effect"},
            {
                "name": "fade",
                "description": "Fade - Recommended for professional presentations",
            },
            {"name": "push", "description": "Push - Dynamic and energetic"},
            {"name": "wipe", "description": "Wipe - Clean and smooth"},
            {"name": "split", "description": "Split - Creative effect"},
            {"name": "zoom", "description": "Zoom - Emphasize key points"},
            {"name": "blinds", "description": "Blinds - Classic effect"},
            {"name": "dissolve", "description": "Dissolve - Gentle transition"},
        ]
        return {
            "success": True,
            "transitions": transitions,
            "total_count": len(transitions),
            "note": "These animation effects can make your presentation more vivid and interesting",
            "recommendation": "Recommended to use 'fade' effect, suitable for most professional presentations",
        }

    def make_presentation_professional(self, file_path: str) -> Dict[str, Any]:
        return self.apply_transition_to_all_slides(file_path, "fade", 1.0)

    def add_smooth_transitions(self, file_path: str) -> Dict[str, Any]:
        return self.apply_transition_to_all_slides(file_path, "fade", 0.8)

    def add_dynamic_effects(self, file_path: str) -> Dict[str, Any]:
        return self.apply_transition_to_all_slides(file_path, "push", 1.2)

    def generate_outline_for_topic(self, file_path: str, topic: str) -> Dict[str, Any]:
        try:
            import json as _json

            outline_data = {
                "slides": [
                    {
                        "title": f"In-depth Discussion on {topic}",
                        "subtitle": "AI-assisted generation",
                    },
                    {
                        "title": "Introduction and Background",
                        "content": [
                            f"Definition and importance of {topic}",
                            "Related historical development",
                            "Main scope of this discussion",
                        ],
                    },
                    {
                        "title": "Core Points Analysis",
                        "content": [
                            "First key aspect",
                            "Second key aspect with examples",
                            "In-depth analysis of third key aspect",
                        ],
                    },
                    {
                        "title": "Case Study or Practical Application",
                        "content": [
                            f"A real-world case about {topic}",
                            "Insights from the case",
                            "How to apply these in practice",
                        ],
                    },
                    {
                        "title": "Summary and Outlook",
                        "content": [
                            f"Summary of core content on {topic}",
                            "Future development trends",
                            "Q&A session",
                        ],
                    },
                ]
            }
            outline_json = _json.dumps(outline_data, ensure_ascii=False, indent=2)
            return {
                "success": True,
                "message": f"Successfully generated outline for topic '{topic}'.",
                "outline_json": outline_json,
            }
        except Exception as e:
            logger.error(f"Error generating outline for topic '{topic}': {e}")
            return {"success": False, "error": str(e)}
