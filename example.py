#!/usr/bin/env python3
"""
PowerPoint Editor Usage Example
Demonstrates how to use PowerPointEditor class to create and edit PPT

Every method now takes file_path as the first parameter. The presentation
is loaded from file_path if it exists, otherwise a new one is created.
After applying changes, the presentation is automatically saved back to file_path.
"""

from datetime import datetime
from tool import PowerPointEditor


def main():
    editor = PowerPointEditor()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = f"c:/temp/_enhanced_presentation_{timestamp}.pptx"

    # 1. Add title slide (creates new presentation since file doesn't exist)
    print("1. Adding title slide...")
    result = editor.add_title_slide(file_path, "My Presentation", "Created with Python")
    print(f"Result: {result}")

    # 2. Add content slide
    print("\n2. Adding content slide...")
    result = editor.add_slide(file_path, 1)
    print(f"Result: {result}")

    # 3. Add bullet points to slide 1
    print("\n3. Adding bullet points...")
    result = editor.add_bullet_points(
        file_path, 1, "Main Content", ["First point", "Second point", "Third point"]
    )
    print(f"Result: {result}")

    # 4. Add another slide
    print("\n4. Adding another slide...")
    result = editor.add_slide(file_path, 1)
    print(f"Result: {result}")

    # 5. Add text box to slide 2
    print("\n5. Adding text box...")
    result = editor.add_text_box(
        file_path,
        slide_index=2,
        text="This is a custom text box",
        left=2,
        top=3,
        width=6,
        height=2,
        font_size=24,
        font_color="FF0000",
    )
    print(f"Result: {result}")

    # 6. Add shape
    print("\n6. Adding shape...")
    result = editor.add_shape(
        file_path,
        slide_index=2,
        shape_type="rectangle",
        left=1,
        top=1,
        width=3,
        height=1.5,
        fill_color="00FF00",
    )
    print(f"Result: {result}")

    # 7. Get presentation info
    print("\n7. Getting presentation info...")
    result = editor.get_presentation_info(file_path)
    print(f"Result: {result}")

    # 8. Duplicate slide 0
    print("\n8. Duplicating slide 0...")
    result = editor.duplicate_slide(file_path, 0)
    print(f"Result: {result}")

    # 9. Add table
    print("\n9. Adding table...")
    result = editor.add_table(
        file_path, slide_index=2, rows=3, cols=4, left=1, top=2, width=8, height=3
    )
    print(f"Result: {result}")

    # 10. Set table cell text
    print("\n10. Setting table cell text...")
    result = editor.set_table_cell_text(
        file_path, slide_index=2, table_index=0, row=0, col=0, text="Header1"
    )
    print(f"Result: {result}")
    result = editor.set_table_cell_text(
        file_path, slide_index=2, table_index=0, row=0, col=1, text="Header2"
    )
    print(f"Result: {result}")
    result = editor.set_table_cell_text(
        file_path, slide_index=2, table_index=0, row=1, col=0, text="Data1"
    )
    print(f"Result: {result}")
    result = editor.set_table_cell_text(
        file_path, slide_index=2, table_index=0, row=1, col=1, text="Data2"
    )
    print(f"Result: {result}")

    # 11. Set slide background color
    print("\n11. Setting slide background color...")
    result = editor.set_slide_background_color(file_path, slide_index=2, color="E6F3FF")
    print(f"Result: {result}")

    # 12. Get slide shapes info
    print("\n12. Getting slide shapes info...")
    result = editor.get_slide_shapes_info(file_path, slide_index=2)
    print(f"Result: {result}")

    # 13. Set text formatting
    print("\n13. Setting text formatting...")
    result = editor.set_text_formatting(
        file_path,
        slide_index=2,
        shape_index=2,
        font_name="Arial",
        font_size=24,
        font_color="FF0000",
        bold=True,
        italic=True,
    )
    print(f"Result: {result}")

    print(f"\nDemo complete! Presentation saved to: {file_path}")


if __name__ == "__main__":
    main()
