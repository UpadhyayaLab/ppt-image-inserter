"""
High-level workflow functions for PowerPoint image replacement.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
import os
from .position import get_image_position
from .slide_utils import duplicate_slide, remove_pictures_from_slide, remove_all_text_from_slide


def copy_slide_replace_image(ppt_path, source_slide_index, new_image_path, position=None,
                             store_metadata=True, add_label=True):
    """
    Copy a slide and replace its image with a new one.

    This is the main high-level function for batch image processing. It:
    1. Duplicates the source slide (creating a new slide at the end)
    2. Removes all pictures from the duplicated slide
    3. Inserts the new image at the specified position
    4. Optionally stores the original image path in the picture's alt text
    5. Optionally adds a visible text label showing filename and folder path

    Args:
        ppt_path (str): Path to the PowerPoint file
        source_slide_index (int): Index of the template slide to copy (0-based)
        new_image_path (str): Path to the new image to insert
        position (dict, optional): Position dict with keys 'left', 'top', 'width', 'height'
                                   If None, auto-detects from first image in source slide
        store_metadata (bool): If True, stores the original image path in the picture's
                              alt text fields (descr = full path, title = filename).
                              This allows tracking image sources after slide reordering.
                              Default: True
        add_label (bool): If True, adds a visible text box in bottom left corner showing
                         the filename and folder path. Useful for visual tracking while
                         viewing slides. Default: True

    Returns:
        int: Index of the newly created slide

    Raises:
        FileNotFoundError: If PPT or image file doesn't exist
        IndexError: If source_slide_index is out of range
        ValueError: If position is None and source slide has no pictures

    Example:
        >>> # Auto-detect position with metadata and label
        >>> new_idx = copy_slide_replace_image(
        ...     'presentation.pptx',
        ...     1,  # Copy slide 2
        ...     'new_image.tif'
        ... )
        >>> print(f"Created slide {new_idx + 1}")

        >>> # Specify exact position without label
        >>> pos = {'left': 0.14, 'top': 0.96, 'width': 5.43, 'height': 2.72}
        >>> new_idx = copy_slide_replace_image(
        ...     'presentation.pptx',
        ...     1,
        ...     'new_image.tif',
        ...     position=pos,
        ...     add_label=False
        ... )
    """
    # Validate input files
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    if not os.path.exists(new_image_path):
        raise FileNotFoundError(f"Image file not found: {new_image_path}")

    # Load presentation
    prs = Presentation(ppt_path)

    # Auto-detect position if not specified
    if position is None:
        position = get_image_position(ppt_path, source_slide_index, image_index=0)

    # Duplicate the source slide
    new_slide = duplicate_slide(prs, source_slide_index)

    # Get the new slide's index (it's added at the end)
    new_slide_index = len(prs.slides) - 1

    # Remove all pictures from the new slide
    num_removed = remove_pictures_from_slide(new_slide)

    # Remove all text (including placeholders) from the new slide
    num_text_removed = remove_all_text_from_slide(new_slide)

    # Insert the new image at the specified position
    picture = new_slide.shapes.add_picture(
        new_image_path,
        Inches(position['left']),
        Inches(position['top']),
        width=Inches(position['width']),
        height=Inches(position['height'])
    )

    # Store metadata in alt text if requested
    if store_metadata:
        # Use XML API to set alt text (python-pptx 1.0.2 doesn't have descr/title properties)
        try:
            picture._element._nvXxPr.cNvPr.set("descr", new_image_path)  # Full path
            picture._element._nvXxPr.cNvPr.set("title", os.path.basename(new_image_path))  # Filename
        except Exception as e:
            print(f"[WARNING] Could not store metadata: {e}")

    # Add visible text label if requested
    if add_label:
        try:
            # Create text box in bottom left corner
            textbox = new_slide.shapes.add_textbox(
                Inches(0.5),    # Left: 0.5" from edge
                Inches(7.0),    # Top: Near bottom (standard slide is 7.5" tall)
                Inches(5.0),    # Width: 5 inches
                Inches(0.4)     # Height: 0.4 inches
            )

            # Set text content
            text_frame = textbox.text_frame
            folder_path = os.path.dirname(new_image_path)
            filename = os.path.basename(new_image_path)
            text_frame.text = f"File: {filename}\nPath: {folder_path}"

            # Format text - small font
            for paragraph in text_frame.paragraphs:
                paragraph.font.size = Pt(8)  # 8pt font
                paragraph.font.name = 'Arial'

        except Exception as e:
            print(f"[WARNING] Could not add text label: {e}")

    # Save the presentation
    prs.save(ppt_path)

    return new_slide_index


def replace_image_on_existing_slide(ppt_path: str, slide_index: int, new_image_path: str,
                                    store_metadata: bool = True, add_label: bool = True) -> None:
    """
    Replace the image on an existing slide with a new one.

    This function:
    1. Gets the position of the first image on the slide
    2. Removes all pictures from the slide
    3. Removes any existing text box labels (to avoid duplicates)
    4. Inserts the new image at the original position
    5. Optionally stores metadata in alt text
    6. Optionally adds a visible text label

    Args:
        ppt_path (str): Path to the PowerPoint file
        slide_index (int): Index of the slide to update (0-based)
        new_image_path (str): Path to the new image to insert
        store_metadata (bool): If True, stores image path in alt text. Default: True
        add_label (bool): If True, adds visible text label in bottom left. Default: True

    Raises:
        FileNotFoundError: If PPT or image file doesn't exist
        IndexError: If slide_index is out of range
        ValueError: If the slide has no pictures to determine position

    Example:
        >>> replace_image_on_existing_slide('presentation.pptx', 1, 'updated_image.tif')
    """
    # Validate input files
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    if not os.path.exists(new_image_path):
        raise FileNotFoundError(f"Image file not found: {new_image_path}")

    # Load presentation
    prs = Presentation(ppt_path)

    # Validate slide index
    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(
            f"Slide index {slide_index} out of range (presentation has {len(prs.slides)} slides)"
        )

    slide = prs.slides[slide_index]

    # Get position from the first image on the slide
    print(f"Getting image position from slide {slide_index} (slide {slide_index + 1} in UI)...")
    position = get_image_position(ppt_path, slide_index, image_index=0)
    print(f"Detected position: left={position['left']:.2f}\", top={position['top']:.2f}\", "
          f"width={position['width']:.2f}\", height={position['height']:.2f}\"")

    # Remove all pictures from the slide
    print(f"Removing existing pictures...")
    num_removed = remove_pictures_from_slide(slide)
    print(f"Removed {num_removed} picture(s)")

    # Remove any existing text boxes (labels) to avoid duplicates
    print(f"Removing existing text labels...")
    shapes_list = list(slide.shapes)
    textboxes_removed = 0
    for shape in shapes_list:
        if shape.has_text_frame:
            # Check if it's a text box (not a placeholder or title)
            try:
                # Text boxes typically don't have placeholder format
                if not hasattr(shape, 'placeholder_format'):
                    parent = shape._element.getparent()
                    if parent is not None:
                        parent.remove(shape._element)
                        textboxes_removed += 1
            except:
                pass  # Skip if we can't determine
    print(f"Removed {textboxes_removed} text box(es)")

    # Insert the new image at the original position
    print(f"Inserting {os.path.basename(new_image_path)}...")
    picture = slide.shapes.add_picture(
        new_image_path,
        Inches(position['left']),
        Inches(position['top']),
        width=Inches(position['width']),
        height=Inches(position['height'])
    )

    # Store metadata in alt text if requested
    if store_metadata:
        try:
            picture._element._nvXxPr.cNvPr.set("descr", new_image_path)
            picture._element._nvXxPr.cNvPr.set("title", os.path.basename(new_image_path))
            print(f"Stored metadata: {new_image_path}")
        except Exception as e:
            print(f"[WARNING] Could not store metadata: {e}")

    # Add visible text label if requested
    if add_label:
        try:
            textbox = slide.shapes.add_textbox(
                Inches(0.5),
                Inches(7.0),
                Inches(5.0),
                Inches(0.4)
            )

            text_frame = textbox.text_frame
            folder_path = os.path.dirname(new_image_path)
            filename = os.path.basename(new_image_path)
            text_frame.text = f"File: {filename}\nPath: {folder_path}"

            for paragraph in text_frame.paragraphs:
                paragraph.font.size = Pt(8)
                paragraph.font.name = 'Arial'

            print(f"Added text label: {filename}")
        except Exception as e:
            print(f"[WARNING] Could not add text label: {e}")

    # Save the presentation
    print(f"Saving presentation...")
    prs.save(ppt_path)

    print(f"[SUCCESS] Updated slide {slide_index} (slide {slide_index + 1} in PowerPoint UI)")
