"""
Position and unit conversion utilities for PowerPoint images.
"""

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os


def cm_to_inches(cm):
    """
    Convert centimeters to inches.

    Args:
        cm (float): Length in centimeters

    Returns:
        float: Length in inches
    """
    return cm / 2.54


def get_image_position(ppt_path, slide_index, image_index=0):
    """
    Extract position and size information from an image on a slide.

    Useful for getting template image parameters to use when inserting
    replacement images at the same location.

    Args:
        ppt_path (str): Path to the PowerPoint file
        slide_index (int): Slide number (0-based index)
        image_index (int): Which image to inspect if multiple exist (default: 0)

    Returns:
        dict: Position info with keys 'left', 'top', 'width', 'height' (all in inches)

    Raises:
        FileNotFoundError: If PPT file doesn't exist
        IndexError: If slide_index or image_index is out of range
        ValueError: If the slide has no pictures

    Example:
        >>> pos = get_image_position('presentation.pptx', 1, 0)
        >>> print(pos)
        {'left': 0.14, 'top': 0.96, 'width': 5.43, 'height': 2.72}
    """
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    prs = Presentation(ppt_path)

    if slide_index < 0 or slide_index >= len(prs.slides):
        raise IndexError(
            f"Slide index {slide_index} out of range. "
            f"Presentation has {len(prs.slides)} slides"
        )

    slide = prs.slides[slide_index]

    # Find all pictures on the slide
    pictures = [shape for shape in slide.shapes
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE]

    if not pictures:
        raise ValueError(f"Slide {slide_index} has no pictures")

    if image_index < 0 or image_index >= len(pictures):
        raise IndexError(
            f"Image index {image_index} out of range. "
            f"Slide has {len(pictures)} pictures (indices 0-{len(pictures)-1})"
        )

    # Get the specified picture
    picture = pictures[image_index]

    # Extract position and size (convert from EMUs to inches)
    # EMU = English Metric Units (914400 EMUs = 1 inch)
    position_info = {
        'left': picture.left / 914400.0,
        'top': picture.top / 914400.0,
        'width': picture.width / 914400.0,
        'height': picture.height / 914400.0
    }

    return position_info
