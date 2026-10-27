"""
Position and unit conversion utilities for PowerPoint images.
"""

from typing import Dict, List
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os

# English Metric Units (EMU) conversion
# PowerPoint uses EMUs internally: 914400 EMUs = 1 inch
EMUS_PER_INCH = 914400.0


def cm_to_inches(cm: float) -> float:
    """
    Convert centimeters to inches.

    Args:
        cm (float): Length in centimeters

    Returns:
        float: Length in inches
    """
    return cm / 2.54


def get_image_position(
    ppt_path: str,
    slide_index: int,
    image_index: int = 0
) -> Dict[str, float]:
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
    position_info = {
        'left': picture.left / EMUS_PER_INCH,
        'top': picture.top / EMUS_PER_INCH,
        'width': picture.width / EMUS_PER_INCH,
        'height': picture.height / EMUS_PER_INCH
    }

    return position_info


def get_all_image_positions(
    ppt_path: str,
    slide_index: int
) -> List[Dict[str, float]]:
    """
    Extract positions and sizes of ALL images on a slide.

    Useful for templates with multiple placeholder images. Returns positions
    sorted by visual reading order: top-to-bottom, then left-to-right within
    each row (a picture shares the row of the row's topmost picture when their vertical
    extents overlap by at least half the shorter height); pictures stacked in one column of
    a row go top to bottom, so a tall picture beside a grid is followed by that grid column
    by column. See _reading_order.

    Args:
        ppt_path (str): Path to the PowerPoint file
        slide_index (int): Slide number (0-based index)

    Returns:
        list: List of position dicts, each with keys 'left', 'top', 'width', 'height' (all in inches)
              Returns empty list if slide has no pictures

    Raises:
        FileNotFoundError: If PPT file doesn't exist
        IndexError: If slide_index is out of range

    Example:
        >>> positions = get_all_image_positions('presentation.pptx', 1)
        >>> print(positions)
        [
            {'left': 0.5, 'top': 1.0, 'width': 4.0, 'height': 3.0},
            {'left': 5.0, 'top': 1.0, 'width': 4.0, 'height': 3.0}
        ]
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

    # Convert all picture positions to dicts
    positions = []
    for picture in pictures:
        position_info = {
            'left': picture.left / EMUS_PER_INCH,
            'top': picture.top / EMUS_PER_INCH,
            'width': picture.width / EMUS_PER_INCH,
            'height': picture.height / EMUS_PER_INCH
        }
        positions.append(position_info)

    return _reading_order(positions)


def _reading_order(positions: List[Dict[str, float]]) -> List[Dict[str, float]]:
    """Sort positions top-to-bottom by row, then left-to-right within a row.

    A picture joins the current row when its vertical extent overlaps the row's TALLEST
    picture (the first of them on ties) by at least half the shorter of the two heights.
    Comparing with one picture, not the row's growing span, keeps a staircase layout from
    chaining several visual rows into one. The anchor was the row's first (topmost) picture,
    so a tall picture sitting 0.05" below a grid's top row joined it but the grid's lower
    row did not: images 3 and 4 swapped between 1.00" and 1.05" (bug hunt 2026-10-26). Rounding tops to a grid (the previous
    rule, 0.5") split one visual row whose tops straddled a rounding boundary: in a 2x2
    template with bottom-row tops 4.215" and 4.299" the bottom-RIGHT image was filled
    before the bottom-LEFT one (bug hunt 2026-10-24).

    Within a row, pictures stacked in one column (horizontal extents overlapping by at least
    half the narrower width, the same test transposed) keep top-to-bottom order. Sorting the
    row by `left` alone put a stack's LOWER picture first when it sat 0.01" further left,
    e.g. beside a tall picture that merges the stack into one row (bug hunt 2026-10-24,
    second pass).
    """
    rows: List[List[Dict[str, float]]] = []
    for p in sorted(positions, key=lambda q: (q['top'], q['left'])):
        if rows:
            a = max(rows[-1], key=lambda q: q['height'])   # first of the tallest on ties
            overlap = min(p['top'] + p['height'], a['top'] + a['height']) - max(p['top'], a['top'])
            if overlap >= 0.5 * max(min(p['height'], a['height']), 1e-9):
                rows[-1].append(p)
                continue
        rows.append([p])
    return [p for row in rows for p in _columns_in_row(row)]


def _columns_in_row(row: List[Dict[str, float]]) -> List[Dict[str, float]]:
    """Order one row's pictures left to right by column, top to bottom within a column."""
    cols: List[List[Dict[str, float]]] = []
    for p in sorted(row, key=lambda q: (q['left'], q['top'])):
        for col in cols:
            a = col[0]
            overlap = min(p['left'] + p['width'], a['left'] + a['width']) - max(p['left'], a['left'])
            if overlap >= 0.5 * max(min(p['width'], a['width']), 1e-9):
                col.append(p)
                break
        else:
            cols.append([p])
    return [p for col in cols for p in sorted(col, key=lambda q: (q['top'], q['left']))]
