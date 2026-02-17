#!/usr/bin/env python
"""
Example: Multi-Image Per Slide (Direct API Usage)

This script demonstrates how to insert multiple images into a single slide
using the ppt_image_inserter API directly — without a YAML config file.

Use this approach when you want more programmatic control, or when the batch
script's config format doesn't fit your workflow.

The key function is copy_slide_replace_images(), which:
  1. Duplicates a template slide
  2. Removes placeholder images
  3. Inserts your images at the detected/specified positions

Usage:
    Edit the paths below and run:
    python example_multi_image.py

Common use case: side-by-side comparison slides (e.g., Control vs. Treatment)
"""

import sys
import os

# Add parent directory to path so we can import ppt_image_inserter
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ppt_image_inserter import copy_slide_replace_images, get_all_image_positions


# ---------------------------------------------------------------------------
# CONFIGURATION — edit these paths to match your files
# ---------------------------------------------------------------------------

PPT_FILE = "presentations/my_presentation.pptx"  # Must be .pptx, must be closed
TEMPLATE_SLIDE = 1  # 0-based: slide 2 in PowerPoint UI

# Each inner list = one slide. Images in the list go left-to-right
# matching the order of placeholder images in your template slide.
IMAGE_SETS = [
    # Slide 1: side-by-side comparison
    [
        "data/images/condition_A_timepoint_1.png",
        "data/images/condition_B_timepoint_1.png",
    ],
    # Slide 2: another comparison
    [
        "data/images/condition_A_timepoint_2.png",
        "data/images/condition_B_timepoint_2.png",
    ],
    # Slide 3: three images across
    # (template must have 3 placeholder images for this to work)
    # [
    #     "data/images/replicate_1.png",
    #     "data/images/replicate_2.png",
    #     "data/images/replicate_3.png",
    # ],
]

# ---------------------------------------------------------------------------


def main():
    # Verify files exist before starting
    if not os.path.exists(PPT_FILE):
        print(f"[ERROR] Presentation not found: {PPT_FILE}")
        sys.exit(1)

    for i, image_set in enumerate(IMAGE_SETS):
        for img_path in image_set:
            if not os.path.exists(img_path):
                print(f"[ERROR] Image not found (set {i+1}): {img_path}")
                sys.exit(1)

    # Show what positions will be used (auto-detected from template)
    positions = get_all_image_positions(PPT_FILE, TEMPLATE_SLIDE)
    print(f"Template slide has {len(positions)} placeholder image(s):")
    for j, pos in enumerate(positions):
        print(f"  [{j}] left={pos['left']:.2f}\", top={pos['top']:.2f}\", "
              f"width={pos['width']:.2f}\", height={pos['height']:.2f}\"")
    print()

    # Insert each set of images as a new slide
    success_count = 0
    for i, image_set in enumerate(IMAGE_SETS):
        if len(image_set) != len(positions):
            print(f"[ERROR] Set {i+1} has {len(image_set)} image(s) but template "
                  f"has {len(positions)} placeholder(s) — skipping")
            continue

        try:
            new_slide_idx = copy_slide_replace_images(
                PPT_FILE,
                TEMPLATE_SLIDE,
                image_set,
                positions=None,   # Auto-detect from template (recommended)
                store_metadata=True,
                add_label=False,  # Labels are noisy for multi-image slides
            )
            filenames = [os.path.basename(p) for p in image_set]
            print(f"  Slide {new_slide_idx + 1}: {' | '.join(filenames)}")
            success_count += 1
        except Exception as e:
            print(f"[ERROR] Failed on set {i+1}: {e}")

    print(f"\nDone: {success_count}/{len(IMAGE_SETS)} slides created in {PPT_FILE}")


# ---------------------------------------------------------------------------
# ALTERNATIVE: Manual position override
#
# If auto-detection gives wrong results, specify positions explicitly:
#
#   positions = [
#       {'left': 0.5, 'top': 1.5, 'width': 4.5, 'height': 5.5},  # left image
#       {'left': 5.2, 'top': 1.5, 'width': 4.5, 'height': 5.5},  # right image
#   ]
#   copy_slide_replace_images(PPT_FILE, TEMPLATE_SLIDE, image_set, positions=positions)
#
# Units are inches. Use PowerPoint's Format Picture dialog to find exact values.
# ---------------------------------------------------------------------------


if __name__ == "__main__":
    main()
