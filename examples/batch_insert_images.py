#!/usr/bin/env python
"""
Example: Batch Insert Images into PowerPoint

This script demonstrates how to use a YAML config file to batch-insert
images into a PowerPoint presentation.

Usage:
    python batch_insert_images.py config.yaml

Requirements:
    - PowerPoint file with title slide (slide 1) and template slide (slide 2)
    - Template slide should have one image at the desired position
    - All images listed in the config file should exist
"""

import sys
import os
import yaml
from pptx import Presentation

# Add parent directory to path to import ppt_image_inserter
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ppt_image_inserter import delete_slide, copy_slide_replace_image


def main(config_path):
    """Process images according to config file."""

    # Load configuration
    print(f"Loading config from: {config_path}")
    with open(config_path, 'r') as f:
        config = yaml.safe_load(f)

    ppt_file = config['presentation']
    base_dir = config.get('base_dir', '')
    images = config['images']
    template_slide_index = config.get('template_slide', 1)

    # Get preserve_slides, default to [0, template_slide] if not specified
    preserve_slides = config.get('preserve_slides', [0, template_slide_index])

    # Get backup directory, default to 'PPT/backups' if not specified
    backup_dir = config.get('backup_dir', 'PPT/backups')

    # Ensure template_slide is in preserve_slides if not explicitly excluded
    if template_slide_index not in preserve_slides:
        print(f"Warning: template_slide {template_slide_index} not in preserve_slides")

    print(f"PowerPoint file: {ppt_file}")
    print(f"Base directory: {base_dir}")
    print(f"Template slide: {template_slide_index + 1} (index {template_slide_index})")
    print(f"Preserve slides: {[idx + 1 for idx in preserve_slides]} (indices {preserve_slides})")
    print(f"Images to process: {len(images)}")

    # Step 1: Delete old slides (except preserved ones)
    print("\nStep 1: Deleting old slides...")
    print("=" * 70)

    prs = Presentation(ppt_file)
    total_slides = len(prs.slides)

    # Collect slides to delete (all except preserved)
    slides_to_delete = [idx for idx in range(total_slides) if idx not in preserve_slides]

    if slides_to_delete:
        # Delete in reverse order to maintain indices
        for slide_idx in reversed(slides_to_delete):
            print(f"Deleting slide {slide_idx + 1}...")
            delete_slide(ppt_file, slide_idx, backup_base=backup_dir)
        print(f"Deleted {len(slides_to_delete)} slide(s)")
    else:
        print("No slides to delete")

    print(f"\n[SUCCESS] Preserved slides: {[idx + 1 for idx in preserve_slides]}\n")

    # Step 2: Create new slides from images
    print("Step 2: Creating slides from images...")
    print("=" * 70)

    # Skip first image (should already be in template slide)
    remaining_images = images[1:]

    print(f"Creating {len(remaining_images)} new slides\n")

    success_count = 0
    error_count = 0

    for i, image_filename in enumerate(remaining_images):
        # Construct full image path
        if isinstance(image_filename, dict):
            image_path = image_filename['path']
        else:
            image_path = os.path.join(base_dir, image_filename)

        slide_num = i + 3  # Slides 3, 4, 5, ...

        print(f"\n[{i+1}/{len(remaining_images)}] Creating slide {slide_num}")
        print(f"Image: {os.path.basename(image_path)}")
        print("-" * 70)

        # Check if image exists
        if not os.path.exists(image_path):
            print(f"[ERROR] Image file not found: {image_path}")
            error_count += 1
            continue

        try:
            # Copy template slide and insert image
            new_idx = copy_slide_replace_image(
                ppt_file,
                template_slide_index,
                image_path,
                position=None,  # Auto-detect from template
                store_metadata=True,
                add_label=True
            )
            print(f"[SUCCESS] Created slide {new_idx + 1}")
            success_count += 1
        except Exception as e:
            print(f"[ERROR] Failed to create slide: {e}")
            error_count += 1

    # Summary
    print(f"\n{'=' * 70}")
    print(f"Batch Processing Complete")
    print(f"{'=' * 70}")
    print(f"Successfully created: {success_count}/{len(remaining_images)} slides")
    if error_count > 0:
        print(f"Errors: {error_count}")
    print(f"{'=' * 70}")

    # Exit with error code if any failures
    if error_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python batch_insert_images.py <config.yaml>")
        print("\nExample:")
        print("  python batch_insert_images.py example_config.yaml")
        sys.exit(1)

    config_path = sys.argv[1]

    if not os.path.exists(config_path):
        print(f"Error: Config file not found: {config_path}")
        sys.exit(1)

    main(config_path)
