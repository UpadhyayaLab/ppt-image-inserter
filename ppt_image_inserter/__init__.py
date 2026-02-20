"""
PPT Batch Image Inserter

A Python toolkit for batch-inserting images into PowerPoint presentations.
Designed for researchers who need to create presentation decks with large
numbers of analysis plots following a consistent template.
"""

__version__ = "1.0.0"

# Import all public functions from submodules
from .core import (
    insert_image,
    insert_image_preserve_aspect,
    list_slides,
)

from .position import (
    cm_to_inches,
    get_image_position,
    get_all_image_positions,
)

from .slide_utils import (
    duplicate_slide,
    remove_pictures_from_slide,
    remove_all_text_from_slide,
    delete_slide,
)

from .backup import (
    backup_presentation,
)

from .metadata import (
    extract_image_metadata,
)

from .workflow import (
    copy_slide_replace_image,
    copy_slide_replace_images,
    replace_image_on_existing_slide,
    add_label_to_existing_slide,
)

# Define public API
__all__ = [
    # Core functions
    'insert_image',
    'insert_image_preserve_aspect',
    'list_slides',

    # Position utilities
    'cm_to_inches',
    'get_image_position',
    'get_all_image_positions',

    # Slide manipulation
    'duplicate_slide',
    'remove_pictures_from_slide',
    'remove_all_text_from_slide',
    'delete_slide',

    # Backup system
    'backup_presentation',

    # Metadata extraction
    'extract_image_metadata',

    # High-level workflows
    'copy_slide_replace_image',
    'copy_slide_replace_images',
    'replace_image_on_existing_slide',
    'add_label_to_existing_slide',
]
