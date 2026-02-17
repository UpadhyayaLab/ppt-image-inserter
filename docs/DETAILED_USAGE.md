# Detailed Usage Guide

## Configuration File Format

### Complete Example

```yaml
# Path to the PowerPoint presentation
presentation: "presentations/experiment_results.pptx"

# Template slide to copy (0-based index)
# Slide 2 in PowerPoint UI = index 1
template_slide: 1

# Slides to preserve (not delete) - 0-based indices
# Default: [0, template_slide] (keeps title slide and template)
preserve_slides: [0, 1]

# Backup directory for automatic backups (optional)
# Default: PPT/backups
# backup_dir: "custom/backup/location"

# Auto-detect position from the first image in the template slide
auto_position: true

# Manual position (only used if auto_position is false)
# Uncomment and set values if you want to override auto-detection
# position:
#   left: 0.5      # inches from left edge
#   top: 1.0       # inches from top edge
#   width: 8.0     # image width in inches
#   height: 6.0    # image height in inches

# Base directory containing all images
base_dir: "C:/data/analysis/experiment_2024/plots"

# List of image filenames (each creates a new slide)
# Images are processed in the order listed
images:
  - nuc_aspect_ratio_grid.tif
  - actin_deform_ratio_grid.tif
  - centrosome_center_z_grid.tif
  - deepest_invag_fraction_chull_volume_grid.tif
```

### Configuration Options

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `presentation` | string | Yes | Path to PowerPoint file (relative or absolute) |
| `template_slide` | integer | Yes | Index of slide to use as template (0-based) |
| `preserve_slides` | list | No | Slide indices to preserve (default: [0, template_slide]) |
| `backup_dir` | string | No | Backup directory path (default: PPT/backups) |
| `auto_position` | boolean | No | Auto-detect image position from template (default: true) |
| `position` | object | No | Manual position settings (left, top, width, height in inches) |
| `base_dir` | string | Yes | Base directory for image paths |
| `images` | list | Yes | List of image filenames to insert |

## Function Reference

### Core Functions

#### `insert_image()`
Insert an image at a specific position on an existing slide.

```python
insert_image(
    ppt_path: str,
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: float,
    height: float
) -> None
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `slide_index`: Index of the slide (0-based)
- `image_path`: Path to the image file
- `left`: Distance from left edge in inches
- `top`: Distance from top edge in inches
- `width`: Image width in inches
- `height`: Image height in inches

#### `copy_slide_replace_image()`
Duplicate a template slide and replace its image.

```python
copy_slide_replace_image(
    ppt_path: str,
    template_slide_index: int,
    new_image_path: str,
    position: dict = None,
    store_metadata: bool = True,
    add_label: bool = True
) -> int
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `template_slide_index`: Index of template slide to copy (0-based)
- `new_image_path`: Path to the new image
- `position`: Optional dict with keys: `left`, `top`, `width`, `height`. If None, auto-detects from template
- `store_metadata`: Whether to store image path in slide notes
- `add_label`: Whether to add image filename as text label

**Returns:** Index of the newly created slide

#### `delete_slide()`
Delete a slide from the presentation. **Note:** Backups are automatically created before deletion.

```python
delete_slide(
    ppt_path: str,
    slide_index: int,
    backup_base: str = 'PPT/backups'
) -> None
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `slide_index`: Index of slide to delete (0-based)
- `backup_base`: Base directory for backups (default: 'PPT/backups')

**Note:** This function automatically creates timestamped backups before deleting the slide.

#### `list_slides()`
Get information about all slides in a presentation.

```python
list_slides(ppt_path: str) -> list
```

**Returns:** List of slide information dictionaries

#### `insert_image_preserve_aspect()`
Insert an image while preserving its aspect ratio. Specify either width or height, and the other dimension will be calculated automatically.

```python
insert_image_preserve_aspect(
    ppt_path: str,
    slide_index: int,
    image_path: str,
    left: float,
    top: float,
    width: float = None,
    height: float = None
) -> None
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `slide_index`: Index of the slide (0-based)
- `image_path`: Path to the image file
- `left`: Distance from left edge in inches
- `top`: Distance from top edge in inches
- `width`: Image width in inches (height will be calculated). Optional.
- `height`: Image height in inches (width will be calculated). Optional.

**Note:** Specify exactly one of `width` or `height` to preserve aspect ratio.

#### `replace_image_on_existing_slide()`
Replace the image on an existing slide with a new one. Auto-detects position from the current image.

```python
replace_image_on_existing_slide(
    ppt_path: str,
    slide_index: int,
    new_image_path: str,
    store_metadata: bool = True,
    add_label: bool = True
) -> None
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `slide_index`: Index of the slide to update (0-based)
- `new_image_path`: Path to the new image to insert
- `store_metadata`: If True, stores image path in alt text (default: True)
- `add_label`: If True, adds visible text label showing filename (default: True)

### Utility Functions

#### `cm_to_inches()`
Convert centimeters to inches.

```python
cm_to_inches(cm: float) -> float
```

#### `get_image_position()`
Extract position and size information from an image on a slide. Useful for auto-detecting template image parameters.

```python
get_image_position(
    ppt_path: str,
    slide_index: int,
    image_index: int = 0
) -> dict
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `slide_index`: Slide number (0-based index)
- `image_index`: Which image to inspect if multiple exist (default: 0)

**Returns:** Dictionary with keys `'left'`, `'top'`, `'width'`, `'height'` (all in inches)

**Example:**
```python
>>> pos = get_image_position('presentation.pptx', 1, 0)
>>> print(pos)
{'left': 0.14, 'top': 0.96, 'width': 5.43, 'height': 2.72}
```

### Slide Manipulation Functions

#### `duplicate_slide()`
Duplicate a slide by creating a new slide and copying all shapes. Internal helper function used by batch workflows.

```python
duplicate_slide(prs: Presentation, slide_index: int) -> Slide
```

**Parameters:**
- `prs`: The Presentation object (not a file path - this is a lower-level function)
- `slide_index`: Index of the slide to duplicate (0-based)

**Returns:** The newly created slide object

**Note:** This is a lower-level function. Most users should use `copy_slide_replace_image()` instead.

#### `remove_pictures_from_slide()`
Remove all picture shapes from a slide. Uses XML manipulation.

```python
remove_pictures_from_slide(slide: Slide) -> int
```

**Parameters:**
- `slide`: The slide to remove pictures from (Slide object, not index)

**Returns:** Number of pictures removed

**Note:** Changes are only saved when you call `prs.save()`.

#### `remove_all_text_from_slide()`
Remove ALL text boxes and placeholders from a slide. This includes both regular text boxes AND placeholders (title, body, etc.).

```python
remove_all_text_from_slide(slide: Slide) -> int
```

**Parameters:**
- `slide`: The slide to remove text from (Slide object, not index)

**Returns:** Number of text elements removed

### Backup and Metadata Functions

#### `backup_presentation()`
Create backups in multiple time-interval categories. Implements a smart backup strategy that maintains ONE backup per time interval category.

```python
backup_presentation(
    ppt_path: str,
    backup_base: str = 'PPT/backups'
) -> dict
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file to backup
- `backup_base`: Base directory for backups (default: 'PPT/backups')

**Returns:** Dictionary mapping category names to backup file paths for backups created

**Backup Categories:**
- `'latest'`: Always creates a backup (0 seconds threshold)
- `'5min'`: Creates backup if >5 minutes since last backup in this category
- `'10min'`: Creates backup if >10 minutes since last backup
- `'30min'`: Creates backup if >30 minutes since last backup
- `'hourly'`: Creates backup if >1 hour since last backup

**Example:**
```python
>>> backups = backup_presentation('presentation.pptx')
>>> print(f"Created backups: {list(backups.keys())}")
Created backups: ['latest', '5min', '10min']
```

#### `extract_image_metadata()`
Extract image source metadata from all slides in a presentation. Retrieves metadata stored in alt text fields.

```python
extract_image_metadata(ppt_path: str) -> list
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file

**Returns:** List of dictionaries, each containing:
- `slide_index` (int): 0-based slide index
- `slide_number` (int): 1-based slide number (UI numbering)
- `original_path` (str): Original image path from alt text (or None)
- `filename` (str): Image filename
- `position` (dict): Position/size dict with 'left', 'top', 'width', 'height'

**Example:**
```python
>>> metadata = extract_image_metadata('presentation.pptx')
>>> for entry in metadata:
...     print(f"Slide {entry['slide_number']}: {entry['filename']}")
...     print(f"  Path: {entry['original_path']}")
Slide 2: nuc_aspect_ratio_grid.tif
  Path: J:/FF/data/plots/nuc_aspect_ratio_grid.tif
```

## Workflow Example: Microscopy Analysis

Generate a presentation with nuclear morphology analysis plots from a microscopy experiment.

### Config File

```yaml
# config_nuclear_morphology_2024.yaml
presentation: "presentations/Nuclear_Morphology_Analysis.pptx"
template_slide: 1
auto_position: true
base_dir: "J:/FF/fixed_cell/results_compilation/analysis_20240210/grid_panels"

images:
  - nuc_aspect_ratio_grid.tif
  - actin_deform_ratio_grid.tif
  - centrosome_center_z_rel_bottom_actin_plane_grid.tif
  - avg_normal_angle_adaptive_region_growth_grid.tif
  - nuc_cent_closest_dist_grid.tif
  - chull_max_D_grid.tif
  - chull_mean_D_cent_ratio_grid.tif
  - centrosome_dist_deepest_real_avg_periphery_ratio_grid.tif
  - concavity_index_around_cent_grid.tif
  - deepest_invag_fraction_chull_volume_grid.tif
  - nuc_resized_solidity_grid.tif
  - nuc_SA_mesh_grid.tif
  - nuc_volume_mesh_grid.tif
  - deepest_region_periph_ratio_025um_grid.tif
```

### Processing Script

```python
#!/usr/bin/env python
"""
Generate nuclear morphology analysis presentation
Processes microscopy data analysis results into a presentation deck
"""
import os
import yaml
from pptx import Presentation
from ppt_image_inserter import delete_slide, copy_slide_replace_image

# Load configuration
with open('config_nuclear_morphology_2024.yaml', 'r') as f:
    config = yaml.safe_load(f)

ppt_file = config['presentation']
base_dir = config['base_dir']
images = config['images']

# Step 1: Delete old slides (keep title and template)
print("Deleting old slides...")
prs = Presentation(ppt_file)
for slide_idx in reversed(range(2, len(prs.slides))):
    delete_slide(ppt_file, slide_idx)

print(f"Slides 1-2 remain.\n")

# Step 2: Create new slides from analysis images
print(f"Creating {len(images)-1} slides from analysis results...\n")

success_count = 0
error_count = 0

for i, img_filename in enumerate(images[1:]):  # Skip first (already in template)
    image_path = os.path.join(base_dir, img_filename)
    slide_num = i + 3  # Slides 3, 4, 5, ...

    print(f"[{i+1}/{len(images)-1}] Creating slide {slide_num}: {img_filename}")

    if not os.path.exists(image_path):
        print(f"  ERROR: Image not found")
        error_count += 1
        continue

    try:
        new_idx = copy_slide_replace_image(
            ppt_file,
            1,  # Template at index 1
            image_path,
            position=None,  # Auto-detect
            store_metadata=True,
            add_label=True
        )
        print(f"  SUCCESS: Created slide {new_idx + 1}")
        success_count += 1
    except Exception as e:
        print(f"  ERROR: {e}")
        error_count += 1

print(f"\nComplete: {success_count}/{len(images)-1} slides created")
if error_count > 0:
    print(f"Errors: {error_count}")
```

## Advanced Usage

### Manual Position Override

If you need precise control over image position:

```yaml
presentation: "my_presentation.pptx"
template_slide: 1
auto_position: false  # Disable auto-detection

# Specify exact position
position:
  left: 0.5      # inches from left
  top: 1.0       # inches from top
  width: 8.0     # inches wide
  height: 6.0    # inches tall

base_dir: "images/"
images:
  - image1.tif
  - image2.tif
```

### Direct Function Usage (Without Config)

For one-off tasks or custom workflows:

```python
from ppt_image_inserter import copy_slide_replace_image

# Process a single image
copy_slide_replace_image(
    ppt_path="PPT/analysis_results.pptx",
    template_slide_index=1,
    new_image_path="plots/nuc_aspect_ratio_grid.tif",
    position=None,
    store_metadata=True,
    add_label=True
)
```
