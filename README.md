# PPT Batch Image Inserter

A Python toolkit for programmatically inserting images into PowerPoint presentations at scale. Perfect for researchers, data scientists, and anyone who needs to create presentation decks with dozens or hundreds of images following a consistent template.

## 🎯 What It Does

- **Batch insert images** into PowerPoint slides using YAML configuration files
- **Template-based slide creation** - copy a template slide and replace images automatically
- **Auto-detect image positions** from existing slides
- **Preserve aspect ratios** or use exact dimensions
- **Metadata tracking** - embeds source paths and labels in slides
- **Backup system** - automatic timestamped backups before modifications

## 🔬 Use Cases

- **Scientific data visualization**: Generate presentation decks with hundreds of analysis plots
- **Report automation**: Create standardized reports with consistent formatting
- **A/B testing presentations**: Quickly swap images across multiple slides
- **Educational materials**: Batch-create slide decks from image sets

## 📋 Prerequisites

- Python 3.7+
- PowerPoint files (.pptx format)
- Images in supported formats (TIF, PNG, JPG, etc.)

## 🚀 Quick Start

### Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/ppt-batch-image-inserter.git
   cd ppt-batch-image-inserter
   ```

2. **Create a conda/mamba environment:**
   ```bash
   mamba create -n ppt_inserter python=3.9
   mamba activate ppt_inserter
   ```

3. **Install dependencies:**
   ```bash
   pip install python-pptx pyyaml
   ```

### Basic Usage

#### Option 1: Using Configuration Files (Recommended)

1. **Create a YAML config file** (`my_config.yaml`):
   ```yaml
   # Path to your PowerPoint file
   presentation: "presentations/my_presentation.pptx"

   # Template slide to copy (0-based index)
   template_slide: 1

   # Auto-detect position from template
   auto_position: true

   # Base directory containing images
   base_dir: "data/images"

   # List of images to insert (each creates a new slide)
   images:
     - image1.png
     - image2.png
     - image3.png
   ```

2. **Run the batch script:**
   ```python
   import yaml
   from ppt_image_inserter import copy_slide_replace_image

   # Load config
   with open('my_config.yaml', 'r') as f:
       config = yaml.safe_load(f)

   ppt_file = config['presentation']
   base_dir = config['base_dir']

   # Process each image
   for img in config['images'][1:]:  # Skip first (already in template)
       image_path = os.path.join(base_dir, img)
       copy_slide_replace_image(
           ppt_file,
           config['template_slide'],
           image_path,
           position=None,  # Auto-detect
           store_metadata=True,
           add_label=True
       )
   ```

#### Option 2: Direct API Usage

```python
from ppt_image_inserter import insert_image, copy_slide_replace_image

# Insert image at specific position
insert_image(
    ppt_path="presentation.pptx",
    slide_index=0,
    image_path="chart.png",
    left=1.5,    # inches from left
    top=2.0,     # inches from top
    width=6.0,   # width in inches
    height=4.0   # height in inches
)

# Copy template slide and replace image (auto-detect position)
copy_slide_replace_image(
    ppt_path="presentation.pptx",
    template_slide_index=1,
    new_image_path="new_chart.png",
    position=None,  # Auto-detect from template
    store_metadata=True,
    add_label=True
)
```

## 📖 Detailed Usage

### Configuration File Format

#### Complete Example

```yaml
# Path to the PowerPoint presentation
presentation: "presentations/experiment_results.pptx"

# Template slide to copy (0-based index)
# Slide 2 in PowerPoint UI = index 1
template_slide: 1

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
base_dir: "/path/to/your/images"

# List of image filenames (each creates a new slide)
# Images are processed in the order listed
images:
  - experiment_01_results.png
  - experiment_02_results.png
  - experiment_03_results.png
  - experiment_04_results.png
```

#### Configuration Options

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `presentation` | string | Yes | Path to PowerPoint file (relative or absolute) |
| `template_slide` | integer | Yes | Index of slide to use as template (0-based) |
| `auto_position` | boolean | No | Auto-detect image position from template (default: true) |
| `position` | object | No | Manual position settings (left, top, width, height in inches) |
| `base_dir` | string | Yes | Base directory for image paths |
| `images` | list | Yes | List of image filenames to insert |

### API Reference

#### Core Functions

##### `insert_image()`
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

##### `copy_slide_replace_image()`
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

##### `delete_slide()`
Delete a slide from the presentation.

```python
delete_slide(
    ppt_path: str,
    slide_index: int,
    backup: bool = True
) -> None
```

**Parameters:**
- `ppt_path`: Path to the PowerPoint file
- `slide_index`: Index of slide to delete (0-based)
- `backup`: Whether to create backup before deletion (default: True)

##### `list_slides()`
Get information about all slides in a presentation.

```python
list_slides(ppt_path: str) -> list
```

**Returns:** List of slide information dictionaries

#### Utility Functions

##### `cm_to_inches()`
Convert centimeters to inches.

```python
cm_to_inches(cm: float) -> float
```

### Workflow Examples

#### Example 1: Scientific Analysis Results

Generate a presentation with 50 analysis plots from an experiment:

```yaml
# config_experiment_2024.yaml
presentation: "presentations/Experiment_Results_2024.pptx"
template_slide: 1
auto_position: true
base_dir: "/data/analysis/experiment_2024/plots"

images:
  - nuclear_aspect_ratio.png
  - actin_deformation.png
  - centrosome_position.png
  # ... 47 more images
```

Run the generation script:

```python
#!/usr/bin/env python
"""Generate experiment results presentation"""
import os
import yaml
from pptx import Presentation
from ppt_image_inserter import delete_slide, copy_slide_replace_image

# Load config
with open('config_experiment_2024.yaml', 'r') as f:
    config = yaml.safe_load(f)

ppt_file = config['presentation']
base_dir = config['base_dir']
images = config['images']

# Delete old slides (keep title and template)
prs = Presentation(ppt_file)
for slide_idx in reversed(range(2, len(prs.slides))):
    delete_slide(ppt_file, slide_idx)

# Create new slides from images
for img in images[1:]:  # First image already in template
    image_path = os.path.join(base_dir, img)
    copy_slide_replace_image(
        ppt_file,
        1,  # Template at index 1
        image_path,
        position=None,
        store_metadata=True,
        add_label=True
    )

print(f"Generated {len(images)} slides!")
```

#### Example 2: Weekly Report Automation

Automatically update weekly report with latest metrics:

```python
from datetime import datetime
from ppt_image_inserter import copy_slide_replace_image

# Generate weekly report filename
week = datetime.now().strftime("%Y-W%W")
ppt_file = f"reports/Weekly_Report_{week}.pptx"

# Metrics to include
metrics = [
    "metrics/user_growth.png",
    "metrics/revenue_chart.png",
    "metrics/engagement_stats.png"
]

# Build presentation
for metric in metrics:
    copy_slide_replace_image(
        ppt_file,
        template_slide_index=0,
        new_image_path=metric,
        position=None
    )
```

## 🎨 Best Practices

### 1. Template Slide Design

- **Create a template slide (slide 2)** with your desired layout, formatting, and a placeholder image
- **Position matters**: The first image in your template determines where all new images will be placed
- **Keep it simple**: Avoid complex layouts that might not work well with batch processing

### 2. Image Organization

```
project/
├── config.yaml
├── script.py
├── presentations/
│   └── my_presentation.pptx
└── images/
    ├── experiment_01/
    │   ├── metric1.png
    │   └── metric2.png
    └── experiment_02/
        ├── metric1.png
        └── metric2.png
```

### 3. Configuration Management

- **Use descriptive names**: `config_experiment_histones.yaml` not `config1.yaml`
- **Version control**: Keep configs in git (but not the .pptx files!)
- **Comment your configs**: Explain what each image represents
- **Relative paths**: Use relative paths when possible for portability

### 4. Error Handling

Always check if images exist before processing:

```python
import os

for img in images:
    img_path = os.path.join(base_dir, img)
    if not os.path.exists(img_path):
        print(f"WARNING: Image not found: {img_path}")
        continue

    # Process image...
```

### 5. Backup Strategy

The tool creates automatic backups, but also:
- **Version your presentations**: `presentation_v1.pptx`, `presentation_v2.pptx`
- **Use git**: Track changes to scripts and configs
- **Test on copies**: Test new scripts on a copy of your presentation first

## 🔧 Troubleshooting

### Common Issues

#### Images not appearing at the right position

**Problem**: Images are inserted but in the wrong location.

**Solution**:
1. Check that `auto_position: true` is set in your config
2. Ensure your template slide (slide 2) has exactly one image
3. Verify the template image is where you want new images to appear

#### "File not found" errors

**Problem**: Script can't find image files.

**Solution**:
1. Check that `base_dir` path is correct (absolute or relative to script)
2. Verify image filenames match exactly (case-sensitive!)
3. Use forward slashes `/` even on Windows, or double backslashes `\\`

#### Presentation becomes corrupted

**Problem**: PowerPoint file won't open after processing.

**Solution**:
1. Check the backup files in `PPT/.backups/`
2. Ensure you're not modifying the presentation in PowerPoint while the script runs
3. Verify you have write permissions to the file

#### Slow performance with many images

**Problem**: Processing hundreds of slides is slow.

**Solution**:
1. This is normal - PowerPoint files are complex
2. Consider batching: process 50 slides at a time
3. Use SSD storage for better I/O performance
4. Close other applications to free up memory

## 🤝 Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

### Development Setup

```bash
# Clone the repo
git clone https://github.com/yourusername/ppt-batch-image-inserter.git
cd ppt-batch-image-inserter

# Create development environment
mamba create -n ppt_dev python=3.9 pytest black flake8
mamba activate ppt_dev

# Install in editable mode
pip install -e .
```

### Running Tests

```bash
pytest tests/
```

## 📄 License

MIT License - see [LICENSE](LICENSE) for details

## 🙏 Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/)
- Developed for biological microscopy data analysis workflows
- Thanks to all contributors and users!

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/yourusername/ppt-batch-image-inserter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/yourusername/ppt-batch-image-inserter/discussions)

## 🗺️ Roadmap

- [ ] Add support for multiple images per slide
- [ ] GUI for config file generation
- [ ] Support for video/animation insertion
- [ ] Template gallery with common layouts
- [ ] Performance optimizations for large batches
- [ ] Integration with Jupyter notebooks
- [ ] Cloud storage support (S3, Google Drive)

---

**Made with ❤️ for researchers who have better things to do than manually insert hundreds of images into PowerPoint**
