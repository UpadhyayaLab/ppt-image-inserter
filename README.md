# PPT Batch Image Inserter

A Python toolkit for batch-inserting images into PowerPoint presentations. Designed for researchers who need to create presentation decks with large numbers of analysis plots following a consistent template.

## Quick Start

### Installation

```bash
# Clone and navigate
git clone https://github.com/UpadhyayaLab/ppt-image-inserter.git
cd ppt-image-inserter

# Create environment
conda create -n ppt_inserter python=3.9
conda activate ppt_inserter

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

1. **Copy and edit an example config**:
   ```cmd
   copy examples_and_configs\configs\example_config.yaml my_config.yaml
   ```
   Or for the scientific microscopy example:
   ```cmd
   copy examples_and_configs\configs\example_config_PLL_aCD3_vim_fixed_cells.yaml my_config.yaml
   ```
   Then edit `my_config.yaml` with your presentation path, image directory, and image list.

2. **Run the batch script**:
   ```bash
   python examples_and_configs/batch_insert_images.py my_config.yaml
   ```

## How It Works

1. **Template slide**: Create a slide with your desired layout and one or more placeholder images. Specify which slide to use as the template in your config (0-indexed: slide 2 in PowerPoint = index 1)
2. **Config file**: List all images you want to insert and which slides to preserve
3. **Batch process**: Script copies the template slide for each image set, replacing the placeholder(s)
4. **Result**: Presentation with consistent formatting across all slides

## Prerequisites

- Python 3.9+
- PowerPoint files (.pptx format)
- Images in common formats (PNG, JPG, TIFF, GIF, BMP)
- Windows OS

**Important**: Close your PowerPoint file before running the script. The script needs exclusive access to modify the file.

## Key Features

- **Template-based** - Consistent formatting across all slides
- **Multi-image per slide** - Side-by-side comparisons using list syntax (e.g., `[control.png, treated.png]`)
- **Config-driven** - YAML configuration for easy batch processing
- **Template preservation** - Use `output_path` to create a new file, leaving the original untouched
- **Metadata tracking** - Displays image filenames on slides and stores paths in alt text
- **Automatic backups** - Creates timestamped backups before modifications
- **Image pre-validation** - Checks all images exist before making any changes

## Best Practices

### Template Design
- Create a template slide with your desired layout and placeholder image(s)
- Specify the template slide index in your config file (0-indexed)
- For single-image slides: one placeholder; position is auto-detected
- For multi-image slides (e.g., side-by-side): add N placeholder images; positions are auto-detected from left to right
- Use `preserve_slides` in config to specify which slides to keep (default: title slide and template)

### File Organization
```
project/
├── examples_and_configs/
│   ├── batch_insert_images.py      # Generic batch script
│   ├── example_multi_image.py      # Multi-image API example
│   └── configs/
│       ├── example_config.yaml     # Basic example
│       └── my_experiment.yaml      # Your research configs
├── presentations/
│   └── analysis_results.pptx
└── data/
    └── analysis/
        ├── control_result.png
        ├── treated_result.png
        └── ...
```

## Troubleshooting

**File not found errors?**
- Check `base_dir` path is correct
- Verify filenames match exactly (case-sensitive)
- Use forward slashes `/` in paths

**Presentation corrupted or slides lost**
- Check backup files in the `backups/` subfolder next to your presentation
- **Close PowerPoint before running the script** - file must not be open
- Verify write permissions

**Images not in the right position?**
- Check that the template slide contains the correct number of placeholder images
- For multi-image slides, placeholder order in the template determines left-to-right image order

## Documentation

- [Detailed Usage Guide](docs/DETAILED_USAGE.md) - Complete function reference and examples
- [Examples](examples_and_configs/) - Sample configs and scripts

## Support

- **Issues**: [GitHub Issues](https://github.com/UpadhyayaLab/ppt-image-inserter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/UpadhyayaLab/ppt-image-inserter/discussions)
