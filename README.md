# PPT Batch Image Inserter

A Python toolkit for batch-inserting images into PowerPoint presentations. Designed for researchers who need to create presentation decks with large numbers of analysis plots following a consistent template.

**Note**: The batch processing workflow currently supports **one image per slide**. Each slide is created from the template with a single image replacement, and metadata (filename, path) is automatically tracked.

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

1. **Copy and edit the example config**:
   ```bash
   # Basic example
   cp examples/example_config.yaml config.yaml

   # Or use the scientific microscopy example
   cp examples/example_config_PLL_aCD3_vim_fixed_cells.yaml config.yaml

   # Edit config.yaml with your presentation path, image directory, and image list
   ```

2. **Run the batch script**:
   ```bash
   python examples/batch_insert_images.py config.yaml
   ```

## How It Works

1. **Template slide**: Create a slide with your desired layout and one placeholder image. Specify which slide to use as the template in your config (0-indexed: slide 2 in PowerPoint = index 1)
2. **Config file**: List all images you want to insert and which slides to preserve
3. **Batch process**: Script copies the template slide for each image, replacing the placeholder
4. **Result**: Presentation with consistent formatting across all slides

## Prerequisites

- Python 3.9+
- PowerPoint files (.pptx format)
- Images in common formats (PNG, JPG, TIFF, GIF, BMP)
- Windows OS

**Important**: Close your PowerPoint file before running the script. The script needs exclusive access to modify the file.

## Key Features

- **Template-based** - Consistent formatting across all slides
- **Config-driven** - YAML configuration for easy batch processing
- **Metadata tracking** - Displays image filenames on slides and stores paths in notes
- **Automatic backups** - Creates timestamped backups before modifications

## Best Practices

### Template Design
- Create a template slide with your desired layout and one placeholder image
- Specify the template slide index in your config file (0-indexed)
- This image's position is auto-detected and used for all new slides
- One image per slide in batch mode
- Use `preserve_slides` in config to specify which slides to keep (default: title slide and template)

### File Organization
```
project/
├── config.yaml
├── presentations/
│   └── analysis_results.pptx
└── data/
    └── analysis/
        ├── nuc_aspect_ratio_grid.tif
        ├── actin_deform_ratio_grid.tif
        └── ...
```

## Troubleshooting

**File not found errors?**
- Check `base_dir` path is correct
- Verify filenames match exactly (case-sensitive)
- Use forward slashes `/` in paths

**Presentation corrupted or slides lost**
- Check backup files in `PPT/backups/`
- **Close PowerPoint before running the script** - file must not be open
- Verify write permissions

## Documentation

- [Detailed Usage Guide](docs/DETAILED_USAGE.md) - Complete function reference and examples
- [Examples](examples/) - Sample configs and scripts

## Support

- **Issues**: [GitHub Issues](https://github.com/UpadhyayaLab/ppt-image-inserter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/UpadhyayaLab/ppt-image-inserter/discussions)
