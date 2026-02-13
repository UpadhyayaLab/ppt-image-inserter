# PPT Batch Image Inserter

A Python toolkit for batch-inserting images into PowerPoint presentations. Designed for researchers who need to create presentation decks with large numbers of analysis plots following a consistent template.

**Note**: The batch processing workflow currently supports **one image per slide**. Each slide is created from the template with a single image replacement.

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
pip install python-pptx pyyaml
```

### Basic Usage

1. **Create a YAML config** (`config.yaml`):
   ```yaml
   presentation: "presentations/my_presentation.pptx"
   template_slide: 1  # Slide 2 in PowerPoint = index 1
   auto_position: true
   base_dir: "data/analysis/plots"

   images:
     - nuc_aspect_ratio_grid.tif
     - actin_deform_ratio_grid.tif
     - centrosome_center_z_grid.tif
   ```

2. **Run the batch script** (see `examples/batch_insert_images.py`):
   ```bash
   python examples/batch_insert_images.py config.yaml
   ```

That's it! Your PowerPoint now has slides with all your analysis plots.

## How It Works

1. **Template slide**: Create slide 2 with your desired layout and one placeholder image
2. **Config file**: List all images you want to insert
3. **Batch process**: Script copies slide 2 for each image, replacing the placeholder
4. **Result**: Presentation with consistent formatting across all slides

## Prerequisites

- Python 3.7+
- PowerPoint files (.pptx format)
- Images in common formats (PNG, JPG, TIFF, GIF, BMP)

## Key Features

- **Auto-detect positions** - Automatically detects image position from template
- **Template-based** - Consistent formatting across all slides
- **Metadata tracking** - Stores image paths and filenames
- **Automatic backups** - Creates timestamped backups before modifications
- **Config-driven** - YAML configuration for easy batch processing

## Best Practices

### Template Design
- Create slide 2 with desired layout and one placeholder image
- The first image position determines where all new images go
- One image per slide in batch mode

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

### Configuration
- Use descriptive config names: `config_experiment_2024.yaml`
- Version control configs (not .pptx files)
- Use relative paths for portability

## Troubleshooting

**Images in wrong position?**
- Check `auto_position: true` is set
- Ensure template slide has exactly one image
- Verify template image is positioned correctly

**File not found errors?**
- Check `base_dir` path is correct
- Verify filenames match exactly (case-sensitive)
- Use forward slashes `/` in paths

**Presentation corrupted?**
- Check backup files in `PPT/.backups/`
- Don't modify PPT in PowerPoint while script runs
- Verify write permissions

**Slow with many images?**
- Normal for PowerPoint files (complex XML)
- Process in batches of 50-100 slides
- Use SSD storage for better performance

## Documentation

- [Detailed Usage Guide](docs/DETAILED_USAGE.md) - Complete function reference and examples
- [Examples](examples/) - Sample configs and scripts
- [CLAUDE.md](CLAUDE.md) - Instructions for Claude Code instances

## Support

- **Issues**: [GitHub Issues](https://github.com/UpadhyayaLab/ppt-image-inserter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/UpadhyayaLab/ppt-image-inserter/discussions)

## Future Features

- Multiple images per slide support
- GUI for config file generation
- Performance optimizations for large batches
- Template gallery with common layouts

## License

MIT License - see [LICENSE](LICENSE)

## Acknowledgments

- Built with [python-pptx](https://python-pptx.readthedocs.io/)
- Developed for biological microscopy data analysis workflows
