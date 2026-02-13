# Examples

This directory contains example configurations and scripts demonstrating how to use the PPT Batch Image Inserter.

## Quick Start

The `batch_insert_images.py` script is a **generic script that works with ANY config file**. You don't need to create separate scripts for each dataset - just pass different config files as command-line arguments.

### Basic Usage

```bash
# Run with the basic example config
python batch_insert_images.py example_config.yaml

# Run with the microscopy analysis example
python batch_insert_images.py example_config_PLL_aCD3_vim_fixed_cells.yaml

# Run with your own custom config
python batch_insert_images.py path/to/your_config.yaml
```

## Available Example Configs

### `example_config.yaml`
Basic configuration demonstrating:
- Simple image list with generic filenames
- Auto-detected image positions
- Preserved slides configuration

**Use this as a starting point** for your own configs.

### `example_config_PLL_aCD3_vim_fixed_cells.yaml`
Real-world microscopy analysis example demonstrating:
- Large-scale batch processing (23 images)
- Organized by metric categories (nuclear, actin, centrosome, deformation, vimentin)
- Actual research workflow from cell biology experiments

**Use this to see** how the tool scales to real research projects.

## How It Works

1. **Copy an example config** or create your own:
   ```bash
   cp example_config.yaml my_experiment_config.yaml
   ```

2. **Edit the config** to match your needs:
   - Set the PowerPoint file path
   - Specify the template slide index
   - List which slides to preserve
   - Add your image filenames

3. **Run the script** with your config:
   ```bash
   python batch_insert_images.py my_experiment_config.yaml
   ```

The script will:
- Delete old slides (except those in `preserve_slides`)
- Create new slides by copying the template
- Insert your images with consistent formatting
- Track metadata (filenames, paths) on each slide

## Common Workflows

### Updating an Existing Presentation

```bash
# First run - creates initial slides
python batch_insert_images.py experiment_v1.yaml

# Update data, regenerate plots, then re-run
python batch_insert_images.py experiment_v1.yaml  # Regenerates all slides
```

### Multiple Experiments/Datasets

```bash
# Different experiments, different configs, SAME script
python batch_insert_images.py configs/experiment_activation.yaml
python batch_insert_images.py configs/experiment_inhibition.yaml
python batch_insert_images.py configs/experiment_control.yaml
```

**No need to create separate Python scripts!** The config files contain all the differences.

## Creating Your Own Config

See the [main README](../README.md#basic-usage) and [detailed usage guide](../docs/DETAILED_USAGE.md) for complete documentation on config file format.

Key parameters:
- `presentation`: Path to your PowerPoint file
- `template_slide`: Which slide to use as template (0-indexed)
- `preserve_slides`: Which slides to keep (e.g., `[0, 1]` = keep title and template)
- `base_dir`: Directory containing your images
- `images`: List of image filenames to process

## Tips

- **Version control your configs**, not your PowerPoint files
- **Use descriptive config names**: `config_experiment_2024-02-10.yaml`
- **Test with a few images first**, then scale up to large batches
- **Check backups** in `PPT/.backups/` if something goes wrong

## Questions?

See the [main README](../README.md) or [detailed usage guide](../docs/DETAILED_USAGE.md) for more information.
