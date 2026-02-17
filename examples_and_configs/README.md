# Examples and Configs

This directory contains scripts and configuration files for using the PPT Batch Image Inserter.

## Scripts

### `batch_insert_images.py` — Generic Batch Script

The main entry point. Reads any YAML config and runs the full workflow: validates images, deletes old slides, creates new ones from the template.

```bash
# Run with any config
python batch_insert_images.py configs/example_config.yaml
python batch_insert_images.py configs/my_experiment.yaml

# Config path can be relative or absolute
python batch_insert_images.py /path/to/my_config.yaml
```

**Use this for**: any batch job — single image per slide or multiple images per slide.

### `example_multi_image.py` — Multi-Image API Example

Demonstrates how to use `copy_slide_replace_images()` directly from Python (no config file needed). Useful when you want programmatic control over which images go on which slide, or when the YAML config format doesn't fit your workflow.

```bash
# Edit the paths at the top of the file, then run
python example_multi_image.py
```

**Use this for**: understanding the API, one-off scripts, or custom workflows.

---

## configs/ — Configuration Files

All YAML configs live here. The batch script accepts any of them as an argument.

### Example Configs (starting points)

| File | Description |
|------|-------------|
| `example_config.yaml` | Basic template — mix of single and multi-image slides |
| `example_config_PLL_aCD3_vim_fixed_cells.yaml` | Real-world microscopy example (23 images, single per slide) |

Copy one of these as a starting point:
```bash
cp configs/example_config.yaml configs/my_experiment.yaml
# Then edit my_experiment.yaml
```

### Research Configs

| File | Description |
|------|-------------|
| `config_ctrl_dznep_montages.yaml` | XZ MIP montages — 4 images/slide (2×2 grid) |
| `config_ctrl_dznep_laminb_hoechst_montages.yaml` | Lamin B fixed slice montages — 2 images/slide |
| `config_ctrl_dznep_btub_hoechst_montages.yaml` | bTub fixed slice montages — 2 images/slide |

---

## Multi-Image Support

Use list syntax in your config to put multiple images on one slide:

```yaml
images:
  - single_image.png                         # one image → one slide
  - [control.png, treated.png]               # two images → one slide (side-by-side)
  - [rep1.png, rep2.png, rep3.png]           # three images → one slide
```

The template slide must have the same number of placeholder images as the list. Positions are auto-detected from the template.

---

## Workflow Tips

- **Different experiments, same script**: only the config changes
  ```bash
  python batch_insert_images.py configs/experiment_A.yaml
  python batch_insert_images.py configs/experiment_B.yaml
  ```

- **Re-run to update**: the script deletes and recreates all content slides, so re-running after updating images gives a fresh result

- **Template preservation**: use `output_path` in the config to write to a new file, leaving the original template untouched

- **Version control your configs**, not your PowerPoint files

See the [main README](../README.md) and [detailed usage guide](../docs/DETAILED_USAGE.md) for full documentation.
