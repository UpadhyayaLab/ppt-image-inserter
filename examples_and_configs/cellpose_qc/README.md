# Cellpose QC deck builder

Build a PowerPoint deck where each slide shows the **XY MIP of a raw FOV
channel** (left) next to the **cellpose mask** (right) at matched
dimensions. Used to visually spot cells the segmentation missed.

## Layout

```
cellpose_qc/
    mip_mask_compositor.py     # pure helpers: load TIFF, MIP, mask, side-by-side
    build_cellpose_qc_deck.py  # CLI: YAML config -> .pptx
    configs/                   # one YAML per dataset/condition
```

## Usage

```bash
conda run -n PPT_editing python examples_and_configs/cellpose_qc/build_cellpose_qc_deck.py \
    examples_and_configs/cellpose_qc/configs/cart_20260607_CAT_5min.yaml
```

The script never writes inside `raw_dir` / `mask_dir`. Composite PNGs go
to `cache_dir`; the deck goes to `output_pptx`.

## Adding a new dataset

Copy a YAML in `configs/` and edit:

- `raw_dir`, `mask_dir` — source directories
- `raw_pattern`, `mask_pattern` — filename templates with `{fov}`
- `actin_channel` — 0-based channel index for the raw stack
- `fov_ids` — list of FOV strings to include
- `output_pptx`, `cache_dir` — under `K:/FF/PPT/PPT_autogeneration/...`

No Python edits are needed for routine new-dataset runs.

## Mask format

Works with both:
- pre-colored RGB mask TIFFs `(Y, X, 3)` (used directly)
- 2-D integer label TIFFs `(Y, X)` (colored on-the-fly via `tab20`)

## Dependencies

`tifffile`, `numpy`, `Pillow`, `PyYAML`, `python-pptx`, `matplotlib` (only
needed for the 2-D label fallback). All present in the `PPT_editing`
conda env on this machine.
