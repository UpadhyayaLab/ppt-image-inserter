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

A composite is redrawn when its raw image or mask has a different file
time than when it was drawn, older or newer, and every FOV is redrawn when
`raw_dir`, `mask_dir` or a display setting in the config changes (the last
complete pass's settings and file times are kept in
`cache_dir/cache_settings.json`). A cache with no such file, or one written
before file times were recorded, is redrawn once. An existing deck is backed
up before it is replaced.

Each pass also writes `cache_dir/current_composites.json`, the composites it
drew or confirmed. `combine_with_nuc_decks.py` and `contact_sheet.py` show only
those and print what they left out, so a FOV that lost its mask or left
`fov_ids` does not come back with its old composite. A cache from before this
list shows everything, with a warning.

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
