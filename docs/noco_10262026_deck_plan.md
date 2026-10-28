# Noco 10/26/2026 comparison deck: plan

Last updated 2026-10-27. Builder: `examples_and_configs/insert_noco_10262026_montage_slides.py`.
Review page: https://claude.ai/artifact/UcXoqApk8cvbPsxxJvnyET

## Goal and rules

- One packed slide type per question. The four conditions sit 2 × 2 (PLL DMSO, CD3 DMSO /
  CD3 Noco 5 μM, CD3 Noco 10 μM), all at one scale, 12 cells per condition per slide.
- Build one slide at a time. Start with CD3 Noco 10 μM (the only well in the new run);
  leave the other blocks blank until tonight's rerun.
- Every tile cropped to its content and packed edge to edge; cell id on every tile.
- Mesh-based slides show EVERY cell. Image-based slides (XZ MIPs etc.) need not; one
  slide per montage type is enough.
- Per-cell strips (DNA / MT / region / hull) stay in per-condition QC decks.

## Slides, in deck order

| # | Question | Slide | Status |
|---|---|---|---|
| 1 | Did the drug work? | MT XZ side view, fixed contrast 10–750, aligned on the coverslip (`--compare`, 10 μM only; first 12 non-excluded cells by id) | review |
| 2 | Centrosome position | Actin + centrin XZ side view (`actin_cent_xz_nolines`), first 12 non-excluded cells by id | review |
| 1b | Did the drug work? | MT broadest slice, fixed contrast 10–500 (`MT_bz_fixedclim`, `COMPARE_XY`); crop box from `MT_nuc_bz`. Too dim at 10–500; FF asked for 10–200 on the next pipeline run. Next: MT + DNA | review |
| 2a | Centrosome position | MT (fixed contrast 10–750, green) + centrin (magenta) XZ side view (`MT_cent_xz_nolines_fixedclim`, pipeline re-render 2026-10-27); crop box from `MT_xz_nolines` | review |
| 2b | Centrosome position | Centrin + DNA XZ side view (`cent_nuc_xz`), shown separately; centrin faint under DNA | review |
| 3 | Nuclear shape | Hull Surf. Dist. meshes, every cell, deepest invagination near centrosome first | review (all 47 10 μM cells, 4 slides) |
| 4 | Nuclear shape | Actin + DNA XZ side view (`actin_nuc_xz_nolines`, `--compare`, 10 μM, new run) | review |
| 5 | Synapse | Actin bottom-slice mask (XY, `actin/bottom_slice_seg`), cell-id order; scale bar pending (px/μm asked of the pipeline session) | review |

## Decisions so far

- Hull camera (FF 2026-10-27, replaces the tilt-to-centrosome rule): each cell gets the
  view that best shows its invagination, any angle (`--angle pocket --pocket-scan -90 90 15`).
  Labels show the camera above (↑NN°) or below (↓NN°) level. 10 μM render:
  `noco1026_hull_10um_pocketscan_20261027`; other wells planned with the same rule.
- Hull blocks ordered by `chull_max_D_by_cent` (compile: "Invagination depth near
  centrosome"), deepest first. Labels: cell · depth (μm) · camera angle.
- Every mesh tile carries its own 5 μm bar; nuclei aligned on their bottom.
- XZ tiles, kept simple (FF 2026-10-27): bottom cut 2 planes below the actin bottom slice
  (`XZ_BELOW_PLANES`), all cells aligned on it; top and sides trimmed to any pixel > 15
  plus 4 px. No noise rules; it need not be perfect.
- Only cells the final compile kept appear on the hull and comparison slides (FF 2026-10-27:
  showing a dropped cell is a major error): `kept_ids()` reads
  `results_compilation_20261027\cell_tracing.csv` (152 cells: 42/39/31/40). It replaced a local
  re-derivation (manual list + mesh genus + centrin ≤ 50) that gave the same 152 cells.
- CD3 DMSO cell 18: tall, unspread, not cut off (checked on the raw field stack); its tile
  reads "strange cell shape" (`TILE_NOTES`).
- Hull dot: radius 1 μm, black rim 0.25 μm (FF option A, 2026-10-27; `HULL_SUBDIR`, *_dot1p0).
  Hull deck in cell-id order only; block headings name the view (best view vs. side view).
- Fixed-contrast MT tiles are cropped on their per-cell-contrast twin (same geometry); the
  fixed-contrast image alone is too dim to crop on.

- Channel colors (FF 2026-10-27): actin yellow, MT magenta, centrin magenta (as in the
  paper), DNA cyan. MT and centrin never share an image. The pipeline PNGs still show actin
  magenta in other wells until tonight's run; 10 μM was re-rendered with the new colors and
  XY MT fixed contrast 10–200 (old images in `B2_10uM_aCD3_\superseded_display_20261027`).

## Tonight (FF 2026-10-27)

- MT at two contrast levels: keep 10–750 (shows how much lower MT signal is in Noco) and add a
  second slide at 10–200 (or 10–250) so Noco MT are visible; DMSO saturating there is fine. Needs a
  pipeline render of the XZ MIP family at the lower contrast (ask the pipeline session).
- Single-plane MT (broadest slice): second slide at fixed contrast 0–50, next to 10–200
  (FF 2026-10-27).

## Flagged cells (kept, FF 2026-10-27)

- CD3 DMSO (A1) cell 21: bilobed nucleus, two lobes joined by a DNA bridge at z67 (separate
  at the broadest slice, z74); one centrosome at the cleft. Hull depth 2.9 μm is that cleft.
  FF: keep for now, flagged.
- Full list: `config.flagged_cells` in the pipeline's Jurkats_Noco_MT_20261026_compilation.m
  (A1 21 plus FF's 11 QC-page flags, which have no notes; still on the slides: A1 24,
  B1 19/22, B2 23/44, A4 13/15/20).

## Open questions for FF

- Ordering metric for the MT slide and the centrosome slide (proposed: centrosome-to-synapse
  distance for the centrosome slide).
- Ordering for the synapse slide (proposed: synapse area).
- Replace the montage slides in the 4-well deck with these packed slides?

## Data and runs

- New run: `L:\FF\Nucleus_centrosome\nocodazole_fixed\10262026_...\prog_fixed_cells\`. B2_10uM_aCD3_,
  A1_DMSO_aCD3_ (excluded 12, 25, 39, 41) and B1_5uM_aCD3_ (excluded 6, 27) done and on the
  comparison slides; A4 PLL running (≈17:35).
- Last night's run (other wells, prototype only): `prog_fixed_cells_superseded`.
- Per-cell images for packing: `physical_scale_images\<family>\cells\`, 5 μm = 104 px.
- Hull renders: `L:\FF\blender_renderings\_raw\noco1026_hull_subset_20261027` (cells 0–3 per
  well, level), `noco1026_hull_10um_all_20261027` (all 47 10 μM cells, tilted; running).
- Full hull run for all wells: on hold until tonight's rerun (FF 2026-10-27).
