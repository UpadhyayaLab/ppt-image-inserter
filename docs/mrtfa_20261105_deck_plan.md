# MRTF-A on gels (CTLs, fixed): deck plan

Builder: `examples_and_configs/build_mrtfa_gels_deck.py` (reads the compilation config; `--smoke`
= Nov 3, every cell, montage slides only, smoke 2026-11-04 in
`J:\FF\fixed_cell\CTL_nucleus\MRTF-A_Hoechst\tests\mrtfa_deck_smoke\`, 8.5 s).

Draft 2026-11-04, from the brief of the TCell-3D-Morphodynamics "MRTF dataset configs" session.
Model: `noco_10262026_deck_plan.md`, `build_noco_two_experiment_montages.py`,
`build_noco_shared_deck_two_experiments.py`. Layout is FF's call; open questions at the end.

## Order of work

1. Full compile, then build the deck from it (preliminary cell set).
2. FF's QC exclusions, then recompile.
3. Rebuild the deck from the config: no cell list kept in the builder; kept cells come from the
   compile (`seg_review_cells.csv`, `kept_by_compile`), as `kept_ids()` does for Noco.

## Data

- Experiments, in order, never pooled (one page or panel per experiment): Nov 3 2023, Feb 5 2024,
  Feb 22 2024.
- Gels, in order: 1.5 kPa (`Gel3_1p5kPa`, well `Gel3_3p0_...`), 6 kPa (`Gel2_6kPa`, `Gel2_5p0_...`),
  12 kPa (`Gel1_12kPa`, `Gel1_8p0_...`). Colors and labels from
  `TCell-3D-Morphodynamics\config\compilation\fixed\CTL\MRTF\CTL_Gels_MRTFA_compilation.m`.
- Channels: MRTF-A (MKL1), actin, Hoechst. No centrosome stain.
- Per-cell tiles: `<well>\bg_sub_fovs\cells\individual_channels\prog_fixed_cells\physical_scale_images\<family>\cells\`
  (Nov 3 already reprocessed, 2026-11-04). Pipeline colors: MRTF-A red, DNA cyan, actin orange.
- Compile: `latest_output_dir` on `J:\FF\fixed_cell\CTL_nucleus\MRTF-A_Hoechst\compiled_results\CTL_Gels_MRTFA`;
  metric panels `<compile>\grid_panels\<field>_per_experiment_grid.png` (3 experiments side by side).
- Ready when `TCell-3D-Morphodynamics\dev\runs\mrtfa_report_20261104.status` reads
  `step end chain | done`.
- Output (FF 2026-11-04): `K:\FF\PPT\PPT_autogeneration\Transcription_factors\`, beside the NFAT / NFκB decks.

## What carries over from Noco

- Packed montage slides, one per question and per experiment, all tiles at one scale, cropped to
  content, cell id on every tile, kept cells only.
- Image slides: 12 kept cells per condition per page. Mesh slides: every cell.
- XZ tiles aligned on the coverslip (bottom cut 2 planes below the actin bottom slice).
- Results slides: one per panel group, plots from the config's titles, so a recompiled panel
  lands on the next build.

## What drops (no centrosome stain)

- Centrin side views (`actin_cent_xz`, `cent_nuc_xz`, `MT_cent_xz`), MT slides.
- Centrosome dot on hull meshes; ordering by `chull_max_D_by_cent`; centrosome-proximal region
  slides; nucleus-centrosome distance and polarization.

## Draft slides, in deck order

| # | Question | Slide | Source |
|---|---|---|---|
| 1 | Where is MRTF-A? | MRTF-A + DNA, broadest slice | `MRTFA_nuc_bz` |
| 1b | Where is MRTF-A? | MRTF-A / DNA / merge split, broadest slice (3 panels per cell, wide) | `MRTFA_nuc_3p_bz` |
| 2 | Nuclear shape | Actin + DNA XZ side view | `actin_nuc_xz_nolines` |
| 2b | Where is MRTF-A, in z? | MRTF-A + DNA XZ side view (replaces the centrin side view) | not rendered: needs a pipeline family |
| 3 | Nuclear shape | Hull Surf. Dist. meshes, every cell, deepest invagination first (`chull_max_D`) | not rendered: needs a Blender run from `mrtfa_mesh_qc` |
| 4 | Synapse | Actin bottom-slice mask (XY) | `actin\bottom_slice_seg` |
| 5–9 | Results | Metric panels, one slide per group of `config.report_extra`: MRTF-A and DNA near the deepest invagination; MRTF-A nucleus vs cytoplasm; invagination size; nuclear shape; synapse | `grid_panels` |

Nucleus QC sheets (`seg_review\nucleus_slices`) and the mesh QC export stay out of the deck:
they are for FF's exclusion pass.

## Decisions (FF 2026-11-04)

- Gel blocks: three columns, 1.5 | 6 | 12 kPa left to right.
- Image slides: first 12 kept cells by id per gel (Noco rule).
- Cell id on every tile, for now (`SHOW_IDS`).
- MRTF-A + DNA side view (2b): not requested; dropped.
- Hull meshes (3): maybe, not essential, must not hold up the deck. Build without them; add the
  slide later if a render exists.

## Tile z planes (MRTF dataset configs session, from pipeline source, 2026-11-04)

- `MRTFA_nuc`, `MRTFA_nuc_3p`, `nucleus`: the XY slice that best shows the deepest invagination
  (most deepest-invagination boundary voxels; `visualize_nucleus_slice.m`).
- `*_bz`: the broadest nuclear slice (largest mask area).
- The z in the filename indexes the resized isotropic stack, not the raw planes.

- MRTF-A + DNA at the deepest-invagination slice (`MRTFA_nuc`): yes (FF 2026-11-04), after the
  broadest-slice slide.
