"""
insert_noco_10262026_montage_slides.py

Montage deck for the 10/26/2026 fixed Jurkat nocodazole experiment
(L:/FF/Nucleus_centrosome/nocodazole_fixed/10262026_jurkats_noco_...), the
4-condition analog of the June CilioD montage decks
(insert_cilio_dmso_nuc_phys_scale_06122026_slides.py and
insert_cilio_actin_synapse_and_xz_combined_slides.py), merged into one deck
and ordered by question (FF 2026-10-27):

  1. Did the drug work?   microtubules first: fixed-contrast MT (placeholders
                          until the pipeline renders them), then the per-cell
                          contrast MT montages
  2. Centrosome position
  3. Nuclear shape
  4. Synapse

This experiment has NO H3K27me3: C1 is beta-tubulin despite the folder name.

Fixed-contrast MT montages (FF 2026-10-27): floor 10 (the montages are drawn
from bg-subtracted images), ceiling = 99.5th percentile of CD3 DMSO MT, the
same for all four conditions. The TCell pipeline will render them into
*_fixedclim folders; until those exist the deck shows a "pending" slide per
block, and the montages fill in on the next build.

Two layouts (--layout):
  grid    all four conditions on one slide, as a 2x2 grid (PLL DMSO | CD3 DMSO
          on top, CD3 Noco 5μM | CD3 Noco 10μM below). One slide per montage
          chunk index (all chunks, so every cell is shown); a condition with
          fewer chunks shows "(none)". The four panels share one PPI
          (PIN_PER_SLIDE), so scalebars match within a slide, not between slides.
  single  one montage sheet per slide, full slide (FF 2026-10-27: "one sheet
          for each condition"). Block -> condition -> chunk order. Every slide
          in a scale group is pinned to one PPI deck-wide, so the whole deck is
          at one physical scale per group.

Every slide carries its channels in their render colors, upper left. Colors
are the pipeline's (TCell-3D-Morphodynamics src/plotting: xz_channel_color.m,
save_*_xz_mip_physical.m, visualize_cytoplasmic_channel_with_nucleus*.m).
The synapse-mask montages carry no scalebar, as in the June actin deck.

PROG_DIR points at the full rerun (prog_fixed_cells; repointed 2026-10-28 from
prog_fixed_cells_superseded, whose 7 PLL cells were cropped to part of the
cell). Strips and packed XZ tiles show only the compile's kept cells
(kept_ids); the pipeline's chunk montages are pre-rendered with every cell, so
the grid/single decks still include cells the compile dropped.

--well <folder> builds a one-well check deck from the new run only
(prog_fixed_cells), one sheet per slide (FF 2026-10-27: 10 µM first).

Usage (base miniforge python, see build-decks skill):
    python examples_and_configs/insert_noco_10262026_montage_slides.py --layout all
    python examples_and_configs/insert_noco_10262026_montage_slides.py --well B2_10uM_aCD3_
"""

import argparse
import io
import os
import re
import sys
from pathlib import Path
from typing import List

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ppt_image_inserter import backup_presentation  # noqa: E402
from insert_cilio_dmso_nuc_phys_scale_06122026_slides import (  # noqa: E402
    BLACK, WHITE, SCALEBAR_PX, SCALEBAR_UM,
    _chunk_start_index, _exists_long, _list_chunk_files, _winlong,
    add_image_at_ppi, add_textbox, compute_slide_ppi, set_slide_background,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXP_ROOT = Path(
    "L:/FF/Nucleus_centrosome/nocodazole_fixed/"
    "10262026_jurkats_noco_h3k27me3640LP35_p549LP45_egfpcentrin2488LP45_h405LP40_"
)
PROG_DIR = "prog_fixed_cells"
# The fixed-contrast MT montages are rendered post hoc from the rerun's outputs,
# into prog_fixed_cells/<well>/physical_scale_images/*_fixedclim (TCell session).
FIXED_CLIM_PROG_DIR = "prog_fixed_cells"

OUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Noco")
OUTPUTS = {
    "grid":   OUT_DIR / "Noco_Jurkats_montages_10262026.pptx",
    "single": OUT_DIR / "Noco_Jurkats_montages_10262026, one condition per slide.pptx",
}

TAG = "10/26/2026"
TP_LABEL = "7 min"          # FF 2026-10-27

# (condition folder, label) in grid order; labels are the compile's labels.
CONDITIONS = [
    ("A4_PLL_DMSO_",  "PLL DMSO"),
    ("A1_DMSO_aCD3_", "CD3 DMSO"),
    ("B1_5uM_aCD3_",  "CD3 Noco 5μM"),
    ("B2_10uM_aCD3_", "CD3 Noco 10μM"),
]

# Channel label colors = the pipeline's render colors (xz_channel_color.m).
ORANGE = RGBColor(0xFF, 0x80, 0x00)
CYAN = RGBColor(0x00, 0xFF, 0xFF)
MAGENTA = RGBColor(0xFF, 0x00, 0xFF)
YELLOW = RGBColor(0xFF, 0xFF, 0x00)
GREEN = RGBColor(0x00, 0xFF, 0x00)
RED = RGBColor(0xFF, 0x00, 0x00)
GRAY = RGBColor(0xBF, 0xBF, 0xBF)

PHYS = "physical_scale_images/{combo}/montages"
SEG = "actin/bottom_slice_seg/montages"

# FF 2026-10-27: XY [10 500], XZ [10 750] on bg-subtracted MT, same for all four
# conditions (above the CD3 DMSO 99.5th percentiles, 415.6 and 658.6); XY lowered
# to [10 200] the same day (params.MT_display_clim in the Noco dataset config), and
# every well's *_fixedclim PNGs were rendered after that (bug hunt 2026-10-28).
FIXED_CLIM_NOTE = "Pending: fixed-contrast MT montages (TCell pipeline)"

MT_NUC = [("MT", MAGENTA), ("DNA", CYAN)]

# Blender nuclear-mesh strips (microscopy-nodes-blender-vis session): per cell a
# 1x4 strip DNA MIP | MT MIP | near-centrosome region | distance to convex hull,
# all tiles one camera, a level side view toward the centrosome (z up), mirrored
# L-R (Figure 1 convention); one px/µm for the whole run. <N> is the pipeline's
# cell id (the same id burned into the montage tiles); _EXCLUDED marks FF's
# excluded cells. Only from the new run, so only in prog_fixed_cells decks.
# Two MT variants per cell: <..>_cell_<N>.png (per-cell MT contrast) and
# <..>_cell_<N>_mtfixed.png (one MT window for every cell).
# {source: {well: (strips dir, filename prefix)}}. "fullcell": the noco-only
# --full-cell framing (FF 2026-10-27: crop by actin bottom/top, no zeroing away
# from the nucleus), 3 cells first for fast iteration.
STRIPS = {
    "nucleus": {"B2_10uM_aCD3_": (
        Path("L:/FF/blender_renderings/_raw/noco1026_strips_10um_16cells_v2_20261027/strips"),
        "noco1026_10um")},
    "fullcell": {"B2_10uM_aCD3_": (
        Path("L:/FF/blender_renderings/_raw/noco1026_strips_10um_3cells_fullcell_v2_20261027/strips"),
        "noco1026_10um")},
}
STRIP_RE = r"^{prefix}_cell_(\d+)(_EXCLUDED)?{suffix}\.png$"
# Hull Surf. Dist. tiles alone (Blender --hull-only), for comparing conditions
# (FF 2026-10-27). All wells pinned to one px/μm by the render (--ppu 56), so
# one PPI deck-wide keeps every cell at the same physical scale. Wells marked
# proto come from the superseded run until tonight's rerun replaces them.
HULL_ROOT = Path("L:/FF/blender_renderings/_raw/noco1026_hull_subset_20261027")
# Fuller runs that replace HULL_ROOT for a prefix once their folder exists
# (all 47 10 μM cells; camera per cell at the view that best shows its
# invagination, any angle: FF 2026-10-27. Replaces noco1026_hull_10um_all_20261027,
# the camera tilted to the centrosome).
HULL_ROOT_OVERRIDE = {
    "noco1026_10um": Path("L:/FF/blender_renderings/_raw/noco1026_hull_10um_pocketscan_20261027"),
    "noco1026_dmso": Path("L:/FF/blender_renderings/_raw/noco1026_hull_dmso_pocketscan_20261027"),
    "noco1026_5um": Path("L:/FF/blender_renderings/_raw/noco1026_hull_5um_pocketscan_20261027"),
    "noco1026_pll": Path("L:/FF/blender_renderings/_raw/noco1026_hull_pll_pocketscan_20261027"),
}
# Tile subfolder under the override root when it is not the prefix itself:
# the larger centrosome dot (radius 1 μm, rim 0.25 μm; FF 2026-10-27, option A).
HULL_SUBDIR = {p: p + "_dot1p0_kept" for p in ("noco1026_10um", "noco1026_dmso", "noco1026_5um")}
HULL = {"A4_PLL_DMSO_": "noco1026_pll", "A1_DMSO_aCD3_": "noco1026_dmso",
        "B1_5uM_aCD3_": "noco1026_5um", "B2_10uM_aCD3_": "noco1026_10um"}
HULL_RE = r"^{prefix}_cell_(\d+)(_EXCLUDED)?_hull\.png$"
HULL_PPU = 56                 # px/μm of every hull tile (the run's --ppu)
# Blocks are ordered by the compile's "Invagination depth near centrosome"
# (FF 2026-10-27), deepest first; with more cells than slots, cells are picked
# at evenly spaced ranks so the block shows the condition's range. Values come
# from the same run as the render (proto wells: the superseded run).
HULL_METRIC = "chull_max_D_by_cent"
RESULTS_MAT = "Jurkats_Noco_MT_20261026_results.mat"
# The final compile's kept cells (kept_ids); condition_name -> well folder.
COMPILE_TRACING = EXP_ROOT / "results_compilation_20261028" / "cell_tracing.csv"
COMPILE_WELLS = {"PLL_DMSO": "A4_PLL_DMSO_", "CD3_DMSO": "A1_DMSO_aCD3_",
                 "CD3_Noco5uM": "B1_5uM_aCD3_", "CD3_Noco10uM": "B2_10uM_aCD3_"}
PACK_SHELVES = 3              # packed decks: rows of cells per condition block
PACK_PER_SHELF = 4            # cells per row: 12 per condition per slide
HULL_PER_SHELF = 5            # hull slides: 15 cells per condition, 3 slides (FF 2026-10-28)
HULL_LABEL_FONT_PT = 18
HULL_FOOT_H = 0.45
HULL_OUT = "Noco_Jurkats_hull_dist_10262026.pptx"
# Segmentation-artifact cells (results.mat cell_id), left off the hull slides
# only (FF 2026-10-28; other montages keep them). Same list as ARTIFACT_CELLS in
# TCell-3D-Morphodynamics dev/analyses/noco_mt_20261026/bycent_vs_global_depth.py.
HULL_ARTIFACT_CELLS = {"A1_DMSO_aCD3_": {37, 7}, "B1_5uM_aCD3_": {35, 22, 15},
                       "B2_10uM_aCD3_": {16, 10, 23, 15, 44}}
HULL_FOOTER = ("Number = invagination depth near centrosome (μm); each nucleus rotated to best "
               "show it. Dot = centrosome. Bars = 5 μm.")
# Blender test, level vs camera tilted to the centrosome (--elev-test).
ELEV_TEST_ROOT = Path("L:/FF/blender_renderings/_raw/noco1026_hull_elev_test_20261027")
ELEV_TEST_PREFIXES = [("noco1026_5um_proto", "5μM"), ("noco1026_10um", "10μM")]
ELEV_TEST_OUT = "Noco_Jurkats_hull_elev_test_10262026.pptx"
# Packed XZ side views (--xz-packed): the per-cell XZ MIPs from each family's
# cells/ folder. Geometry: images are at the pipeline's fixed montage scale
# (5 μm = SCALEBAR_PX px), z = 1 at the bottom row; zstep from the dataset
# config (config/datasets/fixed/Jurkats/Noco/Jurkats_Noco_MT_20261026_dataset.m).
XZ_PPU = SCALEBAR_PX / SCALEBAR_UM
XZ_ZSTEP_UM = 0.3
XZ_BELOW_PLANES = 2           # tiles end this many z planes below the actin bottom slice
# Black figure margin under the image in every physical-scale PNG
# (set_figure_for_physical_scale.m: margin_px 20 at ppum 20 = 1 μm; 20 rows measured).
# Ignoring it put the cut ~3.3 planes lower than XZ_BELOW_PLANES (bug hunt 2026-10-28).
XZ_FIG_MARGIN_UM = 1.0
XZ_ID_BOX = (40, 30)          # (w, h) px of the burned-in cell id, bottom right
# FF 2026-10-27: for XZ MIPs every cell is not needed; one montage type, one
# slide (the first 12 cells per condition).
XZ_PACKED = [
    ("actin_nuc_xz_nolines", "Actin + Nuc XZ MIP", [("Actin", ORANGE), ("DNA", CYAN)]),
]
XZ_PACKED_PAGES = 1
# Comparison deck (--compare), built one slide at a time (FF 2026-10-27; plan:
# docs/noco_10262026_deck_plan.md). Wells come from the new run as they land;
# a well not listed in COMPARE_SOURCES shows a blank block.
COMPARE_SOURCES = {"B2_10uM_aCD3_": "prog_fixed_cells", "A1_DMSO_aCD3_": "prog_fixed_cells",
                   "B1_5uM_aCD3_": "prog_fixed_cells", "A4_PLL_DMSO_": "prog_fixed_cells"}
COMPARE_SLIDES = [
    ("MT_xz_nolines_fixedclim", "Noco effects on MT, side view",
     [("MT contrast 10–750", MAGENTA)]),
    # Noco colors from the 10 μM re-render (FF 2026-10-27): actin yellow, MT and
    # centrin magenta, DNA cyan; MT green only where it meets centrin.
    ("actin_cent_xz_nolines", "Centrosome repositioning, side view",
     [("Actin", YELLOW), ("Centrin", MAGENTA)]),
    ("cent_nuc_xz", "Centrosome and nucleus, side view",
     [("Centrin", MAGENTA), ("DNA", CYAN)]),
    ("actin_nuc_xz_nolines", "Nuclear flattening, side view",
     [("Actin", YELLOW), ("DNA", CYAN)]),
    ("MT_cent_xz_nolines_fixedclim", "MT and centrosome, side view",
     [("MT contrast 10–750", GREEN), ("Centrin", MAGENTA)]),
]
# Comparison families shown with every kept cell, 12 per condition per slide
# (FF 2026-10-27: CD3 DMSO polarization lower than expected, show all cells).
COMPARE_ALL_CELLS = {"actin_cent_xz_nolines"}
# Notes drawn beside a cell's tile (FF 2026-10-27): (well, cell id) -> text.
# A1 18: tall, unspread cell; not cut off (checked on the raw field stack, 06.tif).
TILE_NOTES = {("A1_DMSO_aCD3_", 18): "strange cell shape"}
# Cell-id label on each montage tile (off: FF 2026-10-28, "remove cell #").
TILE_IDS = False
# Top views at the broadest slice (FF 2026-10-27), same packing: (family, family
# with the same geometry to find the crop box in, title, channel legend). Fixed-
# contrast MT is too dim to crop on, so the box comes from the per-cell twin.
COMPARE_XY = [
    ("MT_bz_fixedclim", "MT_nuc_bz", "Noco effects on MT, broadest slice of nucleus",
     [("MT contrast 10–200", MAGENTA)]),
]
# Second MT contrast (FF 2026-10-28: right after the first): fixed-contrast
# family -> (its _fixedclim_low twin, legend), slide 2 of the same deck.
# Windows: params.MT_display_clim_extra in the Noco dataset config.
COMPARE_LOW = {
    "MT_xz_nolines_fixedclim": ("MT_xz_nolines_fixedclim_low", [("MT contrast 10–200", MAGENTA)]),
    "MT_cent_xz_nolines_fixedclim": ("MT_cent_xz_nolines_fixedclim_low",
                                     [("MT contrast 10–200", GREEN), ("Centrin", MAGENTA)]),
    "MT_bz_fixedclim": ("MT_bz_fixedclim_low", [("MT contrast 0–50", MAGENTA)]),
}
COMPARE_OUT = "Noco_Jurkats_compare_{i}_{combo}_10262026.pptx"
# px/μm of actin/bottom_slice_seg PNGs: the pipeline's montage scale, same as the
# XZ MIPs (segment_actin_single_plane.m, ppum 20 at 100 dpi on a 96-dpi session).
SYN_PPU = XZ_PPU
XZ_PACKED_OUT = "Noco_Jurkats_{combo}_packed_10262026.pptx"
# One deck with every comparison slide and the hull slides (--combined), in
# this order of builder tags (the review numbering 1, 1b, 2, 2a, 3, 4, then
# hull beside the nuclear slides, then synapse).
COMBINED_OUT = "Noco_Jurkats_compare_all_10262026.pptx"
COMBINED_ORDER = ["MT_xz_nolines_fixedclim", "MT_bz_fixedclim", "actin_cent_xz_nolines",
                  "MT_cent_xz_nolines_fixedclim", "cent_nuc_xz", "actin_nuc_xz_nolines",
                  "hull", "synapse"]
COLORBAR_DIR = Path("L:/FF/blender_renderings/_raw/noco1026_colorbars")
COLORBAR_DARK = COLORBAR_DIR / "colorbar_chull_dist_invag_depth_0-2um.png"        # black text
COLORBAR_LIGHT = COLORBAR_DIR / "colorbar_chull_dist_invag_depth_0-2um_light.png"  # requested
STRIPS_NOTE = "Pending: Blender nuclear-mesh strips"
# Column headers over the four tiles, Figure 1's names (figure_creation
# CLAUDE.md rule 10: the hull quantity is always "Hull Surf. Dist.").
STRIP_HEADERS = [("DNA", CYAN), ("Microtubules", MAGENTA),
                 ("Near region", YELLOW), ("Hull Surf. Dist.", WHITE)]
# MT window, shown under the Microtubules header (FF: contrast above the MT
# image, not in the title). Blender session 2026-10-27: 10 µM MT peaks at
# ~116-177 counts vs ~619 in DMSO, so fixed-window MT tiles look nearly black.
STRIP_MT_WINDOW = {"_mtfixed": "contrast 10–750", "": "per-cell contrast"}
# Colorbar: only the color ramp is taken from the Blender PNG (inside its
# 1-px frame); title and ticks are drawn here, bold and large (FF 2026-10-27),
# so a label change needs no Blender re-render. Ticks top to bottom; the ramp
# is linear 0-2 μm, clipped at 2.
CBAR_TITLE = "Hull Surf.\nDist. (μm)"
CBAR_TICKS = ["≥2", "1", "0"]
CBAR_FONT_PT = 18
CBAR_RAMP_W = 0.30
STRIP_FOOTER = "Side view toward the centrosome. Magenta dot = centrosome. Bar = 5 μm."

# (section title, [block, ...]); block = (montage subpath under
# prog_fixed_cells/<cond>/, slide title, scale group, [(channel, label color)],
# placeholder note or None). A block with a note shows a "pending" slide while
# its montages don't exist; a block without one is skipped with a warning.
# The MT contrast is named in the channel label above the images (FF: not in
# the title): fixed XY 10–200 / XZ 10–750, or per-cell contrast.
MT_XY_FIXED = ("MT contrast 10–200", MAGENTA)
MT_XZ_FIXED = ("MT contrast 10–750", MAGENTA)
MT_PER_CELL = ("MT per-cell contrast", MAGENTA)
DNA_C = ("DNA", CYAN)
# Per-tile labels: the cell id on every tile, and on multi-panel tiles a name
# over each panel, in the colors the merge panel uses. Keyed by montage family.
TILE_PANELS = {
    "MT_nuc_3p_bz": [("DNA", CYAN), ("MT", MAGENTA), ("Merge", WHITE)],
    "MT_nuc_3p_bz_fixedclim": [("DNA", CYAN), ("MT", MAGENTA), ("Merge", WHITE)],
}
TILE_LABEL_PT = 12
SIDE_NOTE_W = 1.5       # in; text beside a packed montage (build_packed_deck side_notes)
TILE_LABEL_H = 0.26

SECTIONS = [
    ("1. Did the drug work? Microtubules", [
        (PHYS.format(combo="MT_nuc_3p_bz_fixedclim"),
            "DNA | MT | merge, broadest slice", "broad3p", [MT_XY_FIXED, DNA_C], FIXED_CLIM_NOTE),
        (PHYS.format(combo="MT_bz_fixedclim"),
            "MT, broadest slice", "broad", [MT_XY_FIXED], FIXED_CLIM_NOTE),
        (PHYS.format(combo="MT_xz_nolines_fixedclim"),
            "MT XZ MIP", "xz", [MT_XZ_FIXED], FIXED_CLIM_NOTE),
        (PHYS.format(combo="MT_nuc_xz_fixedclim"),
            "MT + Nuc XZ MIP", "xz", [MT_XZ_FIXED, DNA_C], FIXED_CLIM_NOTE),
        (PHYS.format(combo="MT_nuc_3p_bz"),
            "DNA | MT | merge, broadest slice", "broad3p", [MT_PER_CELL, DNA_C], None),
        (PHYS.format(combo="MT_nuc_bz"),
            "MT + Nuc, broadest slice", "broad", [MT_PER_CELL, DNA_C], None),
        (PHYS.format(combo="MT_xz_nolines"),
            "MT XZ MIP", "xz", [MT_PER_CELL], None),
        (PHYS.format(combo="MT_nuc_xz"),
            "MT + Nuc XZ MIP", "xz", [MT_PER_CELL, DNA_C], None),
        (PHYS.format(combo="actin_MT_xz_nolines"),
            "Actin + MT XZ MIP", "xz", [("Actin", MAGENTA), ("MT per-cell contrast", YELLOW)], None),
    ]),
    ("2. Centrosome position", [
        (PHYS.format(combo="cent_nuc_xz"),
            "Cent + Nuc XZ MIP", "xz", [("Centrin", MAGENTA), ("DNA", CYAN)], None),
        (PHYS.format(combo="actin_cent_xz_nolines"),
            "Actin + Cent XZ MIP", "xz", [("Actin", MAGENTA), ("Centrin", YELLOW)], None),
        (PHYS.format(combo="actin_cent_nuc_xz_nolines"),
            "Actin + Cent + Nuc XZ MIP", "xz",
            [("Actin", MAGENTA), ("Centrin", YELLOW), ("DNA", CYAN)], None),
        (PHYS.format(combo="cent_nuc_bz"),
            "Cent + Nuc, broadest slice", "broad", [("Centrin", MAGENTA), ("DNA", CYAN)], None),
    ]),
    ("3. Nuclear shape", [
        (PHYS.format(combo="nucleus_bz"),
            "Nuc (DNA), broadest slice", "broad", [("DNA", GRAY)], None),
        (PHYS.format(combo="actin_nuc_xz_nolines"),
            "Actin + Nuc XZ MIP", "xz", [("Actin", ORANGE), ("DNA", CYAN)], None),
        # subpath "STRIPS<suffix>": Blender strips with that filename suffix
        # subpath "STRIPS:<source>:<suffix>": Blender strips from STRIPS[source]
        # whose filenames carry <suffix> ("_mtfixed" or "")
        ("STRIPS:fullcell:_mtfixed", "Full cell, side view", "strips", [DNA_C, MT_XZ_FIXED], STRIPS_NOTE),
        ("STRIPS:fullcell:", "Full cell, side view", "strips", [DNA_C, MT_PER_CELL], STRIPS_NOTE),
        ("STRIPS:nucleus:_mtfixed", "Nuclear mesh", "strips", [DNA_C, MT_XZ_FIXED], STRIPS_NOTE),
        ("STRIPS:nucleus:", "Nuclear mesh", "strips", [DNA_C, MT_PER_CELL], STRIPS_NOTE),
    ]),
    ("4. Synapse", [
        (SEG, "Actin synapse mask (bottom slice)", "synapse",
            [("Actin", GRAY), ("Mask outline", RED)], None),
    ]),
]
# The legends above are the superseded run's colors. The new run (prog_fixed_cells)
# draws actin yellow, centrin and MT magenta, DNA cyan (params.actin_xz_colors in
# the TCell Noco dataset config, 2026-10-27), so its actin families get these
# legends instead (bug hunt 2026-10-28): family -> channels.
NEW_RUN_DIR = "prog_fixed_cells"
NEW_RUN_LEGENDS = {
    "actin_MT_xz_nolines": [("Actin", YELLOW), ("MT per-cell contrast", MAGENTA)],
    "actin_cent_xz_nolines": [("Actin", YELLOW), ("Centrin", MAGENTA)],
    "actin_cent_nuc_xz_nolines": [("Actin", YELLOW), ("Centrin", MAGENTA), ("DNA", CYAN)],
    "actin_nuc_xz_nolines": [("Actin", YELLOW), ("DNA", CYAN)],
}


def run_legend(combo: str, prog: str, channels: list) -> list:
    """The channel legend for family combo as drawn by run folder prog."""
    return NEW_RUN_LEGENDS.get(combo, channels) if prog == NEW_RUN_DIR else channels


# Slide layout (inches). 13.333 x 7.5 widescreen.
SLIDE_W = 13.333
SLIDE_H = 7.5
MARGIN = 0.10
TITLE_TOP = 0.05
TITLE_HEIGHT = 0.50
TITLE_FONT_PT = 26
CHAN_W = 3.30               # upper-left channel legend; the title sits right of it
CHAN_FONT_PT = 20
SECTION_FONT_PT = 44
NOTE_FONT_PT = 22
GRID_TOP = 0.62
LABEL_H = 0.34
LABEL_FONT_PT = 18

# grid: 2x2 below the title
COL_GAP = 0.15
ROW_GAP = 0.08
CELL_W = (SLIDE_W - 2 * MARGIN - COL_GAP) / 2
CELL_H = (SLIDE_H - GRID_TOP - MARGIN - ROW_GAP) / 2
IMG_H = CELL_H - LABEL_H

# single: one full-width panel below the title
SINGLE_W = SLIDE_W - 2 * MARGIN
SINGLE_IMG_H = SLIDE_H - GRID_TOP - LABEL_H - MARGIN
SINGLE_LABEL_FONT_PT = 20

# strips: STRIPS_PER_SLIDE cells stacked, cell id at left, colorbar at right;
# column headers over the top strip's tiles, one short footer line.
STRIPS_PER_SLIDE = 3
STRIP_HEAD_TOP = 0.62
STRIP_HEAD_FONT_PT = 24
STRIP_HEAD_SUB_FONT_PT = 18
STRIP_TOP = 1.40
STRIP_FOOT_H = 0.40
STRIP_FOOT_FONT_PT = 18
STRIP_LABEL_W = 1.10
STRIP_LABEL_FONT_PT = 22
STRIP_COL_GAP = 0.12          # in, black gap between tile columns
TILE_INSET_PX = 4
CBAR_W = 1.15
CBAR_H = 3.40
STRIP_W = SLIDE_W - 2 * MARGIN - STRIP_LABEL_W - CBAR_W - 0.10
STRIP_ROW_H = (SLIDE_H - STRIP_TOP - STRIP_FOOT_H - MARGIN) / STRIPS_PER_SLIDE

# grid layout only. True: each slide gets its own PPI (the 4 conditions on a
# slide share one scale; scalebar size differs between slides). False: one PPI
# per scale group, deck-wide, as in the June decks; with 4 panels a slide that
# pins every panel to the largest chunk (5132 x 4220 px) and shrinks the rest ~35%.
PIN_PER_SLIDE = True

# ---------------------------------------------------------------------------


def cell_origin(i: int):
    """(left, top) of grid cell i (row-major)."""
    row, col = divmod(i, 2)
    return MARGIN + col * (CELL_W + COL_GAP), GRID_TOP + row * (CELL_H + ROW_GAP)


def chunk_label(p: Path) -> str:
    """'cells 0–15' from montage_cells_0_15.png: the pipeline cell ids, as
    burned into the montage tiles and used for the Blender strips."""
    parts = p.stem.split("_")
    lo, hi = int(parts[2]), int(parts[3])
    return f"cell {lo}" if lo == hi else f"cells {lo}–{hi}"


def montage_slots(montage: Path):
    """Where each cell sits in a physical-scale montage, rebuilt from the
    per-cell PNGs the pipeline composed it from (<family>/cells/), following
    TCell-3D-Morphodynamics create_montages.m (physical_scale_mode): newest
    file per cell id, cells sorted by id, every slot = the batch's max W x H,
    each image centered in its slot, row-major. Returns
    [(cell_id, slot_x, slot_y, x0, y0, w, h)] in montage px (slot corner, then
    the image's corner and size), or None when the per-cell PNGs
    don't reproduce the montage's size (then no tile labels are drawn)."""
    cells_dir = montage.parent.parent / "cells"
    m = re.match(r"montage_cells_(\d+)_(\d+)", montage.stem)
    if not m or not os.path.isdir(_winlong(cells_dir)):
        return None
    lo, hi = int(m.group(1)), int(m.group(2))
    newest = {}
    for name in os.listdir(_winlong(cells_dir)):
        c = re.search(r"[Cc]ell_(\d+)_", name)
        if not c or not name.lower().endswith(".png") or not lo <= int(c.group(1)) <= hi:
            continue
        p = cells_dir / name
        t = os.path.getmtime(_winlong(p))
        if int(c.group(1)) not in newest or t > newest[int(c.group(1))][0]:
            newest[int(c.group(1))] = (t, p)
    if not newest:
        return None
    sizes = []
    for cid in sorted(newest):
        with Image.open(_winlong(newest[cid][1])) as im:
            sizes.append((cid, *im.size))
    max_w = max(w for _, w, _ in sizes)
    max_h = max(h for _, _, h in sizes)
    with Image.open(_winlong(montage)) as im:
        mw, mh = im.size
    n_cols = mw // max_w
    if n_cols < 1 or mw != n_cols * max_w or mh != -(-len(sizes) // n_cols) * max_h:
        return None
    out = []
    for i, (cid, w, h) in enumerate(sizes):
        sx, sy = (i % n_cols) * max_w, (i // n_cols) * max_h
        out.append((cid, sx, sy, sx + (max_w - w) // 2, sy + (max_h - h) // 2, w, h))
    return out


def label_montage_tiles(slide, pic, montage: Path, ppi: float, block: str) -> None:
    """Cell id at each tile slot's top left; on 3-panel tiles (DNA | MT | merge)
    a label centered over each panel at the image's top edge (FF 2026-10-27:
    "add labels"). The id sits at the slot corner, not the image's, so it
    clears the DNA label on narrow cells."""
    slots = montage_slots(montage)
    if slots is None:
        print(f"  (no tile labels: per-cell PNGs don't match {montage.name})")
        return
    left, top = pic.left / 914400, pic.top / 914400       # EMU -> in
    panels = TILE_PANELS.get(Path(block).parent.name)
    for cid, sx, sy, x0, y0, w, h in slots:
        x, y = left + x0 / ppi, top + y0 / ppi
        add_textbox(slide, str(cid), left + sx / ppi, top + sy / ppi, 0.5, TILE_LABEL_H,
                    font_pt=TILE_LABEL_PT, color=WHITE,
                    bold=True).text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        if panels:
            pw = w / len(panels) / ppi
            for k, (name, rgb) in enumerate(panels):
                add_textbox(slide, name, x + k * pw, y, pw, TILE_LABEL_H,
                            font_pt=TILE_LABEL_PT, color=rgb, bold=True)


def add_channel_legend(slide, channels) -> None:
    """Channel names in their render colors, one line, upper left."""
    box = slide.shapes.add_textbox(Inches(MARGIN), Inches(TITLE_TOP),
                                   Inches(CHAN_W), Inches(TITLE_HEIGHT))
    tf = box.text_frame
    tf.margin_left = Inches(0.05)
    tf.word_wrap = True
    para = tf.paragraphs[0]
    for i, (name, rgb) in enumerate(channels):
        run = para.add_run()
        run.text = name if i == 0 else f"  {name}"
        run.font.size = Pt(CHAN_FONT_PT)
        run.font.bold = True
        run.font.color.rgb = rgb


def new_slide(prs, title: str, channels):
    """Black slide with a title; the colored channel legend upper left when
    channels is given (strip slides name their channels in column headers)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, BLACK)
    legend_w = CHAN_W if channels else 0.0
    if channels:
        add_channel_legend(slide, channels)
    # Box symmetric about the slide center, so the title is centered on the slide.
    add_textbox(slide, title, MARGIN + legend_w, TITLE_TOP,
                SLIDE_W - 2 * (MARGIN + legend_w), TITLE_HEIGHT,
                font_pt=TITLE_FONT_PT, color=WHITE, bold=True)
    return slide


def strip_tile_spans(path) -> List[tuple]:
    """(x0, x1) px of each tile in a Blender strip, split at its white gutters."""
    import numpy as np
    with Image.open(_winlong(path)) as im:
        a = np.asarray(im.convert("L"))
    white = (a >= 250).all(axis=0)
    spans, start = [], None
    for x, w in enumerate(white):
        if not w and start is None:
            start = x
        elif w and start is not None:
            spans.append((start, x))
            start = None
    if start is not None:
        spans.append((start, len(white)))
    # Inset past the gutters' anti-aliased edge columns (thin gray lines otherwise).
    return [(x0 + TILE_INSET_PX, x1 - TILE_INSET_PX) for x0, x1 in spans]


def strip_columns(paths) -> tuple:
    """Tile spans of each strip, and each column's width (px) = its widest tile.
    Tiles go into these shared columns so every row lines up under the headers
    (full-cell tiles differ in width from cell to cell)."""
    spans = [strip_tile_spans(p) for p in paths]
    if any(len(s) != len(STRIP_HEADERS) for s in spans):
        return spans, None
    return spans, [max(s[j][1] - s[j][0] for s in spans) for j in range(len(STRIP_HEADERS))]


def strip_layout_ppi(paths) -> float:
    """PPI that fits the column layout of these strips into one strip row."""
    _, cols = strip_columns(paths)
    if cols is None:
        return compute_slide_ppi(paths, STRIP_W, STRIP_ROW_H)
    h = 0
    for p in paths:
        with Image.open(_winlong(p)) as im:
            h = max(h, im.size[1])
    avail_w = STRIP_W - (len(cols) - 1) * STRIP_COL_GAP
    return max(sum(cols) / avail_w, h / STRIP_ROW_H)


def build_section_slide(prs, spec, ppi: float) -> List[str]:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, BLACK)
    add_textbox(slide, spec["title"], MARGIN, SLIDE_H / 2 - 0.5, SLIDE_W - 2 * MARGIN, 1.0,
                font_pt=SECTION_FONT_PT, color=WHITE, bold=True)
    return []


def build_pending_slide(prs, spec, ppi: float) -> List[str]:
    slide = new_slide(prs, spec["title"], spec["channels"])
    box = add_textbox(slide, spec["note"], 1.5, SLIDE_H / 2 - 0.6, SLIDE_W - 3.0, 1.2,
                      font_pt=NOTE_FONT_PT, color=GRAY)
    box.text_frame.word_wrap = True
    return []


def build_grid_slide(prs, spec, ppi: float) -> List[str]:
    slide = new_slide(prs, spec["title"], spec["channels"])
    missing = []
    for i, (label, img) in enumerate(spec["panels"]):
        left, top = cell_origin(i)
        text = f"{label} ({chunk_label(img)})" if img is not None else label
        add_textbox(slide, text, left, top, CELL_W, LABEL_H,
                    font_pt=LABEL_FONT_PT, color=WHITE, bold=True)
        if img is not None and _exists_long(img):
            pic = add_image_at_ppi(slide, img, ppi, left, top + LABEL_H, CELL_W, IMG_H)
            label_montage_tiles(slide, pic, img, ppi, spec["block"])
        else:
            # A pending block (fixed-contrast MT) renders well by well, so an
            # empty cell there means "not rendered yet", not "no more cells".
            add_textbox(slide, "(pending)" if spec.get("note") else "(none)",
                        left, top + LABEL_H + IMG_H / 2 - 0.15,
                        CELL_W, 0.3, font_pt=14, color=WHITE)
            missing.append(label)
    return missing


def build_single_slide(prs, spec, ppi: float) -> List[str]:
    slide = new_slide(prs, spec["title"], spec["channels"])
    label, img = spec["panels"][0]
    add_textbox(slide, f"{label} ({chunk_label(img)})", MARGIN, GRID_TOP,
                SINGLE_W, LABEL_H, font_pt=SINGLE_LABEL_FONT_PT, color=WHITE, bold=True)
    pic = add_image_at_ppi(slide, img, ppi, MARGIN, GRID_TOP + LABEL_H, SINGLE_W, SINGLE_IMG_H)
    label_montage_tiles(slide, pic, img, ppi, spec["block"])
    return []


def build_strips_slide(prs, spec, ppi: float) -> List[str]:
    slide = new_slide(prs, spec["title"], None)
    left = MARGIN + STRIP_LABEL_W
    spans, cols = strip_columns([p for _, _, p in spec["strips"]])
    # Column headers centred over the shared tile columns; the MT window on a
    # second line under "Microtubules".
    if cols is not None:
        col_x = [left + sum(cols[:j]) / ppi + j * STRIP_COL_GAP for j in range(len(cols))]
        for (name, rgb), x, w in zip(STRIP_HEADERS, col_x, cols):
            cx = x + w / ppi / 2
            box = add_textbox(slide, name, cx - 1.3, STRIP_HEAD_TOP, 2.6, 0.42,
                              font_pt=STRIP_HEAD_FONT_PT, color=rgb, bold=True)
            if name == "Microtubules":
                para = box.text_frame.add_paragraph()
                run = para.add_run()
                run.text = spec["mt_window"]
                run.font.size = Pt(STRIP_HEAD_SUB_FONT_PT)
                run.font.color.rgb = rgb
                para.alignment = box.text_frame.paragraphs[0].alignment
    add_textbox(slide, STRIP_FOOTER, MARGIN, SLIDE_H - STRIP_FOOT_H - MARGIN / 2,
                SLIDE_W - 2 * MARGIN, STRIP_FOOT_H,
                font_pt=STRIP_FOOT_FONT_PT, color=GRAY)
    for i, (cell_id, excluded, path) in enumerate(spec["strips"]):
        top = STRIP_TOP + i * STRIP_ROW_H
        add_textbox(slide, f"cell {cell_id}", MARGIN, top + STRIP_ROW_H / 2 - 0.30,
                    STRIP_LABEL_W, 0.40, font_pt=STRIP_LABEL_FONT_PT, color=WHITE, bold=True)
        if excluded:
            add_textbox(slide, "EXCLUDED", MARGIN, top + STRIP_ROW_H / 2 + 0.08,
                        STRIP_LABEL_W, 0.34, font_pt=16, color=RED, bold=True)
        with Image.open(_winlong(path)) as im:
            w_px, h_px = im.size
        img_top = Inches(top + (STRIP_ROW_H - h_px / ppi) / 2)
        if cols is None:
            # Gutters not found: the whole strip, left-aligned.
            slide.shapes.add_picture(_winlong(path), Inches(left), img_top,
                                     width=Inches(w_px / ppi))
            continue
        # Each tile cropped out of the strip and centred in its column.
        for x, w, (x0, x1) in zip(col_x, cols, spans[i]):
            pic = slide.shapes.add_picture(_winlong(path), Inches(x + (w - (x1 - x0)) / ppi / 2),
                                           img_top, width=Inches(w_px / ppi))
            pic.crop_left = x0 / w_px
            pic.crop_right = 1 - x1 / w_px
            pic.width = Inches((x1 - x0) / ppi)
    add_colorbar(slide)
    return []


def colorbar_ramp(path) -> io.BytesIO:
    """The color ramp inside the Blender colorbar's 1-px frame, as a PNG stream.
    Frame = the first opaque run of columns at mid-height, and the rows where
    that run is opaque (the tick labels sit to the right of it). Edges are
    anti-aliased, hence the alpha threshold and the 2-px inset."""
    import numpy as np
    with Image.open(_winlong(path)) as im:
        rgba = im.convert("RGBA")
    alpha = np.asarray(rgba)[..., 3] > 200
    row = alpha[alpha.shape[0] // 2]
    x0 = int(np.argmax(row))
    x1 = x0 + int(np.argmin(row[x0:]))          # one past the frame's right column
    rows = np.nonzero(alpha[:, x0:x1].all(axis=1))[0]
    y0, y1 = int(rows.min()), int(rows.max())
    ramp = rgba.crop((x0 + 2, y0 + 2, x1 - 2, y1 - 1)).convert("RGB")
    buf = io.BytesIO()
    ramp.save(buf, format="PNG")
    buf.seek(0)
    return buf


def add_colorbar(slide) -> None:
    """Hull Surf. Dist. colorbar at the right: Blender's ramp, labels drawn here."""
    cbar = COLORBAR_LIGHT if _exists_long(COLORBAR_LIGHT) else COLORBAR_DARK
    if not _exists_long(cbar):
        return
    cb_left = SLIDE_W - MARGIN - CBAR_W
    title = add_textbox(slide, CBAR_TITLE, cb_left - 0.15, STRIP_TOP, CBAR_W + 0.15, 0.70,
                        font_pt=CBAR_FONT_PT, color=WHITE, bold=True)
    for para in title.text_frame.paragraphs:      # add_textbox formats line 1 only
        para.alignment = title.text_frame.paragraphs[0].alignment
        for run in para.runs:
            run.font.size = Pt(CBAR_FONT_PT)
            run.font.bold = True
            run.font.color.rgb = WHITE
    ramp_top = STRIP_TOP + 0.80
    slide.shapes.add_picture(colorbar_ramp(cbar), Inches(cb_left), Inches(ramp_top),
                             width=Inches(CBAR_RAMP_W), height=Inches(CBAR_H))
    for k, tick in enumerate(CBAR_TICKS):
        y = ramp_top + CBAR_H * k / (len(CBAR_TICKS) - 1)
        add_textbox(slide, tick, cb_left + CBAR_RAMP_W, y - 0.18, CBAR_W - CBAR_RAMP_W, 0.36,
                    font_pt=CBAR_FONT_PT, color=WHITE, bold=True)


def collect_strips(conditions, source: str, suffix: str = "") -> List[tuple]:
    """[(label, [(cell_id, excluded, path), ...])] for wells that have strips;
    only the compile's kept cells (kept_ids)."""
    out = []
    for folder, label in conditions:
        if folder not in STRIPS[source]:
            continue
        d, prefix = STRIPS[source][folder]
        pat = re.compile(STRIP_RE.format(prefix=re.escape(prefix), suffix=re.escape(suffix)))
        keep = kept_ids(folder)
        found = []
        if os.path.isdir(d):
            for name in os.listdir(d):
                m = pat.match(name)
                if m and int(m.group(1)) in keep:
                    found.append((int(m.group(1)), bool(m.group(2)), d / name))
        if found:
            out.append((label, sorted(found)))
    return out


def collect_specs(layout: str, conditions=CONDITIONS, prog_dir: str = PROG_DIR) -> list:
    """Slide specs in deck order. kind: section | pending | montage | strips."""
    specs = []
    # A fixed-contrast slide and its per-cell-contrast twin are paired by
    # chunk index, so both must come from the same run (bug hunt 2026-10-27).
    fixed_twins = {sp.replace("_fixedclim", "") for _, blocks in SECTIONS
                   for sp, *_ in blocks if "_fixedclim" in sp}
    for section, blocks in SECTIONS:
        # Date on the section slides only, so montage titles fit on one line.
        specs.append({"kind": "section", "title": f"{section} ({TAG})", "key": section})
        for subpath, title, group, channels, note in blocks:
            if subpath.startswith("STRIPS"):
                if prog_dir != "prog_fixed_cells":
                    continue          # strips come from the new run only
                _, source, suffix = subpath.split(":")
                wells = collect_strips(conditions, source, suffix)
                if not wells:
                    specs.append({"kind": "pending", "title": title,
                                  "channels": channels, "note": note, "key": f"{subpath} (pending)"})
                for label, cells in wells:
                    pages = [cells[i:i + STRIPS_PER_SLIDE]
                             for i in range(0, len(cells), STRIPS_PER_SLIDE)]
                    for k, page in enumerate(pages):
                        specs.append({"kind": "strips", "channels": channels, "strips": page,
                                      "source": source,
                                      "mt_window": STRIP_MT_WINDOW[suffix],
                                      "title": f"{title} — {label}" + (
                                          f" {k + 1}/{len(pages)}" if len(pages) > 1 else ""),
                                      "key": f"{subpath} {label} #{k + 1}"})
                continue
            prog = (FIXED_CLIM_PROG_DIR if "_fixedclim" in subpath or subpath in fixed_twins
                    else prog_dir)
            channels = run_legend(Path(subpath).parent.name, prog, channels)
            chunks = {}
            for folder, label in conditions:
                d = EXP_ROOT / prog / folder / Path(subpath)
                chunks[label] = sorted(_list_chunk_files(d), key=_chunk_start_index)
            n = max(len(c) for c in chunks.values())
            base = f"{title} ({TP_LABEL})"
            if n == 0:
                if note:
                    specs.append({"kind": "pending", "title": base, "channels": channels,
                                  "note": note, "key": f"{subpath} (pending)"})
                else:
                    print(f"WARNING: no montages for '{subpath}' — block skipped.")
                continue
            if layout == "grid":
                for k in range(n):
                    specs.append({
                        "kind": "montage", "title": f"{base} — {k + 1}/{n}",
                        "channels": channels,
                        "panels": [(label, chunks[label][k] if k < len(chunks[label]) else None)
                                   for _, label in conditions],
                        "group": group, "block": subpath, "key": f"{subpath} #{k + 1}",
                        "note": note,
                    })
            else:
                for _, label in conditions:
                    m = len(chunks[label])
                    for k, img in enumerate(chunks[label]):
                        specs.append({
                            "kind": "montage", "title": f"{base} — {k + 1}/{m}",
                            "channels": channels, "panels": [(label, img)],
                            "group": group, "block": subpath,
                            "key": f"{subpath} {label} #{k + 1}",
                        })
    return pair_fixed_with_auto(specs)


def pair_fixed_with_auto(specs: list) -> list:
    """Put each fixed-contrast MT slide right before its per-cell-contrast twin
    (same view, same cells; the twin's block is the same folder without
    "_fixedclim"), so the drug effect on MT is read slide to slide
    (TCell session 2026-10-27). Twins whose fixed slide is missing or pending
    stay where they are; MT_bz_fixedclim has no per-cell twin."""
    by_key = {s["key"]: s for s in specs}
    twin = {}
    for s in specs:
        if s["kind"] == "montage" and "_fixedclim" in s["key"]:
            t = by_key.get(s["key"].replace("_fixedclim", ""))
            if t is not None:
                twin[s["key"]] = t
    moved = {id(t) for t in twin.values()}
    out = []
    for s in specs:
        if id(s) in moved:
            continue
        out.append(s)
        if s["key"] in twin:
            out.append(twin[s["key"]])
    return out


def pin_ppi(specs: list, layout: str) -> dict:
    """PPI per montage spec key. single: one PPI per scale group, deck-wide.
    grid: one per slide (PIN_PER_SLIDE), floored so a slide with only leftover
    chunks (e.g. PLL's 4th) never renders larger than the block's full
    4-panel slides. Strips: one PPI per Blender run (source), so that run's
    single px/µm carries through and its cells stay to scale; the runs differ
    in px/µm, so one deck-wide PPI would only shrink the smaller run."""
    strip_ppi = {}
    for s in specs:
        if s["kind"] == "strips":
            strip_ppi[s["source"]] = max(strip_ppi.get(s["source"], 0.0),
                                         strip_layout_ppi([p for _, _, p in s["strips"]]))
    out = {s["key"]: strip_ppi[s["source"]] for s in specs if s["kind"] == "strips"}
    out.update(_pin_montage_ppi(specs, layout))
    return out


def _pin_montage_ppi(specs: list, layout: str) -> dict:
    specs = [s for s in specs if s["kind"] == "montage"]
    box_w, box_h = (SINGLE_W, SINGLE_IMG_H) if layout == "single" else (CELL_W, IMG_H)
    for s in specs:
        imgs = [p for _, p in s["panels"] if p is not None and _exists_long(p)]
        s["ppi"] = compute_slide_ppi(imgs, box_w, box_h) if imgs else 0.0
    if layout == "single" or not PIN_PER_SLIDE:
        group = {}
        for s in specs:
            group[s["group"]] = max(group.get(s["group"], 0.0), s["ppi"])
        return {s["key"]: group[s["group"]] for s in specs}
    full_min = {}
    for s in specs:
        if all(p is not None for _, p in s["panels"]):
            full_min[s["block"]] = min(full_min.get(s["block"], float("inf")), s["ppi"])
    return {s["key"]: max(s["ppi"], full_min.get(s["block"], 0.0)) for s in specs}


def build_deck(layout: str, out_path: Path, conditions=CONDITIONS,
               prog_dir: str = PROG_DIR) -> int:
    specs = collect_specs(layout, conditions, prog_dir)
    ppi = pin_ppi(specs, layout)
    montages = [s for s in specs if s["kind"] == "montage"]

    print(f"\n[{layout}] {len(specs)} slides "
          f"({sum(s['kind'] == 'pending' for s in specs)} pending); PPI per scale group:")
    for g in dict.fromkeys(s["group"] for s in montages):
        vals = sorted({round(ppi[s["key"]], 1) for s in montages if s["group"] == g})
        note = ("(no scalebar)" if g == "synapse" else
                f"{SCALEBAR_UM} μm scalebar = "
                + ", ".join(f"{SCALEBAR_PX / v * 2.54:.2f}" for v in vals) + " cm")
        print(f"  {g:>8s}  PPI={vals}  {note}")

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    builders = {"section": build_section_slide, "pending": build_pending_slide,
                "strips": build_strips_slide,
                "montage": build_single_slide if layout == "single" else build_grid_slide}
    missing_total = []
    for s in specs:
        for label in builders[s["kind"]](prs, s, ppi.get(s["key"], 0.0)):
            missing_total.append(f"{s['key']} / {label}")

    if out_path.exists():
        backup_presentation(str(out_path))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print(f"[{layout}] {len(specs)} slides -> {out_path}")
    if missing_total:
        print(f"[{layout}] {len(missing_total)} empty grid cells (conditions with fewer chunks)")
    return len(montages) + sum(s["kind"] == "strips" for s in specs)


def content_crop(im, pad: int = 4):
    """Crop a black-background tile to its non-black content (+ pad px).
    Scale is untouched; only empty margin goes."""
    import numpy as np
    a = np.asarray(im.convert("RGB"))
    ys, xs = np.nonzero(a.max(axis=-1) > 15)
    if len(xs) == 0:
        return im
    return im.crop((max(0, xs.min() - pad), max(0, ys.min() - pad),
                    min(im.width, xs.max() + 1 + pad), min(im.height, ys.max() + 1 + pad)))


def with_scalebar(im, ppu: float, um: float = 5.0):
    """The tile with a strip added below it carrying a white `um` μm bar at
    the left (FF 2026-10-27: a scalebar in every mesh panel). Every tile gets
    the same strip, so bottom alignment is unchanged."""
    from PIL import ImageDraw
    bar = int(round(um * ppu))
    out = Image.new("RGB", (max(im.width, bar + 8), im.height + 26), (0, 0, 0))
    out.paste(im, (0, 0))
    ImageDraw.Draw(out).rectangle([4, im.height + 12, 4 + bar, im.height + 19], fill=(255, 255, 255))
    return out


def angle_tag(elev) -> str:
    """Camera angle for a tile label: "level" or "↑NN°" (camera raised to
    the centrosome's elevation)."""
    return "level" if not elev else f"↑{elev:.0f}°"


def record_elevations(d: Path) -> dict:
    """{cell id: camera elevation (deg)} from a Blender noco_strips_record.json;
    renders made before the field existed were all level (0)."""
    import json
    p = d / "noco_strips_record.json"
    if not os.path.isfile(_winlong(p)):
        return {}
    with open(_winlong(p)) as fh:
        return {int(c["cell"]): c.get("elevation_deg") or 0.0 for c in json.load(fh).get("cells", [])}


def collect_hull() -> dict:
    """{condition label: [(cell_id, excluded, tile image, camera elevation), ...]}:
    the hull-only renders (HULL), then for cells without one, the Hull Surf.
    Dist. tile cut from that well's nuclear-mesh strips (same 56 px/μm, level
    camera). Each tile is cropped to its content."""
    out = {}
    for folder, label in CONDITIONS:
        tiles = {}
        prefix = HULL[folder]
        root = HULL_ROOT_OVERRIDE.get(prefix, HULL_ROOT)
        sub = HULL_SUBDIR.get(prefix, prefix)
        if not os.path.isdir(_winlong(root / sub / "hull")):
            root, sub = HULL_ROOT, prefix
        d = root / sub / "hull"
        elev = record_elevations(root / sub)
        pat = re.compile(HULL_RE.format(prefix=re.escape(prefix)))
        if os.path.isdir(d):
            for name in os.listdir(d):
                m = pat.match(name)
                if m:
                    cid = int(m.group(1))
                    with Image.open(_winlong(d / name)) as im:
                        tiles[cid] = (bool(m.group(2)), im.convert("RGB"), elev.get(cid, 0.0))
        if folder in STRIPS["nucleus"]:
            for cid, excluded, p in next(iter(c for _, c in collect_strips(
                    [(folder, label)], "nucleus", "")), []):
                spans = strip_tile_spans(p)
                if cid in tiles or len(spans) != len(STRIP_HEADERS):
                    continue
                with Image.open(_winlong(p)) as im:
                    x0, x1 = spans[-1]                      # the hull tile
                    tiles[cid] = (excluded, im.convert("RGB").crop((x0, 0, x1, im.height)), 0.0)
        out[label] = [(cid, exc, content_crop(im), e) for cid, (exc, im, e) in sorted(tiles.items())]
    return out


def hull_metric(folder: str) -> dict:
    """{cell_id: HULL_METRIC in μm} from the well's results.mat."""
    import numpy as np
    import scipy.io as sio
    prog = "prog_fixed_cells_superseded" if HULL[folder].endswith("_proto") else "prog_fixed_cells"
    m = sio.loadmat(_winlong(EXP_ROOT / prog / folder / RESULTS_MAT), squeeze_me=True,
                    variable_names=["cell_id", HULL_METRIC])
    ids = np.asarray(m["cell_id"]).ravel().astype(int)
    vals = np.asarray(m[HULL_METRIC], dtype=float).ravel()
    return dict(zip(ids.tolist(), vals.tolist()))


def rank_pick(tiles: list, depth: dict, n: int) -> list:
    """Tiles ordered deepest first (no value: last); more than n -> n picked
    at evenly spaced ranks, keeping the deepest and the shallowest."""
    import math
    key = lambda t: -depth[t[0]] if not math.isnan(depth.get(t[0], math.nan)) else math.inf  # noqa: E731
    tiles = sorted(tiles, key=key)
    if len(tiles) <= n:
        return tiles
    return [tiles[round(i * (len(tiles) - 1) / (n - 1))] for i in range(n)]


def build_packed_deck(out_path: Path, tag: str, title: str, pages: list, headings: dict,
                      ppu: float, footer: str, channels=None, colorbar: bool = False,
                      slide_bar: bool = True, side_notes: list = None,
                      page_channels: list = None, sink: dict = None,
                      per_shelf: int = PACK_PER_SHELF) -> int:
    """Packed montage deck (FF 2026-10-27: 12 cells per condition, large).
    sink: {"prs": Presentation, "tags": []} appends the slides to that deck
    (one tag per slide) instead of saving out_path (--combined).
    pages: one {condition label: [(tile label, tile image), ...]} per slide.
    The four conditions sit 2 x 2 (CONDITIONS order); each block holds
    PACK_SHELVES rows of PACK_PER_SHELF cells, tiles packed edge to edge
    (shelf packing, in the order given) and bottom-aligned on each shelf's
    baseline. One PPI for every tile on every slide, so cells compare at
    physical scale across conditions; one 5 μm bar per slide (ppu = tile px/μm).
    page_channels: one channel legend per slide (default: channels on every slide)."""
    if not pages:
        return 0
    gap = 0.06
    right = (CBAR_W + 0.10) if colorbar else 0.0
    area_w = SLIDE_W - 2 * MARGIN - right
    block_w = (area_w - 0.20) / 2
    block_h = (SLIDE_H - GRID_TOP - HULL_FOOT_H - 0.12) / 2
    shelf_area_h = block_h - LABEL_H

    def shelves(tiles):
        return [tiles[i:i + per_shelf] for i in range(0, len(tiles), per_shelf)]

    def col_widths(sh):
        # Column grid (FF 2026-10-27: rows should align): each column as wide
        # as its widest tile, tiles left-aligned in it, so ids and bars line up.
        return [max(s[j][1].width for s in sh if j < len(s)) for j in range(len(sh[0]))]

    # The largest PPI any block needs (widest grid, tallest stack) sets the scale.
    # A side note takes SIDE_NOTE_W of the block's width.
    ppi = 0.0
    for k, page in enumerate(pages):
        for label, tiles in page.items():
            sh = shelves(tiles)
            if not sh:
                continue
            cw = col_widths(sh)
            note_w = SIDE_NOTE_W if side_notes and side_notes[k].get(label) else 0.0
            widest = sum(cw) / (block_w - note_w - gap * (len(cw) - 1))
            stack = sum(max(im.height for _, im in s) for s in sh)
            ppi = max(ppi, widest, stack / (shelf_area_h - gap * (len(sh) - 1)))
    print(f"\n[{tag}] {len(pages)} slides, PPI {ppi:.1f} (5 μm = {5 * ppu / ppi * 2.54:.2f} cm)")

    if sink is None:
        prs = Presentation()
        prs.slide_width = Inches(SLIDE_W)
        prs.slide_height = Inches(SLIDE_H)
    else:
        prs = sink["prs"]
        sink["tags"] += [tag] * len(pages)
    for k, page in enumerate(pages):
        # title: one string (numbered k/n), or one title per page.
        slide = new_slide(prs, title[k] if isinstance(title, list) else
                          title + (f" {k + 1}/{len(pages)}" if len(pages) > 1 else ""),
                          page_channels[k] if page_channels else channels)
        for b, (_, label) in enumerate(CONDITIONS):
            left, top = MARGIN + (b % 2) * (block_w + 0.20), GRID_TOP + (b // 2) * (block_h + 0.12)
            # headings: one {label: text} for every slide, or one per slide.
            hd = headings[k] if isinstance(headings, list) else headings
            add_textbox(slide, hd.get(label, label), left, top, block_w, LABEL_H,
                        font_pt=HULL_LABEL_FONT_PT, color=WHITE, bold=True)
            sh = shelves(page.get(label, []))
            heights = [max(im.height for _, im in s) / ppi for s in sh]
            y = top + LABEL_H + (shelf_area_h - sum(heights) - gap * (len(sh) - 1)) / 2
            cw = [w / ppi for w in col_widths(sh)] if sh else []
            grid_w = sum(cw) + gap * (len(cw) - 1)
            x0 = left + (block_w - grid_w) / 2
            # side_notes: one {label: text} per slide, set beside the montage
            # (the grid moves left to make room).
            note = (side_notes[k] if side_notes else {}).get(label)
            if note and sh:
                x0 = left + max(0.0, (block_w - grid_w - SIDE_NOTE_W) / 2)
                box = add_textbox(slide, note, x0 + grid_w + 0.15, top + LABEL_H,
                                  SIDE_NOTE_W, shelf_area_h, font_pt=TILE_LABEL_PT,
                                  color=WHITE, bold=True)
                box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                for para in box.text_frame.paragraphs:   # add_textbox styles line 1 only
                    para.alignment = PP_ALIGN.LEFT
                    for run in para.runs:
                        run.font.size, run.font.bold = Pt(TILE_LABEL_PT), True
                        run.font.color.rgb = WHITE
            for s, sh_h in zip(sh, heights):
                for j, (text, im) in enumerate(s):
                    x = x0 + sum(cw[:j]) + gap * j
                    buf = io.BytesIO()
                    im.save(buf, format="PNG")
                    buf.seek(0)
                    slide.shapes.add_picture(buf, Inches(x), Inches(y + sh_h - im.height / ppi),
                                             width=Inches(im.width / ppi))
                    # text: "label", or (left label, right label) both within the tile,
                    # or (left, right, note) with the note beside the tile, in its column.
                    left_t, right_t, note_t = ((tuple(text) + ("", ""))[:3]
                                               if isinstance(text, tuple) else (text, "", ""))
                    if note_t:
                        nb = add_textbox(slide, note_t, x + im.width / ppi, y + sh_h - im.height / ppi,
                                         max(cw[j] - im.width / ppi, 0.6), im.height / ppi,
                                         font_pt=TILE_LABEL_PT, color=WHITE, bold=True)
                        nb.text_frame.word_wrap = True
                        nb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
                        nb.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
                    tw = max(im.width / ppi, 0.6)
                    if left_t:
                        add_textbox(slide, left_t, x - 0.04, y - 0.04, tw + 0.08, TILE_LABEL_H,
                                    font_pt=TILE_LABEL_PT, color=WHITE,
                                    bold=True).text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
                    if right_t:
                        add_textbox(slide, right_t, x - 0.04, y - 0.04, tw + 0.08, TILE_LABEL_H,
                                    font_pt=TILE_LABEL_PT, color=YELLOW,
                                    bold=True).text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
                y += sh_h + gap
        # One 5 μm scalebar for the slide unless the tiles carry their own,
        # then the footer.
        bar_w = 5 * ppu / ppi if slide_bar else 0.0
        y = SLIDE_H - HULL_FOOT_H + 0.17
        if slide_bar:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(MARGIN + 0.1), Inches(y),
                                         Inches(bar_w), Inches(0.05))
            bar.fill.solid()
            bar.fill.fore_color.rgb = WHITE
            bar.line.fill.background()
        # footer: one gray string, or [(text, color), ...] runs.
        if footer:
            runs = [(footer, GRAY)] if isinstance(footer, str) else footer
            box = add_textbox(slide, runs[0][0], MARGIN + 0.2 + bar_w, y - 0.17,
                              SLIDE_W - 2 * MARGIN - 0.3 - bar_w, HULL_FOOT_H - 0.05,
                              font_pt=STRIP_FOOT_FONT_PT, color=runs[0][1])
            para = box.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT
            for text, color in runs[1:]:
                run = para.add_run()
                run.text = text
                run.font.size = Pt(STRIP_FOOT_FONT_PT)
                run.font.color.rgb = color
        if colorbar:
            add_colorbar(slide)
    if sink is not None:
        return len(pages)
    if out_path.exists():
        backup_presentation(str(out_path))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    print(f"[{tag}] {len(pages)} slides -> {out_path}")
    return len(pages)


def build_hull_deck(out_path: Path, sink: dict = None) -> int:
    """Hull Surf. Dist. tiles only, packed: EVERY rendered cell (FF
    2026-10-27: "for the mesh based things I need to see every cell"), each
    condition ordered by HULL_METRIC, deepest first, 12 per condition per
    slide; excluded cells and mesh failures left out, nuclei aligned on their
    bottom. Shown twice (FF 2026-10-27): deepest first, then in cell-id
    order. Each tile carries its own 5 μm bar and its depth; no cell id or
    camera angle (FF: not needed)."""
    import numpy as np
    per_page = PACK_SHELVES * HULL_PER_SHELF
    depth = {f: hull_metric(f) for f, _ in CONDITIONS}
    hull = collect_hull()
    # While views differed between wells the heading named the view (FF
    # 2026-10-27: fine, but say so); only a well still on the old render does now.
    names = {lab: lab + (" (last night's run, side view)" if HULL[f] not in HULL_ROOT_OVERRIDE
                         else "") for f, lab in CONDITIONS}
    pages, titles, notes = [], [], []
    # Cell-id order only (FF 2026-10-27); ("depth", "deepest invagination near
    # centrosome first") restores the ranked pages.
    for order, name in (("id", "cell-id order"),):
        tiles, vals = {}, {}
        for f, lab in CONDITIONS:
            keep = kept_ids(f) - HULL_ARTIFACT_CELLS.get(f, set())
            kept = [c for c in hull[lab] if not c[1] and c[0] in keep]
            if order == "depth":
                kept = rank_pick(kept, depth[f], len(kept))
            vals[lab] = [depth[f].get(cid, float("nan")) for cid, *_ in kept]
            tiles[lab] = [(f"{v:.1f}", with_scalebar(im, HULL_PPU))
                          for v, (_, _, im, _) in zip(vals[lab], kept)]
        n_pages = max(-(-len(t) // per_page) for t in tiles.values())
        pages += [{lab: t[k * per_page:(k + 1) * per_page] for lab, t in tiles.items()}
                  for k in range(n_pages)]
        # Cell-id order is not named in the title (FF 2026-10-27).
        titles += ["Hull Surf. Dist." + (f", {name}" if order != "id" else "")
                   + f" {k + 1}/{n_pages}" for k in range(n_pages)]
        # Mean depth of the well's whole population and of the cells on the
        # slide, beside each montage (FF 2026-10-27).
        for k in range(n_pages):
            nt = {}
            for lab, v in vals.items():
                shown = v[k * per_page:(k + 1) * per_page]
                if shown:
                    nt[lab] = (f"μ={np.nanmean(v):.2f} μm (all)\n"
                               f"μ={np.nanmean(shown):.2f} μm (shown)")
            notes.append(nt)
    return build_packed_deck(out_path, "hull", titles, pages, names, HULL_PPU, HULL_FOOTER,
                             colorbar=True, slide_bar=False, side_notes=notes, sink=sink,
                             per_shelf=HULL_PER_SHELF)


def build_elev_test_deck(out_path: Path) -> int:
    """Level vs tilted hull views for high-centrosome cells (Blender test,
    FF 2026-10-27): one column per cell, level on top, camera tilted to the
    centrosome's elevation below, every tile at one PPI and bottom-aligned."""
    rows = []
    for view in ("level", "cent"):
        tiles = []
        for prefix, label in ELEV_TEST_PREFIXES:
            d = ELEV_TEST_ROOT / view / prefix
            elev = record_elevations(d)
            pat = re.compile(HULL_RE.format(prefix=re.escape(prefix)))
            for name in sorted(os.listdir(d / "hull")):
                m = pat.match(name)
                if m:
                    cid = int(m.group(1))
                    with Image.open(_winlong(d / "hull" / name)) as im:
                        tile = with_scalebar(content_crop(im.convert("RGB")), HULL_PPU)
                    tag = angle_tag(elev.get(cid, 0.0))
                    tiles.append((f"{label} cell {cid} · {tag}", tile))
        rows.append(tiles)
    n_cols = max(len(r) for r in rows)
    col_w = (SLIDE_W - 2 * MARGIN - CBAR_W - 0.10) / n_cols
    row_h = (SLIDE_H - GRID_TOP - HULL_FOOT_H - 0.10) / len(rows)
    ppi = max(max(im.width / (col_w - 0.10), im.height / (row_h - TILE_LABEL_H - 0.05))
              for r in rows for _, im in r)
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    slide = new_slide(prs, "Hull Surf. Dist.: level view vs camera tilted to the centrosome", None)
    for r, tiles in enumerate(rows):
        top = GRID_TOP + r * row_h
        base = top + row_h - 0.05
        for j, (text, im) in enumerate(tiles):
            x = MARGIN + j * col_w
            add_textbox(slide, text, x, top, col_w, TILE_LABEL_H, font_pt=TILE_LABEL_PT,
                        color=WHITE, bold=True)
            buf = io.BytesIO()
            im.save(buf, format="PNG")
            buf.seek(0)
            slide.shapes.add_picture(buf, Inches(x + (col_w - im.width / ppi) / 2),
                                     Inches(base - im.height / ppi), width=Inches(im.width / ppi))
    y = SLIDE_H - HULL_FOOT_H + 0.17
    box = add_textbox(slide, "Top: level side view. Bottom: camera raised to the centrosome's "
                      "elevation (↑). Magenta dot = centrosome. Bars = 5 μm.",
                      MARGIN + 0.1, y - 0.17, SLIDE_W - 2 * MARGIN - 0.2,
                      HULL_FOOT_H - 0.05, font_pt=STRIP_FOOT_FONT_PT, color=GRAY)
    box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    add_colorbar(slide)
    if out_path.exists():
        backup_presentation(str(out_path))
    prs.save(str(out_path))
    print(f"[elev] 1 slide, PPI {ppi:.1f} -> {out_path}")
    return 1


def percell_files(cells_dir: Path) -> dict:
    """{cell id: newest per-cell PNG} in a physical-scale family's cells/
    folder (the files create_montages.m composes the montages from)."""
    newest = {}
    if not os.path.isdir(_winlong(cells_dir)):
        return newest
    for name in os.listdir(_winlong(cells_dir)):
        c = re.search(r"[Cc]ell_(\d+)_", name)
        if not c or not name.lower().endswith(".png"):
            continue
        p = cells_dir / name
        t = os.path.getmtime(_winlong(p))
        if int(c.group(1)) not in newest or t > newest[int(c.group(1))][0]:
            newest[int(c.group(1))] = (t, p)
    return {cid: p for cid, (t, p) in newest.items()}


def xz_tile(path: Path, bottom_slice: float, box_path: Path = None):
    """One XZ MIP cut down to the cell. The image spans the whole stack, z = 1
    at the bottom row (save_xz_mip_physical.m); the cut is at the cell's actin
    bottom slice (the coverslip, converted as original_slice_to_resized.m
    does: (slice - 0.5) * zstep above the stack bottom), so every tile's
    bottom edge is its coverslip. Top and sides are cropped to content; the
    tiny cell id burned in at the bottom right is ignored and blacked out.
    box_path: an image of the same cell and geometry to find the content box
    in (fixed-contrast MT is too dim to crop on; its per-cell-contrast twin
    is used instead)."""
    import numpy as np
    with Image.open(_winlong(path)) as im:
        a = np.asarray(im.convert("RGB")).copy()
    h, w = a.shape[:2]
    a[h - XZ_ID_BOX[1]:, w - XZ_ID_BOX[0]:] = 0
    b = a
    if box_path is not None:
        with Image.open(_winlong(box_path)) as im:
            b = np.asarray(im.convert("RGB")).copy()
        if b.shape != a.shape:
            b = a
        b[h - XZ_ID_BOX[1]:, w - XZ_ID_BOX[0]:] = 0
    # Cut XZ_BELOW_PLANES below the actin bottom slice (FF 2026-10-27: a
    # little margin under the coverslip, not aggressive). The stack bottom sits
    # XZ_FIG_MARGIN_UM above the PNG's bottom row.
    y_cut = min(h, h - int(round((XZ_FIG_MARGIN_UM + (bottom_slice - XZ_BELOW_PLANES - 0.5)
                                  * XZ_ZSTEP_UM) * XZ_PPU)))
    ys, xs = np.nonzero(b[:y_cut].max(axis=-1) > 15)
    if len(xs) == 0:
        return None
    pad = 4
    return Image.fromarray(a[max(0, ys.min() - pad):y_cut,
                             max(0, xs.min() - pad):min(w, xs.max() + 1 + pad)])


def kept_ids(folder: str) -> set:
    """Cell ids the final compile kept in a well (FF 2026-10-27: only kept
    cells anywhere in the deck): local_cell_idx of the compile's
    cell_tracing.csv. FF's manual exclusions, the mesh filters and the centrin
    filter are all applied there."""
    import csv
    with open(_winlong(COMPILE_TRACING), encoding="utf-8-sig", newline="") as fh:
        return {int(float(r["local_cell_idx"])) for r in csv.DictReader(fh)
                if COMPILE_WELLS[r["condition_name"]] == folder}


def xz_tiles(combo: str, sources: dict, kept_only: bool = False) -> dict:
    """{condition label: [(cell id, tile), ...]} for one XZ family, in cell-id
    order. sources maps well folder -> run folder (prog dir); a well missing
    from it gets no tiles (its block stays blank)."""
    import numpy as np
    import scipy.io as sio
    tiles = {}
    for folder, label in CONDITIONS:
        tiles[label] = []
        prog = sources.get(folder)
        if prog is None:
            continue
        m = sio.loadmat(_winlong(EXP_ROOT / prog / folder / RESULTS_MAT), squeeze_me=True,
                        variable_names=["cell_id", "actin_bottom_slice_num"])
        bottoms = dict(zip(np.asarray(m["cell_id"]).ravel().astype(int).tolist(),
                           np.asarray(m["actin_bottom_slice_num"], float).ravel().tolist()))
        keep = kept_ids(folder) if kept_only else None
        phys = EXP_ROOT / prog / folder / "physical_scale_images"
        files = percell_files(phys / combo / "cells")
        # Fixed-contrast families crop on their per-cell-contrast twin (same
        # geometry); MT + centrin has none, so it crops on per-cell MT.
        fixed = combo.removesuffix("_low")
        twin = {"MT_cent_xz_nolines_fixedclim": "MT_xz_nolines"}.get(
            fixed, fixed.replace("_fixedclim", ""))
        twins = percell_files(phys / twin / "cells") if fixed.endswith("_fixedclim") else {}
        for cid in sorted(files):
            if (keep is not None and cid not in keep) or cid not in bottoms or np.isnan(bottoms[cid]):
                continue
            t = xz_tile(files[cid], bottoms[cid], twins.get(cid))
            if t is not None:
                # A 5 μm bar under every tile (FF 2026-10-27); same strip on
                # every tile, so the coverslip baselines stay aligned.
                note = TILE_NOTES.get((folder, cid))
                tiles[label].append(((tile_id(cid), "", note) if note else tile_id(cid),
                                     with_scalebar(t, XZ_PPU)))
    return tiles


def tile_id(cid: int) -> str:
    """A tile's cell-id label, or none (TILE_IDS)."""
    return str(cid) if TILE_IDS else ""


def contrast_in_title(title: str, channels: list):
    """Move an "MT contrast a–b" legend entry into the slide title (FF 2026-10-28):
    the legend then reads plain "MT"."""
    for name, _ in channels:
        if name.startswith("MT contrast"):
            title = f"{title}: {name}"
    return title, [("MT" if name.startswith("MT contrast") else name, color)
                   for name, color in channels]


def build_compare_decks(sink: dict = None) -> int:
    """The comparison deck, one slide type at a time (FF 2026-10-27): each
    COMPARE_SLIDES entry -> its own deck, 12 cells per condition (every cell,
    paged, for COMPARE_ALL_CELLS), excluded cells left out. Wells not yet in
    the new run stay blank."""
    n = 0
    per_page = PACK_SHELVES * PACK_PER_SHELF
    pending = {lab: f"{lab} (tonight's rerun)" for f, lab in CONDITIONS if f not in COMPARE_SOURCES}
    for i, (combo, title, channels) in enumerate(COMPARE_SLIDES, start=1):
        tiles = xz_tiles(combo, COMPARE_SOURCES, kept_only=True)
        n_pages = 1
        if combo in COMPARE_ALL_CELLS:
            n_pages = max(1, max((-(-len(t) // per_page) for t in tiles.values()), default=1))
        pages = [{lab: t[k * per_page:(k + 1) * per_page] for lab, t in tiles.items()}
                 for k in range(n_pages)]
        base = title
        title, channels = contrast_in_title(title, channels)
        titles = [title] if n_pages == 1 else [f"{title} {k + 1}/{n_pages}" for k in range(n_pages)]
        page_channels = [channels] * len(pages)
        if combo in COMPARE_LOW:
            low, low_channels = COMPARE_LOW[combo]
            low_tiles = xz_tiles(low, COMPARE_SOURCES, kept_only=True)
            pages.append({lab: t[:per_page] for lab, t in low_tiles.items()})
            low_title, low_channels = contrast_in_title(base, low_channels)
            titles.append(low_title)
            page_channels.append(low_channels)
        n += build_packed_deck(
            OUT_DIR / COMPARE_OUT.format(i=i, combo=combo), combo, titles, pages, pending,
            XZ_PPU, "", channels=channels,   # no footer (FF 2026-10-28: "remove text")
            slide_bar=False, page_channels=page_channels, sink=sink)
    n += build_synapse_deck(len(COMPARE_SLIDES) + 1, sink)
    for i, (combo, box_combo, title, channels) in enumerate(COMPARE_XY, start=len(COMPARE_SLIDES) + 2):
        families = [(combo, channels)]
        if combo in COMPARE_LOW:
            families.append(COMPARE_LOW[combo])
        pages = []
        for fam, _ in families:
            page = {}
            for folder, label in CONDITIONS:
                page[label] = []
                prog = COMPARE_SOURCES.get(folder)
                if prog is None:
                    continue
                keep = kept_ids(folder)
                phys = EXP_ROOT / prog / folder / "physical_scale_images"
                files, boxes = percell_files(phys / fam / "cells"), percell_files(phys / box_combo / "cells")
                for cid in sorted(c for c in files if c in keep)[:per_page]:
                    page[label].append((tile_id(cid), with_scalebar(xy_tile(files[cid], boxes.get(cid)), XZ_PPU)))
            pages.append(page)
        moved = [contrast_in_title(title, ch) for _, ch in families]
        n += build_packed_deck(
            OUT_DIR / COMPARE_OUT.format(i=i, combo=combo), combo, [t for t, _ in moved], pages,
            pending, XZ_PPU, "",   # no footer (FF 2026-10-28)
            channels=moved[0][1], slide_bar=False, page_channels=[ch for _, ch in moved],
            sink=sink)
    return n


def build_combined_deck(out_path: Path) -> int:
    """Every comparison slide plus the Hull Surf. Dist. slides in one deck
    (FF 2026-10-28), in COMBINED_ORDER; the per-slide decks stay as they are."""
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    sink = {"prs": prs, "tags": []}
    n = build_compare_decks(sink) + build_hull_deck(OUT_DIR / HULL_OUT, sink)
    # Reorder (stable within a family): slides were added in builder order.
    lst = prs.slides._sldIdLst
    els = list(lst)
    for el in els:
        lst.remove(el)
    for k in sorted(range(len(els)), key=lambda k: COMBINED_ORDER.index(sink["tags"][k])):
        lst.append(els[k])
    if out_path.exists():
        backup_presentation(str(out_path))
    prs.save(str(out_path))
    print(f"[combined] {n} slides -> {out_path}")
    return n


def xy_tile(path: Path, box_path: Path = None):
    """A top-view tile: the burned-in cell id blanked, cropped to the content
    of box_path (same geometry; per-cell contrast) or of the image itself."""
    import numpy as np
    with Image.open(_winlong(path)) as im:
        a = np.asarray(im.convert("RGB")).copy()
    h, w = a.shape[:2]
    a[h - XZ_ID_BOX[1]:, w - XZ_ID_BOX[0]:] = 0
    b = a
    if box_path is not None:
        with Image.open(_winlong(box_path)) as im:
            b = np.asarray(im.convert("RGB")).copy()
        if b.shape != a.shape:
            b = a
        b[h - XZ_ID_BOX[1]:, w - XZ_ID_BOX[0]:] = 0
    ys, xs = np.nonzero(b.max(axis=-1) > 15)
    if len(xs) == 0:
        return Image.fromarray(a)
    pad = 4
    return Image.fromarray(a[max(0, ys.min() - pad):min(h, ys.max() + 1 + pad),
                             max(0, xs.min() - pad):min(w, xs.max() + 1 + pad)])


def build_synapse_deck(i: int, sink: dict = None) -> int:
    """Comparison slide i: the actin synapse mask (the pipeline's bottom-slice
    actin with the mask outline, actin/bottom_slice_seg/Cell_<id>_plane_<n>.png),
    packed like the XZ slides: cell-id order, excluded cells left out, each
    tile cropped to content. A 5 μm bar per tile once SYN_PPU is known."""
    per_page = PACK_SHELVES * PACK_PER_SHELF
    pending = {lab: f"{lab} (tonight's rerun)" for f, lab in CONDITIONS if f not in COMPARE_SOURCES}
    page = {}
    for folder, label in CONDITIONS:
        page[label] = []
        prog = COMPARE_SOURCES.get(folder)
        if prog is None:
            continue
        keep = kept_ids(folder)
        files = percell_files(EXP_ROOT / prog / folder / "actin" / "bottom_slice_seg")
        for cid in sorted(c for c in files if c in keep)[:per_page]:
            with Image.open(_winlong(files[cid])) as im:
                t = content_crop(im.convert("RGB"))
            page[label].append((tile_id(cid), with_scalebar(t, SYN_PPU) if SYN_PPU else t))
    bars = "Bars = 5 μm." if SYN_PPU else "Scale bar pending."
    return build_packed_deck(
        OUT_DIR / COMPARE_OUT.format(i=i, combo="synapse_mask"), "synapse", "Synapse, actin bottom slice",
        # "Red = ..." drawn in red (FF 2026-10-28).
        [page], pending, SYN_PPU or 1.0, [("Red = synapse mask outline. ", RED), (bars, GRAY)],
        channels=[("Actin", GRAY), ("Mask outline", RED)], slide_bar=False, sink=sink)


def build_xz_packed_deck(prog_dir: str = PROG_DIR) -> int:
    """XZ side views, packed (FF 2026-10-27), one slide set per XZ_PACKED
    family: 12 cells per condition per slide in cell-id order, every cell cut
    at its coverslip and all cells aligned on it. Only the compile's kept
    cells (kept_ids)."""
    n = 0
    per_page = PACK_SHELVES * PACK_PER_SHELF
    for combo, title, channels in XZ_PACKED:
        tiles = xz_tiles(combo, {f: prog_dir for f, _ in CONDITIONS}, kept_only=True)
        n_pages = min(XZ_PACKED_PAGES, max(-(-len(t) // per_page) for t in tiles.values()))
        pages = [{lab: t[k * per_page:(k + 1) * per_page] for lab, t in tiles.items()}
                 for k in range(n_pages)]
        n += build_packed_deck(
            OUT_DIR / XZ_PACKED_OUT.format(combo=combo), combo, f"{title} ({TP_LABEL})", pages, {},
            XZ_PPU, "Each cell cut at its coverslip (actin bottom slice); all cells aligned on it. "
            "Bars = 5 μm.", channels=run_legend(combo, prog_dir, channels), slide_bar=False)
    return n


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("--layout", choices=["grid", "single", "all"], default="all")
    ap.add_argument("--well", choices=[f for f, _ in CONDITIONS],
                    help="one-well check deck (e.g. B2_10uM_aCD3_): one sheet per slide, "
                         "every montage read from prog_fixed_cells (the new run), "
                         "nothing from PROG_DIR")
    ap.add_argument("--hull", action="store_true",
                    help="Hull Surf. Dist. tiles only, all conditions, one scale")
    ap.add_argument("--xz-packed", action="store_true",
                    help="XZ side views packed 12 per condition, aligned on the coverslip")
    ap.add_argument("--compare", action="store_true",
                    help="comparison deck slides (COMPARE_SLIDES), one deck per slide")
    ap.add_argument("--elev-test", action="store_true",
                    help="hull tiles, level vs camera tilted to the centrosome (Blender test)")
    ap.add_argument("--combined", action="store_true",
                    help="every comparison slide plus the hull slides in one deck")
    args = ap.parse_args()
    if args.combined:
        n = build_combined_deck(OUT_DIR / COMBINED_OUT)
    elif args.hull:
        n = build_hull_deck(OUT_DIR / HULL_OUT)
    elif args.xz_packed:
        n = build_xz_packed_deck()
    elif args.compare:
        n = build_compare_decks()
    elif args.elev_test:
        n = build_elev_test_deck(OUT_DIR / ELEV_TEST_OUT)
    elif args.well:
        # FF 2026-10-27: a 10 µM-only deck from the new run, to iron out every
        # montage before the 4-well run.
        conds = [(f, lab) for f, lab in CONDITIONS if f == args.well]
        label = conds[0][1].replace("μ", "u").replace(" ", "_")
        out = OUT_DIR / f"Noco_Jurkats_montages_10262026, {label} only, new run.pptx"
        n = build_deck("single", out, conds, "prog_fixed_cells")
    else:
        layouts = ["grid", "single"] if args.layout == "all" else [args.layout]
        n = sum(build_deck(lay, OUTPUTS[lay]) for lay in layouts)
    if n == 0:
        sys.exit("ERROR: no montage slides written.")


if __name__ == "__main__":
    main()
