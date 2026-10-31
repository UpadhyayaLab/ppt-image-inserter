#!/usr/bin/env python
"""
Comparison montage decks for the 10/29/2026 H3K27me3 sets (Noco, Ciliobrevin), in
the 10/26 Noco layout (FF 2026-10-31: "format nicely like in 1026 expt"): black
slides, the conditions 2 x 2, 12 of the compile's kept cells per condition per
slide in cell-id order, side views cut at each cell's coverslip, one physical
scale per slide type.

Reuses insert_noco_10262026_montage_slides.py (xz_tiles, xy_tile, build_packed_deck,
kept_ids); only the experiment settings in EXPS differ. Every well reads
prog_fixed_cells (tile restore finished, TCell session, 2026-10-31).

"compile" in EXPS sets the kept cells (its cell_tracing.csv). A recompile under a new
folder name: change "compile" here and base_dir in the two results-deck configs, then
rebuild; otherwise the deck keeps the old kept-cell set. The path prints at build time.
--kept-side-views takes the same cells from the side-view CSVs instead, for when the compile
is being rebuilt and must not be read.

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_compare_decks.py [noco|cilio|all] [--kept-side-views]
"""

import csv
import os
import sys
from pathlib import Path

import numpy as np
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parent))
import insert_noco_10262026_montage_slides as m  # noqa: E402
from insert_noco_10262026_montage_slides import CYAN, MAGENTA, ORANGE, RED, YELLOW  # noqa: E402

PROG ="prog_fixed_cells"
OUT_ROOT = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous"

# wells: (well folder, label, compile condition_name, per-cell PNG source).
# xz: (family, title, channel legend, every kept cell?). Colors are the pipeline's:
# Noco sets params.actin_xz_colors (actin yellow, centrin magenta, DNA cyan) in its
# TCell dataset config; Ciliobrevin keeps the defaults (save_actin_cent_xz_mip_physical:
# actin magenta + centrin yellow; save_nuc_actin_xz_mip_physical: actin orange;
# save_cent_nuc_xz_mip_physical: centrin magenta, DNA cyan). Order as the 10/26
# combined deck, MT and hull slides dropped (no MT stain, no Blender renders yet).
EXPS = {
    "noco": {
        "root": "L:/FF/Nucleus_centrosome/nocodazole_fixed/"
                "10292026_jurkats_noco_h3k27me3640LP45_p549LP45_egfpcentrin2488LP45_h405LP40_",
        "results": "Jurkats_Noco_H3K27me3_20261029_results.mat",
        "compile": "results_compilation_20261031",
        "wells": [("G1A4_h3k27me3_PLL_dmso_", "PLL DMSO", "PLL_DMSO", PROG),
                  ("G1A1_h3k27me3_aCD3_dmso_", "CD3 DMSO", "CD3_DMSO", PROG),
                  ("G1B1_h3k27me3_aCD3_5uM_", "CD3 Noco 5μM", "CD3_Noco5uM", PROG),
                  ("G1B2_h3k27me3_aCD3_10uM_", "CD3 Noco 10μM", "CD3_Noco10uM", PROG)],
        "xz": [("actin_cent_xz_nolines", "Centrosome repositioning, side view",
                [("Actin", YELLOW), ("Centrin", MAGENTA)], True),
               ("cent_nuc_xz", "Centrosome and nucleus, side view",
                [("Centrin", MAGENTA), ("DNA", CYAN)], True),
               ("actin_nuc_xz_nolines", "Nuclear flattening, side view",
                [("Actin", YELLOW), ("DNA", CYAN)], True)],
        "out": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Noco_Jurkats_H3K27me3_compare_10292026.pptx",
        "footer": "PRELIMINARY: no figure review yet",
    },
    "cilio": {
        "root": "L:/FF/Nucleus_centrosome/ciliobrevin_fixed/"
                "10292026_jurkats_ciliobrevin_h3k27me3640LP45_p549LP45_egfpcentrin2488LP45_h405LP40_",
        "results": "Jurkats_Ciliobrevin_H3K27me3_20261029_results.mat",
        "compile": "results_compilation_20261031",
        "wells": [("G0A4_h3k27me3_PLL_dmso_", "PLL DMSO", "PLL_DMSO", PROG),
                  ("G0A1_h3k27me3_aCD3_dmso_001", "CD3 DMSO", "CD3_DMSO", PROG),
                  ("G0B2_h3k27me3_aCD3_cilio_", "CD3 Cilio 50μM", "CD3_Cilio", PROG)],
        "xz": [("actin_cent_xz_nolines", "Centrosome repositioning, side view",
                [("Actin", MAGENTA), ("Centrin", YELLOW)], True),
               ("cent_nuc_xz", "Centrosome and nucleus, side view",
                [("Centrin", MAGENTA), ("DNA", CYAN)], True),
               ("actin_nuc_xz_nolines", "Nuclear flattening, side view",
                [("Actin", ORANGE), ("DNA", CYAN)], True)],
        "out": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/CilioD_Jurkats_H3K27me3_compare_10292026.pptx",
        "footer": "PRELIMINARY: no figure review yet",
    },
}
# Top views, both sets; H3K27me3 last, as in the results decks (FF 2026-10-31).
XY_SLIDES = [
    ("H3K27me3_nuc_bz", "H3K27me3, broadest slice of nucleus",
     [("H3K27me3", RED), ("DNA", CYAN)], True),
]
# Burned-in cell id on the top-view tiles: ~35 x 40 px at the bottom right
# (measured on Noco G1B1 Cell_0_z139.png), larger than the side views' XZ_ID_BOX.
XY_ID_BOX = (60, 50)
XY_BLOB_LEVEL = 15    # top-view crop follows the nucleus (largest blob above this), not stray specks
XZ_MIN_BLOB = 40      # side-view crop ignores specks under this many px (centrin spots are larger)
# Side-view crop also ignores dim blobs below this level, for these cells only: a dim 200 px
# speck pushed this nucleus off-center. As a general rule it clips the dim tops of real cells.
XZ_MIN_PEAK = {("G1A1_h3k27me3_aCD3_dmso_", 9): 64}
# Every view is scaled per cell, so brightness is not comparable between cells: top views
# 1st-99th percentile of the nucleus (visualize_struct_in_nuc_channel.m), side views each
# cell's own ceiling (xz_channel_clim.m). Footer on every slide, in the 10/26 wording.
CONTRAST_NOTE = "Per-cell contrast."
SIDE_TITLE = "H3K27me3, slice near centrosome"
SIDE_FOOT = "Number = cell id. Arrowhead = deepest point near centrosome. Bars = 5 μm."
PER_PAGE = m.PACK_SHELVES * m.PACK_PER_SHELF
SIDE_VIEW_SUB = "h3k27me3_side_views_20261031"
SIDE_VIEW_PPU = 61.5      # px/μm of the side-view tiles
SIDE_VIEW_LEVEL = 40      # the tiles' dark background carries noise above m.content_crop's level
EXCLUDED_CSV = "C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/config/excluded_cells.csv"
# FF exclusions newer than the CSV's last entry; drop a line once the CSV lists it.
EXCLUDED_NOT_IN_CSV = {("G1B2_h3k27me3_aCD3_10uM_", 12)}   # FF 2026-10-31, DNA 4.1x G1 peak


def excluded_cells() -> set:
    """{(well, cell id)} of FF's exclusions; the CSV's header comments start with #."""
    with open(EXCLUDED_CSV, encoding="utf-8-sig", newline="") as fh:
        listed = {(r["source_folder"], int(float(r["cell_id"])))
                  for r in csv.DictReader(line for line in fh if not line.startswith("#"))}
    return listed | EXCLUDED_NOT_IN_CSV


def kept_from_side_views(folder: str) -> set:
    """The well's cells in its side-view CSV, less any exclusion newer than that CSV."""
    path = m.EXP_ROOT / SIDE_VIEW_SUB / f"{folder}_side_view_cells.csv"
    with open(m._winlong(path), encoding="utf-8-sig", newline="") as fh:
        cells = {int(float(r["cell_id"])) for r in csv.DictReader(fh)}
    return {c for c in cells if (folder, c) not in excluded_cells()}


def configure(exp: dict) -> dict:
    """Point the 10/26 module at this experiment; returns {well: PNG source}."""
    m.EXP_ROOT = Path(exp["root"])
    m.CONDITIONS = [(w, lab) for w, lab, _, _ in exp["wells"]]
    m.RESULTS_MAT = exp["results"]
    m.COMPILE_TRACING = m.EXP_ROOT / exp["compile"] / "cell_tracing.csv"
    print(f"Kept cells from {m.COMPILE_TRACING}")
    m.COMPILE_WELLS = {cond: w for w, _, cond, _ in exp["wells"]}
    m.TILE_NOTES = {}
    return {w: src for w, _, _, src in exp["wells"]}


def xy_tiles(combo: str, sources: dict) -> dict:
    """{condition label: [(tile label, tile), ...]}: top views of the kept cells,
    cell-id order, cropped to content, a 5 μm bar under each."""
    out = {}
    for folder, label in m.CONDITIONS:
        keep = m.kept_ids(folder)
        files = m.percell_files(m.EXP_ROOT / sources[folder] / folder / "physical_scale_images"
                                / combo / "cells")
        out[label] = [(m.tile_id(cid),
                       m.with_scalebar(m.xy_tile(files[cid], id_box=XY_ID_BOX, blob_level=XY_BLOB_LEVEL),
                                       m.XZ_PPU))
                      for cid in sorted(c for c in files if c in keep)]
    return out


def side_view_tiles() -> dict:
    """{condition label: [(cell id, tile), ...]}: H3K27me3 + DNA merge of the slice through
    the deepest point near the centrosome, kept cells in cell-id order, a 5 μm bar under each."""
    out = {}
    for folder, label in m.CONDITIONS:
        tiles = []
        for cid in sorted(m.kept_ids(folder)):
            path = m.EXP_ROOT / SIDE_VIEW_SUB / folder / f"{folder}_cell_{cid}_centside_tile_merge.png"
            if not os.path.isfile(m._winlong(path)):
                print(f"  no near-centrosome slice: {label} cell {cid}")
                continue
            with m.Image.open(m._winlong(path)) as im:
                tile = im.convert("RGB")
            ys, xs = np.nonzero(np.asarray(tile).max(axis=-1) > SIDE_VIEW_LEVEL)
            tile = tile.crop((xs.min(), ys.min(), xs.max() + 1, ys.max() + 1))
            tiles.append((str(cid), m.with_scalebar(tile, SIDE_VIEW_PPU)))
        out[label] = tiles
    return out


def paged(tiles: dict, every_cell: bool) -> list:
    """Pages of PER_PAGE cells per condition: all of them, or the first page only."""
    n = max((-(-len(t) // PER_PAGE) for t in tiles.values()), default=1) if every_cell else 1
    return [{lab: t[k * PER_PAGE:(k + 1) * PER_PAGE] for lab, t in tiles.items()}
            for k in range(max(n, 1))]


def build(name: str) -> int:
    exp = EXPS[name]
    sources = configure(exp)
    results = {w: PROG for w in sources}
    out = Path(exp["out"])
    prs = Presentation()
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    n = 0
    families = ([(c, t, ch, e, "xz") for c, t, ch, e in exp["xz"]]
                + [(c, t, ch, e, "xy") for c, t, ch, e in XY_SLIDES])
    for combo, title, channels, every_cell, view in families:
        tiles = (m.xz_tiles(combo, sources, kept_only=True, results_sources=results,
                            min_blob=XZ_MIN_BLOB, min_peak=XZ_MIN_PEAK)
                 if view == "xz" else xy_tiles(combo, sources))
        # A well whose copy lacks this family says so in its heading, not a silent gap.
        headings = {lab: f"{lab} (not rendered yet)" for lab, t in tiles.items() if not t}
        for lab, t in tiles.items():
            print(f"  {name} {combo} {lab}: {len(t)} cells")
        pages = paged(tiles, every_cell)
        titles = [title] if len(pages) == 1 else [f"{title} {k + 1}/{len(pages)}"
                                                 for k in range(len(pages))]
        n += m.build_packed_deck(out, combo, titles, pages, headings, m.XZ_PPU,
                                 f"{CONTRAST_NOTE} {exp['footer']}",
                                 channels=channels, slide_bar=False, sink=sink)
    tiles = side_view_tiles()
    for lab, t in tiles.items():
        print(f"  {name} near-centrosome slice {lab}: {len(t)} cells")
    pages = paged(tiles, True)
    n += m.build_packed_deck(out, "h3k_side", [f"{SIDE_TITLE} {k + 1}/{len(pages)}" for k in range(len(pages))],
                             pages, {}, SIDE_VIEW_PPU, f"{SIDE_FOOT} {exp['footer']}",
                             channels=[("H3K27me3", RED), ("DNA", CYAN)], slide_bar=False, sink=sink)
    if out.exists():
        m.backup_presentation(str(out))
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"[10/29 {name} compare] {n} slides -> {out}")
    return n


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--kept-side-views"]
    if "--kept-side-views" in sys.argv:
        m.kept_ids = kept_from_side_views
    which = args[0] if args else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(0 if all(build(n) for n in names) else 1)
