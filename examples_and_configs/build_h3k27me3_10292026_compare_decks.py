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

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_compare_decks.py [noco|cilio|all]
"""

import csv
import os
import sys
from pathlib import Path

import numpy as np
from PIL import ImageDraw
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parent))
import insert_noco_10262026_montage_slides as m  # noqa: E402
from insert_noco_10262026_montage_slides import CYAN, MAGENTA, ORANGE, RED, YELLOW  # noqa: E402

PROG ="prog_fixed_cells"
OUT_ROOT = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous"

H3K_XZ = ("H3K27me3_nuc_xz", "H3K27me3 and DNA, side view", [("H3K27me3", RED), ("DNA", CYAN)], True)
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
        "compile": "results_compilation_20261101",
        "wells": [("G1A4_h3k27me3_PLL_dmso_", "PLL DMSO", "PLL_DMSO", PROG),
                  ("G1A1_h3k27me3_aCD3_dmso_", "CD3 DMSO", "CD3_DMSO", PROG),
                  ("G1B1_h3k27me3_aCD3_5uM_", "CD3 Noco 5μM", "CD3_Noco5uM", PROG),
                  ("G1B2_h3k27me3_aCD3_10uM_", "CD3 Noco 10μM", "CD3_Noco10uM", PROG)],
        "xz": [("actin_cent_xz_nolines", "Centrosome repositioning, side view",
                [("Actin", YELLOW), ("Centrin", MAGENTA)], True),
               ("cent_nuc_xz", "Centrosome and nucleus, side view",
                [("Centrin", MAGENTA), ("DNA", CYAN)], True),
               ("actin_nuc_xz_nolines", "Nuclear flattening, side view",
                [("Actin", YELLOW), ("DNA", CYAN)], True),
               H3K_XZ],
        "out": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Noco_Jurkats_H3K27me3_compare_10292026.pptx",
        "footer": "",
    },
    "cilio": {
        "root": "L:/FF/Nucleus_centrosome/ciliobrevin_fixed/"
                "10292026_jurkats_ciliobrevin_h3k27me3640LP45_p549LP45_egfpcentrin2488LP45_h405LP40_",
        "results": "Jurkats_Ciliobrevin_H3K27me3_20261029_results.mat",
        "compile": "results_compilation_20261101",
        "wells": [("G0A4_h3k27me3_PLL_dmso_", "PLL DMSO", "PLL_DMSO", PROG),
                  ("G0A1_h3k27me3_aCD3_dmso_001", "CD3 DMSO", "CD3_DMSO", PROG),
                  ("G0B2_h3k27me3_aCD3_cilio_", "CD3 Cilio 50μM", "CD3_Cilio", PROG)],
        "xz": [("actin_cent_xz_nolines", "Centrosome repositioning, side view",
                [("Actin", MAGENTA), ("Centrin", YELLOW)], True),
               ("cent_nuc_xz", "Centrosome and nucleus, side view",
                [("Centrin", MAGENTA), ("DNA", CYAN)], True),
               ("actin_nuc_xz_nolines", "Nuclear flattening, side view",
                [("Actin", ORANGE), ("DNA", CYAN)], True),
               H3K_XZ],
        "out": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/CilioD_Jurkats_H3K27me3_compare_10292026.pptx",
        "footer": "",
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
SIDE_TITLE = "H3K27me3, slice through Cent.-Prox. Reg."
SIDE_FOOT = "Number = cell id. Arrowhead = deepest point of Cent.-Prox. Reg. Bars = 5 μm."
SAMPLE_FOOT = ("DNA, H3K27me3, merge. White = cell id. Yellow = H3K27me3 around Region Nadir "
               "(arrowhead), 25% region; 1 = cell's surface average. Bars = 5 μm.")
SIDE_PER_SLIDE = 6        # FF 2026-11-01: slice samples, 12 cells per condition over the two slides
SIDE_PER_SHELF = 2
SIDE_COL_GAP = 0.22       # inches, so a value does not read as the next cell's id
SIDE_METRIC = "H3K27me3_S2_f025_rel"      # FF 2026-11-01: printed by each slice, to compare with the eye
SEED_DEPTH_CSV = "tests/seed_depth_20261101/seed_depth/seed_depth_percell_Oct_29,_2026.csv"
SIDE_VALUE_GAP = 440      # px of black right of the merge tile, for the value: over the tile it hit the arrowhead
SAMPLE_PER_SLIDE = 3      # cells per condition per slide: three tiles each, so far fewer than PER_PAGE
SAMPLE_SLIDES = 2         # FF 2026-11-01: a sampling, not every cell
SAMPLE_GAP = 14           # px between a cell's three tiles
NATIVE_BAR_PX = 150       # white px in a row of a gray side-view tile: its burned-in bar, not signal
XZ_BAR_STRIP = 32         # px, the strip m.with_scalebar adds at its default thickness
PER_PAGE = m.PACK_SHELVES * m.PACK_PER_SHELF
SIDE_VIEW_SUB = "H3K27me3/side_views"     # of <prog_fixed_cells>/<well>
SIDE_VIEW_PPU = 61.5      # px/μm of the side-view tiles
SIDE_VIEW_LEVEL = 40      # the tiles' dark background carries noise above m.content_crop's level
SIDE_BAR_THICK = 40       # px; these tiles shrink more than the others on the slide
SIDE_LABEL_PT = 15        # cell ids, readable at FF's fit-to-width view
ARROW_SCALE = 1.8         # the native arrowhead is small at montage scale
ARROW_COLOR = (255, 255, 0)   # FF 2026-11-01: not magenta; yellow as the Cent.-Prox. Reg. on the meshes
EXCLUDED_CSV = "C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/config/excluded_cells.csv"


def excluded_cells() -> set:
    """{(well, cell id)} of FF's exclusions; the CSV's header comments start with #."""
    with open(EXCLUDED_CSV, encoding="utf-8-sig", newline="") as fh:
        return {(r["source_folder"], int(float(r["cell_id"])))
                for r in csv.DictReader(line for line in fh if not line.startswith("#"))}


def side_view_dir(well: str) -> Path:
    """The well's side views, as analyze writes them: Cell_<id>_{deepest,centside}_*."""
    return m.EXP_ROOT / PROG / well / SIDE_VIEW_SUB


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


def bigger_arrowhead(tile):
    """Redraw the tile's native magenta arrowhead ARROW_SCALE times larger, same tip and direction,
    in ARROW_COLOR; the larger one covers the native one. plot_invag_side_view.m points it at the
    deepest point from the nucleus side: down for an invagination on the coverslip side, else up."""
    a = np.asarray(tile).astype(int)
    ys, xs = np.nonzero((a[..., 0] > 170) & (a[..., 1] < 90) & (a[..., 2] > 170))
    if not len(ys):
        return tile
    points_down = ys.mean() < (ys.min() + ys.max()) / 2     # most of a triangle lies toward its base
    tip_y = ys.max() if points_down else ys.min()
    tip_x = xs[ys == tip_y].mean()
    half_w = ARROW_SCALE * (xs.max() - xs.min() + 1) / 2
    h = ARROW_SCALE * (ys.max() - ys.min() + 1) * (1 if points_down else -1)
    ImageDraw.Draw(tile).polygon([(tip_x - half_w, tip_y - h), (tip_x + half_w, tip_y - h), (tip_x, tip_y)],
                                 fill=ARROW_COLOR)
    return tile


def side_view_tiles() -> dict:
    """{condition label: [(cell id, tile), ...]}: H3K27me3 + DNA merge of the slice through
    the deepest point of the Cent.-Prox. Reg., kept cells in cell-id order, a 5 μm bar under each."""
    out = {}
    for folder, label in m.CONDITIONS:
        tiles = []
        for cid in sorted(m.kept_ids(folder)):
            path = side_view_dir(folder) / f"Cell_{cid}_centside_tile_merge.png"
            if not os.path.isfile(m._winlong(path)):
                print(f"  no Cent.-Prox. Reg. slice: {label} cell {cid}")
                continue
            with m.Image.open(m._winlong(path)) as im:
                tile = bigger_arrowhead(im.convert("RGB"))
            ys, xs = np.nonzero(np.asarray(tile).max(axis=-1) > SIDE_VIEW_LEVEL)
            tile = tile.crop((xs.min(), ys.min(), xs.max() + 1, ys.max() + 1))
            tiles.append((str(cid), m.with_scalebar(tile, SIDE_VIEW_PPU, thick=SIDE_BAR_THICK)))
        out[label] = tiles
    return out


def without_native_bar(tile, merge):
    """The H3K27me3 side-view tile carries the strip's burned-in bar; put back the merge's red there."""
    a = np.array(tile)
    rows = np.nonzero((a.min(axis=-1) > 220).sum(axis=1) > NATIVE_BAR_PX)[0]
    if len(rows):
        cols = np.nonzero((a[rows].min(axis=-1) > 220).any(axis=0))[0]
        r0, r1, c0, c1 = rows.min() - 8, rows.max() + 9, cols.min() - 8, cols.max() + 9
        a[r0:r1, c0:c1] = np.asarray(merge)[r0:r1, c0:c1, 0][..., None]
    return m.Image.fromarray(a)


def joined(tiles: list):
    """Tiles of one height side by side."""
    strip = m.Image.new("RGB", (sum(t.width for t in tiles) + SAMPLE_GAP * (len(tiles) - 1), tiles[0].height),
                        (0, 0, 0))
    x = 0
    for tile in tiles:
        strip.paste(tile, (x, 0))
        x += tile.width + SAMPLE_GAP
    return strip


def channel_strip(tile):
    """A red + cyan tile with its 5 μm bar strip as DNA, H3K27me3, merge: red and cyan never mix in
    these PNGs, so the channels split exactly. One bar, under the DNA tile."""
    a = np.asarray(tile.convert("RGB"))
    dna, h3k = (np.repeat(a[..., k:k + 1], 3, axis=-1) for k in (1, 0))
    merge = a.copy()
    h3k[-XZ_BAR_STRIP:] = 0
    merge[-XZ_BAR_STRIP:] = 0
    return joined([m.Image.fromarray(x) for x in (dna, h3k, merge)])


def seed_depth_values(field: str) -> dict:
    """{(condition label, cell id): value} from this experiment's seed-depth per-cell table."""
    with open(m._winlong(m.EXP_ROOT / SEED_DEPTH_CSV), encoding="utf-8-sig", newline="") as fh:
        return {(r["condition"], int(float(r["cell_id"]))): float(r[field]) for r in csv.DictReader(fh)}


def side_view_samples() -> dict:
    """{condition label: [((cell id, SIDE_METRIC), strip), ...]}: the first kept cells in cell-id
    order, each as DNA, H3K27me3, merge of the slice through the Cent.-Prox. Reg., all three on the
    merge's crop (it holds no native bar), one 5 μm bar under the strip."""
    values = seed_depth_values(SIDE_METRIC)
    out = {}
    for folder, label in m.CONDITIONS:
        strips = []
        for cid in sorted(m.kept_ids(folder)):
            if len(strips) == SIDE_PER_SLIDE * SAMPLE_SLIDES:
                break
            paths = [side_view_dir(folder) / f"Cell_{cid}_centside_tile_{ch}.png"
                     for ch in ("hoechst", "h3k27me3", "merge")]
            if not all(os.path.isfile(m._winlong(p)) for p in paths):
                print(f"  no Cent.-Prox. Reg. slice: {label} cell {cid}")
                continue
            tiles = []
            for path in paths:
                with m.Image.open(m._winlong(path)) as im:
                    tiles.append(im.convert("RGB"))
            tiles[1] = without_native_bar(tiles[1], tiles[2])
            tiles = [bigger_arrowhead(t) for t in tiles]
            ys, xs = np.nonzero(np.asarray(tiles[-1]).max(axis=-1) > SIDE_VIEW_LEVEL)
            box = (xs.min(), ys.min(), xs.max() + 1, ys.max() + 1)
            tiles = [t.crop(box) for t in tiles]
            strip = m.with_scalebar(joined(tiles), SIDE_VIEW_PPU, thick=SIDE_BAR_THICK)
            room = m.Image.new("RGB", (strip.width + SIDE_VALUE_GAP, strip.height), (0, 0, 0))
            room.paste(strip, (0, 0))
            strips.append(((str(cid), f"{values[label, cid]:.2f}"), room))
        out[label] = strips
    return out


def add_sample_slides(out: Path, tag: str, title: str, tiles: dict, ppu: float, footer: str, sink: dict = None,
                      label_pt: float = m.TILE_LABEL_PT, per_slide: int = SAMPLE_PER_SLIDE,
                      per_shelf: int = 1, col_gap: float = 0.06) -> int:
    """A sampling of the cells as DNA, H3K27me3, merge strips: SAMPLE_SLIDES slides."""
    pages = [{lab: t[k * per_slide:(k + 1) * per_slide] for lab, t in tiles.items()}
             for k in range(SAMPLE_SLIDES)]
    titles = [f"{title} {k + 1}/{SAMPLE_SLIDES}" for k in range(SAMPLE_SLIDES)]
    return m.build_packed_deck(out, tag, titles, pages, not_rendered(tag, tiles, sink if sink is not None else {}),
                               ppu, footer, channels=[("H3K27me3", RED), ("DNA", CYAN)], slide_bar=False,
                               sink=sink, per_shelf=per_shelf, label_pt=label_pt, right_pt=label_pt,
                               col_gap=col_gap)


def xz_samples(tiles: dict) -> dict:
    """The first cells of a red + cyan family as channel strips."""
    n = SAMPLE_PER_SLIDE * SAMPLE_SLIDES
    return {lab: [(text, channel_strip(tile)) for text, tile in t[:n]] for lab, t in tiles.items()}


def paged(tiles: dict, every_cell: bool) -> list:
    """Pages of PER_PAGE cells per condition: all of them, or the first page only."""
    n = max((-(-len(t) // PER_PAGE) for t in tiles.values()), default=1) if every_cell else 1
    return [{lab: t[k * PER_PAGE:(k + 1) * PER_PAGE] for lab, t in tiles.items()}
            for k in range(max(n, 1))]


def not_rendered(family: str, tiles: dict, sink: dict) -> dict:
    """Headings for wells with no tiles: a well lacking this family says so, not a silent
    gap, and build() then exits non-zero."""
    empty = [lab for lab, t in tiles.items() if not t]
    sink.setdefault("empty", []).extend(f"{family}: {lab}" for lab in empty)
    return {lab: f"{lab} (not rendered yet)" for lab in empty}


def add_slides(name: str, sink: dict, one_page: bool = False) -> int:
    """Every family of this set onto sink's deck, tagged by family. one_page: 12 cells per
    condition per family, not every kept cell (the two-experiment Noco deck)."""
    exp = EXPS[name]
    sources = configure(exp)
    results = {w: PROG for w in sources}
    out = Path(exp["out"])
    n = 0
    families = ([(c, t, ch, e, "xz") for c, t, ch, e in exp["xz"]]
                + [(c, t, ch, e, "xy") for c, t, ch, e in XY_SLIDES])
    for combo, title, channels, every_cell, view in families:
        tiles = (m.xz_tiles(combo, sources, kept_only=True, results_sources=results,
                            min_blob=XZ_MIN_BLOB, min_peak=XZ_MIN_PEAK)
                 if view == "xz" else xy_tiles(combo, sources))
        headings = not_rendered(combo, tiles, sink)
        for lab, t in tiles.items():
            print(f"  {name} {combo} {lab}: {len(t)} cells")
        if one_page and combo == H3K_XZ[0]:     # FF 2026-11-01: the merge alone says too little
            n += add_sample_slides(out, combo, title, xz_samples(tiles), m.XZ_PPU,
                                   f"Left to right: DNA, H3K27me3, merge. {CONTRAST_NOTE} Bars = 5 μm.", sink,
                                   SIDE_LABEL_PT)
            continue
        pages = paged(tiles, every_cell and not one_page)
        titles = [title] if len(pages) == 1 else [f"{title} {k + 1}/{len(pages)}"
                                                 for k in range(len(pages))]
        n += m.build_packed_deck(out, combo, titles, pages, headings, m.XZ_PPU,
                                 f"{CONTRAST_NOTE} Bars = 5 μm. {exp['footer']}",
                                 channels=channels, slide_bar=False, sink=sink)
    if one_page:
        return n + add_sample_slides(out, "h3k_side", SIDE_TITLE, side_view_samples(), SIDE_VIEW_PPU,
                                     SAMPLE_FOOT, sink, SIDE_LABEL_PT, SIDE_PER_SLIDE, SIDE_PER_SHELF, SIDE_COL_GAP)
    tiles = side_view_tiles()
    for lab, t in tiles.items():
        print(f"  {name} Cent.-Prox. Reg. slice {lab}: {len(t)} cells")
    pages = paged(tiles, True)
    titles = [SIDE_TITLE] if len(pages) == 1 else [f"{SIDE_TITLE} {k + 1}/{len(pages)}" for k in range(len(pages))]
    n += m.build_packed_deck(out, "h3k_side", titles,
                             pages, not_rendered("h3k_side", tiles, sink), SIDE_VIEW_PPU, f"{SIDE_FOOT} {exp['footer']}",
                             channels=[("H3K27me3", RED), ("DNA", CYAN)], slide_bar=False, sink=sink,
                             label_pt=SIDE_LABEL_PT)
    return n


def build(name: str) -> int:
    out = Path(EXPS[name]["out"])
    prs = Presentation()
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    n = add_slides(name, sink)
    if out.exists():
        m.backup_presentation(str(out))
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"[10/29 {name} compare] {n} slides -> {out}")
    if sink.get("empty"):
        print(f"[MISSING] {len(sink['empty'])} block(s): " + "; ".join(sink["empty"]))
        return 0
    return n


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(0 if all(build(n) for n in names) else 1)
