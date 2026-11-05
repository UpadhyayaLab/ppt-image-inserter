#!/usr/bin/env python
"""
Blender nuclear-mesh montages for the 10/29/2026 H3K27me3 sets (FF 2026-10-31), in the
10/26 hull format: first the overview slides, 15 cells per condition per slide, one view
at a time, then compact pair slides (hull and H3K27me3 of each cell joined, 12 cells per
condition per slide), then per-cell rows with the three views side by side. Cell-id order.

Tiles (microscopy-nodes-blender-vis session), one px/μm for the whole run:
<TILES>/<prefix>/{hull,h3k27me3,cent_proximal}/<prefix>_cell_<id>[_EXCLUDED|_PARTIAL]_<view>.png,
and the same under <DEEPEST>/<prefix>/deepest_by_cent/ (FF 2026-11-01: the centrosome-side pass only).
A condition's tiles count only once its noco_strips_record.json exists (the last file
written), so a folder still being copied is never read; until then its hull tiles come from
the finished gallery, its other views are left out and its heading says so. Rerun when it lands. _EXCLUDED cells are hidden here (FF:
"render everything and just don't show it"), as are cells FF excluded after their render
(excluded_cells()). Partial cells, by tile tag or by partial_exclusions.csv, are hidden too
(FF 2026-10-31: "leave 'partial' out of mesh montages but not others"), so the means beside
the hull slides leave them out.

Cells come from the tile names, depth from the compile. No by-centrosome voxel ratio is printed.

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_mesh_decks.py [noco|cilio|all]
"""

import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import numpy as np
from pptx import Presentation
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_h3k27me3_10292026_compare_decks import EXPS, OUT_ROOT, configure, excluded_cells, m  # noqa: E402
from build_h3k27me3_10292026_side_view_deck import partial_cells, side_cells  # noqa: E402
from insert_noco_10262026_montage_slides import GRAY, WHITE  # noqa: E402

RAW = Path("L:/FF/blender_renderings/_raw")
TILES = RAW / "h3k1029_final_tiles_20261101"
DEEPEST = RAW / "h3k1029_final_deepest_20261101"
HULL_ONLY = RAW / "h3k1029_hull_gallery_20261031"   # same hull pixels, finished earlier
BARS = RAW / "h3k1029_tiles_20261031" / "_colorbars"   # still valid for the final tiles
PREFIX = {"noco": ["noco1029_pll", "noco1029_dmso", "noco1029_5um", "noco1029_10um"],
          "cilio": ["cilio1029_pll", "cilio1029_dmso", "cilio1029_cilio"]}
OUT = {"noco": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Noco_Jurkats_H3K27me3_meshes_10292026.pptx",
       "cilio": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/CilioD_Jurkats_H3K27me3_meshes_10292026.pptx"}
PPU = 56               # px/μm of every tile (the run's --ppu)
TILE_BAR_THICK = 40    # px; overview and pair tiles shrink the most on the slide
MESH_LABEL_PT = 15     # tile numbers and side notes, readable at FF's fit-to-width view
PAIR_LABEL_PT = 16     # the pair tiles shrink more than the overview tiles
SMALL_BAR_PT = 16      # colorbar text in the right column
m.SIDE_NOTE_W = 1.9   # in; fits the hull side notes at MESH_LABEL_PT, clear of the colorbar
PER_PAGE = m.PACK_SHELVES * m.HULL_PER_SHELF
WORKERS = 8
HIDDEN = excluded_cells() | {(well, int(cid)) for well, cid in partial_cells()}
TICKS = ["≥2", "1", "0"]
HULL_BAR = (BARS / "hull_invag_depth_colorbar_light.png", "Hull Surf.\nDist. (μm)")
H3K_BAR = (BARS / "h3k27me3_crimson_teal_colorbar_light.png", "Relative\nH3K27me3")
# (view, slide title, colorbar or None, footer: one line, so about 130 characters at most)
VIEWS = [("hull", "Hull Surf. Dist.", HULL_BAR,
          "Number = Invag. Depth in Cent.-Prox. Reg. (μm). Nucleus rotated to show it. "
          "Dot = centrosome. Bars = 5 μm."),
         ("h3k27me3", "Relative H3K27me3 density", H3K_BAR,
          "Density within 0.25 μm of the nuclear surface, per face / the cell's area-weighted mean. "
          "Dot = centrosome. Bars = 5 μm."),
         ("cent_proximal", "Cent.-Prox. Reg.", None,
          "Yellow = Cent.-Prox. Reg., cyan = away. "
          "Dot = centrosome. Bars = 5 μm."),
         ("deepest_by_cent", "Deepest region near centrosome", None,
          "Yellow = deepest region near the centrosome, green = away. "
          "Dot = centrosome. Bars = 5 μm.")]
# Colored legend, upper left, for the views with no colorbar (FF 2026-11-01). Green as rendered.
VIEW_LEGEND = {"deepest_by_cent": [("Deepest region", m.YELLOW), ("Away", m.RGBColor(0x00, 0xB8, 0x00))]}
PAIR_TITLE = "Hull Surf. Dist. (left), Relative H3K27me3 (right)"
PAIR_FOOT = "Number = Invag. Depth in Cent.-Prox. Reg. (μm). Dot = centrosome. Bars = 5 μm."
PAIR_PER_SHELF = 4
PAIR_PER_PAGE = m.PACK_SHELVES * PAIR_PER_SHELF
PAIR_GAP = 16         # px between the two meshes of a pair
ROW_HEADS = ("Hull Surf. Dist.", "Relative H3K27me3", "Cent.-Prox. Reg.")
ROW_VIEWS = [v for v, _, _, _ in VIEWS[:3]]   # the deepest-region tiles stay on the overview slides
ROWS_PER_SLIDE = 3
ROW_TOP = 1.10
ROW_H = 1.95
ROW_LABEL_W = 1.05
ROW_GAP = 0.08
ROW_BAR_W = 1.20
ROW_FOOT = "Yellow = Cent.-Prox. Reg., cyan = away. Dot = centrosome. Bar = 5 μm."


def tile_dir(prefix: str, view: str) -> Path:
    """The view's finished folder; see the module docstring for what counts as finished."""
    root = DEEPEST if view == "deepest_by_cent" else TILES
    if os.path.isfile(m._winlong(root / prefix / "noco_strips_record.json")):
        return root / prefix / view
    return HULL_ONLY / prefix / view


def tile_files(prefix: str, view: str) -> dict:
    """{cell id: path} of the shown cells; _EXCLUDED and _PARTIAL cells are left out."""
    d = tile_dir(prefix, view)
    if not os.path.isdir(m._winlong(d)):
        return {}
    pat = re.compile(rf"^{re.escape(prefix)}_cell_(\d+)(_EXCLUDED|_PARTIAL)?_{view}\.png$")
    hits = [pat.match(name) for name in os.listdir(m._winlong(d))]
    well = dict(zip(PREFIX["noco"] + PREFIX["cilio"], (w[0] for n in ("noco", "cilio") for w in EXPS[n]["wells"])))[prefix]
    return {int(h.group(1)): d / h.group(0)
            for h in hits if h and not h.group(2) and (well, int(h.group(1))) not in HIDDEN}


def load(path: Path):
    with m.Image.open(m._winlong(path)) as im:
        return im.convert("RGB")


def set_bar(bar: tuple) -> None:
    """Point the 10/26 module's colorbar at this view's bar."""
    m.COLORBAR_LIGHT, m.CBAR_TITLE = bar
    m.CBAR_TICKS = TICKS


def overview(name: str, sink: dict, depth: dict, pool) -> None:
    """One view at a time: 15 cells per condition per slide, in the 10/26 hull format."""
    labels = [lab for _, lab in m.CONDITIONS]
    for view, title, bar, footer in VIEWS:
        tiles, vals = {}, {}
        for prefix, lab in zip(PREFIX[name], labels):
            files = tile_files(prefix, view)
            cids = sorted(files)
            ims = pool.map(lambda c: m.with_scalebar(m.content_crop(load(files[c])), PPU, thick=TILE_BAR_THICK),
                           cids)
            vals[lab] = [depth[lab].get(c, float("nan")) for c in cids]
            tiles[lab] = [(f"{v:.1f}" if view == "hull" else "", im)
                          for v, im in zip(vals[lab], ims)]
        n_pages = max(-(-len(t) // PER_PAGE) for t in tiles.values())
        pages = [{lab: t[k * PER_PAGE:(k + 1) * PER_PAGE] for lab, t in tiles.items()}
                 for k in range(n_pages)]
        heads = {lab: lab if tiles[lab] else f"{lab}: rendering" for lab in labels}
        notes = None
        if view == "hull":      # mean depth beside each montage (FF 2026-10-27)
            notes = [{lab: f"μ={np.nanmean(v):.2f} μm (all)\nμ={np.nanmean(v[k * PER_PAGE:(k + 1) * PER_PAGE]):.2f} μm (shown)"
                      for lab, v in vals.items() if v[k * PER_PAGE:(k + 1) * PER_PAGE]} for k in range(n_pages)]
        if bar:
            set_bar(bar)
        print(f"  {view}: " + ", ".join(f"{lab} {len(t)}" for lab, t in tiles.items()))
        m.build_packed_deck(None, view, [f"{title} {k + 1}/{n_pages}" for k in range(n_pages)], pages, heads,
                            PPU, footer, channels=VIEW_LEGEND.get(view), colorbar=bool(bar), slide_bar=False,
                            side_notes=notes, sink=sink, per_shelf=m.HULL_PER_SHELF, label_pt=MESH_LABEL_PT)


def crop_views(files: list) -> list:
    """Views of one cell, cropped to one box, so they stay in register."""
    ims = [load(p) for p in files]
    mask = np.maximum.reduce([np.asarray(im).max(axis=-1) for im in ims]) > 15
    ys, xs = np.nonzero(mask)
    box = (max(0, xs.min() - 4), max(0, ys.min() - 4), xs.max() + 5, ys.max() + 5)
    return [im.crop(box) for im in ims]


def row_tiles(files: list) -> tuple:
    """The three views of one cell, each with its own 5 μm bar."""
    return tuple(m.with_scalebar(im, PPU) for im in crop_views(files))


def pair_tile(files: list):
    """Two views of one cell joined side by side, with one 5 μm bar."""
    a, b = crop_views(files)
    joined = m.Image.new("RGB", (a.width + PAIR_GAP + b.width, a.height))
    joined.paste(a, (0, 0))
    joined.paste(b, (a.width + PAIR_GAP, 0))
    return m.with_scalebar(joined, PPU, thick=TILE_BAR_THICK)


def pairs(name: str, sink: dict, depth: dict, pool) -> None:
    """Compact slides: hull and H3K27me3 of each cell as one tile, all conditions on a slide."""
    labels = [lab for _, lab in m.CONDITIONS]
    tiles = {}
    for prefix, lab in zip(PREFIX[name], labels):
        hull, h3k = tile_files(prefix, "hull"), tile_files(prefix, "h3k27me3")
        cids = sorted(set(hull) & set(h3k))
        ims = pool.map(lambda c: pair_tile([hull[c], h3k[c]]), cids)
        tiles[lab] = [(f"{depth[lab].get(c, float('nan')):.1f}", im) for c, im in zip(cids, ims)]
    n_pages = max(-(-len(t) // PAIR_PER_PAGE) for t in tiles.values())
    if not n_pages:
        return
    pages = [{lab: t[k * PAIR_PER_PAGE:(k + 1) * PAIR_PER_PAGE] for lab, t in tiles.items()}
             for k in range(n_pages)]
    heads = {lab: lab if tiles[lab] else f"{lab}: rendering" for lab in labels}

    def two_bars(slide) -> None:
        add_small_bar(slide, HULL_BAR, 1.10)
        add_small_bar(slide, H3K_BAR, 3.60)

    print("  pairs: " + ", ".join(f"{lab} {len(t)}" for lab, t in tiles.items()))
    m.build_packed_deck(None, "pairs", [f"{PAIR_TITLE} {k + 1}/{n_pages}" for k in range(n_pages)], pages, heads,
                        PPU, PAIR_FOOT, colorbar=two_bars, slide_bar=False,
                        sink=sink, per_shelf=PAIR_PER_SHELF, label_pt=PAIR_LABEL_PT)


def add_small_bar(slide, bar: tuple, top: float) -> None:
    """A colorbar in the right column of a per-cell slide; ramp from Blender, labels here."""
    path, title = bar
    left = m.SLIDE_W - m.MARGIN - ROW_BAR_W
    m.add_textbox(slide, title.replace("\n", " "), left - 0.05, top, ROW_BAR_W + 0.10, 0.76,
                  font_pt=SMALL_BAR_PT, color=WHITE, bold=True).text_frame.word_wrap = True
    h = 1.55
    y0 = top + 0.80
    slide.shapes.add_picture(m.colorbar_ramp(path), Inches(left + 0.05), Inches(y0),
                             width=Inches(0.28), height=Inches(h))
    for k, tick in enumerate(TICKS):
        m.add_textbox(slide, tick, left + 0.32, y0 + h * k / 2 - 0.17, 0.6, 0.34,
                      font_pt=SMALL_BAR_PT,
                      color=WHITE, bold=True)


def cell_rows(name: str, prs, pool) -> int:
    """Per-cell rows for the conditions with all three views on disk."""
    labels = [lab for _, lab in m.CONDITIONS]
    conds = []
    for prefix, lab in zip(PREFIX[name], labels):
        files = {v: tile_files(prefix, v) for v in ROW_VIEWS}
        cids = sorted(set.intersection(*(set(f) for f in files.values())))
        if not cids:
            continue
        rows = pool.map(lambda c: (c, row_tiles([files[v][c] for v in ROW_VIEWS])), cids)
        conds.append((lab, list(rows)))
    if not conds:
        return 0
    col_w = (m.SLIDE_W - 2 * m.MARGIN - ROW_LABEL_W - ROW_BAR_W - 3 * ROW_GAP) / 3
    every = [t for _, rows in conds for _, ts in rows for t in ts]
    ppi = max(max(t.width for t in every) / col_w, max(t.height for t in every) / (ROW_H - 0.05))
    n = 0
    for lab, rows in conds:
        pages = [rows[i:i + ROWS_PER_SLIDE] for i in range(0, len(rows), ROWS_PER_SLIDE)]
        for k, page in enumerate(pages, start=1):
            slide = m.new_slide(prs, f"{lab}, mesh views {k}/{len(pages)}", None)
            for j, head in enumerate(ROW_HEADS):
                m.add_textbox(slide, head, m.MARGIN + ROW_LABEL_W + j * (col_w + ROW_GAP), 0.66, col_w, 0.36,
                              font_pt=18, color=WHITE, bold=True)
            for i, (cid, tiles) in enumerate(page):
                top = ROW_TOP + i * ROW_H
                m.add_textbox(slide, f"cell {cid}", m.MARGIN, top + ROW_H / 2 - 0.18, ROW_LABEL_W, 0.36,
                              font_pt=16, color=WHITE, bold=True)
                for j, t in enumerate(tiles):
                    x = m.MARGIN + ROW_LABEL_W + j * (col_w + ROW_GAP) + (col_w - t.width / ppi) / 2
                    buf = m.io.BytesIO()
                    t.save(buf, format="PNG")
                    buf.seek(0)
                    slide.shapes.add_picture(buf, Inches(x), Inches(top + (ROW_H - t.height / ppi) / 2),
                                             width=Inches(t.width / ppi))
            add_small_bar(slide, HULL_BAR, 1.10)
            add_small_bar(slide, H3K_BAR, 3.60)
            m.add_textbox(slide, ROW_FOOT, m.MARGIN, 7.08, m.SLIDE_W - 2 * m.MARGIN, 0.36, font_pt=14, color=GRAY)
            n += 1
    print(f"  rows: " + ", ".join(f"{lab} {len(r)}" for lab, r in conds) + f", PPI {ppi:.1f}")
    return n


def cent_depths(name: str) -> dict:
    """{condition label: {cell id: Invag. Depth in Cent.-Prox. Reg.}}."""
    return {lab: {int(c["cell_id"]): float(c["chull_max_D_by_cent"]) for c in side_cells(name, w)}
            for w, lab, _, _ in EXPS[name]["wells"]}


def build(name: str) -> int:
    configure(EXPS[name])
    m.STRIP_FOOT_FONT_PT = 15   # one line under the montage
    depth = cent_depths(name)
    prs = Presentation()
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    with ThreadPoolExecutor(WORKERS) as pool:
        overview(name, sink, depth, pool)
        pairs(name, sink, depth, pool)
        rows = cell_rows(name, prs, pool)
    out = Path(OUT[name])
    if out.exists():
        m.backup_presentation(str(out))
    prs.save(str(out))
    print(f"[meshes {name}] {len(prs.slides)} slides ({rows} per-cell) -> {out}")
    return len(prs.slides)


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(0 if all(build(n) for n in names) else 1)
