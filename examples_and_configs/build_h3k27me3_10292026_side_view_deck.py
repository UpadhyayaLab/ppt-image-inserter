#!/usr/bin/env python
"""
Lamin-style slice deck for the 10/29/2026 H3K27me3 sets, one cell per slide (plan
agreed with the TCell session, 2026-10-31; FF: "slices to best identify deepest
invagination", cell-id order). Top row: deepest invagination overall. Bottom row:
centrosome side. Each row: H3K27me3 | Hoechst | merge | trace along the envelope.
Every kept cell of both sets; a Partial cell (FF's mesh QC) says so in the slide title.

Tiles come from TCell-3D-Morphodynamics dev/analyses/lamin_reslice/plot_invag_side_view.m:
<SIDE_DIR>/<well>/<well>_cell_<id>_{deepest,centside}_tile_{h3k27me3,hoechst,merge}.png
and _profile.png, with <SIDE_DIR>/<well>_side_view_cells.csv beside the well folders.
All image tiles sit at one PPI deck-wide, so they keep one physical scale. Wells without
tiles are skipped.

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_side_view_deck.py [noco|cilio|all]
"""

import csv
import math
import os
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_h3k27me3_10292026_compare_decks import EXPS, m  # noqa: E402
from insert_noco_10262026_montage_slides import CYAN, GRAY, RED, WHITE  # noqa: E402

SIDE_SUB = "h3k27me3_side_views_20261031"
OUT_ROOT = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous"
OUT = {"noco": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Noco_Jurkats_H3K27me3_side_views_10292026.pptx",
       "cilio": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/CilioD_Jurkats_H3K27me3_side_views_10292026.pptx"}
PARTIAL_CSV = "C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/config/partial_exclusions.csv"
FOOTER = "Bars = 5 μm. Arrowhead = deepest point of that row's region. PRELIMINARY: no figure review yet"

TILE_PX_W = 1108       # 18 μm at 61.5 px/μm (TCell session)
TILE_GAP = 0.10
TRACE_W = 2.95         # in; the trace's 24 pt ticks stay ~18 pt at this width
TILE_W = (m.SLIDE_W - 2 * m.MARGIN - TRACE_W - 3 * TILE_GAP) / 3
TILE_PPI = TILE_PX_W / TILE_W
HEAD_TOP = 0.60
HEAD_FONT_PT = 20
ROW_TOPS = (0.98, 4.00)
ROW_H = 3.00
ROW_LABEL_H = 0.32
ROW_LABEL_FONT_PT = 18
FOOT_TOP = 7.03
FOOT_FONT_PT = 16
TILES = ("h3k27me3", "hoechst", "merge")
TILE_HEADS = ("H3K27me3", "Hoechst", "Merge")


def partial_cells() -> set:
    """{(well, cell id)} of FF's Partial cells; the CSV's header comments start with #."""
    with open(PARTIAL_CSV, encoding="utf-8-sig", newline="") as fh:
        return {(r["source_folder"], r["cell_id"])
                for r in csv.DictReader(line for line in fh if not line.startswith("#"))}


def side_cells(side_dir: Path, well: str) -> list:
    """Rows of the well's CSV in cell-id order, or [] when it has no tiles yet."""
    path = side_dir / f"{well}_side_view_cells.csv"
    if not os.path.isfile(m._winlong(path)):
        return []
    with open(m._winlong(path), encoding="utf-8-sig", newline="") as fh:
        return sorted(csv.DictReader(fh), key=lambda r: int(r["cell_id"]))


def fmt(value: str) -> str:
    """A CSV number to 2 decimals, or n/a when it is NaN or empty."""
    v = float(value) if value else math.nan
    return "n/a" if math.isnan(v) else f"{v:.2f}"


def row_labels(cell: dict) -> tuple:
    """(top, bottom) row text. When the deepest point is on the centrosome side,
    both rows show the same slices, so the bottom row says so."""
    same = cell["chull_max_D"] == cell["chull_max_D_by_cent"]
    return (f"Deepest invagination: {fmt(cell['chull_max_D'])} μm deep, "
            f"relative level {fmt(cell['ratio_deepest_025um'])}",
            f"Centrosome side{' (same point as above)' if same else ''}: "
            f"{fmt(cell['chull_max_D_by_cent'])} μm deep, "
            f"relative level {fmt(cell['ratio_centside_025um'])}")


def add_headers(slide) -> None:
    """Column names over the tiles; the trace's names in its two line colors."""
    x = m.MARGIN
    for head in TILE_HEADS:
        m.add_textbox(slide, head, x, HEAD_TOP, TILE_W, 0.36, font_pt=HEAD_FONT_PT,
                      color=WHITE, bold=True)
        x += TILE_W + TILE_GAP
    box = m.add_textbox(slide, "H3K27me3", x, HEAD_TOP, TRACE_W, 0.36, font_pt=HEAD_FONT_PT,
                        color=RED, bold=True)
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "  Hoechst"
    run.font.size = Pt(HEAD_FONT_PT)
    run.font.bold = True
    run.font.color.rgb = CYAN


def add_row(slide, side_dir: Path, stem: str, top: float, label: str) -> list:
    """One row: label over the tiles, tiles centered in the row, trace at the right.
    Returns the files that are missing."""
    box = m.add_textbox(slide, label, m.MARGIN, top, 3 * TILE_W + 2 * TILE_GAP, ROW_LABEL_H,
                        font_pt=ROW_LABEL_FONT_PT, color=WHITE, bold=True)
    box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    missing = []
    area_h = ROW_H - ROW_LABEL_H
    for j, name in enumerate(TILES):
        path = side_dir / f"{stem}_tile_{name}.png"
        if not os.path.isfile(m._winlong(path)):
            missing.append(path.name)
            continue
        with m.Image.open(m._winlong(path)) as im:
            h_in = im.size[1] / TILE_PPI
        slide.shapes.add_picture(m._winlong(path), Inches(m.MARGIN + j * (TILE_W + TILE_GAP)),
                                 Inches(top + ROW_LABEL_H + (area_h - h_in) / 2),
                                 width=Inches(TILE_W))
    trace = side_dir / f"{stem}_profile.png"
    if os.path.isfile(m._winlong(trace)):
        slide.shapes.add_picture(m._winlong(trace), Inches(m.SLIDE_W - m.MARGIN - TRACE_W),
                                 Inches(top + (ROW_H - TRACE_W) / 2), width=Inches(TRACE_W))
    else:
        missing.append(trace.name)
    return missing


def build(name: str) -> int:
    side_dir = Path(EXPS[name]["root"]) / SIDE_SUB
    partial = partial_cells()
    prs = Presentation()
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    n, missing = 0, []
    for well, label, _, _ in EXPS[name]["wells"]:
        cells = side_cells(side_dir, well)
        print(f"  {label}: {len(cells)} cells, {sum((well, c['cell_id']) in partial for c in cells)} Partial")
        for cell in cells:
            cid = cell["cell_id"]
            tag = ", Partial" if (well, cid) in partial else ""
            slide = m.new_slide(prs, f"{label}, cell {cid}{tag}", None)
            add_headers(slide)
            for region, top, text in zip(("deepest", "centside"), ROW_TOPS, row_labels(cell)):
                folder = well + "/" + f"{well}_cell_{cid}_{region}"
                missing += add_row(slide, side_dir, folder, top, text)
            foot = m.add_textbox(slide, FOOTER, m.MARGIN, FOOT_TOP, m.SLIDE_W - 2 * m.MARGIN, 0.36,
                                 font_pt=FOOT_FONT_PT, color=GRAY)
            foot.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
            n += 1
    if missing:
        print(f"[MISSING] {len(missing)} file(s): {missing}")
    out = Path(OUT[name])
    if out.exists():
        m.backup_presentation(str(out))
    prs.save(str(out))
    print(f"[side views {name}] {n} slides, tile PPI {TILE_PPI:.1f} -> {out}")
    return 0 if n and not missing else 1


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(max(build(n) for n in names))
