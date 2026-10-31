#!/usr/bin/env python
"""
Actin bottom and top plane QC for the 10/29/2026 H3K27me3 sets, in the format of the CAR T
actin montage deck (FF 2026-10-31: "look at my car t ppt", option 2:
CART/actin_only/CART_actin_montages_20260717_20260723_20261017.pptx, built by
CART/CART_actin_ppt/build_actin_decks.py): white slides, two conditions side by side,
one montage sheet each, one px/in per view across the deck.

Views, all from saved outputs (no MATLAB):
  Actin at Synapse         actin/bottom_slice_raw/montages (the bottom plane)
  Actin Synapse Mask       actin/bottom_slice_seg/montages
  Actin XZ, four contrasts  <EXP>/actin_planes_qc_20261031b thick-line sheets (TCell session),
                            first in the deck; these replaced the thin-line XZ views
  Actin Axial Profile       no sheets on disk: actin/bottom_top_slices/Axial_Profile plots,
                            tiled 5 x 5 here (cell-id order) beside the deck exports in
                            <EXP>/deck_png_20261031/actin_planes_montages

The sheets hold every cell the pipeline saw, including removed ones. Every page of every
well is shown; a condition with fewer pages leaves its box empty.

Usage (base miniforge python):
    python examples_and_configs/build_actin_planes_montage_decks_10292026.py [noco|cilio|all]
"""

import os
import re
import sys
from pathlib import Path

from PIL import Image
from pptx.enum.text import PP_ALIGN

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "CART" / "CART_pMLC_ppt"))
import build_pmlc_profiles_deck as bpd  # noqa: E402  (shared Deck / add_text)
import insert_CART_CatB_MT_actin_nuc_montages_slides as mont  # noqa: E402  (one-scale placement)
from build_actin_planes_qc_10292026_decks import ledger  # noqa: E402
from build_h3k27me3_10292026_compare_decks import EXPS, PROG, configure, m  # noqa: E402

DASH = "\u2014"
OUT_ROOT = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous"
OUT = {"noco": "Noco/H3K27me3_20261029/Noco_Jurkats_H3K27me3_actin_planes_thick_lines_10292026.pptx",
       "cilio": "CilioD/H3K27me3_20261029/CilioD_Jurkats_H3K27me3_actin_planes_thick_lines_10292026.pptx"}
SET_NAME = {"noco": "Noco 10/29", "cilio": "CilioD 10/29"}
# Conditions shown together (indices into EXPS wells); CAT / FMC in the CAR T deck.
PAIRS = {"noco": [(0, 1), (2, 3)], "cilio": [(0, 1), (2,)]}
# Thick-line side views from the TCell session (FF: "the lines here are so thin"), first in the deck.
QC_ROOT = "actin_planes_qc_20261031b"
THICK = [("Actin XZ, Auto Contrast (per cell, 20 to 99.5 percentile)", "actin_planes_auto"),
         ("Actin XZ, Bright (per cell, 0 to max / 4)", "actin_planes_bright"),
         ("Actin XZ, Fixed Contrast (all cells, 15 to 100)", "actin_planes_fixed"),
         ("Actin XZ, Faint Signal (all cells, 10 to 35)", "actin_planes_faint")]
VIEWS = [("Actin at Synapse", "actin/bottom_slice_raw/montages"),
         ("Actin Synapse Mask", "actin/bottom_slice_seg/montages"),
         ("Actin Axial Profile", ("Axial_Profile", r"Cell_(\d+)_Normalized_Total_Signal\.png"))]
# Cells to look at first, after going through every sheet by eye (FF: "look through full PPT when
# constructed and put problematic ones at front"): (what is wrong, [(well index, cell id), ...]).
PROBLEMS = {"noco": [("Bottom Line on Slice 1", [(0, 16), (0, 17), (0, 18), (0, 55)]),
                     ("Top Line Below Cell Top", [(1, 7), (1, 10), (1, 16), (1, 21), (1, 41), (1, 43),
                                                  (2, 5), (2, 17), (2, 19), (2, 24), (3, 28)]),
                     ("Cell Reaches Stack Top", [(0, 22), (0, 53), (1, 4), (1, 36), (1, 42), (2, 0), (2, 3), (2, 18),
                                                 (2, 21), (2, 30), (2, 34), (2, 36), (2, 48), (2, 51), (2, 57),
                                                 (3, 10), (3, 13), (3, 29), (3, 45), (3, 46), (3, 47), (3, 54)]),
                     ("Mask Has Two Lobes or a Gap", [(0, 37), (2, 52), (2, 54), (3, 52), (3, 54)]),
                     ("Mask Has an Unusual Shape", [(2, 56)])],
            "cilio": [("Bottom Line Suspect", [(0, 7), (0, 24)]),
                      ("Top Line Below Cell Top", [(0, 34), (1, 3), (1, 7), (1, 13), (1, 24)]),
                      ("Cell Reaches Stack Top", [(0, 27), (2, 17), (2, 18), (2, 20), (2, 23), (2, 28), (2, 33),
                                                  (2, 36), (2, 39), (2, 40), (2, 42)]),
                      ("Mask Has Two Lobes or a Gap", [(0, 42), (1, 28), (2, 9), (2, 13)]),
                      ("Mask Has an Unusual Shape", [(0, 44), (0, 49)]),
                      ("Mask Misses a Faint Edge or Fragment", [(2, 17), (2, 20)])]}
PROBLEM_COLS = [("Auto", "actin_planes_auto"), ("Bright", "actin_planes_bright"), ("Faint", "actin_planes_faint")]
PROBLEM_ROWS = 4
PROBLEM_LEFT = 1.3     # room for the condition, cell id and status
PROBLEM_TOP = 1.0
TILE_COLS = TILE_ROWS = 5
TILE_PX = (438, 328)  # half size of the saved 875 x 656 plots
# Boxes as in build_actin_decks.py (MONT_TOP, MONT_HALF, MONT_H).
MONT_TOP = 1.05
MONT_HALF = (bpd.SLIDE_W - 0.3) / 2
MONT_H = bpd.SLIDE_H - 1.12


def sheets(folder: Path) -> list:
    """Montage sheets in a folder, by first cell id."""
    if not os.path.isdir(mont._winlong(folder)):
        return []
    names = [n for n in os.listdir(mont._winlong(folder)) if re.fullmatch(r"montage_cells_\d+_\d+\.png", n)]
    return [folder / n for n in sorted(names, key=lambda n: int(n.split("_")[2]))]


def tiled_sheets(root: Path, well: str, sub: str, pattern: str) -> list:
    """Tile the well's per-cell plots in actin/bottom_top_slices/<sub> into sheets, cell-id order."""
    src = root / PROG / well / "actin" / "bottom_top_slices" / sub
    out_dir = root / "deck_png_20261031" / "actin_planes_montages" / well / sub
    os.makedirs(mont._winlong(out_dir), exist_ok=True)
    files = {int(m.group(1)): n for n in os.listdir(mont._winlong(src)) if (m := re.fullmatch(pattern, n))}
    ids = sorted(files)
    n_per = TILE_COLS * TILE_ROWS
    out = []
    for k in range(0, len(ids), n_per):
        page = ids[k:k + n_per]
        sheet = Image.new("RGB", (TILE_COLS * TILE_PX[0], TILE_ROWS * TILE_PX[1]), "white")
        for i, cid in enumerate(page):
            with Image.open(mont._winlong(src / files[cid])) as im:
                tile = im.convert("RGB").resize(TILE_PX, Image.LANCZOS)
            row, col = divmod(i, TILE_COLS)
            sheet.paste(tile, (col * TILE_PX[0], row * TILE_PX[1]))
        path = out_dir / f"montage_cells_{page[0]}_{page[-1]}.png"
        sheet.save(mont._winlong(path))
        out.append(path)
    return out


def problem_slides(deck, name: str, root: Path, wells: list, conds: list, status: dict) -> None:
    """Problem cells first: one row per cell, three contrasts and the synapse mask, one scale.
    Each row's status comes from the compile's ledger, so a recompile cannot leave it stale."""
    w = (bpd.SLIDE_W - PROBLEM_LEFT - 0.1) / 4
    h = (bpd.SLIDE_H - PROBLEM_TOP - 0.05) / PROBLEM_ROWS
    segs = {i: m.percell_files(root / PROG / well / "actin" / "bottom_slice_seg")   # newest of a rerun
            for i, (well, _) in enumerate(wells)}

    def tile(i, cid, sub):
        return root / QC_ROOT / wells[i][0] / "physical_scale_images" / sub / "cells" / f"Cell_{cid}_xz.png"

    every = [c for _, cells in PROBLEMS[name] for c in cells]
    xz_ppi = mont.compute_group_ppi([tile(i, cid, sub) for i, cid in every for _, sub in PROBLEM_COLS], w, h)
    seg_ppi = mont.compute_group_ppi([segs[i][cid] for i, cid in every], w, h)
    for what, cells in PROBLEMS[name]:
        pages = [cells[i:i + PROBLEM_ROWS] for i in range(0, len(cells), PROBLEM_ROWS)]
        for k, page in enumerate(pages, start=1):
            s = deck.slide(f"To Check: {what} {DASH} {SET_NAME[name]}, {k}/{len(pages)}")
            for c, text in enumerate([t for t, _ in PROBLEM_COLS] + ["Synapse Mask"]):
                bpd.add_text(s, PROBLEM_LEFT + c * w, 0.62, w, 0.36, text, 18, bold=True, color=bpd.BLACK,
                             align=PP_ALIGN.CENTER)
            for r, (i, cid) in enumerate(page):
                well, label = wells[i]
                top = PROBLEM_TOP + r * h
                bpd.add_text(s, 0.05, top + 0.1, PROBLEM_LEFT - 0.05, 1.45,
                             f"{label}, cell {cid}, {status.get((conds[i], cid), 'not in ledger')}", 14, bold=True,
                             color=bpd.BLACK, align=PP_ALIGN.LEFT)
                paths = [tile(i, cid, sub) for _, sub in PROBLEM_COLS] + [segs[i][cid]]
                for c, path in enumerate(paths):
                    mont.add_image_at_ppi(s, path, xz_ppi if c < 3 else seg_ppi, PROBLEM_LEFT + c * w, top, w, h)


def build(name: str) -> int:
    exp = EXPS[name]
    root = Path(exp["root"])
    wells = [(w, lab) for w, lab, _, _ in exp["wells"]]
    conds = [c for _, _, c, _ in exp["wells"]]
    configure(exp)
    status = ledger(exp)
    pages = {}
    views = THICK + VIEWS
    for stem, sub in THICK:
        for w, _ in wells:
            pages[stem, w] = sheets(root / QC_ROOT / w / "physical_scale_images" / sub / "montages")
    for stem, sub in VIEWS:
        for w, _ in wells:
            pages[stem, w] = (sheets(root / PROG / w / Path(sub)) if isinstance(sub, str)
                              else tiled_sheets(root, w, *sub))
        missing = [w for w, _ in wells if not pages[stem, w]]
        if missing:
            print(f"[MISSING] {stem}: {missing}")
    ppi = {}
    for stem, _ in views:
        pngs = [p for w, _ in wells for p in pages[stem, w]]
        ppi[stem] = mont.compute_group_ppi(pngs, MONT_HALF, MONT_H)
        print(f"  {stem}: {len(pngs)} sheets, {ppi[stem]:.1f} px/in")
    bpd.OUT = Path(f"{OUT_ROOT}/{OUT[name]}")
    deck = bpd.Deck()
    problem_slides(deck, name, root, wells, conds, status)
    for stem, _ in views:
        n_pages = sum(max(len(pages[stem, wells[i][0]]) for i in pair) for pair in PAIRS[name])
        k = 0
        for pair in PAIRS[name]:
            for p in range(max(len(pages[stem, wells[i][0]]) for i in pair)):
                k += 1
                s = deck.slide(f"{stem} {DASH} {SET_NAME[name]}, {k}/{n_pages}")
                for col, i in enumerate(pair):
                    well, label = wells[i]
                    left = 0.1 + col * (MONT_HALF + 0.1)
                    bpd.add_text(s, left, 0.62, MONT_HALF, 0.4, label, 20, bold=True, color=bpd.BLACK,
                                 align=PP_ALIGN.CENTER)
                    if p < len(pages[stem, well]):
                        mont.add_image_at_ppi(s, pages[stem, well][p], ppi[stem], left, MONT_TOP,
                                              MONT_HALF, MONT_H)
    deck.save()
    print(f"[actin planes montages {name}] {len(deck.prs.slides)} slides -> {bpd.OUT}")
    return len(deck.prs.slides)


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(0 if all(build(n) for n in names) else 1)
