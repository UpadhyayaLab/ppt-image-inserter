#!/usr/bin/env python
"""
QC decks of the actin bottom and top planes for the 10/29/2026 H3K27me3 sets (FF
2026-10-31, via the TCell session: "prepare QC of actin bottom and top", in the CART
format). Four PLL cells are flagged by actin_below_FOV. FF reviewed the planes and the
filter is now off in both compile configs, so a flagged cell's fate comes from the ledger.

CART's plane QC (morphology/actin_qc/generate_cell_montages.m) is the XZ MIP with the
planes plus the z-profile; the same views are on disk here, so this renders nothing:
  flagged cells   one slide each: XZ (actin + DNA, physical scale, red lines = planes),
                  brightened XZ with the planes labeled, axial profile, raw bottom plane
  all other cells 6 per slide, by condition, in cell-id order

Plane values are read from each well's results.mat and printed on every cell; status
comes from the compile's cell_drop_ledger.csv. A [CHECK] line prints each cell whose
planes disagree with its nucleus. Cells with no row in results.mat are left out.

Usage (base miniforge python):
    python examples_and_configs/build_actin_planes_qc_10292026_decks.py [noco|cilio|all]
"""

import csv
import os
import sys
from pathlib import Path

import numpy as np
import scipy.io as sio
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_h3k27me3_10292026_compare_decks import EXPS, PROG, configure, m  # noqa: E402
from insert_noco_10262026_montage_slides import CYAN, GRAY, ORANGE, RED, WHITE, YELLOW  # noqa: E402

Z_UM = 0.3             # μm per raw slice (params.zstep)
ISO_UM = 0.065         # μm per iso slice (nuc_bottom/top_slice_num)
OUT = {"noco": "Noco/H3K27me3_20261029/Noco_Jurkats_H3K27me3_actin_planes_QC_10292026.pptx",
       "cilio": "CilioD/H3K27me3_20261029/CilioD_Jurkats_H3K27me3_actin_planes_QC_10292026.pptx"}
OUT_ROOT = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous"
SET_NAME = {"noco": "Noco", "cilio": "CilioD"}
# actin_nuc_xz colors: params.actin_xz_colors on Noco, the default orange on Ciliobrevin.
ACTIN_COLOR = {"noco": YELLOW, "cilio": ORANGE}

# Shown first, one slide each (well, cell id, note). The cells flagged by actin_below_FOV
# and the known-wrong plane.
FLAGGED = {
    "noco": [("G1A4_h3k27me3_PLL_dmso_", 16, ""), ("G1A4_h3k27me3_PLL_dmso_", 17, ""),
             ("G1A4_h3k27me3_PLL_dmso_", 18, ""), ("G1A4_h3k27me3_PLL_dmso_", 55, "")],
    "cilio": [("G0A4_h3k27me3_PLL_dmso_", 24, ""),
              ("G0A4_h3k27me3_PLL_dmso_", 7, "Known wrong: bottom plane far above the nucleus bottom")],
}
# cell_drop_ledger.csv Stage -> short status, in the compile's own terms.
STATUS = {"kept": "kept", "actin_below_FOV": "removed: actin below FOV",
          "cent_found": "removed: centrosome not found", "stage0_exclusions": "removed: nucleus mesh has handles",
          "mesh_unusable": "removed: nuclear mesh unusable", "cell_found": "not analyzed",
          "periph_ratio": "removed: periph ratio"}

PER_SLIDE = 6
COLS, ROWS = 3, 2
GRID_TOP = 0.62
LABEL_H = 0.32
LABEL_FONT_PT = 16
FOOT_TOP = 7.05
FOOT_FONT_PT = 14
CELL_W = (m.SLIDE_W - 2 * m.MARGIN) / COLS
CELL_H = (FOOT_TOP - GRID_TOP) / ROWS
TILE_PPI = 200.0       # one scale for every grid tile: the widest tile (850 px) fits CELL_W


def cell_values(folder: str) -> dict:
    """{cell id: fields} from the well's results.mat (raw slices; nucleus in iso slices)."""
    path = m.EXP_ROOT / PROG / folder / m.RESULTS_MAT
    names = ["cell_id", "actin_bottom_slice_num", "actin_top_slice_num", "actin_below_FOV",
             "nuc_bottom_slice_num", "nuc_top_slice_num"]
    r = sio.loadmat(m._winlong(path), squeeze_me=True, variable_names=names)
    ids = np.asarray(r["cell_id"], float).ravel()
    col = {n: np.asarray(r[n], float).ravel() for n in names[1:]}
    return {int(c): {n: float(v[i]) for n, v in col.items()}
            for i, c in enumerate(ids) if not np.isnan(c)}


def ledger(exp: dict) -> dict:
    """{(compile condition, cell id): status text} from cell_drop_ledger.csv.
    Stage 0 holds FF's own exclusions and the automatic mesh check; Detail tells them apart."""
    path = m.EXP_ROOT / exp["compile"] / "cell_drop_ledger.csv"
    with open(m._winlong(path), encoding="utf-8-sig", newline="") as fh:
        return {(r["Condition"], int(r["CellID"])): "removed: FF exclusion" if "excluded_cells.csv" in r["Detail"]
                else STATUS.get(r["Stage"], f"removed: {r['Stage']}")
                for r in csv.DictReader(fh) if r["CellID"] != "NaN"}   # not-analyzed rows have no id


def slices(v: float) -> str:
    return "n/a" if np.isnan(v) else f"{v:g}"


def um(v: float) -> str:
    return "n/a" if np.isnan(v) else f"{v * Z_UM:.1f} μm"


def check(label: str, cid: int, v: dict) -> None:
    """Print a cell whose planes disagree with its nucleus: bottom more than 1 μm above
    the nucleus bottom, or top below the nucleus top."""
    bot, top = v["actin_bottom_slice_num"] * Z_UM, v["actin_top_slice_num"] * Z_UM
    nb, nt = v["nuc_bottom_slice_num"] * ISO_UM, v["nuc_top_slice_num"] * ISO_UM
    if bot > nb + 1:
        print(f"[CHECK] {label} cell {cid}: actin bottom {bot:.1f} μm, nucleus bottom {nb:.1f} μm")
    if top < nt:
        print(f"[CHECK] {label} cell {cid}: actin top {top:.1f} μm, nucleus top {nt:.1f} μm")


def views(folder: str, cid: int) -> dict:
    """The cell's saved QC images; a view missing on disk is None."""
    d = m.EXP_ROOT / PROG / folder
    planes = d / "actin" / "bottom_top_slices"
    xz = m.percell_files(d / "physical_scale_images" / "actin_nuc_xz" / "cells").get(cid)
    raw = m.percell_files(d / "actin" / "bottom_slice_raw").get(cid)   # newest of a rerun

    def found(p):
        return p if os.path.isfile(m._winlong(p)) else None

    return {"xz": xz, "bright": found(planes / "side_views" / f"Cell_{cid}_XZ_MIP_Brightened.png"),
            "axial": found(planes / "Axial_Profile" / f"Cell_{cid}_Normalized_Total_Signal.png"),
            "raw": raw}


def place(slide, path, left: float, top: float, w: float, h: float) -> None:
    """A picture fitted inside the box, top left, keeping its aspect ratio."""
    if path is None:
        return
    with m.Image.open(m._winlong(path)) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    slide.shapes.add_picture(m._winlong(path), Inches(left), Inches(top), width=Inches(iw * scale))


def left_text(slide, text, left, top, w, h, font_pt, color, bold=True):
    box = m.add_textbox(slide, text, left, top, w, h, font_pt=font_pt, color=color, bold=bold)
    box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
    box.text_frame.word_wrap = True
    return box


def flagged_slide(prs, name, label, cid, v, status, note, img, channels) -> None:
    """One cell: XZ views at left, axial profile in the middle, raw bottom plane at right."""
    slide = m.new_slide(prs, f"{SET_NAME[name]} {label}, cell {cid}", channels)
    info = left_text(slide, f"Bottom slice {slices(v['actin_bottom_slice_num'])} "
                     f"({um(v['actin_bottom_slice_num'])})  ·  Top slice "
                     f"{slices(v['actin_top_slice_num'])} ({um(v['actin_top_slice_num'])})  ·  ",
                     m.MARGIN, 0.58, m.SLIDE_W - 2 * m.MARGIN, 0.36, 20, WHITE)
    run = info.text_frame.paragraphs[0].add_run()
    flag = v["actin_below_FOV"] == 1 and "below FOV" not in status
    run.text = status + ("  ·  flagged: actin below FOV" if flag else "")
    run.font.size, run.font.bold = Pt(20), True
    run.font.color.rgb = WHITE if status == "kept" else RED
    place(slide, img["xz"], m.MARGIN, 1.02, 4.05, 3.0)
    if img["xz"] is None:
        left_text(slide, "(no XZ image)", m.MARGIN, 2.3, 4.05, 0.32, 16, GRAY, bold=False)
    place(slide, img["bright"], m.MARGIN, 4.08, 4.05, 3.0)
    place(slide, img["axial"], 4.25, 1.02, 6.2, 4.65)
    left_text(slide, "Bottom plane, actin", 10.5, 1.02, 2.75, 0.32, 16, WHITE)
    place(slide, img["raw"], 10.5, 1.36, 2.73, 3.3)
    text = "Top left: red lines = bottom and top planes."
    left_text(slide, f"{note}. {text}" if note else text, 4.25, 5.75, 8.9, 0.7, 18, GRAY, bold=False)


def grid_slide(prs, name, label, cells, k, n, status, values, channels) -> None:
    """Up to 6 cells: the label over each XZ tile, one physical scale."""
    first, last = cells[0][0], cells[-1][0]
    slide = m.new_slide(prs, f"{SET_NAME[name]} {label}, cells {first}–{last}, {k}/{n}", channels)
    for i, (cid, tile) in enumerate(cells):
        row, col = divmod(i, COLS)
        x, y = m.MARGIN + col * CELL_W, GRID_TOP + row * CELL_H
        v = values[cid]
        box = left_text(slide, f"Cell {cid}: bottom {slices(v['actin_bottom_slice_num'])}, "
                        f"top {slices(v['actin_top_slice_num'])}", x, y, CELL_W, LABEL_H,
                        LABEL_FONT_PT, WHITE)
        if status[cid] != "kept":
            run = box.text_frame.paragraphs[0].add_run()
            run.text = f"  {status[cid]}"
            run.font.size, run.font.bold = Pt(LABEL_FONT_PT - 2), True
            run.font.color.rgb = RED
        if tile is None:
            left_text(slide, "(no XZ image)", x, y + 1.2, CELL_W, 0.3, 14, GRAY, bold=False)
            continue
        with m.Image.open(m._winlong(tile)) as im:
            slide.shapes.add_picture(m._winlong(tile), Inches(x), Inches(y + LABEL_H),
                                     width=Inches(min(im.size[0] / TILE_PPI, CELL_W - 0.05)))
    left_text(slide, f"Plane values are slices ({Z_UM} μm each). Red lines = bottom and top planes.",
              m.MARGIN, FOOT_TOP, m.SLIDE_W - 2 * m.MARGIN, 0.34, FOOT_FONT_PT, GRAY, bold=False)


def build(name: str) -> int:
    exp = EXPS[name]
    configure(exp)
    drops = ledger(exp)
    channels = [("Actin", ACTIN_COLOR[name]), ("DNA", CYAN)]
    prs = Presentation()
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    wells = {w: (lab, cond) for w, lab, cond, _ in exp["wells"]}
    values = {w: cell_values(w) for w in wells}
    n = 0
    for well, cid, note in FLAGGED[name]:
        lab, cond = wells[well]
        status = drops.get((cond, cid), "not in ledger")
        flagged_slide(prs, name, lab, cid, values[well][cid], status, note, views(well, cid), channels)
        n += 1
    done = {(w, c) for w, c, _ in FLAGGED[name]}
    for well, (lab, cond) in wells.items():
        files = m.percell_files(m.EXP_ROOT / PROG / well / "physical_scale_images" / "actin_nuc_xz" / "cells")
        for cid, v in sorted(values[well].items()):
            check(lab, cid, v)
        cells = [(c, files.get(c)) for c in sorted(values[well]) if (well, c) not in done]
        status = {c: drops.get((cond, c), "not in ledger") for c, _ in cells}
        no_tile = [c for c, t in cells if t is None]
        print(f"  {name} {lab}: {len(cells)} cells on grid slides, no XZ image: {no_tile}")
        pages = [cells[i:i + PER_SLIDE] for i in range(0, len(cells), PER_SLIDE)]
        for k, page in enumerate(pages, start=1):
            grid_slide(prs, name, lab, page, k, len(pages), status, values[well], channels)
            n += 1
    out = Path(f"{OUT_ROOT}/{OUT[name]}")
    if out.exists():
        m.backup_presentation(str(out))
    prs.save(str(out))
    print(f"[actin planes QC {name}] {n} slides -> {out}")
    return n


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(0 if all(build(n) for n in names) else 1)
