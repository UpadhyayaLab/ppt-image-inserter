"""
insert_vimkd_centrosome_slide.py

One-slide deck for FF's advisor (2026-10-04): VimKD fixed cells, siCtrl vs siVim, PERICENTRIN
experiments only, data from the old vim-paper pipeline (original Aug 2024 run), plotted with the
TCell by-day plotters:
  left   centrosome-synapse distance, one violin pair per experiment
  right  % polarized (< 2 um) per experiment
plus one small line summarizing the beta-tubulin experiments (different pipeline), computed from
the earlier 7-experiment run's stats (BH across those 4 experiments, from the raw rank-sum p).

Style = the MCF7 channel-misalignment group-meeting deck (16:9), as FF adjusted it: 48 pt bold
title in FF's wording, no footer line.

Panels and stats come from
  fixed_cells_general_20240801_vim_paper/vimkd_centrosome_by_day_20261004/gen_vimkd_cent_pericentrin_by_day.m
  (tubulin line: .../gen_vimkd_cent_by_day_oldpipeline.m run of 2026-10-04)

Usage:
    conda run -n PPT_editing python examples_and_configs/insert_vimkd_centrosome_slide.py
"""

import csv
import os
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import insert_channel_misalignment_slides as CM

BASE = Path("M:/FF/FF_4TB_2_Backup_fullHD/Vimentin_Project_2ndharddrive/VimentinKD_NucleusData_Fixed/"
            "results_compilation")
RUN = BASE / "VimKD_pericentrin_oldpipeline_by_day_20261004"
PANELS = [RUN / "centrosome_center_z_rel_bottom_actin_plane_by_day.png",
          RUN / "polarized_fraction_by_day.png"]
STATS = RUN / "stats.csv"
TUB_STATS = (BASE / "VimKD_cent_by_day_pericentrin_old_tubulin_new_20261004" / "by_day_panels"
             / "centrosome_center_z_rel_bottom_actin_plane_by_day_stats.csv")
TUB_DAYS = ["May 23, 2022", "Oct 24, 2023", "Jan 17, 2024", "Jan 29, 2024"]

OUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/VimentinKD")   # FF: VimKD PI decks live here
OUT = OUT_DIR / "VimKD_centrosome_polarization_20261004.pptx"
TITLE_PT = 48   # FF 2026-10-04: "somewhat larger" than the 40 pt the deck's title_pt() gave
TITLE_H = 1.65  # two lines at 48 pt
TITLE = "Vimentin knockdown may slightly impair centrosome polarization"   # FF 2026-10-04
NOTE_PT, NOTE_H = 20, 0.45
GAP = 0.25


def bh(p):
    m = len(p)
    order = sorted(range(m), key=lambda i: p[i])
    adj = [0.0] * m
    prev = 1.0
    for rank in range(m, 0, -1):
        i = order[rank - 1]
        prev = min(prev, p[i] * m / rank)
        adj[i] = prev
    return adj


def tubulin_note():
    with open(TUB_STATS, newline="", encoding="utf-8") as fh:
        rows = [r for r in csv.DictReader(fh) if r["day"] in TUB_DAYS]
    assert len(rows) == 4, f"expected the 4 tubulin days in {TUB_STATS}, found {len(rows)}"
    higher = sum(float(r["median_siVim"]) > float(r["median_siCtrl"]) for r in rows)
    tied = sum(float(r["median_siVim"]) == float(r["median_siCtrl"]) for r in rows)
    sig = sum(q < 0.05 for q in bh([float(r["p_ranksum"]) for r in rows]))
    print("tubulin:", [(r["day"], r["median_siCtrl"], r["median_siVim"], r["p_ranksum"]) for r in rows])
    tie_txt = f" ({tied} tied)" if tied else ""
    sig_txt = "n.s." if sig == 0 else f"p<0.05 in {sig}"
    return (f"\u03b2-tubulin-based experiments (different pipeline): siVim higher in {higher} of 4"
            f"{tie_txt}, {sig_txt}")


def main():
    with open(STATS, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            print(r)
    note = tubulin_note()
    print(note)

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(CM.W), Inches(CM.H)
    s = CM.M._new_slide(prs)
    CM.M.add_textbox(s, TITLE, CM.MARGIN, CM.TITLE_TOP, CM.IMG_W, TITLE_H,
                     font_pt=TITLE_PT, color=CM.BLACK, bold=True)
    s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # two panels side by side, same height, centered as a pair
    src = OUT_DIR / "source_images"
    src.mkdir(parents=True, exist_ok=True)
    imgs, aspects = [], []
    for p in PANELS:
        out = src / p.name
        CM.trim(Image.open(p)).save(out)
        with Image.open(out) as im:
            aspects.append(im.width / im.height)
        imgs.append(out)
    top = CM.TITLE_TOP + TITLE_H + 0.08
    box_h = CM.H - top - CM.MARGIN   # FF 2026-10-04: tubulin note removed from the slide (still printed above)
    h = min(box_h, (CM.IMG_W - GAP) / sum(aspects))
    widths = [h * a for a in aspects]
    left = CM.MARGIN + (CM.IMG_W - sum(widths) - GAP) / 2
    for img, w in zip(imgs, widths):
        s.shapes.add_picture(str(img), Inches(left), Inches(top), Inches(w), Inches(h))
        left += w + GAP

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        CM.M.backup_presentation(str(OUT), backup_base=str(OUT_DIR / "backups"))
    prs.save(str(OUT))
    print(f"Done -> {OUT}")


if __name__ == "__main__":
    main()
