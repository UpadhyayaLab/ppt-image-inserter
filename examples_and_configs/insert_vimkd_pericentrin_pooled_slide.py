"""
insert_vimkd_pericentrin_pooled_slide.py

PI deck slides 3-4 (FF / AU 2026-10-04): VimKD fixed cells, pericentrin experiments, square paper-mode
panels, one slide per group (FF: four panels on one slide were too thin):
  slide A  Pooled, 3 experiments: [violin | % polarized]
  slide B  Apr 6, 2022:           [violin | % polarized]
from fixed_cells_general_20240801_vim_paper/vimkd_centrosome_by_day_20261004/gen_vimkd_cent_pericentrin_paper_panels.m
(old vim-paper pipeline, 0.75 cell-bottom threshold run). Same style as the other slides
(16:9, 48 pt bold title, no footer).

Usage:
    conda run -n PPT_editing python examples_and_configs/insert_vimkd_pericentrin_pooled_slide.py
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

RUN = Path("M:/FF/FF_4TB_2_Backup_fullHD/Vimentin_Project_2ndharddrive/VimentinKD_NucleusData_Fixed/"
           "results_compilation/VimKD_pericentrin_pooled_and_apr6_paper_panels_20261004")
GROUPS = [("apr6", "Fixed cells, pericentrin: Apr 6, 2022"),             # FF: Apr 6 alone first
          ("may20", "Fixed cells, pericentrin: May 20, 2022"),           # PI 2026-10-04: other two
          ("jun28", "Fixed cells, pericentrin: Jun 28, 2022"),           #   replicates, slides 4-5
          ("pooled", "Fixed cells, pericentrin: pooled, 3 experiments")]  # draft titles; FF to edit
OUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/VimentinKD")
OUT = OUT_DIR / "VimKD_pericentrin_pooled_and_apr6_20261004.pptx"
TITLE_PT, TITLE_H, GAP = 48, 1.65, 0.6


def main():
    with open(RUN / "stats.csv", newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            print(r)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(CM.W), Inches(CM.H)
    src = OUT_DIR / "source_images"
    src.mkdir(parents=True, exist_ok=True)
    for tag, title in GROUPS:
        s = CM.M._new_slide(prs)
        CM.M.add_textbox(s, title, CM.MARGIN, CM.TITLE_TOP, CM.IMG_W, TITLE_H,
                         font_pt=TITLE_PT, color=CM.BLACK, bold=True)
        s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        pair = []
        for kind in ("violin", "polarized"):
            p = RUN / f"{tag}_{kind}.png"
            out = src / f"pericentrin_{p.name}"
            CM.trim(Image.open(p)).save(out)
            with Image.open(out) as im:
                pair.append((out, im.width / im.height))
        img_top = CM.TITLE_TOP + TITLE_H + 0.08
        h = min(CM.H - img_top - CM.MARGIN, (CM.IMG_W - GAP) / sum(a for _, a in pair))
        left = CM.MARGIN + (CM.IMG_W - h * sum(a for _, a in pair) - GAP) / 2
        for img, a in pair:
            s.shapes.add_picture(str(img), Inches(left), Inches(img_top), Inches(h * a), Inches(h))
            left += h * a + GAP

    if OUT.exists():
        CM.M.backup_presentation(str(OUT), backup_base=str(OUT_DIR / "backups"))
    prs.save(str(OUT))
    print(f"Done -> {OUT}")


if __name__ == "__main__":
    main()
