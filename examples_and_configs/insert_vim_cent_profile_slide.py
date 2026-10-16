"""
insert_vim_cent_profile_slide.py

One-slide deck for FF's advisor (2026-10-04): vimentin radial profile around the
centrosome, 3 vs 12 min, two panels side by side:
  left  = aCD3 3 vs 12 min (Figure 6B, formerly 5B)
  right = PLL 3 vs 12 min (supplement)

Style = insert_vimkd_centrosome_slide.py (the siVim PI slide): the MCF7
channel-misalignment deck's layout, no footer line, 48 pt bold title, images
trimmed of white margins. Two columns as in CM.pair_slide (one bold label per column).

Panels come from
  fixed_cells_general_20240801_vim_paper/vim_paper_20260629/gen_fig6_vim_cent_profile.m
n per condition is read from that run's SOURCES.txt (never typed).

Usage:
    conda run -n PPT_editing python examples_and_configs/insert_vim_cent_profile_slide.py
"""

import os
import re
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import insert_channel_misalignment_slides as CM

RUN = Path("K:/FF/Papers/vimentin_paper/Figures/postprocessing_output/Figure6/B_vim_cent_profile")
PANELS = [("B_aCD3_3_vs_12min.png", "aCD3 (Figure 6B)"),
          ("S_PLL_3_vs_12min.png", "PLL (supplement)")]

OUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/VimentinKD")   # FF 2026-10-04: the existing VimKD folder
OUT = OUT_DIR / "Vim_centrosome_profile_20261004.pptx"
TITLE_PT = 40   # FF 2026-10-04: smaller than the siVim slide's 48 pt; 40 keeps the forced two-line title on two lines
TITLE_H = 1.65  # two lines at 48 pt
LABEL_PT = 28
LABEL_H = 0.55
GAP = 0.3


def n_text():
    """'n = a / b' per panel from SOURCES.txt, in the order the generator wrote it."""
    src = (RUN / "SOURCES.txt").read_text(encoding="utf-8")
    out = {}
    for tag, n in re.findall(r"^(\S+)\s+.+?\s+n=(\d+)", src, flags=re.M):
        out.setdefault(tag, []).append(n)
    return out


def main():
    ns = n_text()
    print(ns)
    # vertical tab = line break inside one PowerPoint paragraph (break before the timepoints)
    title = "Vimentin is enriched around the centrosome\vat 3 and 12 min in activated and non-activated cells"

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(CM.W), Inches(CM.H)
    s = CM.M._new_slide(prs)
    CM.M.add_textbox(s, title, CM.MARGIN, CM.TITLE_TOP, CM.IMG_W, TITLE_H,
                     font_pt=TITLE_PT, color=CM.BLACK, bold=True)
    s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for para in s.shapes[-1].text_frame.paragraphs:   # the helper formats only the first run;
        for r in para.runs:                           # the line break starts a new one
            r.font.size, r.font.bold, r.font.color.rgb = Pt(TITLE_PT), True, CM.BLACK

    top = CM.TITLE_TOP + TITLE_H + 0.08
    col_w = (CM.IMG_W - GAP) / 2
    img_h = CM.H - top - LABEL_H - CM.MARGIN
    side = min(col_w, img_h)
    x0 = CM.MARGIN + (CM.IMG_W - (2 * side + GAP)) / 2   # center the pair
    src_dir = OUT_DIR / "source_images" / OUT.stem
    src_dir.mkdir(parents=True, exist_ok=True)
    for i, (png, lab) in enumerate(PANELS):
        tag = Path(png).stem
        n = ns.get(tag, [])
        label = f"{lab}: n = {' / '.join(n)}" if n else lab
        left = x0 + i * (side + GAP)
        CM.M.add_textbox(s, label, left, top, side, LABEL_H, font_pt=LABEL_PT, color=CM.BLACK, bold=True)
        img = src_dir / png
        CM.trim(Image.open(RUN / "for_paper" / png)).save(img)
        CM.M.add_image_in_box(s, str(img), left, top + LABEL_H, side, img_h)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        CM.M.backup_presentation(str(OUT), backup_base=str(OUT_DIR / "backups"))
    prs.save(str(OUT))
    print(f"Done -> {OUT}")


if __name__ == "__main__":
    main()
