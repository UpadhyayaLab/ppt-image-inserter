#!/usr/bin/env python3
r"""build_pmlc_paper_figure.py  Assemble the pMLC component of Figure 1 as ONE
extensive slide and graft it onto a COPY of the CAR-T draft deck.

The original draft (CAR T cell_short_figures_draft_20260827.pptx) is only read +
copied; a new dated copy gets the appended figure slide. An existing same-day output
is backed up (ppt_image_inserter.backup_presentation) before the copy replaces it.

The slide (portrait, matched to the draft's 7.5 x 10.83 in page) is data-driven:
  Panel A  representative micrographs, matched contrast -- CAT row over FMC row,
           columns actin (green) / pMLC (magenta) / merge, from paper_reps_*.
  Panels B..  pooled CAT-vs-FMC quantification from pooled_paper_panels_* (the
           MATLAB run tests/run_pmlc_pooled_paper_panels.m): radial profile +
           gradient, per-region MFI-norm violins (10 min), and pooled timecourses.
Everything is FF-swappable: change REPS/QUANT below. Panels whose PNG is not yet
on disk are drawn as a labelled placeholder box, so the deck builds before AND
after the pooled run (re-run to fill them in).

The pooled and rep folders are auto-resolved to the NEWEST matching dir, so this
works whatever date the MATLAB / rep steps ran.

Output: K:\FF\PPT\PPT_autogeneration\CART\pMLC\CAR T cell_short_figures_draft_<yyyymmdd>.pptx
Run:    python build_pmlc_paper_figure.py
"""
from __future__ import annotations

import datetime
import shutil
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# This repo's package for the tiered backup (was a copy under PPT_editing\code whose
# failed import silently disabled backups; bug hunt 2026-10-24).
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from ppt_image_inserter import backup_presentation  # noqa: E402
sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_pmlc_profiles_deck import newest_run, one_match  # noqa: E402  (shared run / rep pickers)

NAS = Path(r'\\10.225.56.176\ULab_data\User_data\Kiet\results_compiled\pmlc_joint')
PPT_DIR = Path(r'K:\FF\PPT\PPT_autogeneration\CART\pMLC')
DRAFT = PPT_DIR / 'CAR T cell_short_figures_draft_20260827.pptx'
OUT = PPT_DIR / f'CAR T cell_short_figures_draft_{datetime.date.today():%Y%m%d}.pptx'
# pMLC display range of the rep panels: tests/run_pmlc_rep_render.m sets ropts.pmlc_clim =
# [0 500], overriding the dataset config's fixed_clim [0 1000] (FF 2026-09-14). Keep in sync.
PMLC_CLIM = (0, 500)

BLACK = RGBColor(0, 0, 0); GREY = RGBColor(0x55, 0x55, 0x55)
GREEN = RGBColor(0x11, 0x99, 0x11); MAGENTA = RGBColor(0xAA, 0x11, 0xAA)


def newest(pattern):
    return newest_run(NAS, pattern)


REPS = newest('paper_reps_*')
POOLED = newest('pooled_paper_panels_*')

# Representative cells shown (tag -> row label). Files found by glob in REPS.
REP_ROWS = [('CAT_10min', 'CAT (low-affinity)'), ('FMC_10min', 'FMC (high-affinity)')]
REP_COLS = [('actin', 'Actin', GREEN), ('pmlc', 'pMLC', MAGENTA), ('merge', 'Merge', BLACK)]

# Pooled quantification panels: (pooled PNG filename, letter, caption).
# FF 2026-09-14: profiles as the green actin / magenta pMLC overlays with the
# central/peripheral/distal zones marked (combined\ subfolder, shared y-axis).
# The blue/red CAT-vs-FMC single-axis profiles, synapse MFI and inner-outer
# timecourses remain in pooled_paper_panels_* for the supplement.
QUANT = [
    # combined panels carry their own title (condition + Pooled) -> caption blank
    ('combined/combined_profile_CAT_10_min_Pooled_.png',        'B', ''),
    ('combined/combined_profile_FMC_10_min_Pooled_.png',        'C', ''),
    ('pMLC_synapse_peripheral_MFI_norm_pooled_10min.png',       'D', 'Peripheral pMLC (10 min)'),
    ('combined/combined_gradient_CAT_10_min_Pooled_.png',       'E', ''),
    ('combined/combined_gradient_FMC_10_min_Pooled_.png',       'F', ''),
    ('pMLC_synapse_distal_MFI_norm_pooled_10min.png',           'G', 'Distal pMLC (10 min)'),
    ('pMLC_synapse_central_MFI_norm_pooled_10min.png',          'H', 'Central pMLC (10 min)'),
    ('pMLC_synapse_peripheral_MFI_norm_pooled_timecourse.png',  'I', 'Peripheral pMLC (5/10/15 min)'),
    # CAT-normalized, not raw: raw synapse MFI tracks staining batch (CAT 10 min medians
    # differ ~2.3x across the 3 days), so pooling it showed **** at 5 min where the
    # normalized data are n.s. (plot_pmlc_pooled_paper_panels.m; bug hunt 2026-10-24).
    # (the panel's own y-axis reads 'pMLC MFI / CAT Median'; a longer caption wraps into it)
    ('pMLC_synapse_MFI_normCAT_pooled_timecourse.png',          'J', 'Synapse pMLC (5/10/15 min)'),
]


def add_text(slide, left, top, width, height, text, size, *, bold=False,
             color=BLACK, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
    return tb


def fit_in_box(slide, png, left, top, box_w, box_h):
    png = Path(png)
    if not png.is_file():
        return False
    with Image.open(png) as im:
        iw, ih = im.size
    ar = iw / ih
    if ar >= box_w / box_h:
        w = box_w; h = w / ar
    else:
        h = box_h; w = h * ar
    slide.shapes.add_picture(str(png), Inches(left + (box_w - w) / 2),
                             Inches(top + (box_h - h) / 2), width=Inches(w))
    return True


def placeholder(slide, left, top, w, h, label):
    add_text(slide, left, top + h / 2 - 0.3, w, 0.6,
             ['[ pending pooled run ]', label], 10, color=GREY, align=PP_ALIGN.CENTER)


def rep_file(tag, chan):
    return one_match(REPS, f'rep_*_{tag}_*_{chan}.png')


def main():
    if not DRAFT.is_file():
        print(f'[error] draft not found: {DRAFT}')
        return 1
    # Back up an existing same-day OUT BEFORE the draft copy overwrites it (it held FF's
    # hand edits; backing up after the copy only ever saved the unedited draft).
    if OUT.exists():
        try:
            backup_presentation(str(OUT), backup_base=str(PPT_DIR / 'backups'))
            print(f'tiered backup of the existing {OUT.name} written')
        except Exception as e:
            print(f'[error] backup of the existing {OUT} failed ({e}); not overwriting it')
            return 1
    shutil.copy2(DRAFT, OUT)
    print(f'copied draft -> {OUT}')

    prs = Presentation(str(OUT))
    W = prs.slide_width / 914400.0
    H = prs.slide_height / 914400.0
    s = prs.slides.add_slide(prs.slide_layouts[6])
    print(f'appended figure slide (page {W:.2f} x {H:.2f} in); reps={REPS}; pooled={POOLED}')

    m = 0.3
    add_text(s, m, 0.1, W - 2 * m, 0.5,
             'pMLC Distribution at the CAR-T Synapse',
             15, bold=True, align=PP_ALIGN.CENTER)

    # ---- Panel A: representative micrographs (CAT row over FMC row) ----------
    a_top = 0.75
    lab_w = 0.95
    grid_left = m + lab_w
    grid_w = (W - m) - grid_left
    col_w = grid_w / len(REP_COLS)
    row_h = 1.55
    add_text(s, m, a_top - 0.28, 0.4, 0.3, 'A', 16, bold=True)
    # column headers (channel names, coloured)
    for ci, (chan, clab, ccol) in enumerate(REP_COLS):
        add_text(s, grid_left + ci * col_w, a_top - 0.28, col_w, 0.25, clab, 12,
                 bold=True, color=ccol, align=PP_ALIGN.CENTER)
    for ri, (tag, rlab) in enumerate(REP_ROWS):
        ry = a_top + ri * row_h
        add_text(s, m, ry + row_h / 2 - 0.2, lab_w, 0.5, rlab, 11, bold=True)
        for ci, (chan, clab, ccol) in enumerate(REP_COLS):
            p = rep_file(tag, chan)
            if not (p and fit_in_box(s, p, grid_left + ci * col_w + 0.03, ry + 0.03,
                                     col_w - 0.06, row_h - 0.06)):
                placeholder(s, grid_left + ci * col_w, ry, col_w, row_h, f'{rlab} {clab}')
    add_text(s, grid_left, a_top + 2 * row_h - 0.02, grid_w, 0.22,
             f'Jun 07, 10 min; pMLC matched contrast (CLim {PMLC_CLIM[0]}-{PMLC_CLIM[1]}); scale bar 5 '
             '\u03bcm. Representative; swappable.', 8, color=GREY)

    # ---- Panels B..I: pooled CAT-vs-FMC quantification (3-col flow) ----------
    q_top = a_top + 2 * row_h + 0.28
    ncol = 3
    nrow = -(-len(QUANT) // ncol)
    cell_w = (W - 2 * m) / ncol
    cell_h = (H - q_top - 0.15) / nrow
    for i, (fname, letter, cap) in enumerate(QUANT):
        r, c = divmod(i, ncol)
        lx = m + c * cell_w
        ly = q_top + r * cell_h
        add_text(s, lx, ly - 0.02, 0.3, 0.25, letter, 13, bold=True)
        if cap:
            add_text(s, lx + 0.3, ly - 0.02, cell_w - 0.3, 0.25, cap, 9, bold=True, color=GREY)
        p = POOLED / fname if POOLED else None
        if not (p and fit_in_box(s, p, lx + 0.05, ly + 0.24, cell_w - 0.1, cell_h - 0.3)):
            placeholder(s, lx, ly + 0.24, cell_w, cell_h - 0.3, cap)

    # text boxes to front
    tree = s.shapes._spTree
    for sh in list(s.shapes):
        if sh.has_text_frame:
            el = sh._element; tree.remove(el); tree.append(el)

    prs.save(str(OUT))
    print(f'saved {OUT} ({len(prs.slides)} slides; figure is the last slide)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
