#!/usr/bin/env python3
r"""build_grant_deck.py  Two-slide grant-report deck (FF 2026-09-29): CAT vs FMC at the
CAR-T synapse, 10 min -- representative cells + POOLED quantification (FF explicitly
asked to combine experiments here).

  Slide 1  Actin (Kiet's Jul 17 + Jul 23, 2026 confocal 3/5/7/10 min set)
           A reps (Actin | Regions)   <- newest New folder\results_compiled\actin_reps_*
           B pooled actin profile     <- actin_pooled_panels_*\profiles
           C pooled synapse area      <- pooled_10min_violins_*
           D pooled zone violins      <- actin_pooled_panels_*\violins
           (4th slot left blank, FF 2026-09-29)
  Slide 2  pMLC (Mar 12 / May 10 / Jun 07, 2026)
           A reps (Actin | pMLC | Merge | Regions) <- pmlc_joint\paper_reps_*
           B pooled actin + pMLC profile           <- pmlc_joint\grant_pmlc_panels_*\profiles (no title)
           C pMLC MFI / CAT median, D-F zone MFI / synapse mean
                                                   <- pmlc_joint\pooled_paper_panels_*
Layout = sketch_grant_wireframe.py (FF-approved). Sources auto-resolve to the NEWEST
matching folder; a missing PNG is drawn as a labelled placeholder, so the deck builds
before the actin representative cells are rendered.

Helpers (add_text / fit_in_box / Deck / export_pngs) are reused from
CART_pMLC_ppt\build_pmlc_profiles_deck.py; tiered backup via ppt_image_inserter.

Output: K:\FF\PPT\PPT_autogeneration\CART\grant\CART_grant_actin_pMLC_<yyyymmdd>.pptx
Run:    python build_grant_deck.py [--export <dir>]
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'CART_pMLC_ppt'))
import build_pmlc_profiles_deck as bpd  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN  # noqa: E402
from pptx.util import Inches, Pt  # noqa: E402
import numpy as np  # noqa: E402
from PIL import Image  # noqa: E402

# This repo's package for the tiered backup (was a copy under PPT_editing\code whose
# failed import silently disabled backups; bug hunt 2026-10-24).
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from ppt_image_inserter import backup_presentation  # noqa: E402

NF = Path(r'\\10.225.56.176\ULab_data\Data_to_share\FF\New folder\results_compiled')
PJ = Path(r'\\10.225.56.176\ULab_data\User_data\Kiet\results_compiled\pmlc_joint')
OUT_DIR = Path(r'K:\FF\PPT\PPT_autogeneration\CART\grant')
OUT = OUT_DIR / f'CART_grant_actin_pMLC_{datetime.date.today():%Y%m%d}.pptx'

# Results folder the hand-written bullets were last checked against (None = not recorded);
# bpd.flag_text_version warns when the newest panels come from a different run.
# 2026-10-24: all five bullets hold. Actin: actin_pooled_panels_v4_20260929\zone_stats.csv +
# compiled_20260717_20260723_20260917\statistics_summary.csv (FMC larger and distal-shifted in
# 2/2 experiments). pMLC: grant_pmlc_panels_v4_20260929 / extended_panels_20260917 zone_stats.csv
# (FMC more total, more peripheral, less distal pMLC in 3/3). The keys are the panel folders
# the check compares against (PNG-only folders; the stats live beside them as listed).
GRANT_TEXT_VERIFIED = {('grant', 'actin'): 'pooled_10min_violins_20260929',
                       ('grant', 'pMLC'): 'pooled_paper_panels_v2_20260929'}

X0 = 0.38                                   # wireframe left margin
GREEN = RGBColor(0x11, 0x99, 0x11); MAGENTA = RGBColor(0xB0, 0x10, 0xB0)
GREY = RGBColor(0x55, 0x55, 0x55)


def newest(root, pattern):
    d = bpd.newest_run(root, pattern)
    print(f'  {pattern:<28s} -> {d}')
    return d


def rep_png(rep_dir, pattern):
    return bpd.one_match(rep_dir, pattern)


def letter(s, x, y, ch):
    bpd.add_text(s, x - 0.08, y - 0.12, 0.4, 0.4, ch, 20, bold=True)


def axes_rows(png):
    """(top, bottom) of the plot's axes frame as fractions of the PNG height: the rows
    where > 40 % of pixels are dark (the frame's horizontal lines)."""
    a = np.asarray(Image.open(png).convert('L'))
    rows = np.where((a < 80).mean(axis=1) > 0.4)[0]
    if rows.size < 2:                    # no frame found (re-styled panel): caller falls back
        print(f'[warning] no axes frame found in {Path(png).name}; panels C and D not height-matched')
        return None
    return rows.min() / a.shape[0], rows.max() / a.shape[0]


def bar_label(s, png, left, top, tile):
    """'5 μm' centered just above the renderer's scale bar (FF 2026-09-29). The bar
    is located from its white pixels in the bottom-right corner of the tile, not assumed."""
    a = np.asarray(Image.open(png).convert('L'))
    h, w = a.shape
    r0, c0 = int(0.9 * h), int(0.6 * w)
    rr, cc = np.where(a[r0:, c0:] > 250)
    if rr.size == 0:
        return
    bar_top = top + (rr.min() + r0) / h * tile
    cx = left + ((cc.min() + cc.max()) / 2 + c0) / w * tile
    bw = (cc.max() - cc.min()) / w * tile
    size = 16 if tile > 2 else 11
    tb = s.shapes.add_textbox(Inches(cx - bw / 2 - 0.3), Inches(bar_top - 0.4), Inches(bw + 0.6), Inches(0.38))
    tf = tb.text_frame
    tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
    tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    para = tf.paragraphs[0]; para.alignment = PP_ALIGN.CENTER
    r = para.add_run(); r.text = '5 μm'
    r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def reps_block(s, rep_dir, cond_fmt, cols, x, y_head, y0, tile, gap=0.06, lab_w=0.95):
    """Rows CAT / FMC, columns = (suffix, header, colour); header row above the tiles."""
    for j, (_, name, col) in enumerate(cols):
        bpd.add_text(s, x + lab_w + j * (tile + gap), y_head, tile, 0.4, name, 18, bold=True,
                     color=col, align=PP_ALIGN.CENTER)
    for i, c in enumerate(('CAT', 'FMC')):
        top = y0 + i * (tile + gap)
        bpd.add_text(s, x, top + tile / 2 - 0.25, lab_w - 0.08, 0.5, c, 22, bold=True,
                     color=bpd.CAT_BLUE if c == 'CAT' else bpd.FMC_RED, align=PP_ALIGN.RIGHT)
        for j, (suffix, _, _) in enumerate(cols):
            png = rep_png(rep_dir, cond_fmt.format(c=c, suffix=suffix))
            if png is None:
                bpd.add_text(s, x + lab_w + j * (tile + gap), top + tile / 2 - 0.3, tile, 0.6,
                             f'{c} {suffix}\n(pick pending)', 12, color=GREY, align=PP_ALIGN.CENTER)
            else:
                bpd.fit_in_box(s, png, x + lab_w + j * (tile + gap), top, tile, tile)
                if i == 1 and j == len(cols) - 1:          # the one scale bar: last tile
                    bar_label(s, png, x + lab_w + j * (tile + gap), top, tile)


def footer(s, text, x=X0, y=7.08, w=12.6, h=0.35):
    tb = bpd.add_text(s, x, y, w, h, text, 12, color=GREY)
    for p in tb.text_frame.paragraphs:                    # FF 2026-09-29: footer italic
        for r in p.runs:
            r.font.italic = True


def bullets(s, x, y, w, h, lines, size=15):
    """Concise 10 min takeaways in the slide's blank area (FF 2026-09-29)."""
    bpd.add_text(s, x, y, w, h, lines, size, bold=True)


def calibri(prs):
    """FF 2026-09-29: Calibri for all builder text EXCEPT the slide title (first shape)."""
    for slide in prs.slides:
        for shp in list(slide.shapes)[1:]:
            if shp.has_text_frame:
                for p in shp.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.name = 'Calibri'


def actin_slide(deck):
    reps = newest(NF, 'actin_reps_*')
    ext = newest(NF, 'actin_pooled_panels_*')
    vio = newest(NF, 'pooled_10min_violins_*')
    s = deck.slide('Actin at the CAR T Synapse, 10 min')        # FF: no hyphen in CAR T
    # A: reps 2x2 in the 6.66 x 5.86 in box (headers above the tiles)
    tile = 2.775
    reps_block(s, reps, 'rep_Jul_17_{c}_10min_*_{suffix}.png',
               [('actin', 'Actin', GREEN), ('zones_masks', 'Regions', bpd.BLACK)],   # masks only (FF 2026-09-29)
               X0, 0.72, 1.12, tile)
    letter(s, X0, 0.84, 'A')
    # B-D: 2x2 slots of 2.9 in (bottom-right left blank)
    slots = [('B', ext / 'profiles' / 'both_catfmc_profile_Pooled_10min.png' if ext else None, 6.86, 0.9),
             ('C', vio / 'actin_bottom_mask_area_pooled_10min.png' if vio else None, 9.86, 0.9),
             ('D', ext / 'violins' / 'zones_actin_Pooled_10min.png' if ext else None, 6.86, 3.86)]
    for ch, png, dx, y in slots:
        bpd.fit_in_box(s, png, X0 + dx, y, 2.9, 2.9)
        letter(s, X0 + dx - (0.2 if ch == 'D' else 0), y, ch)   # zone panel: y-label near its edge
    # pooled 10 min medians (actin_only_results.mat, Jul 17 + Jul 23; FF 2026-09-29)
    bullets(s, X0 + 9.86, 4.0, 2.95, 2.8,
            ['• FMC synapses larger',
             '• FMC actin shifted to the edge: more distal, less central'])
    bpd.flag_text_version(s, ('grant', 'actin'), vio, GRANT_TEXT_VERIFIED)
    footer(s, 'Quantification pooled across 2 experiments (Jul 17 + Jul 23, 2026); '
              'representative cells from Jul 17.')
    return s


def pmlc_slide(deck):
    reps = newest(PJ, 'paper_reps_*')
    pool = newest(PJ, 'pooled_paper_panels_*')
    s = deck.slide('pMLC Distribution at the CAR T Synapse, 10 min')   # FF: no hyphen in CAR T
    # A: reps 2x4 in the 7.96 x 3.55 in box
    tile = 1.62
    reps_block(s, reps, 'rep_Jun_07_{c}_10min_*_{suffix}.png',
               [('actin', 'Actin', GREEN), ('pmlc', 'pMLC', MAGENTA),
                ('merge', 'Merge', bpd.BLACK), ('zones_masks', 'Regions', bpd.BLACK)],   # masks only (FF 2026-09-29)
               X0, 0.68, 1.05, tile)
    letter(s, X0, 0.8, 'A')
    # B: pooled actin + pMLC profile (1.19:1) -- from the grant render (no panel title,
    # FF 2026-09-29); same pooled data as extended_panels_*.
    zones = newest(PJ, 'grant_pmlc_panels_*')
    bpd.fit_in_box(s, zones / 'profiles' / 'both_catfmc_profile_Pooled_10min.png' if zones else None,
                   X0 + 8.11, 0.8, 3.55 * 1.19, 3.55)
    letter(s, X0 + 7.85, 0.8, 'B')            # left of the profile's y-label
    # C: CAT-normalized pMLC MFI; D: the three zones in ONE pooled panel, the same style
    # as the actin slide's D (FF 2026-09-29); rest blank.
    row = [('C', pool / 'pMLC_synapse_MFI_normCAT_pooled_10min.png' if pool else None),
           ('D', zones / 'violins' / 'zones_pmlc_Pooled_10min.png' if zones else None)]
    # D is sized so its axes box has C's height with the tops aligned (FF 2026-09-29),
    # MEASURED from both PNGs (axes frame rows), so a cropped / re-rendered panel still lines up.
    c_png, d_png = row[0][1], row[1][1]
    x_d = X0 + (12.58 - 2.62) / 3
    bpd.fit_in_box(s, c_png, X0, 4.45, 2.62, 2.62)
    letter(s, X0, 4.45, 'C')
    c_rows = axes_rows(c_png) if c_png and Path(c_png).is_file() else None
    d_rows = axes_rows(d_png) if d_png and Path(d_png).is_file() else None
    if c_rows and d_rows and d_rows[1] > d_rows[0]:
        (ct, cb), (dt, db) = c_rows, d_rows
        with Image.open(c_png) as im:
            cw, chh = im.size
        c_h = 2.62 * min(1, chh / cw)                  # C's displayed height in its 2.62 box
        c_top = 4.45 + (2.62 - c_h) / 2
        with Image.open(d_png) as im:
            dw, dh = im.size
        d_h = (cb - ct) * c_h / (db - dt)              # D axes height == C axes height
        d_top = c_top + ct * c_h - dt * d_h            # axes tops aligned
        bpd.fit_in_box(s, d_png, x_d, d_top, d_h * dw / dh, d_h)
    else:
        bpd.fit_in_box(s, d_png, x_d, 4.45, 2.62, 2.62)
    letter(s, x_d - 0.2, 4.45, 'D')                    # zone panel: y-label near its edge
    # pooled 10 min medians (CART_results.mat, Mar 12 / May 10 / Jun 07; FF 2026-09-29)
    bullets(s, X0 + 7.1, 4.9, 5.5, 1.5,          # FF 2026-09-29: moved down
            ['• FMC: more pMLC than CAT (in each of 3 experiments)',
             '• pMLC peaks in the peripheral zone, inside the actin edge',
             '• FMC: more peripheral, less distal pMLC'])
    bpd.flag_text_version(s, ('grant', 'pMLC'), pool, GRANT_TEXT_VERIFIED)
    # footer in the blank area right of D (D's legend now reaches the usual footer line)
    footer(s, 'Quantification pooled across 3 experiments (Mar 12, May 10, Jun 07, 2026); '
              'pMLC MFI normalized to each experiment\u2019s CAT median. Representative cells from Jun 07.',
           x=X0 + 7.1, y=6.45, w=5.5, h=0.9)
    return s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--export', help='also export one PNG per slide to this dir')
    a = ap.parse_args()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if OUT.exists() and backup_presentation is not None:
        backup_presentation(str(OUT), backup_base=str(OUT_DIR / 'backups'))
    deck = bpd.Deck()
    actin_slide(deck)                       # FF 2026-09-29: actin before pMLC
    pmlc_slide(deck)
    calibri(deck.prs)
    deck.prs.save(str(OUT))
    print(f'saved {OUT} ({len(deck.prs.slides)} slides)')
    if a.export:
        bpd.OUT = OUT                       # export_pngs opens bpd.OUT
        bpd.export_pngs(a.export, deck.titles)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
