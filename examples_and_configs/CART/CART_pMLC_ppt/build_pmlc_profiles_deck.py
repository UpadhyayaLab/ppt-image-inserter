#!/usr/bin/env python3
r"""build_pmlc_profiles_deck.py  pMLC / actin synapse-distribution deck (the
non-paper companion to the Figure-1 slide in build_pmlc_paper_figure.py).

Supersedes build_pmlc_extended_deck.py (too repetitive) and
build_pmlc_condensed_deck.py (FF 2026-09-14). One slide per distinct view:
  - zones drawn on representative cells       (Rey-Suarez et al. 2021 Fig 1D analogue)
  - actin (green) + pMLC (magenta) radial profile AND gradient, CAT solid / FMC
    dashed, one slide per dataset (pooled + each replicate)      (Fig 1B/C analogue)
  - FMC - CAT difference curves, actin + pMLC, every replicate + pooled
  - zone-intensity violins, CAT/FMC side by side per zone          (Fig 1E-G analogue)
  - effect-size tables from the stats CSV; titles and sparse labels only (no prose).

Inputs (nothing is plotted here):
  newest ...\pmlc_joint\extended_panels_*\  <- tests/run_pmlc_extended_panels.m
  newest ...\pmlc_joint\paper_reps_*\       <- make_pmlc_rep_panels.py
  compiled_pmlc_v2_20260831\statistics_summary.csv (within-replicate total pMLC MFI)

Output: K:\FF\PPT\PPT_autogeneration\CART\pMLC\CART_pMLC_profiles_<yyyymmdd>.pptx
        + one PNG per slide, named by slide title, for the review gallery.
Run:    python build_pmlc_profiles_deck.py [--export <dir>]

--marker pZap70 (FF 2026-09-15): the same deck for the 4 TIRF actin+pZap70 experiments
(panels from tests/run_pzap70_extended_panels.m; NO pooled dataset, no rep cells yet,
plus band-pass puncta slides) ->
  K:\FF\PPT\PPT_autogeneration\CART\pZap70\CART_pZap70_profiles_<yyyymmdd>.pptx
"""
from __future__ import annotations

import argparse
import csv
import datetime
import glob
import re
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

NAS = Path(r'\\10.225.56.176\ULab_data\User_data\Kiet\results_compiled\pmlc_joint')


def newest(pattern):
    hits = sorted(glob.glob(str(NAS / pattern)))
    return Path(hits[-1]) if hits else None


EXT = newest('extended_panels_*')
REPS = newest('paper_reps_*')
PROF = EXT / 'profiles' if EXT else None
VIO = EXT / 'violins' if EXT else None
COMPILE_STATS = NAS / 'compiled_pmlc_v2_20260831' / 'statistics_summary.csv'
OUT = Path(r'K:\FF\PPT\PPT_autogeneration\CART\pMLC') / \
    f'CART_pMLC_profiles_{datetime.date.today():%Y%m%d}.pptx'

SLIDE_W, SLIDE_H = 13.333, 7.5
BLACK = RGBColor(0, 0, 0); GREY = RGBColor(0x55, 0x55, 0x55)
FMC_RED = RGBColor(0xD9, 0x40, 0x40); CAT_BLUE = RGBColor(0x33, 0x73, 0xD9)
MINUS, DASH = '\u2212', '\u2014'
TPS = ['5min', '10min', '15min']
TP_LBL = ['5 min', '10 min', '15 min']
DATASETS = [('Pooled', 'Pooled'), ('Mar_12', 'Mar 12, 2026'), ('May_10', 'May 10, 2026'),
            ('Jun_07', 'Jun 07, 2026')]
REPLICATES = ['Mar 12, 2026', 'May 10, 2026', 'Jun 07, 2026']
ZONES = ['Central', 'Peripheral', 'Distal']
ZONE_FIELD = {'pMLC': 'pMLC_synapse_{}_MFI_norm', 'Actin': 'actin_bottom_slice_{}_MFI_norm'}
MARKER = 'pMLC'
AVAIL = {}          # dataset tag -> timepoint slots it has (absent = all TPS)
REP_TITLE = 'Representative Synapses, 10 min'
REP_GLOB = 'rep_*_{c}_10min_*_{suffix}.png'
REP_BULLETS = ['• pMLC concentrates in the peripheral zone, inside the actin-rich edge (CAT and FMC)',
               '• FMC: more peripheral, less distal pMLC in every replicate; more total pMLC at 10 min']

# pZap70 TIRF set: 2023 experiments imaged one 6 min point (stored in the 5min slot).
PZAP70_ROOT = Path(r'L:\FF\CAR T\CAR T data for actin area trouble shoot\results_compiled_tirf')
PZAP70_DATASETS = [('Sep_2023', 'Sep 2023', 'Sep 21, 2023 (6 min)'),
                   ('Oct_2023', 'Oct 2023', 'Oct 03, 2023 (6 min)'),
                   ('Feb_2025_D3', 'Feb 2025 D3', 'Feb 17, 2025 (D3)'),
                   ('Feb_2025_D5', 'Feb 2025 D5', 'Feb 17, 2025 (D5)')]
# Punctum area dropped (FF 2026-09-16): per-cell median of pixel-count areas sits at 35-36 px in
# nearly every cell -- set by the detector's band-pass scale, not biology.
PUNCTA = [('puncta_per_um2', 'Density'), ('peak_median', 'Peak / Mean'),
          ('puncta_to_nonpuncta', 'Puncta / Non-Puncta')]
# Per-zone puncta (FF 2026-09-17, after Kiet's reprocess stored each punctum's zone):
# density = puncta whose PEAK pixel falls in the zone / zone area; signal fraction =
# marker signal inside puncta pixels of the zone / all marker signal in the zone.
ZONE_PUNCTA = [('puncta_per_um2', 'Puncta Density'),
               ('peak_median', 'Puncta Peak / Mean'),
               ('puncta_signal_frac', 'Puncta Signal Fraction')]


def configure_pzap70():
    global EXT, PROF, VIO, OUT, TPS, TP_LBL, DATASETS, REPLICATES, ZONE_FIELD, MARKER, AVAIL
    global REPS, REP_TITLE, REP_GLOB, REP_BULLETS
    hits = sorted(glob.glob(str(PZAP70_ROOT / 'pzap70_extended_panels_*')))
    EXT = Path(hits[-1]) if hits else None
    rhits = sorted(glob.glob(str(PZAP70_ROOT / 'pzap70_reps_*')))
    REPS = Path(rhits[-1]) if rhits else None
    REP_TITLE = f'Representative Synapses, 15 min {DASH} Feb 2025 D3'
    REP_GLOB = 'rep_D3_{c}_15min_*_{suffix}.png'           # rank-1 picks (alternates are rep_alt*)
    # FF 2026-09-15: super basic summary; numbers are in the effect-table slides.
    # FF 2026-09-16: the zone shift is NOT a takeaway (differs by timepoint); main result =
    # higher synapse MFI in FMC at every timepoint of every experiment (6/6 panels).
    REP_BULLETS = ['• FMC: higher pZap70 MFI at the synapse at every timepoint in every experiment',
                   '• Actin zones: no consistent CAT vs FMC difference']
    PROF = EXT / 'profiles' if EXT else None
    VIO = EXT / 'violins' if EXT else None
    OUT = Path(r'K:\FF\PPT\PPT_autogeneration\CART\pZap70') / \
        f'CART_pZap70_profiles_{datetime.date.today():%Y%m%d}.pptx'
    TPS, TP_LBL = ['5min', '15min'], ['5/6 min', '15 min']
    DATASETS = [(t, n) for t, n, _ in PZAP70_DATASETS]
    REPLICATES = [full for _, _, full in PZAP70_DATASETS]
    ZONE_FIELD = {'pZap70': 'pZap70_synapse_{}_MFI_norm', 'Actin': 'actin_bottom_slice_{}_MFI_norm'}
    MARKER = 'pZap70'
    AVAIL = {'Sep_2023': ['5min'], 'Oct_2023': ['5min']}


def has(tag, tp):
    return tp in AVAIL.get(tag, TPS)


# ---- layout helpers ---------------------------------------------------------------
def add_text(slide, left, top, width, height, text, size, *, bold=False, color=BLACK,
             align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True
    for i, line in enumerate(text if isinstance(text, list) else [text]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
        p.space_after = Pt(6)
    return tb


def fit_in_box(slide, png, left, top, box_w, box_h):
    """Aspect-fit, never stretched (same rule as the vimentin Fig4 builder)."""
    if png == '':                              # intentionally empty cell (condition not imaged)
        return
    png = Path(png) if png else None
    if png is None or not png.is_file():
        add_text(slide, left, top + box_h / 2 - 0.2, box_w, 0.4,
                 f'[missing] {png.name if png else ""}', 10, color=GREY, align=PP_ALIGN.CENTER)
        return
    with Image.open(png) as im:
        iw, ih = im.size
    ar = iw / ih
    w, h = (box_w, box_w / ar) if ar >= box_w / box_h else (box_h * ar, box_h)
    slide.shapes.add_picture(str(png), Inches(left + (box_w - w) / 2),
                             Inches(top + (box_h - h) / 2), width=Inches(w))


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(SLIDE_W); self.prs.slide_height = Inches(SLIDE_H)
        self.titles = []

    def slide(self, title, subtitle=''):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        add_text(s, 0.3, 0.1, SLIDE_W - 0.6, 0.55, title, 26, bold=True, align=PP_ALIGN.CENTER)
        # FF 2026-09-14 ("vastly reduce text"): subtitles are no longer drawn.
        self.titles.append(title)
        return s

    def grid(self, title, subtitle, pngs, nrow, ncol):
        s = self.slide(title, subtitle)
        top = 0.8
        bw = (SLIDE_W - 0.4) / ncol
        bh = (SLIDE_H - top - 0.1) / nrow
        for i, p in enumerate(pngs):
            r, c = divmod(i, ncol)
            fit_in_box(s, p, 0.2 + c * bw, top + r * bh, bw - 0.08, bh - 0.06)
        return s

    def labeled_grid(self, title, pngs, row_labels, col_labels, key=True, fs=18):
        """Grid with dataset / timepoint headers on the slide (panels carry no titles)
        and one CAT / FMC colour key, so the panels get the space. fs = header font size."""
        s = self.slide(title)
        nrow, ncol = len(row_labels), len(col_labels)
        lab_w = (1.35 if fs <= 18 else 2.0) if any(row_labels) else 0.2   # big headers need room ("Gradient")
        top = 1.1 if fs <= 18 else 1.2
        bw = (SLIDE_W - lab_w - 0.2) / ncol
        bh = (SLIDE_H - top - 0.08) / nrow
        for c, lab in enumerate(col_labels):
            add_text(s, lab_w + c * bw, 0.66, bw, 0.5, lab, fs, bold=True, align=PP_ALIGN.CENTER)
        for r, lab in enumerate(row_labels):
            add_text(s, 0.1, top + r * bh + bh / 2 - 0.25, lab_w - 0.1, 0.5, lab, fs, bold=True)
        for i, p in enumerate(pngs):
            r, c = divmod(i, ncol)
            fit_in_box(s, p, lab_w + c * bw, top + r * bh, bw - 0.06, bh - 0.04)
        if key:
            add_text(s, SLIDE_W - 2.4, 0.12, 1.0, 0.5, 'CAT', 22, bold=True, color=CAT_BLUE, align=PP_ALIGN.RIGHT)
            add_text(s, SLIDE_W - 1.35, 0.12, 1.0, 0.5, 'FMC', 22, bold=True, color=FMC_RED)
        return s

    def save(self):
        for sl in self.prs.slides:                 # text above pictures
            tree = sl.shapes._spTree
            for sh in list(sl.shapes):
                if sh.has_text_frame:
                    el = sh._element; tree.remove(el); tree.append(el)
        self.prs.save(str(OUT))


# ---- stats ------------------------------------------------------------------------
def read_csv(path):
    if not Path(path).is_file():
        return []
    with open(path, encoding='utf-8-sig', newline='') as fh:
        return list(csv.DictReader(fh))


def stars(p):
    """Thresholds identical to sigstar_within_ylim.m."""
    p = float(p)
    if p != p or p > 0.05:
        return 'n.s.'
    return '****' if p <= 1e-4 else '***' if p <= 1e-3 else '**' if p <= 1e-2 else '*'


def fmc_delta(r):
    """Cliff's delta re-expressed as FMC relative to CAT (positive = FMC higher)."""
    return -float(r['CliffsDelta_CAT_minus_FMC'])


def effect_table_slide(deck, zrows, chan):
    s = deck.slide(f'{chan} Zone Intensity: FMC vs CAT Effect Sizes',
                   'Cliff\u2019s \u03b4 (FMC relative to CAT; + = FMC higher), BH over the 3 zones of '
                   'each panel (as drawn); red = FMC higher, blue = CAT higher, grey = n.s.')
    metrics = [(ZONE_FIELD[chan].format(z.lower()), z) for z in ZONES]
    _effect_table(s, zrows, metrics)


def puncta_table_slide(deck, prows):
    s = deck.slide(f'{MARKER} Puncta: FMC vs CAT Effect Sizes')
    _effect_table(s, prows, [(f'{MARKER}_synapse_PtoMbp_{f}', lab) for f, lab in PUNCTA])


def _effect_table(s, zrows, metrics):
    dsets = REPLICATES + (['Pooled'] if any(x['Dataset'] == 'Pooled' for x in zrows) else [])
    pairs = []
    for d in dsets:                               # timepoints in the order the CSV lists them
        tps = list(dict.fromkeys(x['Timepoint'] for x in zrows if x['Dataset'] == d)) or TP_LBL
        pairs += [(d, tp) for tp in tps]
    nrow = len(pairs) + 1
    cols = ['Dataset', 'Timepoint'] + [lab for _, lab in metrics]
    mw = (SLIDE_W - 1.2 - 2.6 - 1.5) / len(metrics)
    tbl = s.shapes.add_table(nrow, len(cols), Inches(0.6), Inches(1.15), Inches(SLIDE_W - 1.2),
                             Inches(0.46 * nrow)).table
    for j, w in enumerate([2.6, 1.5] + [mw] * len(metrics)):
        tbl.columns[j].width = Inches(w)
    for j, c in enumerate(cols):
        tbl.cell(0, j).text = c
    prev = None
    for i, (d, tp) in enumerate(pairs, start=1):
        tbl.cell(i, 0).text = d if d != prev else ''
        prev = d
        tbl.cell(i, 1).text = tp
        for j, (met, _) in enumerate(metrics, start=2):
            r = next((x for x in zrows if x['Metric'] == met and x['Dataset'] == d
                      and x['Timepoint'] == tp), None)
            if r is None:
                tbl.cell(i, j).text = 'n/a'; continue
            st = stars(r['p_BH_panel'])
            tbl.cell(i, j).text = f'{fmc_delta(r):+.2f}  {st}'
            col = GREY if st == 'n.s.' else FMC_RED if fmc_delta(r) > 0 else CAT_BLUE
            for para in tbl.cell(i, j).text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = col
    for a in range(nrow):
        for b in range(len(cols)):
            for para in tbl.cell(a, b).text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(14); run.font.bold = (a == 0 or b >= 2)


def reps_slide(deck):
    """Representative cells (near-median picks, same field of view; actin per-cell
    montage contrast, pMLC at the pipeline's fixed CLim): rows CAT / FMC, columns
    Actin / pMLC / Merge + zones, one scale bar (last image), two takeaways below."""
    s = deck.slide(REP_TITLE)
    img, gap, lab_w = 2.45, 0.08, 1.3
    cols = [('actin', 'Actin', RGBColor(0x11, 0x99, 0x11)), (MARKER.lower(), MARKER, RGBColor(0xB0, 0x10, 0xB0)),
            ('merge', 'Merge', BLACK), ('zones', 'Regions', BLACK)]
    left0 = (SLIDE_W - (lab_w + len(cols) * img + (len(cols) - 1) * gap)) / 2 + lab_w
    top0 = 1.2
    for j, (_, name, col) in enumerate(cols):
        add_text(s, left0 + j * (img + gap), 0.68, img, 0.5, name, 24, bold=True, color=col,
                 align=PP_ALIGN.CENTER)
    for i, c in enumerate(('CAT', 'FMC')):
        top = top0 + i * (img + gap)
        add_text(s, left0 - lab_w, top + img / 2 - 0.35, lab_w - 0.1, 0.7, c, 32, bold=True,
                 color=CAT_BLUE if c == 'CAT' else FMC_RED, align=PP_ALIGN.RIGHT)
        for j, (suffix, _, _) in enumerate(cols):
            png = next(iter(sorted(REPS.glob(REP_GLOB.format(c=c, suffix=suffix)))), None) if REPS else None
            fit_in_box(s, png, left0 + j * (img + gap), top, img, img)
    add_text(s, 0.9, top0 + 2 * (img + gap) + 0.02, SLIDE_W - 1.8, 1.25, REP_BULLETS,
             17 if len(REP_BULLETS) > 2 else 18, bold=True)
    return s


def title_slide(deck, title=None):
    s = deck.slide(title or f'{MARKER} and Actin Distribution at the CAR-T Synapse: CAT vs FMC')
    nrows = read_csv(EXT / 'n_per_condition.csv') and list(csv.reader(open(
        EXT / 'n_per_condition.csv', encoding='utf-8-sig', newline='')))
    if nrows:
        tbl = s.shapes.add_table(len(nrows), len(nrows[0]), Inches(1.5), Inches(2.2),
                                 Inches(SLIDE_W - 3.0), Inches(0.42 * len(nrows))).table
        for i, row in enumerate(nrows):
            for j, v in enumerate(row):
                tbl.cell(i, j).text = ('Cells per condition' if (i == 0 and j == 0) else
                                       v.replace('_', ' ').replace('min', ' min') if i == 0 else
                                       ('–' if v == '0' else v))   # not imaged, not "0 cells"
                for para in tbl.cell(i, j).text_frame.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(14); run.font.bold = (i == 0 or j == 0)


def export_pngs(export_dir, titles):
    """Render each slide via PowerPoint and name the PNG after its slide title.

    PowerPoint is single-instance: Dispatch attaches to FF's open window. NEVER call
    app.Quit() (destroys FF's unsaved work); only Close() the presentation this function
    opened, and skip the export if FF already has this deck open."""
    import os
    import win32com.client as w
    app = w.Dispatch('PowerPoint.Application')
    target = os.path.normcase(os.path.abspath(str(OUT)))
    for k in range(1, app.Presentations.Count + 1):
        try:                                      # protected-view / locked decks deny FullName
            full = app.Presentations(k).FullName
        except Exception:                         # noqa: BLE001 -- unreadable: cannot be ours to skip
            continue
        if os.path.normcase(full) == target:
            print(f'[skip export] {OUT.name} is open in PowerPoint; close it there to export slide PNGs')
            return
    export_dir = Path(export_dir); export_dir.mkdir(parents=True, exist_ok=True)
    for f in export_dir.glob('*.png'):
        f.unlink()
    pres = app.Presentations.Open(str(OUT), WithWindow=False)
    try:
        for i, t in enumerate(titles, start=1):
            name = re.sub(r'[^A-Za-z0-9+\-, ]+', ' ', t.replace(MINUS, '-').replace(DASH, '-'))
            name = re.sub(r'\s+', ' ', name).strip()
            pres.Slides(i).Export(os.path.join(str(export_dir), f'{i:02d} {name}.png'), 'PNG', 1920, 1080)
    finally:
        pres.Close()                              # our presentation only; the app stays running
    print(f'exported {len(titles)} slide PNGs -> {export_dir}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--export', help='also export one PNG per slide (named by title) to this dir')
    ap.add_argument('--marker', default='pMLC', choices=['pMLC', 'pZap70'])
    ap.add_argument('--puncta', action='store_true',
                    help='pMLC band-pass puncta deck (CART_pMLC_puncta_<date>.pptx) instead of profiles')
    args = ap.parse_args()
    if args.marker == 'pZap70':
        configure_pzap70()
    if args.puncta:
        if EXT is None:
            print('[error] run the extended-panels driver for this marker first'); return 1
        return build_puncta(args)
    if args.marker == 'pZap70':
        if EXT is None:
            print('[error] run tests/run_pzap70_extended_panels.m first'); return 1
        return build_pzap70(args)
    if EXT is None:
        print('[error] run tests/run_pmlc_extended_panels.m first'); return 1
    print(f'panels: {EXT}\nreps:   {REPS}')
    zrows = read_csv(EXT / 'zone_stats.csv')
    deck = Deck()
    P = lambda n: PROF / n  # noqa: E731

    # FF 2026-09-14: titles and sparse labels only -- no subtitles, findings or methods text.
    title_slide(deck)
    reps_slide(deck)
    deck.grid('pMLC Synapse MFI', '', [VIO / f'synapse_mfi_{tag}.png' for tag, _ in DATASETS[1:]], 1, 3)
    for tag, name in DATASETS:
        deck.grid(f'Actin + pMLC Profile and Gradient {DASH} {name}', '',
                  [P(f'both_catfmc_profile_{tag}_{t}.png') for t in TPS] +
                  [P(f'both_catfmc_grad_{tag}_{t}.png') for t in TPS], 2, 3)
    deck.grid(f'FMC {MINUS} CAT', '',
              [P(f'diff_both_profile_{t}.png') for t in TPS] +
              [P(f'diff_both_grad_{t}.png') for t in TPS], 2, 3)
    # FF 2026-09-14: all pMLC zone panels on ONE slide, one shared y-axis.
    deck.labeled_grid('pMLC by Zone: Normalized by MFI over Full Synapse',   # FF 2026-09-15 title
                      [VIO / f'zones_pmlc_{tag}_{t}.png' for t in TPS for tag, _ in DATASETS],
                      TP_LBL, [name.replace(', 2026', '') for _, name in DATASETS])
    # FF 2026-09-15: raw zone MFI too -- per replicate only (raw intensity is never pooled).
    deck.labeled_grid('pMLC Zone MFI (Raw)',
                      [VIO / f'zones_raw_pmlc_{tag}_{t}.png' for t in TPS for tag, _ in DATASETS[1:]],
                      TP_LBL, [name.replace(', 2026', '') for _, name in DATASETS[1:]])
    deck.labeled_grid(f'Actin Zone Intensity {DASH} 10 min',
                      [VIO / f'zones_actin_{tag}_10min.png' for tag, _ in DATASETS],
                      [''], [name.replace(', 2026', '') for _, name in DATASETS])
    if zrows:
        effect_table_slide(deck, zrows, 'pMLC')
        effect_table_slide(deck, zrows, 'Actin')

    deck.save()
    print(f'saved {OUT} ({len(deck.prs.slides)} slides)')
    if args.export:
        export_pngs(args.export, deck.titles)
    return 0


def puncta_reps_slide(deck, bullets):
    """Band-pass puncta QC overlays (pipeline render: detected puncta outlined on the
    pMLC synapse plane, count / density / peak-to-mean printed above) for the SAME
    representative cells FF chose for the profiles deck. Both images keep one inches-
    per-pixel scale, so the burned-in scale bars stay physically comparable."""
    s = deck.slide(PUNCTA_REP_TITLE[MARKER])
    picks = {r['tag']: r for r in read_csv(REPS / 'picks.csv')      # rank-1 picks (skip rep_alt* rows)
             if not r['cond'].startswith('alt')} if REPS else {}
    files = []
    for tag in ('CAT', 'FMC'):
        r = picks.get(tag)
        f = (Path(r['cond_dir']) / 'prog_fixed_cells' / MARKER / 'synapse' / 'peak_to_mean_bandpass'
             / f"cell_{r['cell']}_peak_to_mean_bandpass.png") if r else None
        files.append(f)
    sizes = []
    for f in files:
        with Image.open(f) as im:
            sizes.append(im.size)
    box_w, box_h, gap, top = 5.6, 4.9, 0.4, 1.2   # room under the images for 3 bullet lines
    scale = min(box_w / max(w for w, _ in sizes), box_h / max(h for _, h in sizes))   # in/px, shared
    left0 = (SLIDE_W - (2 * box_w + gap)) / 2
    for j, (tag, f, (w, h)) in enumerate(zip(('CAT', 'FMC'), files, sizes)):
        bl = left0 + j * (box_w + gap)
        add_text(s, bl, 0.68, box_w, 0.5, tag, 28, bold=True,
                 color=CAT_BLUE if tag == 'CAT' else FMC_RED, align=PP_ALIGN.CENTER)
        s.shapes.add_picture(str(f), Inches(bl + (box_w - w * scale) / 2), Inches(top + (box_h - h * scale) / 2),
                             width=Inches(w * scale))
    if bullets:
        add_text(s, 0.9, top + box_h + 0.05, SLIDE_W - 1.8, 1.3, bullets,
                 16 if len(bullets) > 2 else 18, bold=True)
    return s


def build_puncta(args):
    """Band-pass puncta deck, kept SEPARATE from the profiles decks (FF 2026-09-16).
    pMLC: replaces the one-grid-per-slide CART_pMLC_puncta_results deck; pZap70: moved out
    of the pZap70 profiles deck. One slide per metric (1 x datasets), per replicate /
    experiment only (never pooled)."""
    global OUT
    OUT = OUT.with_name(f'CART_{MARKER}_puncta_{datetime.date.today():%Y%m%d}.pptx')
    print(f'panels: {EXT}\nreps:   {REPS}')
    prows = read_csv(EXT / 'puncta_stats.csv')
    reps = [d for d in DATASETS if d[0] != 'Pooled']  # no pooled column
    deck = Deck()
    title_slide(deck, f'{MARKER} Puncta at the CAR-T Synapse: CAT vs FMC')
    puncta_reps_slide(deck, PUNCTA_BULLETS[MARKER])
    # FF 2026-09-16: one slide per metric (1 x datasets) -- bigger plots, clear titles.
    titles = {'puncta_per_um2': f'{MARKER} Puncta Density (Puncta per µm²)',
              'peak_median': f'{MARKER} Puncta Peak / Synapse Mean',
              'puncta_to_nonpuncta': f'{MARKER} Puncta / Non-Puncta MFI'}
    for f, _ in PUNCTA:
        deck.grid(titles[f], '', [VIO / f'puncta_PtoMbp_{f}_{tag}.png' for tag, _ in reps], 1, len(reps))
    # FF 2026-09-17: the same puncta split by synapse zone (one slide per metric per zone;
    # the 3 zones of a metric share a y-axis, so the panels compare zone to zone by eye).
    for f, lab in ZONE_PUNCTA:
        for z in ZONES:
            pngs = [VIO / f'puncta_PtoMbp_{z.lower()}_{f}_{tag}.png' for tag, _ in reps]
            if not all(q.is_file() for q in pngs):
                print(f'skip zone slide (missing panels): {z} {f}'); continue
            deck.grid(f'{MARKER} {lab}: {z} Zone', '', pngs, 1, len(reps))
    if prows:
        puncta_table_slide(deck, prows)
        for f, lab in ZONE_PUNCTA:
            metrics = [(f'{MARKER}_synapse_PtoMbp_{z.lower()}_{f}', z) for z in ZONES]
            if not any(x['Metric'] == metrics[0][0] for x in prows):
                continue
            s = deck.slide(f'{MARKER} {lab} by Zone: FMC vs CAT Effect Sizes')
            _effect_table(s, prows, metrics)
    deck.save()
    print(f'saved {OUT} ({len(deck.prs.slides)} slides)')
    if args.export:
        export_pngs(args.export, deck.titles)
    return 0


# FF: super basic summary; from extended_panels_20260916\puncta_stats.csv (per replicate).
PUNCTA_BULLETS = {
    'pMLC': ['• Whole-synapse puncta density (~1.2 per µm²): nearly the same in CAT and FMC',
             '• By zone: FMC has more peripheral and fewer distal puncta in all 9 comparisons',
             '• Puncta peak / mean: distal puncta dimmer in FMC in all 9 comparisons; central and peripheral mixed'],
    # pZap70 (pzap70_extended_panels_20260915\puncta_stats.csv): no metric agrees in all experiments.
    'pZap70': ['• No consistent CAT vs FMC difference in puncta density or puncta brightness across experiments']}
PUNCTA_REP_TITLE = {'pMLC': 'Representative pMLC Puncta, 10 min',
                    'pZap70': f'Representative pZap70 Puncta, 15 min {DASH} Feb 2025 D3'}


def build_pzap70(args):
    """pZap70 TIRF deck: every panel is ONE experiment (no pooling); the 2023
    experiments have only the 6 min point, so their 15 min cells stay empty."""
    print(f'panels: {EXT}')
    zrows = read_csv(EXT / 'zone_stats.csv')
    prows = read_csv(EXT / 'puncta_stats.csv')
    deck = Deck()
    P = lambda n: PROF / n  # noqa: E731
    ds_labels = [name for _, name in DATASETS]

    title_slide(deck)
    reps_slide(deck)
    deck.grid(f'{MARKER} Synapse MFI', '', [VIO / f'synapse_mfi_{tag}.png' for tag, _ in DATASETS], 1, 4)
    # FF 2026-09-16: every column must say at a glance what it is -> large timepoint headers
    # over the columns, Profile / Gradient row labels (no CAT/FMC colour key: line style
    # encodes CAT/FMC in these panels).
    for tag, name in DATASETS:
        tps = [t for t in TPS if has(tag, t)]
        tp_hdr = ['6 min' if tag in AVAIL and t == '5min' else t.replace('min', ' min') for t in tps]
        ttl = f'Actin + {MARKER} Profile and Gradient {DASH} {name}'
        if len(tps) == 1:          # single timepoint (2023): profile | gradient side by side
            deck.labeled_grid(ttl, [P(f'both_catfmc_profile_{tag}_{tps[0]}.png'),
                                    P(f'both_catfmc_grad_{tag}_{tps[0]}.png')],
                              [''], [f'{tp_hdr[0]}: Profile', f'{tp_hdr[0]}: Gradient'], key=False, fs=26)
        else:
            deck.labeled_grid(ttl, [P(f'both_catfmc_profile_{tag}_{t}.png') for t in tps] +
                                   [P(f'both_catfmc_grad_{tag}_{t}.png') for t in tps],
                              ['Profile', 'Gradient'], tp_hdr, key=False, fs=26)
    deck.labeled_grid(f'FMC {MINUS} CAT',
                      [P(f'diff_both_profile_{t}.png') for t in TPS] +
                      [P(f'diff_both_grad_{t}.png') for t in TPS],
                      ['Profile', 'Gradient'], TP_LBL, key=False, fs=26)
    for chan, ftag in ((MARKER, MARKER.lower()), ('Actin', 'actin')):
        deck.labeled_grid(f'{MARKER} by Zone: Normalized by MFI over Full Synapse' if chan == MARKER
                          else f'{chan} Zone Intensity',   # FF 2026-09-15 title (marker slide)
                          [VIO / f'zones_{ftag}_{tag}_{t}.png' if has(tag, t) else ''
                           for t in TPS for tag, _ in DATASETS],
                          TP_LBL, ds_labels)
        if chan == MARKER:          # FF 2026-09-15: raw zone MFI right after the normalized zones
            deck.labeled_grid(f'{MARKER} Zone MFI (Raw)',
                              [VIO / f'zones_raw_{ftag}_{tag}_{t}.png' if has(tag, t) else ''
                               for t in TPS for tag, _ in DATASETS],
                              TP_LBL, ds_labels)
    # FF 2026-09-16: NO puncta in the profiles deck (misleading next to the zone results);
    # puncta live in their own deck: --marker pZap70 --puncta.
    if zrows:
        effect_table_slide(deck, zrows, MARKER)
        effect_table_slide(deck, zrows, 'Actin')

    deck.save()
    print(f'saved {OUT} ({len(deck.prs.slides)} slides)')
    if args.export:
        export_pngs(args.export, deck.titles)
    return 0


if __name__ == '__main__':
    sys.exit(main())
