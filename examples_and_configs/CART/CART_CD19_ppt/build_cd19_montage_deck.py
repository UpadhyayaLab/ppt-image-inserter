#!/usr/bin/env python3
r"""build_cd19_montage_deck.py  CD19 coculture montage deck, every experiment in one deck.

Modeled directly on the CART GzmB overview deck (FF 2026-10-26: "model directly on those"):
build_gzmb_decks.build_overview / ov_montage_slide. A title slide, then per view every
timepoint, one slide each: the FIRST montage page of CAT (left) | FMC (right), i.e. the
first 25 cells of each, never every page; one px/in for the whole deck (all tiles are 26 um
squares); the contrast in the title as "(Fixed Contrast low–high)"; the '5 μm scalebar' key.

Views, one channel family each as the actin / pMLC / GzmB decks (CAR, CD19, CAR + CD19
merge), at three planes, then the puncta-ranked sheets:
    Synapse Plane                      the z with the most synapse voxels (volscope
                                       cd19_synapse/objects; was Widest Plane until 2026-10-28)
    Plane with Most CD19 >1 μm Deep    the z with the most CD19 > 1 um inside the membrane
    Plane with Most Puncta             the z with the most internalized-CD19 voxels
    Cells with Most Puncta (up to 25)          the 'most' sheet, same plane as above (CD19 and merge
                                       only: FF 2026-10-26 "there is no point showing CAR
                                       alone there")

Sheets come from volscope dev/run_cd19_montage_families.m (dev/show_internalized_cells.m,
layout 'montage'), under <experiment>/cd19_montages/<plane>/:
    <cond>/montages_<style>/montage_cells_<a>_<b>.png      (cell order, 25 a page)
    <cond>_montage_<style>_most<n>.png                      (puncta plane only)
    display_windows.csv

Two decks (FF 2026-10-26: "QC should be its own deck"): --deck montages (CAR, CD19, merge;
default) and --deck qc (CAR with the T cell outline).
One deck over every experiment given (FF 2026-10-26: "there should also not be SEPARATE
ppts for the two different experiments"), as the CAR T all-experiment decks
(insert_MT_qc_CAT_vs_FMC_202311_202406_slides.py): per view, each experiment oldest first,
then its timepoints; every title ends "<experiment>, <tp> min"; one px/in for the deck.
Run:  python build_cd19_montage_deck.py <root> [<root> ...] --exp "Mar 10, 2026" ["Oct 1, 2026" ...]
          [--deck qc] [--export DIR]
Out:  K:\FF\PPT\PPT_autogeneration\CART\CD19\CART_coculture_fixed_{montages,mask_qc}_<yyyymmdd>_<yyyymmdd>_<today>.pptx
"""
from __future__ import annotations

import argparse
import csv
import datetime
import re
import sys
from pathlib import Path

from PIL import ImageFont
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
sys.path.insert(0, r'K:\FF\Code\Scripts\PPT_image_inserter\examples_and_configs')
sys.path.insert(0, str(HERE.parent / 'CART_GzmB_ppt'))
sys.path.insert(0, str(HERE.parent / 'CART_pMLC_ppt'))
sys.path.insert(0, str(HERE.parent / 'CART_actin_ppt'))
import build_gzmb_decks as gz  # noqa: E402  (ov_montage_slide, OV_CELL_W / OV_IMG_H, OvDeck)
import insert_CD19_internalization_montages_slides as cd19  # noqa: E402  (page files, color_title)
from pptx.enum.text import PP_ALIGN  # noqa: E402

bpd, bad = gz.bpd, gz.bad
OUT_DIR = Path(r'K:\FF\PPT\PPT_autogeneration\CART\CD19')
TODAY = f'{datetime.date.today():%Y%m%d}'
DASH = gz.DASH

# The montage deck's first plane is the synapse plane, not the widest (FF 2026-10-28: "our xy
# synapse plane is more appropriate than 'widest'", replacing it); the QC deck keeps the
# widest plane, which shows the whole body outline.
PLANES = [('synapse', 'Synapse Plane'),
          ('deep1um', 'Plane with Most CD19 >1 μm Deep'),
          ('puncta', 'Plane with Most Puncta')]
QC_PLANES = [('widest', 'Plane of Maximal Cross-Sectional Area')]   # FF 2026-10-29: maximal area, not widest
# The contact views (FF 2026-10-28: "pls add montages of this nature to the montage ppt" -- the
# CAR + CD19 + Hoechst views of the contact check), sheets from volscope
# dev/contact_check_sheet.py --montage-root, style folder 'contact', after the planes above.
CONTACT_PLANES = [('contact_synapse', 'Synapse Plane'),
                  ('contact_xz', 'Synapse Slice (XZ)'),
                  ('contact_mipxy', 'XY MIP'),
                  ('contact_mipxz', 'XZ MIP')]
CONTACT_FAMILY = [('contact', 'CAR, CD19 and Hoechst', ['CD19'])]   # CAR / Hoechst: per view / per cell
# (style folder, title word, channels whose window the title states)
# CAR first (FF 2026-10-26: "CAR single channel should come first").
# car_mask = mask QC, the CAR T synapse-slice QC style (body-mask edge red), worded as the
# GzmB deck's 'MT 3-Slice MIP, Actin Outline'.
# Two decks (FF 2026-10-26: "QC should be its own deck"): --deck montages | qc.
DECKS = {
    'montages': [('car', 'CAR', ['CAR']),
                 ('cd19', 'CD19', ['CD19']),
                 ('merge', 'CAR and CD19', ['CAR', 'CD19'])],
    # car_mask_cell: each cell at its own window ONLY (FF 2026-10-26: "for QC montages remove
    # fixed contrast entirely"; the fixed-window 'car_mask' sheets are no longer shown);
    # names=None -> the title states the per-cell rule instead
    'qc': [('car_mask_cell', 'CAR, T Cell Outline', None)],
}
# The QC deck (FF 2026-10-26: "just broadest slice with 25 per page", all usable cells) shows
# EVERY page of the widest plane, CAT page k beside FMC page k; its sheets come from
# volscope dev/run_cd19_mask_qc_widest.m under <experiment>/cd19_mask_qc_widest/.
# How each view's plane (and cells) were chosen, one line under the sheets (FF 2026-10-28: "the
# notes on how the planes were chosen should be added to montage ppts"); from
# dev/show_internalized_cells.m pick_z and run_cd19_montage_families.m.
FIRST25 = 'Cells: the first 25 usable cells, in cell order.'
# No parentheses (FF 2026-10-28: "parentheses is super confusing"); every usable cell has a
# synapse patch, so the synapse views need no fallback (checked over both experiments).
PLANE_NOTE = {
    'Synapse Plane': 'Plane: the z of maximal synapse cross-sectional area. ' + FIRST25,
    'Plane with Most CD19 >1 μm Deep': ('Plane: the z with the most CD19 more than 1 μm inside '
                                        'the T cell. Cells without any: the z of maximal T cell '
                                        'cross-sectional area. ' + FIRST25),
    'Plane with Most Puncta': ('Plane: the z with the most internalized CD19. Cells without any: '
                               'the z of maximal T cell cross-sectional area. ' + FIRST25),
    'Cells with Most Puncta (up to 25)': ('Cells: up to 25 with the most internalized CD19 puncta; cells without '
                                  'any are left out. Plane: '
                                  'the z with the most internalized CD19.'),
    'Plane of Maximal Cross-Sectional Area': 'Plane: the z of maximal T cell cross-sectional area.',
}
CONTACT_WIN = 'CAR: 110 to its 95th percentile in the T cell. Hoechst: 110 to its 99.5th percentile in the cell. '
CONTACT_NOTE = {   # the contact views' notes (style 'contact')
    'Synapse Plane': 'Plane: the z of maximal synapse cross-sectional area. ' + CONTACT_WIN + FIRST25,
    'Synapse Slice (XZ)': ('Slice: the y of maximal synapse cross-sectional area; z up, coverslip at the '
                           'bottom. ' + CONTACT_WIN + FIRST25),
    'XY MIP': 'Maximum over all z. ' + CONTACT_WIN + FIRST25,
    'XZ MIP': 'Maximum over y; z up, coverslip at the bottom. ' + CONTACT_WIN + FIRST25,
}
NOTE_H = 0.35                                  # in, taken off the sheets' height for the note
DECK_TITLE = {'montages': 'CAR T Cell Project: CD19 Internalization',
              'qc': 'CAR T Cell Project: CD19 Internalization, T Cell Outline QC'}
# Merge slides: the channel names in the title in their image colors (the merge's key).
TITLE_COLORS = {'merge': {'CAR': '40E0FF', 'CD19': 'FF4FD8'},
                'contact': {'CAR': '40E0FF', 'CD19': 'FF4FD8', 'Hoechst': 'FFDD00'}}
TITLE_PT, TITLE_W_IN = 24, bpd.SLIDE_W - 0.1      # ov_montage_slide's title box
_FONT = r'C:\Windows\Fonts\calibrib.ttf'


def conditions(plane_dir):
    """{minutes: (CAT folder, FMC folder)} from the condition folders on disk."""
    tps = {}
    for d in sorted(p for p in plane_dir.iterdir() if p.is_dir()):
        m = re.match(r'^(CAT|FMC)_(\d+)min', d.name)
        if m:
            tps.setdefault(int(m.group(2)), {})[m.group(1)] = d.name
    return {t: (v.get('CAT'), v.get('FMC')) for t, v in sorted(tps.items())}


def current_pages(found):
    """The newest render's pages only. A rerun after an exclusion renames its pages
    (montage_cells_1_33 -> _1_35) and leaves the old ones beside them, so a range sort put
    the excluded cell back on the first page and doubled the QC pages (2026-10-29). Newest
    first, a page is kept unless its cell range overlaps one already kept."""
    kept = []
    for p in sorted(found, key=lambda p: p.stat().st_mtime, reverse=True):
        a, b = cd19._parse_chunk_range(p)
        if all(b < c or a > d for c, d in (cd19._parse_chunk_range(q) for q in kept)):
            kept.append(p)
    return sorted(kept, key=lambda p: cd19._parse_chunk_range(p)[0])


def first_page(plane_dir, cond, style):
    """The first 25 cells: the page with the lowest first-cell index (bad.first_page rule)."""
    pages = all_pages(plane_dir, cond, style)
    return pages[0] if pages else None


def all_pages(plane_dir, cond, style):
    if cond is None:
        return []
    return current_pages(cd19._list_chunk_files(plane_dir / cond / f'montages_{style}'))


def most_sheet(plane_dir, cond, style):
    if cond is None:
        return None
    # the NEWEST: since the sheet stops at the last cell with puncta (2026-10-29) its name
    # carries its own count (_most19), and a name sort would still pick an older _most25
    found = sorted(plane_dir.glob(f'{cond}_montage_{style}_most*.png'),
                   key=lambda p: p.stat().st_mtime)
    return found[-1] if found else None


def windows(plane_dir):
    with open(plane_dir / 'display_windows.csv', newline='') as f:
        return {r['channel']: (float(r['low']), float(r['high'])) for r in csv.DictReader(f)}


def contrast(win, names):
    """GzmB-deck wording: '(Fixed Contrast 10–100)'; a merge names each channel."""
    if len(names) == 1:
        lo, hi = win[names[0]]
        return f'Fixed Contrast {lo:g}–{hi:g}'
    return 'Fixed Contrast ' + ', '.join(f'{n} {win[n][0]:g}–{win[n][1]:g}' for n in names)


def fit_title(slide, title):
    """One line: shrink the 24 pt title only when it would wrap (long merge titles)."""
    pt = TITLE_PT
    # floor 14 pt (FF 2026-10-29: the QC title 'Plane of Maximal Cross-Sectional Area' ran off
    # both slide edges at the old 16 pt floor; 14 pt is 22 px at FF's 1500 px)
    while pt > 14 and ImageFont.truetype(_FONT, pt * 10).getlength(title) / 720 > TITLE_W_IN - 0.2:
        pt -= 1
    if pt != TITLE_PT:
        tb = next(sh for sh in slide.shapes if sh.has_text_frame)
        for r in tb.text_frame.paragraphs[0].runs:
            r.font.size = Pt(pt)


def experiment_specs(root, exp, deck_kind, families):
    """(title, style, [(label, png)], group) for ONE experiment, and its timepoints. The
    experiment ends every title (the CAR T all-experiment decks put the date there). group
    = (view, family) for a montage slide; None for QC, whose pages stay in experiment order
    with each own-window page beside its fixed-window twin."""
    specs, all_tps = [], set()
    plan = ([(p, w, families) for p, w in QC_PLANES] if deck_kind == 'qc' else
            # the 3-colour views right after the synapse plane's 2-colour merge (FF 2026-10-28:
            # "move 3 color merge after 2 color merge"), then the other planes
            [(p, w, families) for p, w in PLANES[:1]] +
            [(p, w, CONTACT_FAMILY) for p, w in CONTACT_PLANES] +
            [(p, w, families) for p, w in PLANES[1:]])
    for plane, plane_words, fams in plan:
        pdir = root / plane
        if not pdir.is_dir():
            print(f'  no {pdir} -- skipped')
            continue
        win, tps = windows(pdir), conditions(pdir)
        all_tps |= set(tps)
        if deck_kind == 'qc':
            for tp, (cat, fmc) in tps.items():
                pages = {st: (all_pages(pdir, cat, st), all_pages(pdir, fmc, st))
                         for st, _, _ in fams}
                n = max(max(len(pc), len(pf)) for pc, pf in pages.values())
                for k in range(n):
                    for style, word, names in fams:   # own window only
                        pc, pf = pages[style]
                        sides = [(cd19.page_label(c, pg[k]), pg[k]) if k < len(pg) else ('', None)
                                 for c, pg in ((cat, pc), (fmc, pf))]
                        how = (contrast(win, names) if names else
                               "Each Cell's Own Contrast, 110 to 95th Percentile in Outline")
                        specs.append((f'{word}, {plane_words} ({how}) {DASH} '
                                      f'{exp}, {tp} min ({k + 1} of {n})', style, sides,
                                      None))
            continue
        views = [(plane_words, first_page)]
        if plane == 'puncta':
            views.append(('Cells with Most Puncta (up to 25)', most_sheet))
        for view_words, pick in views:
            for style, word, names in fams:
                if pick is most_sheet and style == 'car':
                    continue                         # puncta-ranked: no CAR-only slides
                if plane in ('deep1um', 'puncta') and style == 'car':
                    continue                         # FF 2026-10-28: CAR alone at the deep / puncta planes removed
                for tp, (cat, fmc) in tps.items():
                    # CD19 alone is fixed on the contact slides; CAR and Hoechst are scaled
                    # per cell (CONTACT_WIN), which "Fixed Contrast" alone hid (audit 2026-10-29)
                    how = (f'Fixed Contrast CD19 {win["CD19"][0]:g}–{win["CD19"][1]:g}; '
                           'CAR, Hoechst per Cell'
                           if style == 'contact' else contrast(win, names))
                    specs.append((f'{word}, {view_words} ({how}) {DASH} '
                                  f'{exp}, {tp} min',
                                  style, [('CAT', pick(pdir, cat, style)),
                                          ('FMC', pick(pdir, fmc, style))],
                                  (view_words, style)))
    return specs, all_tps


def interleave(per_exp):
    """The CAR T all-experiment order (insert_MT_qc_CAT_vs_FMC_202311_202406: view, then
    date, then timepoint): each view's slides of every experiment, oldest first, before the
    next view. QC slides (group None) keep their experiment order."""
    if all(spec[3] is None for specs in per_exp for spec in specs):
        return [spec for specs in per_exp for spec in specs]
    order, groups = [], {}
    for specs in per_exp:
        for spec in specs:
            key = spec[3]
            if key not in groups:
                order.append(key)
                groups[key] = []
            groups[key].append(spec)
    return [spec for key in order for spec in groups[key]]


def folder_date(root: Path):
    """The date in the experiment folder name nearest the root (03102026_... = MMDDYYYY,
    20261001_... = YYYYMMDD), as (folder, {candidate dates}); (None, set()) if no folder on
    the path carries an 8-digit date. Only a LEADING date counts: experiment folders start with
    theirs, while dated output folders (cd19_montages_20261027) end with the render date."""
    for part in reversed(Path(root).parts):
        found = set()
        for tok in re.findall(r'^(\d{8})(?!\d)', part):
            for fmt in ('%Y%m%d', '%m%d%Y'):
                try:
                    found.add(datetime.datetime.strptime(tok, fmt).date())
                except ValueError:
                    pass
        if found:
            return part, found
    return None, set()


def check_exp_labels(dates, labels, roots) -> None:
    """Abort when an --exp label's date is not the date in its root's experiment folder.
    Labels were paired with roots by argument order only, so swapped labels titled and named
    the deck for the wrong experiment with nothing on screen (bug hunt 2026-10-27)."""
    for d, lab, root in zip(dates, labels, roots):
        folder, found = folder_date(root)
        if not found:
            print(f'[WARNING] no date in {root}; cannot check --exp "{lab}" against it')
        elif d not in found:
            raise SystemExit(f'--exp "{lab}" does not match {folder} ({root}); '
                             'give the labels in the order of the roots')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('roots', nargs='+', help='<experiment>/cd19_montages, one per experiment')
    ap.add_argument('--exp', required=True, nargs='+',
                    help='experiment labels in the order of the roots, e.g. "Mar 10, 2026"')
    ap.add_argument('--deck', default='montages', choices=sorted(DECKS))
    ap.add_argument('--export', default=None, help='also export slide PNGs to this folder')
    args = ap.parse_args()
    if len(args.exp) != len(args.roots):
        raise SystemExit('one --exp label per root')
    families = DECKS[args.deck]
    # FF 2026-10-26: "there should also not be SEPARATE ppts for the two different
    # experiments" -- one deck over every experiment, oldest first, as the CAR T decks.
    dates = [datetime.datetime.strptime(e, '%b %d, %Y').date() for e in args.exp]
    check_exp_labels(dates, args.exp, args.roots)
    exps = sorted(zip(dates, args.exp, map(Path, args.roots)))
    tag = '_'.join(f'{d:%Y%m%d}' for d, _, _ in exps)

    per_exp, all_tps = [], set()
    for _, exp, root in exps:
        specs, tps = experiment_specs(root, exp, args.deck, families)
        per_exp.append(specs)
        all_tps |= tps
    specs = interleave(per_exp)
    pngs = [p for _, _, sides, _ in specs for _, p in sides if p is not None]
    if not pngs:
        raise SystemExit(f'no montage sheets under {args.roots}')
    missing = [f'{t}: {lab}' for t, _, sides, _ in specs for lab, p in sides if p is None and lab]
    # One scale group: every sheet is 26 um tiles at native pixels. The sheets give up NOTE_H
    # at the bottom for the plane note (ov_montage_slide reads gz.OV_IMG_H at call time).
    gz.OV_IMG_H = bpd.SLIDE_H - gz.OV_IMG_TOP - 0.05 - NOTE_H
    ppi = bad.mont.compute_group_ppi(pngs, gz.OV_CELL_W, gz.OV_IMG_H)

    name = 'montages' if args.deck == 'montages' else 'mask_qc'
    bpd.OUT = OUT_DIR / f'CART_coculture_fixed_{name}_{tag}_{TODAY}.pptx'
    if (bpd.OUT.parent / f'~${bpd.OUT.name}').exists():   # as build_cd19_replicates_deck
        raise SystemExit(f'{bpd.OUT.name} is open in PowerPoint -- close it first')
    deck = gz.OvDeck()
    s = deck.prs.slides.add_slide(deck.prs.slide_layouts[6])
    deck.titles.append('Title')
    bpd.add_text(s, 0.83, 2.4, 11.67, 1.2, DECK_TITLE[args.deck], 40,
                 bold=True, align=PP_ALIGN.CENTER)
    tps_txt = ', '.join(str(t) for t in sorted(all_tps))
    exp_txt = ' and '.join(e for _, e, _ in exps)
    bpd.add_text(s, 0.83, 3.75, 11.67, 0.5,
                 f'{exp_txt} {"Experiments" if len(exps) > 1 else "Experiment"} {DASH} '
                 f'CAT vs FMC at {tps_txt} min', 24, align=PP_ALIGN.CENTER)
    today = datetime.date.today()
    bpd.add_text(s, 4.37, 4.6, 4.6, 0.5, f'{today:%B} {today.day}, {today.year}', 20,
                 align=PP_ALIGN.CENTER)
    for title, style, sides, group in specs:
        gz.ov_montage_slide(deck, title, sides, ppi)
        slide = deck.prs.slides[-1]
        # top-align: a short sheet (a puncta sheet stopping at its last cell, a QC last page)
        # was centered, starting a row below its neighbor
        for sh in slide.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                sh.top = Inches(gz.OV_IMG_TOP)
        fit_title(slide, title)
        notes = CONTACT_NOTE if style == 'contact' else PLANE_NOTE
        note = notes.get(group[0] if group else QC_PLANES[0][1], '')
        if note:
            # one line, never cut off (FF 2026-10-28): shrink from 13 pt until it fits
            pt = 13
            while pt > 9 and (ImageFont.truetype(_FONT, int(pt * 10)).getlength(note) / 720
                              > bpd.SLIDE_W - 0.4):
                pt -= 0.5
            bad.mont.add_textbox(slide, note, 0.05, bpd.SLIDE_H - NOTE_H, bpd.SLIDE_W - 0.1,
                                 NOTE_H - 0.03, font_pt=pt, color=bad.mont.WHITE, bold=True)
        if style in TITLE_COLORS:
            cd19.color_title(slide, TITLE_COLORS[style],   # view names keep white
                             exclude=[w for _, w in PLANES + QC_PLANES + CONTACT_PLANES])
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    deck.save()
    print(f'saved {bpd.OUT} ({len(deck.prs.slides)} slides, {ppi:.1f} px/in)')
    for m in missing:
        print(f'  missing: {m}')
    if args.export:
        bpd.export_pngs(Path(args.export), deck.titles)


if __name__ == '__main__':
    main()
