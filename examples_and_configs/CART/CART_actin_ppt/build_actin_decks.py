#!/usr/bin/env python3
r"""build_actin_decks.py  Rebuild the two 2026-07 confocal actin-only CAR-T decks (Kiet;
Jul 17 + Jul 23, 2026; CAT / FMC x 3/5/7/10 min) from the Sep 2-3 2026 NAS reprocess.

The Aug 11 decks (CART_actin_summary_20260717_20260723.pptx, CART_actin_montages_...pptx)
pulled panels from H:\results_compiled\actin_only\..., which is offline; their builder was
never saved. This rebuilds both with the same slide sets from the relocated data under
  \\10.225.56.176\ULab_data\Data_to_share\FF\New folder\
and adds (FF 2026-09-15): actin ray profiles + gradients CAT vs FMC, time-course profile
overlays within each cell type, normalized and RAW zone MFI violins (central / peripheral /
distal; per experiment, never pooled), the Dalby clearance / perimeter / foci-density
metrics, and actin clearance montages. Titles + sparse labels only (no path captions).

Inputs (nothing is plotted here):
  newest New folder\results_compiled\compiled_20260717_20260723_2*\   <- compile_cart(
        'd20260717_20260723_actin_only_newfolder_compilation_config')
        (tests/run_actin_newfolder_compile.m)
  newest New folder\results_compiled\actin_extended_panels_*\         <- tests/run_actin_extended_panels.m
  <dataset>\<COND>\split_channels\prog_fixed_cells\...\montages_filtered\ (first page)

Output: K:\FF\PPT\PPT_autogeneration\CART\actin_only\
          CART_actin_summary_20260717_20260723_<yyyymmdd>.pptx
          CART_actin_montages_20260717_20260723_<yyyymmdd>.pptx
Run:    python build_actin_decks.py [--deck summary|montages|both] [--export <dir>]
PowerPoint is only touched by --export (never quit; skipped if FF has the deck open).
"""
from __future__ import annotations

import argparse
import datetime
import glob
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'CART_pMLC_ppt'))
sys.path.insert(0, r'K:\FF\Code\Scripts\PPT_image_inserter\examples_and_configs')
import build_pmlc_profiles_deck as bpd  # noqa: E402  (shared Deck / fit_in_box / export_pngs)
import insert_CART_CatB_MT_actin_nuc_montages_slides as mont  # noqa: E402  (one-scale placement)
from pptx.enum.text import PP_ALIGN  # noqa: E402

# mont._winlong prefixes '\\?\' to a UNC path, which Windows rejects (UNC needs '\\?\UNC\');
# the NAS montage paths stay under 260 chars, so use them as-is.
mont._winlong = str

# Montage slides: CAT left / FMC right boxes under the labels.
MONT_TOP = 1.05
MONT_HALF = (bpd.SLIDE_W - 0.3) / 2
MONT_H = bpd.SLIDE_H - 1.12


def montage_group_ppi(groups):
    """groups: {scale group: [png, ...]} -> {scale group: px/in}. ONE px/in per group, pinned
    to the largest montage in it (insert_CART_CatB_MT_actin_nuc_montages_slides.py rule), so
    every slide of a group -- CAT, FMC, every dataset and timepoint -- shows one physical
    scale. Aspect-fitting each montage to its box instead shrinks each by its own factor."""
    out = {}
    for group, pngs in groups.items():
        pngs = [p for p in pngs if p is not None]
        if pngs:
            out[group] = mont.compute_group_ppi(pngs, MONT_HALF, MONT_H)
            print(f'  scale group {group:<28s} {out[group]:7.1f} px/in')
    return out


def place_montage(slide, png, left, ppi, top=MONT_TOP, box_w=MONT_HALF, box_h=MONT_H):
    """Montage centered in its box at ppi; a missing one gets fit_in_box's [missing] note."""
    if png is None or ppi is None:
        bpd.fit_in_box(slide, png, left, top, box_w, box_h)
    else:
        mont.add_image_at_ppi(slide, png, ppi, left, top, box_w, box_h)

NF = Path(r'\\10.225.56.176\ULab_data\Data_to_share\FF\New folder')
RC = NF / 'results_compiled'
OUT_DIR = Path(r'K:\FF\PPT\PPT_autogeneration\CART\actin_only')
TODAY = f'{datetime.date.today():%Y%m%d}'
DASH, MINUS = '\u2014', '\u2212'

TPS = ['3min', '5min', '7min', '10min']
TP_LBL = ['3 min', '5 min', '7 min', '10 min']
DATASETS = [('Jul_17', 'Jul 17', '20260717', '20260717_CART_3_5_7_10_actin'),   # tag, label, date, folder
            ('Jul_23', 'Jul 23', '20260723', '20260723_fix_actin_CAR_3_5_7_10')]

# Aug 11 summary deck slide set (title, grid panel relative to <compile>\grid_panels\).
TS = 'timecourse_scatter_plots'
SUMMARY_OLD = [
    ('Synapse Area', f'{TS}/actin_bottom_mask_area_grid.png'),
    ('Synapse Area (IRM)', f'{TS}/IRM_mask_area_grid.png'),
    ('Actin Synapse Inner/Outer Ratio (50% Eff Rad Thresh)', f'{TS}/actin_bottom_slice_inner_outer_ratio_50pct_grid.png'),
    ('Actin Synapse Inner/Outer Ratio (70% Eff Rad Thresh)', f'{TS}/actin_bottom_slice_inner_outer_ratio_70pct_grid.png'),
    ('Actin Synapse R_median', f'{TS}/actin_bottom_slice_r_median_grid.png'),
    ('Actin Synapse Radial Profile AUC\u2081 (All Cells + Avg)',
     'CAT_vs_FMC_by_timepoint/actin_bottom_slice_rad_profile_auc1_all_cells_with_average_grid.png'),
    ('Actin Synapse MFI', f'{TS}/actin_bottom_MFI_grid.png'),
    ('Actin Synapse Total Signal', f'{TS}/actin_bottom_total_sig_grid.png'),
    ('Actin Synapse Perimeter', f'{TS}/actin_bottom_slice_perimeter_grid.png'),
    ('Actin Synapse Circularity', f'{TS}/actin_bottom_slice_circularity_grid.png'),
    ('Actin Synapse Solidity', f'{TS}/actin_bottom_slice_solidity_grid.png'),
    ('Actin Synapse Eccentricity', f'{TS}/actin_bottom_slice_eccentricity_grid.png'),
    ('Synapse (3-Slice MIP) Area', f'{TS}/actin_bottom_3slice_mask_area_grid.png'),
    ('Actin Synapse (3-Slice MIP) Inner/Outer Ratio (50% Eff Rad Thresh)', f'{TS}/actin_bottom_3slice_mip_inner_outer_ratio_50pct_grid.png'),
    ('Actin Synapse (3-Slice MIP) Inner/Outer Ratio (70% Eff Rad Thresh)', f'{TS}/actin_bottom_3slice_mip_inner_outer_ratio_70pct_grid.png'),
    ('Actin Synapse (3-Slice MIP) R_median', f'{TS}/actin_bottom_3slice_mip_r_median_grid.png'),
    ('Actin Synapse (3-Slice MIP) Radial Profile AUC\u2081 (All Cells + Avg)',
     'CAT_vs_FMC_by_timepoint/actin_bottom_3slice_mip_rad_profile_auc1_all_cells_with_average_grid.png'),
    ('Actin Synapse (3-Slice MIP) MFI', f'{TS}/actin_bottom_3slice_MFI_grid.png'),
    ('Actin Synapse (3-Slice MIP) Total Signal', f'{TS}/actin_bottom_3slice_total_sig_grid.png'),
    ('Actin Synapse (3-Slice MIP) Perimeter', f'{TS}/actin_bottom_3slice_perimeter_grid.png'),
    ('Actin Synapse (3-Slice MIP) Circularity', f'{TS}/actin_bottom_3slice_circularity_grid.png'),
    ('Actin Synapse (3-Slice MIP) Solidity', f'{TS}/actin_bottom_3slice_solidity_grid.png'),
    ('Actin Synapse (3-Slice MIP) Eccentricity', f'{TS}/actin_bottom_3slice_eccentricity_grid.png'),
    ('Actin Foci Count', f'{TS}/actin_foci_count_grid.png'),
    ('Actin Foci Area (\u03bcm\u00b2)', f'{TS}/actin_foci_area_um2_grid.png'),
    ('Actin Foci Area Fraction (of Synapse)', f'{TS}/actin_foci_area_fraction_grid.png'),
    ('Actin Foci Mean Intensity', f'{TS}/actin_foci_mean_intensity_grid.png'),
    ('Actin Foci Total Intensity', f'{TS}/actin_foci_total_intensity_grid.png'),
    ('Actin Foci Max Normalized Score', f'{TS}/actin_foci_max_norm_score_grid.png'),
    ('Actin Foci Mean Normalized Score', f'{TS}/actin_foci_mean_norm_score_grid.png'),
    ('Actin Foci Detection Threshold Used', f'{TS}/actin_foci_threshold_used_grid.png'),
]
# Metrics added since Aug 11 (Dalby et al. 2026 clearance / perimeter / foci density).
SUMMARY_NEW = [
    ('Actin Clearance Area', f'{TS}/actin_bottom_slice_clearance_area_um2_grid.png'),
    ('Actin Clearance Area Fraction', f'{TS}/actin_bottom_slice_clearance_frac_grid.png'),
    ('Actin Low-Intensity Area Fraction', f'{TS}/actin_bottom_slice_low_frac_grid.png'),
    ('Normalized Synapse Perimeter', f'{TS}/actin_bottom_slice_norm_perimeter_grid.png'),
    ('Actin Clearance Area (3-Slice MIP)', f'{TS}/actin_bottom_3slice_mip_clearance_area_um2_grid.png'),
    ('Actin Clearance Area Fraction (3-Slice MIP)', f'{TS}/actin_bottom_3slice_mip_clearance_frac_grid.png'),
    ('Actin Low-Intensity Area Fraction (3-Slice MIP)', f'{TS}/actin_bottom_3slice_mip_low_frac_grid.png'),
    ('Normalized Synapse Perimeter (3-Slice MIP)', f'{TS}/actin_bottom_3slice_mip_norm_perimeter_grid.png'),
    ('Actin Foci Density', f'{TS}/actin_foci_density_um2_grid.png'),
    ('Actin Foci per Clearance Area', f'{TS}/actin_foci_density_clearance_um2_grid.png'),
    ('Actin Foci Median Area', f'{TS}/actin_foci_median_area_um2_grid.png'),
]
# Montage families: (title stem, prog_fixed_cells sub-folder). The Aug 11 mapping was
# recovered by matching its slide images to these folders' first pages (r >= 0.9 except IRM).
MONTAGES = [
    ('Actin at Synapse', r'actin\synapse\inner_outer\1slice_combined'),
    ('Actin: 0, 2, and 4 Slices above Synapse', r'actin\synapse\bottom_slices_xy'),
    ('IRM Segmentation', r'IRM\seg'),
    ('Actin XZ MIP', r'actin\xz_mip'),
    ('Actin Axial Profile', r'actin\axial\total_signal'),
    ('Actin Foci Overlay', r'foci\final_overlay'),
    ('Actin Clearance', r'actin\synapse\clearance'),       # new since Aug 11
]


def newest(pattern):
    hits = sorted(p for p in glob.glob(str(RC / pattern)) if Path(p).is_dir())
    return Path(hits[-1]) if hits else None


def tight_grid(deck, title, pngs, col_labels, row_labels=None, key=False):
    """Panels as large as the slide allows: column headers above every column, optional narrow
    row labels, no subtitle. Rows = len(pngs) / len(col_labels)."""
    s = deck.slide(title)
    ncol = len(col_labels)
    nrow = -(-len(pngs) // ncol)
    lab_w = 0.95 if row_labels else 0.05
    head_h, gap, top0 = 0.38, 0.06, 0.58
    bw = (bpd.SLIDE_W - lab_w - 0.05) / ncol
    bh = (bpd.SLIDE_H - top0 - head_h - 0.03) / nrow
    # Panels share one aspect: size them to the tighter of width / height, then pack the
    # header + rows as one block under the title so wide panels leave no gaps between rows.
    ars = []
    for p in pngs:
        if p and Path(p).is_file():
            with bpd.Image.open(p) as im:
                ars.append(im.size[0] / im.size[1])
    ar = max(ars) if ars else bw / bh
    pw = min(bw - 0.04, (bh - gap) * ar)
    ph = pw / ar
    top0 += max(0.0, (bpd.SLIDE_H - top0 - 0.03 - head_h - nrow * ph - (nrow - 1) * gap) / 2)
    top = top0 + head_h
    for c, lab in enumerate(col_labels):
        bpd.add_text(s, lab_w + c * bw, top0 - 0.02, bw, head_h, lab, 18, bold=True, align=PP_ALIGN.CENTER)
    for r, lab in enumerate(row_labels or []):
        bpd.add_text(s, 0.02, top + r * (ph + gap) + ph / 2 - 0.25, lab_w, 0.5, lab, 18, bold=True)
    for i, p in enumerate(pngs):
        r, c = divmod(i, ncol)
        bpd.fit_in_box(s, p, lab_w + c * bw + (bw - pw) / 2, top + r * (ph + gap), pw, ph)
    if key:
        bpd.add_text(s, bpd.SLIDE_W - 2.3, 0.1, 1.0, 0.5, 'CAT', 22, bold=True, color=bpd.CAT_BLUE, align=PP_ALIGN.RIGHT)
        bpd.add_text(s, bpd.SLIDE_W - 1.25, 0.1, 1.0, 0.5, 'FMC', 22, bold=True, color=bpd.FMC_RED)
    return s


def single(deck, title, png):
    s = deck.slide(title)
    bpd.fit_in_box(s, png, 0.3, 0.75, bpd.SLIDE_W - 0.6, bpd.SLIDE_H - 0.85)


def build_summary(compile_dir, panels, export):
    bpd.OUT = OUT_DIR / f'CART_actin_summary_20260717_20260723_{TODAY}.pptx'
    print(f'compile: {compile_dir}\npanels:  {panels}')
    deck = bpd.Deck()
    P = (lambda n: panels / 'profiles' / n) if panels else (lambda n: None)  # noqa: E731
    V = (lambda n: panels / 'violins' / n) if panels else (lambda n: None)  # noqa: E731
    G = lambda rel: compile_dir / 'grid_panels' / rel  # noqa: E731

    # New (FF 2026-09-15): actin profiles across conditions, then zone MFIs. FF 2026-09-16:
    # timepoint / condition headers above every column, experiments as rows, minimal margins.
    tcols = [lab for _, lab, _, _ in DATASETS]
    for tag, lab, _, _ in DATASETS:
        tight_grid(deck, f'Actin Profile and Gradient {DASH} {lab}',
                   [P(f'both_catfmc_profile_{tag}_{t}.png') for t in TPS] +
                   [P(f'both_catfmc_grad_{tag}_{t}.png') for t in TPS], TP_LBL)
    tight_grid(deck, 'Actin Profile Time Course',
               [P(f'tc_{ct}_{kind}_{tag}.png') for tag, _, _, _ in DATASETS
                for kind in ('profile', 'grad') for ct in ('CAT', 'FMC')],
               ['CAT Profile', 'FMC Profile', 'CAT Gradient', 'FMC Gradient'], tcols)
    tight_grid(deck, f'FMC {MINUS} CAT Actin Profile',
               [P(f'diff_both_profile_{t}.png') for t in TPS] +
               [P(f'diff_both_grad_{t}.png') for t in TPS], TP_LBL)
    tight_grid(deck, 'Actin by Zone: Normalized by MFI over Full Synapse',
               [V(f'zones_actin_{tag}_{t}.png') for tag, _, _, _ in DATASETS for t in TPS], TP_LBL, tcols, key=True)
    tight_grid(deck, 'Actin Zone MFI (Raw)',
               [V(f'zones_raw_actin_{tag}_{t}.png') for tag, _, _, _ in DATASETS for t in TPS], TP_LBL, tcols, key=True)
    tight_grid(deck, 'Actin Synapse MFI by Experiment', [V(f'synapse_mfi_{tag}.png') for tag, _, _, _ in DATASETS],
               ['', ''])   # panels carry their own experiment title + CAT / FMC legend

    for title, rel in SUMMARY_OLD + SUMMARY_NEW:
        single(deck, title, G(rel))
    for _, lab, _, _ in DATASETS:
        name = lab.replace(' ', '_') + '_2026'
        single(deck, f'Actin Axial Profile Overlay {DASH} per Slice ({lab}, 2026)',
               compile_dir / 'axial_profile_overlays' / f'axial_overlay_perslice_{name}.png')

    deck.save()
    print(f'saved {bpd.OUT} ({len(deck.prs.slides)} slides)')
    if export:
        bpd.export_pngs(Path(export) / 'summary', deck.titles)


def first_page(folder):
    """First montage page = lowest first-cell index in montage_cells_<fov>_<idx>_..."""
    fs = glob.glob(str(folder / 'montages_filtered' / '*.png'))
    if not fs:
        return None
    key = lambda f: [int(x) for x in re.findall(r'\d+', Path(f).stem)]  # noqa: E731
    return Path(min(fs, key=key))


def build_montages(export):
    bpd.OUT = OUT_DIR / f'CART_actin_montages_20260717_20260723_{TODAY}.pptx'
    # Pass 1: find every montage; one scale group per view (title stem).
    pngs = {(stem, folder, ct, t): first_page(NF / folder / f'{ct}_{t}' / 'split_channels' / 'prog_fixed_cells' / sub)
            for stem, sub in MONTAGES for _, _, _, folder in DATASETS for t in TPS for ct in ('CAT', 'FMC')}
    ppi = montage_group_ppi({stem: [p for k, p in pngs.items() if k[0] == stem] for stem, _ in MONTAGES})
    # Pass 2: place.
    deck = bpd.Deck()
    for stem, _ in MONTAGES:
        for _, _, date, folder in DATASETS:
            for t, tl in zip(TPS, TP_LBL):
                s = deck.slide(f'{stem} {DASH} {date}, {tl}')
                for k, (ct, col) in enumerate((('CAT', bpd.CAT_BLUE), ('FMC', bpd.FMC_RED))):
                    left = 0.1 + k * (MONT_HALF + 0.1)
                    bpd.add_text(s, left, 0.62, MONT_HALF, 0.4, ct, 20, bold=True, color=col, align=PP_ALIGN.CENTER)
                    place_montage(s, pngs[(stem, folder, ct, t)], left, ppi.get(stem))
    deck.save()
    print(f'saved {bpd.OUT} ({len(deck.prs.slides)} slides)')
    if export:
        bpd.export_pngs(Path(export) / 'montages', deck.titles)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--deck', default='both', choices=['summary', 'montages', 'both'])
    ap.add_argument('--compile', help='compile dir (default: newest compiled_20260717_20260723_2*)')
    ap.add_argument('--export', help='also export one PNG per slide (named by title) under this dir')
    args = ap.parse_args()
    if args.deck in ('summary', 'both'):
        compile_dir = Path(args.compile) if args.compile else newest('compiled_20260717_20260723_2*')
        panels = newest('actin_extended_panels_*')
        if compile_dir is None:
            print('[error] no compile found; run tests/run_actin_newfolder_compile.m'); return 1
        build_summary(compile_dir, panels, args.export)
    if args.deck in ('montages', 'both'):
        build_montages(args.export)
    return 0


if __name__ == '__main__':
    sys.exit(main())
