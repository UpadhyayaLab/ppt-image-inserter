#!/usr/bin/env python3
r"""make_pmlc_rep_panels.py  Pick the representative CAT / FMC synapses (data-
driven) for the pMLC decks. Rendering is done by the pipeline-faithful MATLAB
renderer (render_rep_synapse_panels.m via tests/run_pmlc_rep_render.m), which
uses the SAME contrast as the pipeline's composites / montages.

SELECTION (FF 2026-09-14: "reflective of general trends"): within one replicate
and timepoint, each condition's QC-passing cells are scored by how close they
sit to that condition's median on the metrics the deck reports -- pMLC central /
peripheral / distal MFI-norm, raw pMLC synapse MFI and synapse (actin mask) area
-- using robust z = |x - median| / (1.4826 MAD), summed. The lowest score wins;
the ranked candidates go to candidates_<tag>.csv (audit / swap via OVERRIDE).

Output: <NAS>\results_compiled\pmlc_joint\paper_reps_<yyyymmdd>\
          picks.csv (tag, cond, cond_dir, cell, config), candidates_<tag>.csv
Run:    python make_pmlc_rep_panels.py   then   run_pmlc_rep_render (MATLAB)
"""
from __future__ import annotations

import csv
import datetime
import sys
from pathlib import Path

import numpy as np
from scipy.io import loadmat

NAS = Path(r'\\10.225.56.176\ULab_data\User_data\Kiet')
OUT = NAS / 'results_compiled' / 'pmlc_joint' / f'paper_reps_{datetime.date.today():%Y%m%d}'
ROOTS = {'Mar_12': ('03122026_pMLC_actin_CAR_T', 'd20260312_actin_pmlc_nucleus_pmlc_config'),
         'May_10': ('20260510_pMLC_Actin561_nucleus_CAR_Tcell_', 'd20260510_actin_pmlc_nucleus_pmlc_config'),
         'Jun_07': ('20260607_pMLC_CART_actin_hoescht', 'd20260607_actin_pmlc_nucleus_pmlc_config')}

DATE = 'Jun_07'                    # replicate used for the representative pair
TIMEPOINT = '10min'
TAGS = ['CAT', 'FMC']
# FF 2026-09-14: chosen by FF from the candidate picker gallery (ranks: CAT #1, FMC #3).
OVERRIDE = {'CAT': '043_166', 'FMC': '061_324'}
SCORE_FIELDS = ['pMLC_synapse_central_MFI_norm', 'pMLC_synapse_peripheral_MFI_norm',
                'pMLC_synapse_distal_MFI_norm', 'pMLC_synapse_MFI', 'actin_bottom_mask_area']


def checked_pick(tag, override, rows):
    """The override for `tag` if FF set one, else the rank-1 cell -- and the override MUST be
    in `rows`, the QC-passing ranked set (cell_found, not skipped, finite scores). Unchecked,
    a cell dropped or renumbered by a reprocess still rendered as the 'representative' cell
    (bug hunt 2026-10-24, second pass)."""
    ranked = [r[0] for r in rows]
    if tag not in override:
        return ranked[0]
    cell = override[tag]
    if cell not in ranked:
        raise SystemExit(f'[ERROR] override {tag}={cell!r} is not in the QC-passing ranked set '
                         f'({len(ranked)} cells); re-pick it from the candidate gallery')
    print(f'  override {tag}={cell} (rank {ranked.index(cell) + 1} of {len(ranked)})')
    return cell


def pick_then_write(out, ranked, fields, override):
    """Pick every tag (checked_pick), and only then create `out` with candidates_<tag>.csv.

    ranked: {tag: (rows, med)}. `out` was created before the picks could fail, so a failed
    override left an empty paper_reps_<today> that newest_run then took as the newest reps
    folder (bug hunt 2026-10-26). On failure the candidates still go to disk, for re-picking,
    under failed_<out name>, which the decks' 'paper_reps_*' globs never match."""
    try:
        picks = {tag: checked_pick(tag, override, rows) for tag, (rows, _) in ranked.items()}
    except SystemExit:
        failed = out.with_name('failed_' + out.name)
        _write_candidates(failed, ranked, fields)
        print(f'candidates for re-picking: {failed}')
        raise
    _write_candidates(out, ranked, fields)
    return picks


def _write_candidates(out, ranked, fields):
    out.mkdir(parents=True, exist_ok=True)
    for tag, (rows, med) in ranked.items():
        with open(out / f'candidates_{tag}.csv', 'w', newline='') as fh:
            w = csv.writer(fh)
            w.writerow(['cell', 'score_sum_robust_z'] + fields)
            w.writerow(['MEDIAN', ''] + [f'{x:.4g}' for x in med])
            for r in rows[:40]:
                w.writerow([r[0], f'{r[1]:.3f}'] + [f'{x:.4g}' for x in r[2:]])


def cond_dir(date, cond):
    return NAS / ROOTS[date][0] / cond / 'converted' / 'cropped' / 'split_channels'


def rank_cells(date, cond):
    m = loadmat(str(cond_dir(date, cond) / 'CART_results.mat'),
                variable_names=['cell_label', 'cell_found', 'cell_skipped'] + SCORE_FIELDS)
    labels = [str(np.ravel(x)[0]) if np.size(x) else '' for x in m['cell_label'].ravel()]
    ok = (m['cell_found'].ravel() == 1) & (m['cell_skipped'].ravel() == 0)
    ok &= np.array([lab != '' for lab in labels])
    X = np.column_stack([m[f].ravel().astype(float) for f in SCORE_FIELDS])
    ok &= np.all(np.isfinite(X), axis=1)
    med = np.median(X[ok], axis=0)
    mad = np.median(np.abs(X[ok] - med), axis=0) * 1.4826
    score = np.where(ok, (np.abs(X - med) / np.where(mad > 0, mad, 1)).sum(axis=1), np.inf)
    rows = [(labels[i], score[i], *X[i]) for i in np.argsort(score) if np.isfinite(score[i])]
    return rows, med


def main_pzap70():
    """--marker pZap70 (FF 2026-09-15): same robust-z near-median pick for the TIRF
    pZap70 set, Feb 2025 D3 at 15 min (largest n; the CAT/FMC zone shift is clearest at
    15 min). Rank 1 goes to rep_D3_<tag>_15min_*; ranks 2-4 are rendered as
    rep_alt<r>_D3_* so FF can swap without a new pick run. Rank-1 rows are written LAST
    (FMC last) so the renderer's single scale bar lands on the FMC zones image."""
    root = Path(r'L:\FF\CAR T\CAR T data for actin area trouble shoot')
    out = root / 'results_compiled_tirf' / f'pzap70_reps_{datetime.date.today():%Y%m%d}'
    base = root / 'pZap70-Actin_CAT_FMC' / '20250217_fixed_D3_488pzap70_CART_TIRF'
    folders = {'CAT': 'CAT_15min', 'FMC': 'FMC63_15min'}          # D3 uses FMC63_* folders
    fields = ['pZap70_synapse_central_MFI_norm', 'pZap70_synapse_peripheral_MFI_norm',
              'pZap70_synapse_distal_MFI_norm', 'pZap70_synapse_MFI', 'actin_bottom_mask_area']
    override = {'CAT': '23', 'FMC': '68'}   # FF 2026-09-15: chosen from the pZap70 picker gallery
    cdirs = {tag: base / folders[tag] / 'converted' / 'cells' / 'channels' for tag in TAGS}
    ranked_rows = {tag: rank_rows(cdirs[tag], fields) for tag in TAGS}
    chosen = pick_then_write(out, ranked_rows, fields, override)
    alts, firsts = [], []
    for tag in TAGS:
        cdir = cdirs[tag]
        ranked = [r[0] for r in ranked_rows[tag][0]]
        first = chosen[tag]
        mk = lambda cond, cell: {'tag': tag, 'cond': cond, 'cond_dir': str(cdir), 'cell': cell,  # noqa: E731
                                 'config': 'd20250217_actin_pzap70_pzap70_config'}
        firsts.append(mk(f'D3_{tag}_15min', first))
        alts += [mk(f'alt{k}_D3_{tag}_15min', c) for k, c in
                 enumerate([c for c in ranked if c != first][:3], start=2)]
        print(f'[pick] {tag}: {first}  (alternates {[c for c in ranked if c != first][:3]})')
    with open(out / 'picks.csv', 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=['tag', 'cond', 'cond_dir', 'cell', 'config'])
        w.writeheader(); w.writerows(alts + firsts)
    for old in out.glob('rep_*.png'):
        old.unlink()
    print(f'-> {out / "picks.csv"}')
    return 0


ACTIN_OVERRIDE = {'CAT': '028_140', 'FMC': '72_606'}   # FF 2026-09-29: chosen from the actin picker gallery


def main_actin():
    """--marker actin (FF 2026-09-29): actin-only representative CAT / FMC pair for the
    Jul 17 2026 confocal 3/5/7/10 min set (Kiet; NAS 'New folder' reprocess), 10 min.
    Robust-z near-median pick on the RELATIVE zone actin (zone MFI / whole-synapse MFI,
    actin_bottom_slice_<zone>_MFI_norm) + synapse area. Writes candidates_<tag>.csv (top
    40, for the picker gallery) and picks.csv (rank 1 until FF chooses; ACTIN_OVERRIDE)."""
    nf = Path(r'\\10.225.56.176\ULab_data\Data_to_share\FF\New folder')
    out = nf / 'results_compiled' / f'actin_reps_{datetime.date.today():%Y%m%d}'
    base = nf / '20260717_CART_3_5_7_10_actin'
    fields = ['actin_bottom_slice_central_MFI_norm', 'actin_bottom_slice_peripheral_MFI_norm',
              'actin_bottom_slice_distal_MFI_norm', 'actin_bottom_mask_area']
    cdirs = {tag: base / f'{tag}_10min' / 'split_channels' for tag in TAGS}
    ranked_rows = {tag: rank_rows(cdirs[tag], fields, 'actin_only_results.mat') for tag in TAGS}
    chosen = pick_then_write(out, ranked_rows, fields, ACTIN_OVERRIDE)
    picks = []
    for tag in TAGS:
        cell = chosen[tag]
        picks.append({'tag': tag, 'cond': f'Jul_17_{tag}_10min', 'cond_dir': str(cdirs[tag]), 'cell': cell,
                      'config': 'd20260717_actin_irm_newfolder_config'})
        print(f'[pick] {tag}: {cell}  (n ranked {len(ranked_rows[tag][0])})')
    with open(out / 'picks.csv', 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=['tag', 'cond', 'cond_dir', 'cell', 'config'])
        w.writeheader(); w.writerows(picks)
    # Stale renders must not be re-adopted (as the pMLC / pZap70 branches): a same-day
    # re-pick left two cells' rep_*.png side by side (bug hunt 2026-10-24).
    for old in out.glob('rep_*.png'):
        old.unlink()
    print(f'-> {out / "picks.csv"}')
    return 0


def rank_rows(cdir, fields, results='CART_results.mat'):
    """Shared robust-z ranking (same as rank_cells) for an explicit condition dir."""
    m = loadmat(str(cdir / results),
                variable_names=['cell_label', 'cell_found', 'cell_skipped'] + fields)
    labels = [str(np.ravel(x)[0]) if np.size(x) else '' for x in m['cell_label'].ravel()]
    ok = np.array([lab != '' for lab in labels])
    if 'cell_found' in m:
        ok &= m['cell_found'].ravel() == 1
    if 'cell_skipped' in m:
        ok &= m['cell_skipped'].ravel() == 0
    X = np.column_stack([m[f].ravel().astype(float) for f in fields])
    ok &= np.all(np.isfinite(X), axis=1)
    med = np.median(X[ok], axis=0)
    mad = np.median(np.abs(X[ok] - med), axis=0) * 1.4826
    score = np.where(ok, (np.abs(X - med) / np.where(mad > 0, mad, 1)).sum(axis=1), np.inf)
    return [(labels[i], score[i], *X[i]) for i in np.argsort(score) if np.isfinite(score[i])], med


def main():
    if '--marker' in sys.argv and sys.argv[sys.argv.index('--marker') + 1] == 'pZap70':
        return main_pzap70()
    if '--marker' in sys.argv and sys.argv[sys.argv.index('--marker') + 1] == 'actin':
        return main_actin()
    ranked_rows = {tag: rank_cells(DATE, f'{tag}_{TIMEPOINT}') for tag in TAGS}
    chosen = pick_then_write(OUT, ranked_rows, SCORE_FIELDS, OVERRIDE)
    picks = []
    for tag in TAGS:
        cond = f'{tag}_{TIMEPOINT}'
        cell = chosen[tag]
        picks.append({'tag': tag, 'cond': f'{DATE}_{cond}', 'cond_dir': str(cond_dir(DATE, cond)),
                      'cell': cell, 'config': ROOTS[DATE][1]})
        print(f'[pick] {tag}: {cell}')
    with open(OUT / 'picks.csv', 'w', newline='') as fh:
        w = csv.DictWriter(fh, fieldnames=['tag', 'cond', 'cond_dir', 'cell', 'config'])
        w.writeheader(); w.writerows(picks)
    for old in OUT.glob('rep_*.png'):          # stale renders must not be re-adopted
        old.unlink()
    print(f'-> {OUT / "picks.csv"}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
