#!/usr/bin/env python3
r"""build_cd19_synapse_k_compare.py  The methods-deck cell's synapse at several cuts, same view.

FF 2026-10-31 (methods deck slide 7: "the synapse detection is extending a bit far"; "show me
mesh at that angle (direct comparison)"): CD19 relative density beside the detected synapse at
median + k x MAD for each k of a volscope synapse sweep (volscope matlab/volscope_cd19_synapse.m
with opts.high_k or opts.cut = 'otsu', run into <exp>/tests/synapse_k_sweep_<date>/
<variant>/<cond>/, variant k<k> or otsu). Every mesh and
color is build_cd19_methods_deck's (mesh_glbs, add_models, colorbar), at the deck's opening
view, so the panels compare directly with the deck.

Run:  python build_cd19_synapse_k_compare.py [--sweep DIR] [--variants k2,k4,otsu]
Out:  <sweep>/compare/synapse_k_compare.pptx + deck_png/
"""
from __future__ import annotations

import argparse
import csv
from pathlib import Path

import build_cd19_methods_deck as md
from build_cd19_methods_deck import bpd

SWEEP = md.EXP / 'volscope' / 'tests' / 'synapse_k_sweep_20261031'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--sweep', type=Path, default=SWEEP)
    ap.add_argument('--variants', default='k2,k4,otsu', help='sweep subfolders: k<k> or otsu')
    ap.add_argument('--rx', type=float, default=20)       # the methods deck's opening view
    ap.add_argument('--ry', type=float, default=40)
    ap.add_argument('--rz', type=float, default=90)
    args = ap.parse_args()
    ks = args.variants.split(',')
    if len(ks) != 3:
        raise SystemExit('--variants: three (one reference + three cuts fill the 2 x 2 grid)')
    out = args.sweep / 'compare'
    bpd.OUT = out / 'synapse_k_compare.pptx'
    if (out / f'~${bpd.OUT.name}').exists():
        raise SystemExit(f'{bpd.OUT.name} is open in PowerPoint -- close it first')

    glbs, rows, g = {}, {}, {}
    for k in ks:
        d = args.sweep / k / md.COND
        glbs[k], g[k] = md.mesh_glbs(_mkdir(out / 'glb' / k),
                                  syn_mat=d / 'faces' / f'{md.LABEL}_synapse.mat')
        rows[k] = next(r for r in csv.DictReader(open(d / 'cd19_synapse.csv', newline=''))
                       if r['label'] == md.LABEL)

    deck = bpd.Deck()
    s = deck.slide('Detected synapse at three thresholds, same cell, same view')
    names = {k: ('Otsu' if k == 'otsu' else f'median + {k[1:]} × MAD') for k in ks}
    xl, xr, w, h = 0.2, 7.5, 5.6, 2.85                   # 2 x 2: reference + three cuts
    cells = [(xl, 1.15), (xr, 1.15), (xl, 4.45), (xr, 4.45)]
    md.header(s, xl, 0.7, w, 0.45, [('CD19', md.RGB['CD19']), (' relative density', None)], 24)
    md.colorbar(s, 6.05, 1.95, 2.0, [(0, '≤1×'), (1 / 3, '2×'),
                                     (2 / 3, '3×'), (1, '≥4×')],
                ['CD19 relative', 'density'])
    for (x, y), k in zip(cells[1:], ks):
        r = rows[k]
        md.header(s, x, y - 0.45, w, 0.45,
                  [(f'{names[k]} = {float(r["synapse_face_thresh"]):.0f}: ', None),
                   (f'{float(r["synapse_area_um2"]):.0f} μm², '
                    f'{100 * float(r["synapse_area_frac"]):.0f}%', md.RGB['CD19'])], 24)
    # 2: raw vs smoothed CD19 on the methods deck's slide-6 geometry and 0..vmax scale, the
    # last variant's cut on the bar (FF 2026-10-31: is the smoothing what overreaches?)
    last = ks[-1]
    s = deck.slide('CD19 near the cell surface, raw and smoothed, same view')
    xl2, wl2, xb2, xr2, wr2 = 0.3, 5.3, 5.8, 7.1, 5.9
    md.header(s, xl2, 0.95, wl2, 0.45, [('CD19', md.RGB['CD19']), (' near the cell surface', None)], 28)
    md.header(s, xr2, 0.95, wr2, 0.45, [('Smoothed ', None), ('CD19', md.RGB['CD19'])], 28)
    vmax = g[last]['vmax']
    md.colorbar(s, xb2, 2.8, 3.25, [(0, '0'), (0.5, f'{vmax / 2:.0f}'), (1, f'≥{vmax:.0f}')],
                ['CD19 (a.u.)'], thresh_frac=min(1.0, g[last]['thresh'] / vmax),
                thresh_label=names[last])
    deck.save()

    box = dict(width=w * 72, height=h * 72, rx=args.rx, ry=args.ry, rz=args.rz)
    paths = [glbs[ks[0]]['cd19']] + [glbs[k]['synapse'] for k in ks]
    md.add_models(bpd.OUT, 1, [dict(path=str(p), left=x * 72, top=y * 72, **box)
                               for p, (x, y) in zip(paths, cells)])
    box2 = dict(top=1.5 * 72, height=4.8 * 72, rx=args.rx, ry=args.ry, rz=args.rz)
    md.add_models(bpd.OUT, 2, [dict(path=str(glbs[last]['raw']), left=xl2 * 72, width=wl2 * 72, **box2),
                               dict(path=str(glbs[last]['smooth']), left=xr2 * 72, width=wr2 * 72,
                                    **box2)])
    bpd.export_pngs(out / 'deck_png', deck.titles)
    print(f'saved {bpd.OUT}')


def _mkdir(p):
    p.mkdir(parents=True, exist_ok=True)
    return p


if __name__ == '__main__':
    main()
