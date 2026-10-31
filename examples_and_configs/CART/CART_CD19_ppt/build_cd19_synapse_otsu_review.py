#!/usr/bin/env python3
r"""build_cd19_synapse_otsu_review.py  Review deck: the synapse at the old MAD cut vs Otsu.

FF 2026-10-31 (switching the volscope synapse cut to Otsu: "we need a smoke with a small
number of cells first ... visually verify every single one ... prepare something for me to
review"): one slide per smoke cell -- CD19 relative density | the pipeline's synapse
(median + 2 x MAD) | the Otsu synapse from the smoke -- as native rotatable 3D models, the
mesh turned so the two patches face the viewer. Every mesh and color is
build_cd19_methods_deck's (mesh_glbs, colorbar, add_models).

Inputs: <smoke>/picks.json (cells, written with the smoke) and <smoke>/otsu/<exp>/<cond>/
(volscope matlab/volscope_cd19_synapse.m, default cut, on copies of those cells' meshes);
the MAD synapse is each condition's pipeline cd19_synapse/faces/.

Run:  python build_cd19_synapse_otsu_review.py [--smoke DIR] [--notes FILE]
Out:  <smoke>/review/synapse_otsu_review.pptx + deck_png/
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

import numpy as np
from scipy.io import loadmat

import build_cd19_methods_deck as md
from build_cd19_methods_deck import bpd

ROOT = Path(r'Y:\User_data\Kiet')
EXPS = {'Oct1': (ROOT / '20261001_coculture_10_30_60', 'Oct 1, 2026'),
        'Mar10': (ROOT / '03102026_coculture_hek_CD19gfp_Hoescht_CAR_HA_APC', 'Mar 10, 2026')}
SMOKE = ROOT / '20261001_coculture_10_30_60' / 'tests' / 'synapse_otsu_smoke_20261031'
TAGS = {'min': 'smallest', 'median': 'median', 'max': 'largest'}


def cond_label(cond):
    m = re.match(r'(CAT|FMC)_(\d+)min', cond)
    return f'{m.group(1)} {m.group(2)} min' if m else cond


def faces(path):
    return np.asarray(loadmat(str(path), variable_names=['synapse_faces'])
                      ['synapse_faces']).reshape(-1) > 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--smoke', type=Path, default=SMOKE)
    ap.add_argument('--notes', type=Path, default=None,
                    help='JSON {"<exp>/<cond>/<label>": "note"}: a line under each slide, '
                         'plus a summary slide of the noted cells')
    ap.add_argument('--reuse-glb', action='store_true',
                    help='use the GLBs a previous build left in review/glb (skips ~10 min of '
                         'meshing; the 3D models themselves take ~16 s each in PowerPoint)')
    args = ap.parse_args()
    picks = json.load(open(args.smoke / 'picks.json'))
    notes = json.load(open(args.notes, encoding='utf-8')) if args.notes else {}
    out = args.smoke / 'review'
    bpd.OUT = out / 'synapse_otsu_review.pptx'
    if (out / f'~${bpd.OUT.name}').exists():
        raise SystemExit(f'{bpd.OUT.name} is open in PowerPoint -- close it first')

    deck = bpd.Deck()
    if notes:
        s = deck.slide('Cells to look at')
        lines = [f'{EXPS[p["exp"]][1]}, {cond_label(p["cond"])}, {p["label"]}: '
                 f'{notes[k]}' for p in picks
                 if (k := f'{p["exp"]}/{p["cond"]}/{p["label"]}') in notes]
        bpd.add_text(s, 0.5, 0.9, bpd.SLIDE_W - 1, 6.4, lines or ['none'], 18, bold=True)
    w, h, top = 3.75, 4.6, 1.75                     # rel. density | bar | MAD | Otsu
    xs = (0.1, 5.55, 9.5)
    models = []
    for p in picks:
        vdir = EXPS[p['exp']][0] / p['cond'] / 'converted' / 'cells' / 'channels' / 'volscope'
        mad_mat = vdir / 'cd19_synapse' / 'faces' / f'{p["label"]}_synapse.mat'
        otsu_mat = args.smoke / 'otsu' / p['exp'] / p['cond'] / 'faces' / f'{p["label"]}_synapse.mat'
        view = faces(mad_mat) | faces(otsu_mat)       # both patches toward the viewer
        key = f'{p["exp"]}/{p["cond"]}/{p["label"]}'
        g_mad, g_otsu = (_glbs(out / 'glb' / key / v, p['label'], m, vdir, view, args.reuse_glb)
                         for v, m in (('mad', mad_mat), ('otsu', otsu_mat)))
        s = deck.slide(f'{EXPS[p["exp"]][1]}, {cond_label(p["cond"])}, {p["label"]} '
                       f'({TAGS[p["tag"]]} old synapse fraction)')
        n = len(deck.prs.slides)
        md.header(s, xs[0], 1.0, w, 0.45, [('CD19', md.RGB['CD19']), (' relative density', None)], 24)
        md.header(s, xs[1], 1.0, w, 0.45, [('Old (MAD): ', None),
                                           (f'{100 * p["mad_frac"]:.0f}%', md.RGB['CD19'])], 24)
        md.header(s, xs[2], 1.0, w, 0.45, [('Otsu: ', None),
                                           (f'{100 * p["otsu_frac"]:.0f}%', md.RGB['CD19'])], 24)
        md.colorbar(s, 4.45, 3.0, 2.6, [(0, '\u22641\u00d7'), (1 / 3, '2\u00d7'),
                                       (2 / 3, '3\u00d7'), (1, '\u22654\u00d7')],
                    ['CD19 relative', 'density'])
        if key in notes:
            bpd.add_text(s, 0.4, 6.6, bpd.SLIDE_W - 0.8, 0.6, notes[key], 20, bold=True,
                         color=bpd.GREY)
        for x, path in zip(xs, (g_mad['cd19'], g_mad['synapse'], g_otsu['synapse'])):
            models.append(dict(path=str(path), slide=n, left=x * 72, top=top * 72,
                               width=w * 72, height=h * 72, rx=0, ry=0, rz=0))
    deck.save()
    md.add_models(bpd.OUT, 1, models)
    bpd.export_pngs(out / 'deck_png', deck.titles)
    print(f'saved {bpd.OUT} ({len(deck.prs.slides)} slides)')


def _glbs(d, label, syn_mat, vdir, view, reuse):
    """md.mesh_glbs into D, or (REUSE) the GLBs it wrote there on an earlier build."""
    old = {k: d / f'{label}_mesh_{k}.glb' for k in ('cd19', 'synapse')}
    if reuse and all(f.exists() for f in old.values()):
        return old
    d.mkdir(parents=True, exist_ok=True)
    return md.mesh_glbs(d, syn_mat=syn_mat, vdir=vdir, label=label, face_view=view)[0]


if __name__ == '__main__':
    main()
