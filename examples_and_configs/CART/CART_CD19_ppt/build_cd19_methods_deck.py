#!/usr/bin/env python3
r"""build_cd19_methods_deck.py  Methods deck: segmentation and 3D analysis of one CAR T cell.

FF 2026-10-28 (for Kiet's Friday talk; Kiet: "showing the process of segmenting and 3D
visualization would be nice to show"): ONE example cell, not an outlier -- Oct 1, 2026
FMC_60min C2-cell72, the 60 min representative of the CAT vs FMC deck -- and the 3D mesh as a
NATIVE, rotatable PowerPoint 3D model.

Slides: title; raw data (CAR | CD19 | Hoechst | merge, synapse plane over synapse slice);
segmentation (ilastik vs final body, volscope matlab/volscope_mask_qc.m COMPARE sheet); the
same on every cell (a mask QC deck page); the 3D mesh (two rotatable models: CD19 near the
surface, and the detected synapse); what is measured (the synapse stage's QC render and the
internalized CD19 objects).

Inputs (all exports of the pipeline, drawn, never re-measured):
  tiles    volscope dev/contact_check_sheet.py --split   -> <exp>/cd19_figures/methods_<date>/
  seg      volscope_mask_qc(..., only={cond/label}, compare=true) -> <exp>/tests/methods_deck/
  mesh     <V>/meshes/<label>_mesh.mat (V µm (y,x,z), F, face_sig_sig_0_05um_skin)
  synapse  <V>/cd19_synapse/faces/<label>_synapse.mat (synapse_faces), qc/<label>_synapse.png
3D: PPT_image_inserter ppt_image_inserter/glb.py (build_colored_glb, lit)
and the Add3DModel COM pattern of models3d.py. The deck must be closed in PowerPoint; this
never kills POWERPNT.

Run:  python build_cd19_methods_deck.py [--export DIR]
Out:  K:\FF\PPT\PPT_autogeneration\CART\CD19\CART_coculture_fixed_methods_<today>.pptx
"""
from __future__ import annotations

import argparse
import csv
import datetime
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from scipy.io import loadmat

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / 'CART_pMLC_ppt'))
sys.path.insert(0, r'K:\FF\Code\Scripts\PPT_image_inserter')
import build_pmlc_profiles_deck as bpd  # noqa: E402
from ppt_image_inserter.glb import build_colored_glb  # noqa: E402

# The volscope widget's own functions, so the deck's colors are the widget's numbers (one
# implementation): the .mat seam, the relative density, the smoothed-field scale.
VOLSCOPE = Path(r'C:\Users\Upadhyaya_Lab\Code\volscope')
sys.path[:0] = [str(VOLSCOPE / 'viewers'), str(VOLSCOPE / 'src')]
from _matlab_mesh import load_mesh_mat  # noqa: E402
from mesh_widget import (RAMP, SIG_LIM, UNMEASURED, _face_enrichment,  # noqa: E402
                         _syn_field_bins)

EXP = Path(r'Y:\User_data\Kiet\20261001_coculture_10_30_60')
COND, LABEL, EXP_LABEL = 'FMC_60min', 'C2-cell72', 'Oct 1, 2026'
V = EXP / COND / 'converted' / 'cells' / 'channels' / 'volscope'
TILES = EXP / 'cd19_figures' / 'methods_20261028'
SEG = EXP / 'tests' / 'methods_deck' / COND / 'parts'
MIP_ROOT = EXP / 'tests' / 'methods_deck' / 'mip_contrast'    # volscope_mask_qc(..., MIP_PCT)
MIP_PCTS = (70,)               # FF 2026-10-29: the 80th slide dropped, "we do not need 3 slides"
# label-free renders (no cell number in any image; TCell dev/runs/run_methods_deck_nolabels.m)
NOLAB = EXP / 'tests' / 'methods_deck' / 'nolabel'
REL_TOP = 2 ** SIG_LIM         # relative density scale top, 4x (the widget's top end)
MIP_TITLE_FRAC = 0.17          # volscope_mask_qc panels: the title + legend band, top 17%
QC_PAGE_DIR =EXP / 'cd19_mask_qc_widest_20261027' / 'widest' / COND / 'montages_car_mask_cell'
DECK_DIR = Path(r'K:\FF\PPT\PPT_autogeneration\CART\CD19')
RGB = {'CAR': RGBColor(0x40, 0xE0, 0xFF), 'CD19': RGBColor(0xFF, 0x4F, 0xD8),
       'Hoechst': RGBColor(0xE6, 0xC2, 0x00)}             # Hoechst darker than the tile yellow, readable on white
CD19_MAG = np.array([1.0, 79 / 255, 216 / 255])            # config #ff4fd8
GREY = np.array([0.42, 0.42, 0.42])
# the segmentation slides' outline colors (FF 2026-10-29: the legend words in them): ilastik =
# volscope_mask_qc's ylw [1 0.85 0], the T cell outline = BODY_COLOR [1 0 0]
# (run_methods_deck_onebar.m); on the white caption the darker Hoechst yellow, readable
SEG_YELLOW, SEG_RED = RGBColor(0xFF, 0xD9, 0x00), RGBColor(0xFF, 0x00, 0x00)


def header(slide, left, top, width, height, parts, size, align=PP_ALIGN.CENTER, base=None,
           wrap=False):
    """One bold line, centred by default; (text, RGBColor or None) parts, None = BASE
    (default black). WRAP: wrap at the box width (the headers never did)."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if wrap:
        tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    for text, col in parts:
        r = p.add_run()
        r.text = text
        r.font.size, r.font.bold = Pt(size), True
        r.font.color.rgb = col or base or bpd.BLACK
    return tb


def chan(name):
    return [(name, RGB.get(name))]


# ---------------------------------------------------------------- 3D ----------------
def mesh_glbs(out_dir):
    """Three colored GLBs of the cell's mesh: CD19 relative density, the synapse patch, and
    the smoothed CD19 the synapse cut runs on. glTF is Y-up, so the cell's z (up from the
    coverslip) becomes Y; centred at the origin."""
    om, curv = load_mesh_mat(V / 'meshes' / f'{LABEL}_mesh.mat')
    f = np.asarray(om.faces, np.int64) - 1
    xyz = om.vertices_um[:, [2, 1, 0]].copy()                 # package (z,y,x) -> (x,y,z)
    xyz -= xyz.mean(axis=0)
    gl = np.column_stack([xyz[:, 0], xyz[:, 2], -xyz[:, 1]])  # (x, z, -y): z up
    sf = loadmat(str(V / 'cd19_synapse' / 'faces' / f'{LABEL}_synapse.mat'),
                 variable_names=['synapse_faces', 'field_raw', 'field_smooth', 'thresh'])
    syn = np.asarray(sf['synapse_faces']).reshape(-1) > 0
    smooth = np.asarray(sf['field_smooth'], float).reshape(-1)
    raw = np.asarray(sf['field_raw'], float).reshape(-1)
    thresh = float(np.asarray(sf['thresh']).reshape(-1)[0])
    # CD19 relative density (FF 2026-10-28): the 0-0.5 um band INCLUDING the skin (the
    # synapse stage's input, as on the detection slide) over the cell's
    # area-weighted surface mean (_face_enrichment); gray at <= 1x -> magenta at the
    # widget's top end, 2^SIG_LIM = 4x, on a LINEAR scale (FF 2026-10-28: "why is colorbar
    # nonlinear???? that is a huge huge mistake" -- the widget's log2 ramp is not used here).
    # Unmeasured faces gray.
    skin = np.asarray(loadmat(str(V / 'meshes' / f'{LABEL}_mesh.mat'),
                              variable_names=['face_sig_sig_0_05um_skin'])
                      ['face_sig_sig_0_05um_skin'], float).reshape(-1)
    ratio = _face_enrichment(skin, curv.face_areas_um2)
    t_rel = np.nan_to_num(np.clip((ratio - 1) / (REL_TOP - 1), 0, 1))[:, None]
    # smoothed CD19: the widget's synapse-field scale, 0 .. vmax (its p99, >= the cut)
    idx, vmax = _syn_field_bins(smooth, thresh)
    t_sm = np.where(idx == UNMEASURED, 0, idx / (RAMP - 1)).astype(float)[:, None]
    # raw field (the stage's field_raw) on the SAME 0..vmax, so a color is one value on both
    # models; raw faces above vmax saturate (the colorbar's top reads >= vmax)
    t_raw = np.nan_to_num(np.clip(raw / vmax, 0, 1))[:, None]
    if not (syn.size == smooth.size == raw.size == t_rel.shape[0] == f.shape[0]):
        raise SystemExit(f'{LABEL}: face counts disagree ({f.shape[0]} faces)')
    out = {}
    for name, col in (('cd19', ramp(t_rel)), ('synapse', np.where(syn[:, None], CD19_MAG, GREY)),
                      ('smooth', ramp(t_sm)), ('raw', ramp(t_raw))):
        out[name] = out_dir / f'{COND}_{LABEL}_mesh_{name}.glb'
        # lit material: PowerPoint shades it, so the shape reads in 3D. glTF COLOR_0 is
        # linear, so the sRGB colors are linearized or they render washed out.
        build_colored_glb(gl, f, col ** 2.2, out[name], unlit=False)
    return out, dict(vmax=float(vmax), thresh=thresh, nsyn=int(syn.sum()), nf=int(f.shape[0]))


def ramp(t):
    """THE color map of every colored mesh and its colorbar: gray (t = 0) -> magenta (t = 1)."""
    return (1 - t) * GREY + t * CD19_MAG


def colorbar(slide, left, top, height, ticks, title, thresh_frac=None):
    """A vertical gray -> magenta bar EXPORTED as an image from ramp(), the same function that
    colors the mesh (FF 2026-10-28: "you should export it not draw in ppt"), with TICKS =
    [(frac, text)] from the bottom, TITLE above and an optional red threshold tick at
    THRESH_FRAC as slide text (labels stay editable overlays)."""
    w = 0.28
    t = np.linspace(1, 0, 512)[:, None]                  # top = 1
    px = np.repeat((ramp(t) * 255).round().astype(np.uint8)[:, None, :], 24, axis=1)
    px[[0, -1], :] = 0
    px[:, [0, -1]] = 0                                    # 1 px black frame
    png = TILES / f'colorbar_{COND}.png'
    Image.fromarray(px).save(png)
    slide.shapes.add_picture(str(png), Inches(left), Inches(top), Inches(w), Inches(height))
    for frac, text in ticks:
        y = top + height * (1 - frac)
        bpd.add_text(slide, left + w + 0.05, y - 0.25, 1.1, 0.5, text, 24, bold=True)
    if thresh_frac is not None:
        y = top + height * (1 - thresh_frac)
        red = RGBColor(0xE0, 0x10, 0x10)
        # right of the bar, clear of the model on the left (no overlay text on data)
        ln = slide.shapes.add_connector(1, Inches(left - 0.05), Inches(y),
                                        Inches(left + w + 0.15), Inches(y))
        ln.line.color.rgb = red
        ln.line.width = Pt(4)
        bpd.add_text(slide, left + w + 0.2, y - 0.27, 1.7, 0.54, 'threshold', 24, bold=True,
                     color=red)
    tb = bpd.add_text(slide, left - 1.0, top - 1.15, w + 2.0, 1.1, title, 24, bold=True,
                      align=PP_ALIGN.CENTER)
    return tb


_PS_ADD_MODELS = r"""
$ErrorActionPreference = 'Stop'
$data = Get-Content -LiteralPath '{manifest}' -Raw | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open($data.pptx, 0, 0, 0)
try {{
    $slide = $pres.Slides.Item([int]$data.slide)
    foreach ($md in $data.models) {{
        $s = $slide.Shapes.Add3DModel($md.path, 0, -1, [single]$md.left, [single]$md.top,
                                      [single]$md.width, [single]$md.height)
        $s.Model3D.RotationX = [single]$md.rx
        $s.Model3D.RotationY = [single]$md.ry
        $s.Model3D.RotationZ = [single]$md.rz
        $s.AlternativeText = ''     # the GLB's file name carries the cell number
    }}
    $pres.Save()
}} finally {{
    $pres.Close()
    if ($app.Presentations.Count -eq 0) {{ $app.Quit() }}   # never close the user's decks
}}
"""


def add_models(pptx, slide_no, models):
    """Add native 3D models to an existing deck (models3d.py's Add3DModel pattern)."""
    man = {'pptx': str(pptx), 'slide': slide_no, 'models': models}
    with tempfile.NamedTemporaryFile('w', suffix='.json', delete=False, encoding='utf-8') as h:
        json.dump(man, h)
    res = subprocess.run(['powershell', '-NoProfile', '-Command',
                          _PS_ADD_MODELS.format(manifest=h.name.replace("'", "''"))],
                         capture_output=True, text=True)
    Path(h.name).unlink(missing_ok=True)
    if res.returncode:
        raise SystemExit(f'PowerPoint COM failed:\n{res.stdout}\n{res.stderr}')


def arrow(slide, x0, y0, x1, y1, rgb=RGBColor(0xFF, 0xDD, 0x00)):
    """A straight arrow, head at (x1, y1), inches: a connector with a triangle tail end."""
    from pptx.oxml.ns import qn
    ln = slide.shapes.add_connector(1, Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    ln.line.color.rgb = rgb
    ln.line.width = Pt(3)
    end = ln.line._get_or_add_ln().makeelement(qn('a:tailEnd'), {'type': 'triangle',
                                                                  'w': 'med', 'len': 'med'})
    ln.line._get_or_add_ln().append(end)
    return ln

# ---------------------------------------------------------------- deck --------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--export', type=Path, default=None)
    # the shared deck's opening view (every build since 2026-10-28 passed --rx 20 --ry 40
    # --rz 90; a rebuild with the old defaults 35/-40/0 turned the models, 2026-10-29)
    ap.add_argument('--rx', type=float, default=20, help='opening tilt of the 3D models, deg')
    ap.add_argument('--ry', type=float, default=40)
    ap.add_argument('--rz', type=float, default=90)
    args = ap.parse_args()
    today = datetime.date.today().strftime('%Y%m%d')
    bpd.OUT = DECK_DIR / f'CART_coculture_fixed_methods_{today}.pptx'
    if (bpd.OUT.parent / f'~${bpd.OUT.name}').exists():
        raise SystemExit(f'{bpd.OUT.name} is open in PowerPoint -- close it first')
    row = next(r for r in csv.DictReader(open(EXP / 'comparison' / 'data' / COND / 'cells.csv',
                                              newline='')) if r['cell'] == LABEL)
    # no cell number on any slide (FF 2026-10-28: "do not mention cell number anywhere")
    who = COND.replace('_', ' ').replace('min', ' min')    # no date (FF 2026-10-28)
    deck = bpd.Deck()

    s = deck.prs.slides.add_slide(deck.prs.slide_layouts[6])
    deck.titles.append('Title')
    bpd.add_text(s, 0.5, 3.0, bpd.SLIDE_W - 1, 1.0,
                 'CAR T–target conjugates: segmentation and 3D analysis', 36, bold=True,
                 align=PP_ALIGN.CENTER)

    # 2 raw data
    s = deck.slide(f'Raw data: {who}')
    cols = [('CAR', 'CAR'), ('CD19', 'CD19'), ('Hoechst', 'Hoechst'), ('merge', None)]
    left0, top, bw = 2.15, 1.45, 2.7
    for j, (key, name) in enumerate(cols):
        # FF 2026-10-29: no ' (target)' after CD19 (added after the audit, then removed)
        parts = (chan(name) if name else
                 [('CAR', RGB['CAR']), (' + ', None), ('CD19', RGB['CD19']), (' + ', None),
                  ('Hoechst', RGB['Hoechst'])])
        hx, hw = (left0 + j * bw, bw - 0.06) if name else (left0 + j * bw - 0.4, bw + 0.74)
        header(s, hx, 0.9, hw, 0.45, parts, 20)
        for i, view in enumerate(('contact_synapse', 'contact_xz')):
            bpd.fit_in_box(s, TILES / f'{COND}_{LABEL}_{view}_{key}.png',
                           left0 + j * bw, top + i * (bw + 0.05), bw - 0.06, bw - 0.06)
    for i, lab in enumerate(('Synapse plane (XY)', 'Synapse side view (XZ)')):
        bpd.add_text(s, 0.05, top + i * (bw + 0.05) + bw / 2 - 0.3, left0 - 0.12, 0.6, lab, 20,
                     bold=True, align=PP_ALIGN.RIGHT)
    # the windows as contact_check_sheet.py :: montage_tiles sets them (Hoechst over the
    # whole image, not the cell -- the review of 2026-10-28)
    bpd.add_text(s, 0.4, 7.0, bpd.SLIDE_W - 0.8, 0.4,
                 'CD19 110–300. CAR 110 to its 95th percentile inside the T cell, per view. '
                 'Hoechst 110 to its 99.5th percentile in the image. 5 μm bar.', 13,
                 bold=True, color=bpd.GREY)

    # 3 segmentation, then the same slide with the projection at lower CAR ceilings (FF
    # 2026-10-28: "you can hardly see part of the cell"; "make a second slide with different
    # thresholds there"; "why in the world are images in 4 not in the same location?"): one
    # layout for every version, so the images do not move between slides. volscope_mask_qc
    # MIP_PCT renders; the single-plane row is the 95th in each. Left out until rendered.
    num = LABEL.split('cell')[-1]
    for p in (95,) + MIP_PCTS:
        d = NOLAB / f'seg_p{p}' / COND / 'parts'
        parts = [d / f'{COND}__cell{num}__{part}.png' for part in ('MIP', 'broadest')]
        if not all(f.is_file() for f in parts):
            print(f'segmentation slide, projection at the {p}th, left out: no renders in {d}')
            continue
        # FF 2026-10-28: "you have to say that segmentation is being done on CAR"; no jargon
        # FF 2026-10-29: the lower-ceiling version says so in its title
        s = deck.slide('Segmenting the T cell from its CAR signal'
                       + ('' if p == 95 else ' (different CAR contrast)'))
        # FF 2026-10-29: "specify Plane of maximal cross-sectional area" (each view's own
        # slice of largest outline area, WIDEST_EACH); two lines, the label column is 3.9 in
        rows = [(['All planes, projected'], p),
                (['Plane of maximal', 'cross-sectional area'], 95)]
        lab_w, gap, ph, head = 3.9, 0.12, 2.55, 0.42
        for i, (f, (what, pct)) in enumerate(zip(parts, rows)):
            # volscope_mask_qc LABELS false: no text in the image (no cell number); the three
            # views sit in equal tiles, so their names go at the thirds, as slide text
            y = 0.9 + head + i * (ph + head + 0.05)
            iw = ph * Image.open(f).size[0] / Image.open(f).size[1]
            x_img = (bpd.SLIDE_W - iw + lab_w + gap) / 2
            bpd.fit_in_box(s, f, x_img, y, iw, ph)
            if i == 0:
                # the outline legend in the top image's empty black band, over XY (FF
                # 2026-10-29): 16 pt is 2.03 in wide, the XY third 2.30 in; the band above
                # every pixel of the XY third is >= 0.35 in on the 95th and 70th panels, the
                # ink 0.08-0.28 in below the image top (zero insets)
                lg = header(s, x_img + 0.1, y + 0.03, 2.15, 0.3,
                            [('ilastik', SEG_YELLOW), ('   ', None), ('postprocessed', SEG_RED)],
                            16, align=PP_ALIGN.LEFT)
                tf = lg.text_frame
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
            for k, v in enumerate(('XY', 'XZ', 'YZ')):
                bpd.add_text(s, x_img + iw * k / 3, y - head, iw / 3, head, v, 20, bold=True,
                             align=PP_ALIGN.CENTER)
            bpd.add_text(s, x_img - gap - lab_w, y + 0.5, lab_w, 1.5,
                         [*what, f'CAR 110 to the {pct}th percentile',
                          'of CAR inside the outline'], 20, bold=True, align=PP_ALIGN.RIGHT)
        # FF 2026-10-29: the caption stays, its color words in their colors
        header(s, 0.4, 6.95, bpd.SLIDE_W - 0.8, 0.5,
               [('Gray: CAR. ', None), ('Yellow', RGB['Hoechst']),
                (': ilastik\u2019s classification of the CAR image. ', None), ('Red', SEG_RED),
                (': the final T cell outline, used for all measurements. 5 \u03bcm bar.', None)],
               16, align=PP_ALIGN.LEFT, base=bpd.GREY, wrap=True)

    # 4 example cells, CAT and FMC 60 min (FF 2026-10-28: "you may as well show CAT 60 min
    # too"); each condition's FIRST mask-QC page = its first 25 cells, NOT every cell
    s = deck.slide('T cell outlines, the first 25 cells of each condition, 60 min')
    for j, cond in enumerate(('CAT_60min', 'FMC_60min')):
        qdir = NOLAB / 'outlines' / cond / 'montages_car_mask_cell'   # no tile labels
        # lowest first cell; on a tie the NEWEST (an exclusion rerun renames 1_36 -> 1_37 and
        # leaves the old page beside it)
        page = min(qdir.glob('montage_cells_*.png'),
                   key=lambda p: (int(p.stem.split('_')[2]), -p.stat().st_mtime))
        x0 = 0.4 + j * 6.35
        bpd.add_text(s, x0, 0.85, 6.15, 0.55, cond.replace('_', ' ').replace('min', ' min'),
                     26, bold=True, align=PP_ALIGN.CENTER)
        bpd.fit_in_box(s, page, x0, 1.45, 6.15, 5.4)
    bpd.add_text(s, 0.4, 6.95, bpd.SLIDE_W - 0.8, 0.45,
                 'CAR in each T cell\u2019s widest plane, 110 to its 95th percentile inside the '
                 'outline. Red: T cell outline. 5 \u03bcm bars.', 16, bold=True, color=bpd.GREY)

    glbs, g = mesh_glbs(TILES)
    area = float(row['synapse_area_um2'])
    # ONE geometry for both mesh slides, so nothing shifts between them (review 2026-10-28)
    top_rel, h_rel = 2.8, 3.25                          # bar title clear of the 28 pt headers
    xl, wl, xb, xr, wr = 0.3, 5.3, 5.8, 7.1, 5.9        # left model, bar, right model

    # 5 synapse detection, raw vs smoothed (models added by COM below; FF 2026-10-28: the
    # old MATLAB QC figure's text was far too small; "compare smoothed and raw")
    s = deck.slide('Finding the synapse on the 3D T cell surface')
    n_syn = len(deck.prs.slides)
    header(s, xl, 0.95, wl, 0.45, [('CD19', RGB['CD19']), (' near the cell surface', None)], 28)
    header(s, xr, 0.95, wr, 0.45, [('Smoothed ', None), ('CD19', RGB['CD19'])], 28)
    colorbar(s, xb, top_rel, h_rel,
             [(0, '0'), (0.5, f'{g["vmax"] / 2:.0f}'), (1, f'≥{g["vmax"]:.0f}')],
             ['CD19 (a.u.)'], thresh_frac=min(1.0, g['thresh'] / g['vmax']))
    bpd.add_text(s, 0.4, 6.35, bpd.SLIDE_W - 0.8, 0.9,
                 ['Left: CD19 in the T cell’s outer 0.5 μm. Right: the same, smoothed '
                  'along the surface.',
                  f'Threshold (red tick): this cell\u2019s median smoothed CD19 + 2 \u00d7 its '
                  f'median absolute deviation = {g["thresh"]:.0f}. Synapse: the largest '
                  f'connected region above it, {area:.0f} \u03bcm\u00b2.'],
                 18, bold=True, color=bpd.GREY)     # 18: at 20 the third line hit the edge

    # 6 3D: relative density + the detected synapse (after the detection slide,
    # FF 2026-10-28: "you should show slide 7 before 6 right?")
    s = deck.slide('CD19 over the T cell surface, and the detected synapse')
    n_3d = len(deck.prs.slides)
    header(s, xl, 0.95, wl, 0.45, [('CD19', RGB['CD19']), (' near the cell surface', None)], 28)
    header(s, xr, 0.95, wr, 0.45, [('Detected synapse', RGB['CD19'])], 28)
    colorbar(s, xb, top_rel, h_rel, [(0, '\u22641\u00d7'), (1 / 3, '2\u00d7'),
                                     (2 / 3, '3\u00d7'), (1, '\u22654\u00d7')],
             ['CD19 relative', 'density'])
    bpd.add_text(s, 0.4, 6.35, bpd.SLIDE_W - 0.8, 0.9,
                 ['Left: CD19 in the T cell’s outer 0.5 μm, relative to the cell’s '
                  'average (1×).',
                  f'Right: magenta = the detected synapse, {area:.0f} μm².'],
                 20, bold=True, color=bpd.GREY)

    # 7 internalized CD19 (FF 2026-10-28: smaller, a scale bar, the plane named below it; no
    # height above the coverslip, no object outlines, "way too much text"). Arrowheads mark
    # the spots in this plane (the exported objects' centroids), off the spots themselves.
    # FF 2026-10-29: CAR | CD19 | merge of that plane, no T cell outline
    s = deck.slide('Internalized CD19')
    oz = next(csv.DictReader(open(TILES / f'{COND}_{LABEL}_objects.csv', newline='')))
    cents = [tuple(map(float, c.split(';'))) for c in oz['centroids_frac'].split('|') if c]
    w_t, gap_t, y_t = 3.7, 0.25, 1.45                   # three square tiles, centered
    x0 = (bpd.SLIDE_W - 3 * w_t - 2 * gap_t) / 2
    for j, (ch, parts) in enumerate((('CAR', chan('CAR')), ('CD19', chan('CD19')),
                                     ('merge', [('CAR', RGB['CAR']), (' + ', None),
                                                ('CD19', RGB['CD19'])]))):
        x_t = x0 + j * (w_t + gap_t)
        header(s, x_t, 0.95, w_t, 0.45, parts, 24)
        bpd.fit_in_box(s, TILES / f'{COND}_{LABEL}_objects_{ch}.png', x_t, y_t, w_t, w_t)
        for k, (r, c, rad) in enumerate(cents):
            # alternate sides; the tip stops just short of the spot's edge (FF 2026-10-28:
            # "make sure arrowheads are closer to puncta"), so each arrow names one spot
            ux, uy = (0.77, -0.64) if k % 2 == 0 else (-0.77, 0.64)
            tx, ty = x_t + c * w_t, y_t + r * w_t
            gap = rad * w_t + 0.03
            arrow(s, tx + ux * (gap + 0.45), ty + uy * (gap + 0.45), tx + ux * gap, ty + uy * gap)
    bpd.add_text(s, 0.4, y_t + w_t + 0.15, bpd.SLIDE_W - 0.8, 1.0,
                 ['Plane with the most internalized CD19',
                  # FF 2026-10-29: "remove 'not connected to its surface'"
                  f'{oz["n_objects"]} CD19 spots inside this T cell'],
                 24, bold=True, align=PP_ALIGN.CENTER)
    bpd.add_text(s, 0.4, 6.95, bpd.SLIDE_W - 0.8, 0.45,
                 f'Arrowheads: {len(cents)} spots in this plane. CD19 110–300; CAR 110 to its '
                 '95th percentile inside the T cell. 5 μm bars.', 16, bold=True, color=bpd.GREY)
    # No alt text: python-pptx writes each picture's file name into it, and those names carry
    # the cell number, which must not travel with the shared deck (audit 2026-10-29).
    for sl in deck.prs.slides:
        for el in sl.shapes._spTree.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr'):
            el.set('descr', '')
    deck.save()

    box = dict(top=1.5 * 72, height=4.8 * 72, rx=args.rx, ry=args.ry, rz=args.rz)
    assert (xr + wr) * 72 <= bpd.SLIDE_W * 72
    add_models(bpd.OUT, n_3d, [dict(path=str(glbs['cd19']), left=xl * 72, width=wl * 72, **box),
                               dict(path=str(glbs['synapse']), left=xr * 72, width=wr * 72,
                                    **box)])
    add_models(bpd.OUT, n_syn, [dict(path=str(glbs['raw']), left=xl * 72, width=wl * 72, **box),
                                dict(path=str(glbs['smooth']), left=xr * 72, width=wr * 72,
                                     **box)])
    print(f'saved {bpd.OUT} ({len(deck.prs.slides)} slides); smoothed CD19 0-{g["vmax"]:.0f}, '
          f'cut {g["thresh"]:.1f}; synapse {g["nsyn"]} of {g["nf"]} faces')
    if args.export:
        bpd.export_pngs(args.export, deck.titles)


if __name__ == '__main__':
    main()
