#!/usr/bin/env python3
r"""update_irm_deck_20260917.py  Second pass on FF's IRM deck (FF 2026-09-17).

Source: CART_actin_IRM_3_5_7_10_20260916_newcompile.pptx (output of update_irm_deck.py; stays untouched).
  * every compile grid picture (timecourse scatter grids, radial profile grid, axial overlays) is swapped
    IN PLACE for the Sep 17 recompile (compiled_20260717_20260723_20260917: one 20 pt CAT/FMC legend
    outside the panels, foci-density labels fixed). Same data + stars as Sep 15. Montages and the
    profile / zone slides (49-53, from actin_extended_panels) are left alone. NO text edits.
  * two new slides after slide 56 (Actin foci): clearance area + fraction, and foci density.
  * clearance QC montages (8 slides, CAT left / FMC right, page 1 of actin\synapse\clearance), styled like
    FF's Actin Foci Overlay slides (black background, same text boxes).
  * all montage slides together (FF 2026-09-17): FF's Actin Foci Overlay block, then the clearance QC
    block, right after the original montages (slide 43); results slides follow.
  * an extra slice montage block ("Actin: 2 Below, at Synapse, 2 Above", synapse_slices_xy) right after
    FF's 0 / +2 / +4 block.
  * FF's old radial profile slide ("Actin distribution at synapse: radial profiles") is removed.
Verifies like update_irm_deck.py: non-picture XML + notes of every original slide unchanged, picture
set unchanged, swapped blobs == source files, no orphan image rels, source untouched.

Run: python update_irm_deck_20260917.py
"""
from __future__ import annotations

import copy
import sys
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_actin_decks as bad  # noqa: E402
import update_irm_deck as uid  # noqa: E402

SRC = Path(r'K:\FF\PPT\PPT_autogeneration\CART\actin_only\CART_actin_IRM_3_5_7_10_20260916_newcompile.pptx')
OUT = SRC.with_name('CART_actin_IRM_3_5_7_10_20260917.pptx')
OLD_COMPILE = bad.RC / 'compiled_20260717_20260723_20260915'
uid.COMPILE = bad.RC / 'compiled_20260717_20260723_20260917'   # plot_source() reads this global
INSERT_AFTER = 56                 # 1-based: "Actin foci" slide
# New QC montage blocks after FF's Actin Foci Overlay block (title stem, prog_fixed_cells subfolder).
# FF 2026-09-17: the overlay's 1-px red foci outlines vanish at slide size -> the binary foci mask shows detection.
# FF 2026-09-17 (later): foci mask block removed again ("not useful").
QC_BLOCKS = [('Actin Clearance', r'actin\synapse\clearance')]
# FF 2026-09-17: extra slice montage block, placed right after FF's 0 / +2 / +4 block (same family,
# sibling folder synapse_slices_xy: sxy_offsets = [-2 0 2] in actin_qc_per_cell).
SLICE_BLOCK = ('Actin: 2 Below, at Synapse, 2 Above', r'actin\synapse\synapse_slices_xy')
SLICE_AFTER_STEM = 'Actin: 0, 2, and 4 slices above synapse'
# FF 2026-09-17: original slides to drop (the old all-cells radial profile grid; superseded by the profile slides).
DROP_TITLES = {'Actin distribution at synapse: radial profiles'}
TS = uid.COMPILE / 'grid_panels' / 'timecourse_scatter_plots'
# FF 2026-09-17: raw zone MFI violins re-rendered (zone_raw_xtick_fontsize 17; "CentralPeripheral" collided).
RAW_ZONES_OLD = sorted((bad.RC / 'actin_extended_panels_20260916' / 'violins').glob('zones_raw_actin_*.png'))
RAW_ZONES_NEW = bad.RC / 'actin_extended_panels_20260917' / 'violins'
NEW_SLIDES = [
    ('Actin Clearance Area', [TS / 'actin_bottom_slice_clearance_area_um2_grid.png',
                              TS / 'actin_bottom_slice_clearance_frac_grid.png'], 2),
    # FF 2026-09-17: density only -- the raw count grid is already on FF's "Actin foci" slide (per-clearance
    # density: outlier-driven; median focus area: at the detection limit).
    ('Actin Foci Density', [TS / 'actin_foci_density_um2_grid.png'], 1),
]


def plan_swaps(prs):
    """Only pictures that came from the Sep 15 compile (blob == that file) are swapped, plus the raw zone
    MFI violins from the Sep 16 extended panels (re-rendered Sep 17 with smaller zone names)."""
    plan, errors = [], []
    for i, s in enumerate(prs.slides):
        if uid.TITLE_RE.match(uid.title_of(s)):
            continue                                              # montage slide
        for pic in (sh for sh in s.shapes if sh.shape_type == 13):
            blob_sha = uid.sha(pic.image.blob)
            raw = next((f for f in RAW_ZONES_OLD if uid.sha(f.read_bytes()) == blob_sha), None)
            if raw is not None:
                new = RAW_ZONES_NEW / raw.name
                if new.is_file():
                    plan.append((i, pic, new))
                else:
                    errors.append(f'slide {i + 1}: missing {new}')
                continue
            name = uid.descr_of(pic).get('descr') or ''
            new = uid.plot_source(name)
            old = Path(str(new).replace(str(uid.COMPILE), str(OLD_COMPILE)))
            if not old.is_file() or uid.sha(old.read_bytes()) != uid.sha(pic.image.blob):
                continue                                          # extended-panel slides etc.
            if not new.is_file():
                errors.append(f'slide {i + 1}: missing {new}')
                continue
            plan.append((i, pic, new))
    return plan, errors


def add_slides(prs):
    """New slides + final order. FF 2026-09-17: all montage (QC) slides sit together after the original
    montage block (slide 43): FF's Actin Foci Overlay slides move there, followed by the new QC_BLOCKS
    (foci mask, clearance); the results slides follow, with the two new metric slides after "Actin foci".
    Returns ({old idx0: new idx0}, n_new, slide-number ranges for the report)."""
    layout = next(sl.slide_layout for sl in prs.slides if sl.slide_layout.name == 'Blank')
    old_ids = list(prs.slides._sldIdLst)
    titles = [uid.title_of(s) for s in prs.slides]
    foci = [i for i, t in enumerate(titles) if t.startswith('Actin Foci Overlay ')]
    if len(foci) != 8 or foci != list(range(foci[0], foci[0] + 8)) or foci[0] <= INSERT_AFTER - 1:
        raise RuntimeError(f'unexpected Actin Foci Overlay slides: {foci}')
    if len(old_ids) != foci[-1] + 1:
        raise RuntimeError('slides after the foci overlay block')
    deck = uid.Shim(prs, layout)
    for title, pngs, ncol in NEW_SLIDES:
        bad.tight_grid(deck, title, pngs, [''] * ncol)
    for s in deck.new:                                            # text above pictures
        tree = s.shapes._spTree
        for sh in list(s.shapes):
            if sh.has_text_frame:
                tree.remove(sh._element); tree.append(sh._element)
    n_metric = len(deck.new)
    qc = []
    for stem, sub in QC_BLOCKS:
        qc += add_qc_montages(prs, layout, [prs.slides[i] for i in foci], stem, sub)
    n_qc = len(qc)
    slices = add_qc_montages(prs, layout, [prs.slides[i] for i in foci], *SLICE_BLOCK)
    sl_after = max(i for i, t in enumerate(titles) if t.startswith(SLICE_AFTER_STEM)) + 1
    ids = list(prs.slides._sldIdLst)
    metric_ids = ids[len(old_ids):len(old_ids) + n_metric]
    qc_ids = ids[len(old_ids) + n_metric:len(old_ids) + n_metric + n_qc]
    slice_ids = ids[len(old_ids) + n_metric + n_qc:]
    MONT_END = 43                                                 # original montage block = slides 4-43
    order = (old_ids[:sl_after] + slice_ids + old_ids[sl_after:MONT_END] + [old_ids[i] for i in foci]
             + qc_ids + old_ids[MONT_END:foci[0]] + metric_ids)
    lst = prs.slides._sldIdLst
    for el in ids:
        lst.remove(el)
    for el in order:
        lst.append(el)
    drop = [el for i, el in enumerate(old_ids) if titles[i] in DROP_TITLES]
    if len(drop) != len(DROP_TITLES):
        raise RuntimeError(f'expected {len(DROP_TITLES)} slides to drop, found {len(drop)}')
    for el in drop:                                               # remove from the deck, drop the part rel
        order.remove(el); lst.remove(el); prs.part.drop_rel(el.rId)
    pos = {id(el): k for k, el in enumerate(order)}
    index_map = {i: pos[id(el)] for i, el in enumerate(old_ids) if id(el) in pos}
    ranges = {SLICE_BLOCK[0]: (sl_after + 1, sl_after + len(slices)),
              'foci overlay': (MONT_END + 1 + len(slices), MONT_END + 8 + len(slices))}
    for k, (stem, _) in enumerate(QC_BLOCKS):
        ranges[stem] = (MONT_END + 9 + len(slices) + 8 * k, MONT_END + 16 + len(slices) + 8 * k)
    ranges['metric'] = (len(order) - n_metric + 1, len(order))
    return index_map, n_metric + n_qc + len(slices) - len(drop), ranges


def add_qc_montages(prs, layout, foci_slides, stem, sub):
    """One slide per dataset x timepoint, styled as FF's Actin Foci Overlay slide for the same
    dataset/timepoint: its black background and title / CAT / FMC text boxes are copied verbatim
    (title text renamed); CAT and FMC montages aspect-fit under the labels."""
    new = []
    box_top, box_h, box_w = 0.72, bad.bpd.SLIDE_H - 0.77, 6.6
    # ONE px/in for the whole block (2026-10-17: aspect-fitting shrank CAT and FMC by
    # different factors), pinned to the largest montage, as in bad.montage_group_ppi.
    pngs = [bad.first_page(bad.NF / folder / f'{ct}_{t}' / 'split_channels' / 'prog_fixed_cells' / sub)
            for _, _, _, folder in bad.DATASETS for t in bad.TPS for ct in ('CAT', 'FMC')]
    ppi = bad.mont.compute_group_ppi([p for p in pngs if p is not None], box_w, box_h)
    for _, _, date, folder in bad.DATASETS:
        for t, tl in zip(bad.TPS, bad.TP_LBL):
            title = f'{stem} {bad.DASH} {date}, {tl}'
            ref = next(s for s in foci_slides if uid.title_of(s) == f'Actin Foci Overlay {bad.DASH} {date}, {tl}')
            s = prs.slides.add_slide(layout)
            for ph in list(s.placeholders):
                ph._element.getparent().remove(ph._element)
            c_sld = s._element.find(qn('p:cSld'))
            c_sld.insert(0, copy.deepcopy(ref._element.find(qn('p:cSld')).find(qn('p:bg'))))
            texts = [sh for sh in ref.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
            if [sh.text_frame.text for sh in texts[1:]] != ['CAT', 'FMC']:
                raise RuntimeError(f'unexpected text boxes on {uid.title_of(ref)!r}')
            for sh in texts:
                s.shapes._spTree.append(copy.deepcopy(sh._element))
            t_run = s.shapes[0].text_frame.paragraphs[0].runs[0]
            t_run.text = title
            for sh in list(s.shapes)[1:]:
                ct = sh.text_frame.text
                png = bad.first_page(bad.NF / folder / f'{ct}_{t}' / 'split_channels' / 'prog_fixed_cells'
                                     / sub)
                if png is None:
                    raise FileNotFoundError(f'no {sub} montage: {folder} {ct}_{t}')
                bad.place_montage(s, png, sh.left / uid.EMU_IN, ppi, box_top, box_w, box_h)
            new.append(s)
    return new


def main():
    missing = [p for _, pngs, _ in NEW_SLIDES for p in pngs if not p.is_file()]
    if missing:
        print('[abort] missing:\n  ' + '\n  '.join(map(str, missing))); return 1
    if OUT.exists() or (OUT.parent / ('~$' + OUT.name)).exists():
        print(f'[abort] {OUT} (or its lock) already exists'); return 1
    src_sha = uid.sha(SRC.read_bytes())
    prs = Presentation(str(SRC))
    plan, errors = plan_swaps(prs)
    if errors:
        print('[abort] unresolved:\n  ' + '\n  '.join(errors)); return 1
    print(f'{len(plan)} compile pictures to swap')
    expected = []
    for i, pic, new in plan:
        old_sha = uid.sha(pic.image.blob)
        print(f'  slide {i + 1:2d} {new.name}')
        uid.swap(prs.slides[i], pic, new)
        expected.append((i, pic.shape_id, str(new), old_sha))
    index_map, n_new, ranges = add_slides(prs)
    prs.save(str(OUT))
    print(f'saved {OUT} ({len(prs.slides)} slides); ' + '; '.join(f'{k} {a}-{b}' for k, (a, b) in ranges.items()))
    uid.SRC, uid.OUT, uid.INSERT_AFTER = SRC, OUT, INSERT_AFTER   # reuse the original verifier
    exp = [(index_map[i], sid, p, old) for i, sid, p, old in expected if i in index_map]
    fails = uid.verify(exp, n_new, src_sha, index_map=index_map)
    print('VERIFY: ' + ('PASS' if not fails else 'FAIL\n  ' + '\n  '.join(fails)))
    return 0 if not fails else 2


if __name__ == '__main__':
    sys.exit(main())
