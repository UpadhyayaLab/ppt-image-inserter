"""Cathepsin B | Granzyme B deck: the same polarization metrics for the three CTSB/MT experiments
(Nov 27 2023, Jun 20 2024, Jun 24 2024) and Kiet's Oct 8, 2026 GzmB set. Every result slide is ONE
axis with the experiments in date order, named above it; never pooled (FF + Kiet 2026-10-29).

Sources (no figure is drawn here): <CTSB root>\\gzmb_vs_ctsb_violins_<newest>\\ from
tests/run_gzmb_vs_ctsb_violins.m
  <metric>.png              violins (the GzmB deck's style)
  <metric>_le<cutoff>.png   fraction of cells <= cutoff
  <metric>_le3_stats.csv    n per condition; the deck's cell counts come from these
Cells as in d20261008_gzmb_compilation_config and
d20231127_20240620_20240624_ctsb_vs_gzmb_compilation_config (July 2026 CTSB results compiled with
the GzmB cell QC, cutoffs and tests).

  python build_gzmb_vs_ctsb_deck.py --export <dir>
-> K:\\FF\\PPT\\PPT_autogeneration\\CART\\GzmB\\CART_GzmB_vs_CTSB_<yyyymmdd>.pptx
"""
from __future__ import annotations

import argparse
import csv
import datetime
import sys
from pathlib import Path

from PIL import Image
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import build_gzmb_decks as bgd          # also puts CART_pMLC_ppt / CART_actin_ppt on sys.path
import build_pmlc_profiles_deck as bpd

MU, DASH, SUB50, SUB75 = bgd.MU, bgd.DASH, bgd.SUB50, bgd.SUB75
CTSB_ROOT = Path(r'J:\FF\fixed_cell\CAR_TCell\results_compiled\MT_CATB_20231127_20240620_20240624')
OUT = bgd.OUT_DIR / f'CART_GzmB_vs_CTSB_{datetime.date.today():%Y%m%d}.pptx'

CENT = 'Centrosome = brightest point of the smoothed MT stack within the actin MIP outline.'
VIOLIN_NOTE = ('White dot: median; bar: interquartile range. Stars: Wilcoxon rank-sum test, '
               'CAT vs FMC, Benjamini-Hochberg corrected within each experiment.')
FRACTION_NOTE = ('Error bars: 95% Clopper-Pearson binomial confidence intervals; numbers: '
                 "polarized / all cells. Stars: Fisher's exact test, Benjamini-Hochberg corrected "
                 'within each experiment.')

# Granule metrics in two versions of the GzmB+ gate (FF 2026-10-29: 'include both versions ...
# indicate clearly the differences'). CTSB cells are the same in both versions.
GZMB_EXP = 'Oct 8, 2026'
GATE_TAG = {'': '10\u2076', '_gzmb1e5': '10\u2075'}
CENT_NAME = 'Centrosome\u2013Synapse Distance'
CENT_CELLS = 'All cells passing QC; no Granzyme B or Cathepsin B gating.'
GRANULE = {'z50': f'Granule {SUB50}', 'z75': f'Granule {SUB75}'}

# Main story, then backup (FF 2026-10-29: one gate in the main slides, the rest at the back;
# granule fractions 3 um first, the primary comparison; related fraction / gate panels share a
# slide). Slide = (title, panels); panel = ('cent', cutoff or None) or
# (granule stem, gate suffix, cutoff or None); None = the violin.
# Centrosome slides never mix with granule panels: their cells are all QC cells, no GzmB / CTSB
# gate (FF 2026-10-29), and both cutoffs are main results.
MAIN = [
    (CENT_NAME, [('cent', None)]),
    (f'{CENT_NAME}: Fraction Polarized', [('cent', 2), ('cent', 3)]),
    (f'{GRANULE["z50"]} ({MU}m from Cell Bottom): GzmB+ > 10⁶', [('z50', '', None)]),
    # 2 um above 3 um, as on the centrosome slide (FF 2026-10-29)
    (f'{GRANULE["z50"]} Fraction Polarized: GzmB+ > 10⁶', [('z50', '', 2), ('z50', '', 3)]),
]
BACKUP = [
    (f'{GRANULE["z50"]} ({MU}m from Cell Bottom): GzmB+ > 10⁵', [('z50', '_gzmb1e5', None)]),
    (f'{GRANULE["z50"]} Fraction Polarized: GzmB+ > 10⁵', [('z50', '_gzmb1e5', 2), ('z50', '_gzmb1e5', 3)]),
    (f'{GRANULE["z75"]} ({MU}m from Cell Bottom)', [('z75', '', None), ('z75', '_gzmb1e5', None)]),
    (f'{GRANULE["z75"]} Fraction Polarized (≤ 3 {MU}m)', [('z75', '', 3), ('z75', '_gzmb1e5', 3)]),
]

MISSING: list[Path] = []


def check(p: Path) -> Path:
    if not p.is_file():
        MISSING.append(p)
    return p


def run_counts(d: Path) -> dict:
    """Cell counts of the run in d, from its stats files: cells with a value / all QC cells.
    They were typed in, so a newer run put the old counts beside its new figures (bug hunt
    2026-10-31)."""
    def totals(metric):
        with open(d / f'{metric}_le3_stats.csv', newline='', encoding='utf-8-sig') as f:
            out: dict = {}
            for r in csv.DictReader(f):
                out[r['experiment']] = out.get(r['experiment'], 0) + int(r['n_cat']) + int(r['n_fmc'])
        return out
    qc, e6, e5 = totals('centrosome_distance'), totals('z50'), totals('z50_gzmb1e5')
    ctsb = [100 * e6[e] / qc[e] for e in qc if e != GZMB_EXP]
    return {'qc': qc[GZMB_EXP], 'e6': e6[GZMB_EXP], 'e5': e5[GZMB_EXP],
            'ctsb_lo': min(ctsb), 'ctsb_hi': max(ctsb)}


def gate_notes(c: dict) -> tuple:
    """({gate suffix: note}, the note when a slide shows both gates)."""
    notes = {
        '': f'GzmB+ = total signal > 10⁶ ({c["e6"]:,} of {c["qc"]:,} GzmB cells); CTSB+ = peak signal > 20.',
        '_gzmb1e5': (f'GzmB+ = total signal > 10⁵ ({c["e5"]:,} of {c["qc"]:,} GzmB cells; '
                     f'{c["e6"]:,} at 10⁶); CTSB+ = peak signal > 20, as at 10⁶.'),
    }
    both = (f'GzmB+ = total signal > 10⁶ ({c["e6"]:,} of {c["qc"]:,} GzmB cells) or > 10⁵ '
            f'({c["e5"]:,}); CTSB+ = peak signal > 20.')
    return notes, both


def bullet_list(s, items, size=24):
    """Bulleted (lead, text[, level]) items: bold lead, spaced paragraphs, level-1 items as
    indented dash sub-bullets (FF 2026-10-29: larger text, nicer formatting)."""
    tb = s.shapes.add_textbox(Inches(0.9), Inches(1.1), Inches(bpd.SLIDE_W - 1.8), Inches(5.9))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        lead, text = item[0], item[1]
        level = item[2] if len(item) > 2 else 0
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(16 if level == 0 else 8)
        pPr = p._p.get_or_add_pPr()                 # real bullet with a hanging indent
        pPr.set('marL', str(Inches(0.4 + 0.5 * level)))
        pPr.set('indent', str(-Inches(0.4)))
        bu = pPr.makeelement(qn('a:buChar'), {'char': '•' if level == 0 else '–'})
        pPr.append(bu)
        for t, bold in ((lead, True), (text, False)):
            r = p.add_run()
            r.text = t
            r.font.size = Pt(size - 2 * level)
            r.font.bold = bold


def panels_slide(deck, title, panels, note):
    """One or two one-axis figures of all four experiments (stacked, each under a bold caption),
    with a large CAT / FMC key (FF 2026-10-29: 'CAT and FMC labels should be larger')."""
    s = deck.slide(title)
    # Top-aligned under the key, spare room below (FF 2026-10-29: 'better to leave space below
    # than above'); fit_in_box would center the wide figure vertically.
    top, bottom = 1.2, bpd.SLIDE_H - 0.72
    cap_h = 0.38 if len(panels) > 1 else 0.0
    box_w = bpd.SLIDE_W - 0.2
    box_h = (bottom - top) / len(panels) - cap_h
    y = top
    for caption, png in panels:
        if caption:
            bpd.add_text(s, 0.1, y, box_w, cap_h, caption, 20, bold=True, align=PP_ALIGN.CENTER)
        if png.is_file():
            with Image.open(png) as im:
                ar = im.size[0] / im.size[1]
            w = min(box_w, box_h * ar)
            s.shapes.add_picture(str(png), Inches(0.1 + (box_w - w) / 2), Inches(y + cap_h), width=Inches(w))
            y += cap_h + w / ar + 0.1
        else:
            bpd.fit_in_box(s, png, 0.1, y + cap_h, box_w, box_h)   # draws the [missing] placeholder
            y += cap_h + box_h
    bpd.add_text(s, bpd.SLIDE_W / 2 - 1.6, 0.68, 1.5, 0.55, 'CAT', 28, bold=True,
                 color=bpd.CAT_BLUE, align=PP_ALIGN.RIGHT)
    bpd.add_text(s, bpd.SLIDE_W / 2 + 0.1, 0.68, 1.5, 0.55, 'FMC', 28, bold=True, color=bpd.FMC_RED)
    tb = bpd.add_text(s, 0.15, bpd.SLIDE_H - 0.55, bpd.SLIDE_W - 0.3, 0.5, note, 14,
                      align=PP_ALIGN.CENTER)
    for r in tb.text_frame.paragraphs[0].runs:
        r.font.italic = True


def build(export):
    d = bpd.newest_run(CTSB_ROOT, 'gzmb_vs_ctsb_violins_2*')
    if d is None:
        sys.exit(f'[error] no gzmb_vs_ctsb_violins_* under {CTSB_ROOT}; run tests/run_gzmb_vs_ctsb_violins.m')
    print(f'figures: {d}')
    counts = run_counts(d)
    gate_note, gate_both = gate_notes(counts)
    bpd.OUT = OUT
    deck = bgd.OvDeck()

    s = deck.prs.slides.add_slide(deck.prs.slide_layouts[6])
    deck.titles.append('Title')
    bpd.add_text(s, 0.83, 2.2, 11.67, 1.2, 'Cathepsin B and Granzyme B: Polarization to the Synapse',
                 40, bold=True, align=PP_ALIGN.CENTER)
    bpd.add_text(s, 0.83, 3.55, 11.67, 0.5, 'CAR T Cells, CAT vs FMC: Each Experiment Shown Separately',
                 24, align=PP_ALIGN.CENTER)
    today = datetime.date.today()
    bpd.add_text(s, 4.37, 4.4, 4.6, 0.5, f'{today:%B} {today.day}, {today.year}', 20,
                 align=PP_ALIGN.CENTER)

    s = deck.slide('Data and Cell Selection')
    bullet_list(s, [
        ('Cathepsin B experiments: ', 'Nov 27, 2023 (15 min); Jun 20, 2024 and Jun 24, 2024 (5, 15 min).'),
        ('Granzyme B experiment: ', 'Oct 8, 2026 (5, 10, 15 min).'),
        ('Centrosome: ', f'{CENT[len("Centrosome = "):-1]}; distance = height above the synapse slice.'),
        (f'{SUB50} / {SUB75}: ', ('height above the synapse slice below which 50% / 75% of the cell\u2019s '
                                  'granule signal lies; same computation for both proteins.')),
        # counts: positive cells with a Z50 / QC cells (run_counts)
        ('Positive cells (granule slides): ', 'the criteria differ.'),
        ('CTSB+: ', f'peak signal > 20 → {counts["ctsb_lo"]:.0f}–{counts["ctsb_hi"]:.0f}% of cells per experiment', 1),
        ('GzmB+: ', f'total signal > 10⁶ → {100 * counts["e6"] / counts["qc"]:.0f}% of cells '
                    '(many express little granzyme B)', 1),
        ('', 'Main slides use GzmB+ > 10⁶; the backup repeats them at GzmB+ > 10⁵ '
             f'({100 * counts["e5"] / counts["qc"]:.0f}% of cells)', 1),
    ])

    # Super brief (FF 2026-10-29); the occasional *** with large n is not called a difference.
    s = deck.slide('Summary')
    bullet_list(s, [
        ('', 'No meaningful CAT vs FMC differences in centrosome or granule polarization.'),
        ('Exception: ', 'Jun 24, 2024, where many FMC cells failed to polarize.'),
    ], size=30)

    def panel(entry):
        """(caption, png, note parts) of one panel; the gate follows a colon or comma, never an
        em dash (FF 2026-10-29)."""
        if entry[0] == 'cent':
            t = entry[1]
            if t is None:
                return '', check(d / 'centrosome_distance.png'), [CENT, CENT_CELLS]
            return (f'Centrosome\u2013Synapse Distance \u2264 {t} {MU}m',
                    check(d / f'centrosome_distance_le{t}.png'), [CENT_CELLS])
        stem, suffix, t = entry
        g = GATE_TAG[suffix]
        sub = SUB50 if stem == 'z50' else SUB75
        cap = f'GzmB+ > {g}' if t is None else f'Granule {sub} \u2264 {t} {MU}m, GzmB+ > {g}'
        png = check(d / (f'{stem}{suffix}.png' if t is None else f'{stem}{suffix}_le{t}.png'))
        return cap, png, [gate_note[suffix]]

    def result_slide(title, entries):
        parts = [panel(e) for e in entries]
        notes = list(dict.fromkeys(n for _, _, ns in parts for n in ns))   # unique, in order
        if set(gate_note.values()) <= set(notes):                         # both gates: one line
            notes = [n for n in notes if n not in gate_note.values()] + [gate_both]
        is_fraction = any(e[-1] is not None for e in entries)
        note = ' '.join(notes + [FRACTION_NOTE if is_fraction else VIOLIN_NOTE])
        caps = [(c if len(parts) > 1 else '', p) for c, p, _ in parts]
        panels_slide(deck, title, caps, note)

    for title, entries in MAIN:
        result_slide(title, entries)
    s = deck.prs.slides.add_slide(deck.prs.slide_layouts[6])
    deck.titles.append('Backup')
    bpd.add_text(s, 0.83, 3.0, 11.67, 1.2, 'Backup', 44, bold=True, align=PP_ALIGN.CENTER)
    bpd.add_text(s, 0.83, 4.1, 11.67, 0.6, f'Granule slides at GzmB+ > 10\u2075; {SUB75}', 22,
                 align=PP_ALIGN.CENTER)
    for title, entries in BACKUP:
        result_slide(title, entries)

    deck.save()
    print(f'saved {OUT} ({len(deck.prs.slides)} slides)')
    if export:
        bpd.export_pngs(Path(export), deck.titles)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--export', help='also export one PNG per slide (named by title) under this dir')
    args = ap.parse_args()
    build(args.export)
    if MISSING:
        print(f'\n[MISSING] {len(MISSING)} panel(s):')
        for m in MISSING:
            print(f'  {m}')
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
