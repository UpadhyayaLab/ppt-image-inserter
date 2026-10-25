#!/usr/bin/env python3
r"""build_rep_picker.py  Clickable candidate gallery for FF to choose representative
CAT / FMC synapses (one self-contained HTML, published as an Artifact).

Inputs: <rep_dir>\candidates_{CAT,FMC}.csv (make_pmlc_rep_panels.py) and the renders in
<rep_dir>\gallery_candidates\cand_<cond>_<cell>_{merge,<marker>}.png
(tests/run_pzap70_rep_candidates.m). Tiles are ranked by closeness to the condition
medians; "Copy choice" yields the sentence to paste back into the chat.

Run: python build_rep_picker.py --marker pZap70 --out <scratch>\rep_picker_pzap70.html
"""
from __future__ import annotations

import argparse
import base64
import csv
import io
import sys
from pathlib import Path

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_pmlc_profiles_deck import newest_run  # noqa: E402  (same folder pick as the decks)

SETS = {
    'pZap70': dict(root=r'L:\FF\CAR T\CAR T data for actin area trouble shoot\results_compiled_tirf',
                   pattern='pzap70_reps_*', cond='D3_{tag}_15min', sub='15 min · Feb 2025 D3',
                   suffix='pzap70', clim='[0 500]'),
    # FF 2026-09-29: actin-only Jul 17 2026 confocal set (make_pmlc_rep_panels.py --marker
    # actin + tests/run_actin_rep_candidates.m); tiles = actin + zones, last metric = area.
    'actin': dict(root=r'\\10.225.56.176\ULab_data\Data_to_share\FF\New folder\results_compiled',
                  pattern='actin_reps_*', cond='Jul_17_{tag}_10min', sub='10 min · Jul 17, 2026',
                  imgs=('actin', 'zones'), last=('Area', '{:.0f} μm²'),
                  note='Actin (per-cell contrast) + zones; ranked by closeness to the condition '
                       'medians of relative zone actin (zone / synapse MFI) and synapse area.'),
}


def thumb(p, size=230):
    im = Image.open(p).convert('RGB').resize((size, size), Image.LANCZOS)
    b = io.BytesIO(); im.save(b, 'JPEG', quality=84)
    return 'data:image/jpeg;base64,' + base64.b64encode(b.getvalue()).decode()


def section(rd, cfg, marker, tag, current):
    rows = list(csv.reader(open(rd / f'candidates_{tag}.csv')))
    med = rows[1]
    cand = rd / 'gallery_candidates'
    tiles = []
    lname, lfmt = cfg.get('last', ('MFI', '{:.0f}'))
    for rank, r in enumerate(rows[2:], start=1):
        cell = r[0]
        stem = cand / f"cand_{cfg['cond'].format(tag=tag)}_{cell}"
        im1, im2 = cfg.get('imgs', ('merge', cfg.get('suffix')))
        if not Path(f'{stem}_{im1}.png').is_file():
            continue
        cen, per, dist, mfi = (float(x) for x in r[2:6])
        cur = ' is-current' if current.get(tag) == cell else ''
        tiles.append(f"""
      <button class="tile{cur}" data-tag="{tag}" data-cell="{cell}" type="button" id="t-{tag}-{cell}">
        <span class="imgs"><img src="{thumb(f'{stem}_{im1}.png')}" alt="{tag} {cell} {im1}"><img src="{thumb(f'{stem}_{im2}.png')}" alt="{tag} {cell} {im2}"></span>
        <span class="cap"><span class="id">{cell}</span><span class="rank">#{rank}</span></span>
        <span class="m"><span>cen <b>{cen:.2f}</b></span><span>per <b>{per:.2f}</b></span><span>dist <b>{dist:.2f}</b></span><span>{lname} <b>{lfmt.format(mfi)}</b></span></span>
      </button>""")
    medline = (f'median: cen {float(med[2]):.2f} &middot; per {float(med[3]):.2f} &middot; '
               f'dist {float(med[4]):.2f} &middot; {lname} {lfmt.format(float(med[5]))}')
    return f"""
  <section class="grp {tag.lower()}" aria-labelledby="h-{tag}">
    <header class="grp-h"><h2 id="h-{tag}">{tag} <span class="sub">{cfg['sub']}</span></h2><p class="med">{medline}</p></header>
    <div class="grid">{''.join(tiles)}
    </div>
  </section>"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--marker', default='pZap70', choices=list(SETS))
    ap.add_argument('--out', required=True)
    a = ap.parse_args()
    cfg = SETS[a.marker]
    note = cfg.get('note', f"Merge + {a.marker} at fixed {a.marker} contrast {cfg.get('clim')}; "
                           'ranked by closeness to the condition medians.')
    # newest_run, not sorted()[-1]: the decks rank folders by date + v<N> tag, so a name-sort
    # could show FF a different rep folder than the deck uses (bug hunt 2026-10-24, 2nd pass).
    rd = newest_run(cfg['root'], cfg['pattern'])
    if rd is None:
        raise SystemExit(f"[ERROR] no {cfg['pattern']} folder under {cfg['root']}")
    current = {}
    with open(rd / 'picks.csv') as fh:
        for r in csv.DictReader(fh):
            if not r['cond'].startswith('alt'):
                current[r['tag']] = r['cell']
    html = f"""<title>{a.marker} Representative Cell Picker</title>
<style>
:root {{ --ground:#0f1213; --surface:#181c1d; --line:#2a3031; --text:#e4eae8; --muted:#8b9693;
  --cat:#3a78dc; --fmc:#dc4747; --pick:#f2f5f4; --sans:system-ui,"Segoe UI",sans-serif; --mono:ui-monospace,Consolas,monospace; }}
* {{ box-sizing:border-box; }}
body {{ background:var(--ground); color:var(--text); font-family:var(--sans); margin:0; padding-inline:20px; padding-block:18px 40px; }}
.bar {{ position:sticky; top:0; z-index:2; background:var(--ground); padding-block:10px; border-bottom:1px solid var(--line);
  display:flex; flex-wrap:wrap; gap:10px 22px; align-items:baseline; }}
.bar h1 {{ font-size:20px; font-weight:600; margin:0; }}
.bar .note {{ color:var(--muted); font-size:14px; margin:0; }}
.sel {{ font-family:var(--mono); font-size:14px; display:flex; gap:14px; flex-wrap:wrap; }}
.sel .c {{ color:var(--cat); }} .sel .f {{ color:var(--fmc); }}
#copy {{ font:500 14px var(--sans); color:var(--ground); background:var(--pick); border:0; border-radius:4px; padding:6px 12px; cursor:pointer; }}
#copy:focus-visible, .tile:focus-visible {{ outline:2px solid var(--pick); outline-offset:2px; }}
.grp {{ margin-block:22px; }}
.grp-h {{ display:flex; flex-wrap:wrap; gap:4px 18px; align-items:baseline; margin-bottom:10px; }}
.grp h2 {{ font-size:26px; font-weight:600; margin:0; }}
.cat h2 {{ color:var(--cat); }} .fmc h2 {{ color:var(--fmc); }}
.grp h2 .sub {{ color:var(--muted); font-size:15px; font-weight:400; }}
.med {{ font-family:var(--mono); font-size:13px; color:var(--muted); margin:0; font-variant-numeric:tabular-nums; }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fill, minmax(270px, 1fr)); gap:12px; }}
.tile {{ all:unset; box-sizing:border-box; display:flex; flex-direction:column; gap:6px; padding:8px; background:var(--surface);
  border:1px solid var(--line); border-radius:6px; cursor:pointer; }}
.tile:hover {{ border-color:var(--muted); }}
.imgs {{ display:grid; grid-template-columns:1fr 1fr; gap:4px; }}
.imgs img {{ width:100%; aspect-ratio:1; display:block; max-width:100%; }}
.cap {{ display:flex; justify-content:space-between; align-items:baseline; }}
.id {{ font-family:var(--mono); font-size:16px; font-weight:600; }}
.rank {{ font-family:var(--mono); font-size:12px; color:var(--muted); }}
.m {{ display:grid; grid-template-columns:repeat(4, auto); justify-content:space-between; font-family:var(--mono); font-size:12px;
  color:var(--muted); font-variant-numeric:tabular-nums; }}
.m b {{ color:var(--text); font-weight:400; }}
.is-current {{ box-shadow:inset 0 0 0 1px var(--muted); }}
.is-current .rank::after {{ content:" current"; }}
.is-picked {{ border-color:var(--pick); box-shadow:inset 0 0 0 2px var(--pick); }}
@media (max-width:420px) {{ .grid {{ grid-template-columns:1fr; }} }}
</style>
<div class="bar">
  <h1>{a.marker} Representative Cell Picker</h1>
  <p class="note">{note} Click one tile per condition.</p>
  <div class="sel" aria-live="polite"><span class="c">CAT <b id="s-CAT">{current.get('CAT', '')}</b></span><span class="f">FMC <b id="s-FMC">{current.get('FMC', '')}</b></span></div>
  <button id="copy" type="button">Copy choice</button>
</div>
{section(rd, cfg, a.marker, 'CAT', current)}
{section(rd, cfg, a.marker, 'FMC', current)}
<script>
const picked = {{CAT: "{current.get('CAT', '')}", FMC: "{current.get('FMC', '')}"}};
document.querySelectorAll('.tile').forEach(t => t.addEventListener('click', () => {{
  const tag = t.dataset.tag;
  document.querySelectorAll('.tile[data-tag="' + tag + '"]').forEach(x => x.classList.remove('is-picked'));
  t.classList.add('is-picked');
  picked[tag] = t.dataset.cell;
  document.getElementById('s-' + tag).textContent = t.dataset.cell;
}}));
const btn = document.getElementById('copy');
btn.addEventListener('click', async () => {{
  const txt = 'Use {a.marker} CAT ' + picked.CAT + ' and FMC ' + picked.FMC + ' as representative cells';
  try {{ await navigator.clipboard.writeText(txt); btn.textContent = 'Copied'; }}
  catch (e) {{ btn.textContent = txt; }}
  setTimeout(() => btn.textContent = 'Copy choice', 1800);
}});
</script>
"""
    Path(a.out).write_text(html, encoding='utf-8')
    print(a.out, round(Path(a.out).stat().st_size / 1e6, 2), 'MB')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
