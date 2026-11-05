#!/usr/bin/env python3
r"""build_cd19_synapse_otsu_review_page.py  One-page review of the Otsu synapse smoke.

FF 2026-10-31 ("make it as easy as possible for me to review"): the review deck's slide
renders (build_cd19_synapse_otsu_review.py) as one card per cell, the cells Claude flagged
first, each with Claude's one-line read and three verdict buttons (Otsu OK / too much /
too little) plus a note. Verdicts live in the artifact's db (collection `verdicts`, one doc
per cell, id `<exp>__<cond>__<label>`) so they can be read back from another PC.

Inputs: <smoke>/picks.json, <smoke>/review/deck_png/ (NN title.png, slide order = picks
order) and --notes JSON {"<exp>/<cond>/<label>": {"flag": bool, "note": str}}.

Run:  python build_cd19_synapse_otsu_review_page.py --notes FILE --out DIR
Out:  DIR/index.html (slide JPEGs embedded) + DIR/img/*.jpg
"""
from __future__ import annotations

import argparse
import base64
import html
import json
import re
from pathlib import Path

from PIL import Image

ROOT = Path(r'Y:\User_data\Kiet')
SMOKE = ROOT / '20261001_coculture_10_30_60' / 'volscope' / 'tests' / 'synapse_otsu_smoke_20261031'
EXP = {'Oct1': 'Oct 1, 2026', 'Mar10': 'Mar 10, 2026'}
TAGS = {'min': 'smallest', 'median': 'median', 'max': 'largest'}
DECK = (r'Y:\User_data\Kiet\20261001_coculture_10_30_60\volscope\tests\synapse_otsu_smoke_20261031'
        r'\review\synapse_otsu_review.pptx')


def cond_label(cond):
    m = re.match(r'(CAT|FMC)_(\d+)min', cond)
    return f'{m.group(1)} {m.group(2)} min' if m else cond


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--smoke', type=Path, default=SMOKE)
    ap.add_argument('--notes', type=Path, required=True)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--summary', default='', help="Claude's read of the whole set, shown at the top")
    args = ap.parse_args()
    with open(args.smoke / 'picks.json', encoding='utf-8') as h:
        picks = json.load(h)
    with open(args.notes, encoding='utf-8') as h:
        notes = json.load(h)
    pngs = sorted((args.smoke / 'review' / 'deck_png').glob('*.png'))
    if len(pngs) != len(picks):
        raise SystemExit(f'{len(pngs)} slide PNGs for {len(picks)} cells: rebuild the deck first')
    (args.out / 'img').mkdir(parents=True, exist_ok=True)
    cards = []
    for i, (p, png) in enumerate(zip(picks, pngs)):
        key = f'{p["exp"]}/{p["cond"]}/{p["label"]}'
        # notes may name the condition without the folder's suffix (Mar 10: CAT_10min_001)
        short = f'{p["exp"]}/{cond_label(p["cond"]).replace(" min", "min").replace(" ", "_")}/{p["label"]}'
        n = notes.get(key) or notes.get(short) or {'flag': False, 'note': ''}
        img = f'img/{i + 1:02d}.jpg'
        im = Image.open(png).convert('RGB')
        im.thumbnail((1600, 1600))
        im.save(args.out / img, quality=85)
        # embedded too: the page is one file (a publish of img/ files was refused)
        img = 'data:image/jpeg;base64,' + base64.b64encode((args.out / img).read_bytes()).decode()
        cards.append(dict(id=key.replace('/', '__'), key=key, flag=bool(n['flag']), note=n['note'],
                          title=f'{EXP[p["exp"]]} \u00b7 {cond_label(p["cond"])} \u00b7 {p["label"]}',
                          tag=TAGS[p['tag']], old=round(100 * p['mad_frac']),
                          otsu=round(100 * p['otsu_frac']), img=img, w=im.width, h=im.height))
    cards.sort(key=lambda c: not c['flag'])                  # flagged first, then slide order
    nflag = sum(c['flag'] for c in cards)
    page = (TEMPLATE.replace('__CARDS__', json.dumps(cards))
            .replace('__NFLAG__', str(nflag)).replace('__N__', str(len(cards)))
            .replace('__DECK__', html.escape(DECK))
            .replace('__SUMMARY__', f'<p class="summary">{html.escape(args.summary)}</p>' if args.summary else ''))
    (args.out / 'index.html').write_text(page, encoding='utf-8')
    print(f'wrote {args.out / "index.html"} ({len(cards)} cells, {nflag} flagged)')


TEMPLATE = r"""<title>Otsu Synapse Smoke Review</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono&family=IBM+Plex+Sans:wght@400;600&display=swap">
<style>
/* Layout: one column of full-width slide cards, verdict bar under each; flagged cells first. */
:root {
  --bg: #f6f7f9; --card: #ffffff; --fg: #1d2330; --muted: #5b6475; --line: #dde1e8;
  --accent: #b5267a;            /* CD19 magenta, as on the slides */
  --ok: #1f7a4d; --warn: #a15c00; --flag-bg: #fff4e5;
  --sans: "IBM Plex Sans", "Segoe UI", system-ui, sans-serif;
  --mono: "IBM Plex Mono", Consolas, monospace;
}
@media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
  --bg: #12151b; --card: #1b1f27; --fg: #e7eaf0; --muted: #9aa3b4; --line: #2c3240;
  --accent: #e766b0; --ok: #4cc38a; --warn: #f0a640; --flag-bg: #2a2114; color-scheme: dark } }
:root[data-theme="dark"] {
  --bg: #12151b; --card: #1b1f27; --fg: #e7eaf0; --muted: #9aa3b4; --line: #2c3240;
  --accent: #e766b0; --ok: #4cc38a; --warn: #f0a640; --flag-bg: #2a2114; color-scheme: dark }
body { background: var(--bg); color: var(--fg); font: 16px/1.5 var(--sans);
  padding-left: max(16px, env(safe-area-inset-left, 0px));
  padding-right: max(16px, env(safe-area-inset-right, 0px)); padding-block: 20px 60px; }
main { max-width: 1240px; margin: 0 auto; display: grid; gap: 22px; }
h1 { font-size: 1.6rem; margin: 0; text-wrap: balance; }
.lede { color: var(--muted); margin: 4px 0 0; max-width: 75ch; }
.summary { margin: 12px 0 0; padding: 12px 16px; max-width: 75ch; background: var(--flag-bg);
  border-left: 4px solid var(--warn); border-radius: 4px; }
.lede code { font-family: var(--mono); font-size: .85em; word-break: break-all; }
.bar { position: sticky; top: env(safe-area-inset-top, 0px); z-index: 5; background: var(--bg);
  padding-block: 10px; display: flex; flex-wrap: wrap; gap: 10px 18px; align-items: center;
  border-bottom: 1px solid var(--line); font-variant-numeric: tabular-nums; }
body::before { content: ""; position: fixed; inset: 0 0 auto; height: env(safe-area-inset-top, 0px);
  z-index: 10; background: var(--bg); }
.bar strong { font-size: 1.1rem; }
.bar label { color: var(--muted); display: flex; gap: 6px; align-items: center; }
.status { color: var(--muted); font-size: .9rem; }
h2 { font-size: 1.15rem; margin: 10px 0 0; color: var(--muted); text-transform: uppercase;
  letter-spacing: .06em; }
.card { background: var(--card); border: 1px solid var(--line); border-radius: 8px; overflow: hidden; }
.card.flag { border-color: var(--warn); }
.head { padding: 12px 16px; display: flex; flex-wrap: wrap; gap: 4px 16px; align-items: baseline; }
.head h3 { margin: 0; font-size: 1.15rem; }
.nums { font-variant-numeric: tabular-nums; color: var(--muted); }
.nums b { color: var(--accent); }
.note { padding: 10px 16px; border-top: 1px solid var(--line); }
.card.flag .note { background: var(--flag-bg); }
.note .who { font-weight: 600; margin-right: 6px; }
.card.flag .note .who { color: var(--warn); }
.shot { display: block; width: 100%; height: auto; max-width: 100%; background: #fff; }
.vote { padding: 12px 16px; border-top: 1px solid var(--line); display: flex; flex-wrap: wrap;
  gap: 8px; align-items: center; }
.vote button { font: inherit; font-weight: 600; padding: 7px 14px; border-radius: 6px; cursor: pointer;
  border: 1.5px solid var(--line); background: var(--card); color: var(--fg); }
.vote button:focus-visible, .vote input:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.vote button[aria-pressed="true"][data-v="ok"] { background: var(--ok); border-color: var(--ok); color: #fff; }
.vote button[aria-pressed="true"]:not([data-v="ok"]) { background: var(--warn); border-color: var(--warn); color: #fff; }
.vote input { flex: 1 1 220px; min-width: 0; font: inherit; padding: 7px 10px; border-radius: 6px;
  border: 1.5px solid var(--line); background: var(--bg); color: var(--fg); }
.saved { color: var(--muted); font-size: .85rem; min-width: 5ch; }
.card[data-done="1"] .head h3::after { content: " \2713"; color: var(--ok); }
</style>

<main>
  <header>
    <h1>Otsu synapse cut: smoke review</h1>
    <p class="lede">36 cells (3 per condition: smallest, median and largest synapse under the old
    cut). Each picture: CD19 relative density | old synapse (median + 2 &times; MAD) | Otsu synapse,
    turned so both patches face you. Is the Otsu patch a reasonable synapse? Rotatable deck:
    <code>__DECK__</code></p>
    __SUMMARY__
  </header>
  <div class="bar">
    <strong id="count">0 / __N__ reviewed</strong>
    <label><input type="checkbox" id="todo"> Hide reviewed</label>
    <span class="status" id="status">Connecting&hellip;</span>
  </div>
  <h2 id="h-flag">Claude flagged (__NFLAG__)</h2>
  <div id="flagged" style="display:grid;gap:22px"></div>
  <h2 id="h-rest">Looked reasonable to Claude</h2>
  <div id="rest" style="display:grid;gap:22px"></div>
</main>

<script>
const CARDS = __CARDS__;
const OPTS = [['ok', 'Otsu OK'], ['too_much', 'Too much'], ['too_little', 'Too little']];
const state = {};          // id -> {verdict, comment}
let col = null;

function esc(s) { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; }

function render() {
  for (const c of CARDS) {
    const el = document.getElementById('c-' + c.id);
    const s = state[c.id] || {};
    el.dataset.done = s.verdict ? '1' : '0';
    for (const b of el.querySelectorAll('button')) b.setAttribute('aria-pressed', String(b.dataset.v === s.verdict));
    const inp = el.querySelector('input');
    if (document.activeElement !== inp) inp.value = s.comment || '';
    el.hidden = document.getElementById('todo').checked && !!s.verdict;
  }
  const n = CARDS.filter(c => (state[c.id] || {}).verdict).length;
  document.getElementById('count').textContent = `${n} / ${CARDS.length} reviewed`;
}

function build() {
  for (const c of CARDS) {
    const el = document.createElement('article');
    el.className = 'card' + (c.flag ? ' flag' : '');
    el.id = 'c-' + c.id;
    el.innerHTML = `
      <div class="head"><h3>${esc(c.title)}</h3>
        <span class="nums">${esc(c.tag)} old synapse &middot; old <b>${c.old}%</b> &rarr; Otsu <b>${c.otsu}%</b> of surface</span></div>
      ${c.note ? `<div class="note"><span class="who">${c.flag ? 'Claude flags:' : 'Claude:'}</span>${esc(c.note)}</div>` : ''}
      <img class="shot" src="${c.img}" width="${c.w}" height="${c.h}" loading="lazy" alt="${esc(c.title)}: relative density, old and Otsu synapse">
      <div class="vote">
        ${OPTS.map(([v, t]) => `<button type="button" data-v="${v}" aria-pressed="false">${t}</button>`).join('')}
        <input id="n-${c.id}" type="text" placeholder="Note (optional, saved when you leave the box)" aria-label="Note for ${esc(c.title)}">
        <span class="saved"></span>
      </div>`;
    (c.flag ? document.getElementById('flagged') : document.getElementById('rest')).appendChild(el);
    for (const b of el.querySelectorAll('button')) b.onclick = () => save(c, {verdict: (state[c.id] || {}).verdict === b.dataset.v ? '' : b.dataset.v});
    el.querySelector('input').onchange = e => save(c, {comment: e.target.value});
  }
  if (!CARDS.some(c => c.flag)) document.getElementById('h-flag').hidden = true;
  document.getElementById('todo').onchange = render;
}

async function save(c, patch) {
  const next = Object.assign({verdict: '', comment: ''}, state[c.id], patch);
  state[c.id] = next;
  render();
  const tag = document.querySelector('#c-' + c.id + ' .saved');
  if (!col) { tag.textContent = 'not saved (offline)'; return; }
  tag.textContent = 'saving…';
  try {
    await col.doc(c.id).set({key: c.key, verdict: next.verdict, comment: next.comment, at: new Date().toISOString()});
    tag.textContent = 'saved';
  } catch (e) { tag.textContent = 'not saved: ' + (e && e.code || 'error'); }
}

build(); render();
(async () => {
  const st = document.getElementById('status');
  const db = window.claude ? await claude.use('db') : null;
  if (!db) { st.textContent = 'Verdicts cannot be saved in this view; open the page signed in to claude.ai.'; return; }
  col = db.collection('verdicts');
  col.onSnapshot(snap => {
    for (const d of snap.docs) { const v = d.data(); state[d.id] = {verdict: v.verdict || '', comment: v.comment || ''}; }
    st.textContent = 'Verdicts save automatically.';
    render();
  }, e => { st.textContent = 'Lost connection to saved verdicts (' + e.code + '). Reload the page.'; });
})();
</script>
"""

if __name__ == '__main__':
    main()
