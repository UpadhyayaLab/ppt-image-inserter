#!/usr/bin/env python
"""
Review pages for the 10/29/2026 H3K27me3 side-view decks. Slides are embedded in the page
as JPEGs, and a page must stay under 16 MB, so each set splits into two pages (PAGES). Each
page has condition tabs, cell-id buttons and Partial cells marked.

Reads the slide PNGs that export_slides wrote to <EXP>/deck_png_20261031/side_views_all/
and checks each one's cell against the deck's slide title before it goes on a page.

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_side_view_pages.py <output folder>
Writes <output folder>/<page name>.html, ready for the Artifact tool.
"""

import base64
import io
import json
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_h3k27me3_10292026_compare_decks import EXPS, configure, m  # noqa: E402
from build_h3k27me3_10292026_side_view_deck import OUT, SIDE_SUB, partial_cells, side_cells  # noqa: E402

PNG_SUB = "deck_png_20261031/side_views_all"
JPEG_W = 1500          # FF reads fit-to-width at about this size
JPEG_QUALITY = 80      # keeps the largest page under 16 MB
WORKERS = 8
PAGE_LIMIT_MB = 16
NOCO_A = "https://claude.ai/artifact/2CNT5DPJBi8dg3kAQBsSKd"
NOCO_B = "https://claude.ai/artifact/J117zD3UN96ARdntNx3L5r"
CILIO_A = "https://claude.ai/artifact/AnSEkZhEBE4iKwj6adqotW"
CILIO_B = "https://claude.ai/artifact/RqahBx1pEVQnbqKHkFy3sf"
# (file name, set, condition indices, page title, (link, link text) or None)
PAGES = [("noco_pll_dmso", "noco", (0, 1), "Noco 10/29 side views, PLL and DMSO", (NOCO_B, "5 and 10 μM page")),
         ("noco_5_10um", "noco", (2, 3), "Noco 10/29 side views, 5 and 10 μM", (NOCO_A, "PLL and DMSO page")),
         ("cilio_pll_dmso", "cilio", (0, 1), "CilioD 10/29 side views, PLL and DMSO", (CILIO_B, "50 μM page")),
         ("cilio_50um", "cilio", (2,), "CilioD 10/29 side views, 50 μM", (CILIO_A, "PLL and DMSO page"))]

PAGE = """<title>__TITLE__</title>
<meta name="claude:viewport-fit-sides" content="cover">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600&family=IBM+Plex+Mono:wght@500&display=swap">
<style>
/* Light-table layout: controls pinned on top, slides black on black below, one column at reading width. */
:root {
  color-scheme: dark;
  --bg: #07090b; --panel: #12161a; --line: #2a3238; --fg: #e8edf0; --dim: #9aa6ad;
  --hoechst: #2cc8d8; --h3k: #ff5a52; --ink: #041014;
  --sans: "IBM Plex Sans", "Segoe UI", system-ui, sans-serif;
  --mono: "IBM Plex Mono", Consolas, monospace;
}
body { background: var(--bg); color: var(--fg); font: 15px/1.45 var(--sans); }
body::before { content: ""; position: fixed; inset: 0 0 auto; height: env(safe-area-inset-top, 0px); z-index: 10; background: var(--panel); }
.top { position: sticky; top: env(safe-area-inset-top, 0px); z-index: 5; background: var(--panel); border-bottom: 1px solid var(--line); }
.bar, main { max-width: 1500px; margin-inline: auto; padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); }
.bar { display: grid; gap: 8px; padding-block: 10px; }
h1 { margin: 0; font-size: 20px; font-weight: 600; text-wrap: balance; }
.note { margin: 0; color: var(--dim); font-size: 14px; }
.row { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }
a { color: var(--hoechst); }
button { font: inherit; color: var(--fg); background: transparent; border: 1px solid var(--line); border-radius: 3px; padding: 4px 10px; cursor: pointer; }
button:hover { border-color: var(--dim); }
button:focus-visible { outline: 2px solid var(--hoechst); outline-offset: 1px; }
button[aria-pressed="true"] { background: var(--hoechst); border-color: var(--hoechst); color: var(--ink); font-weight: 600; }
.ids { max-height: 76px; overflow-y: auto; }
.ids button { font: 500 13px var(--mono); padding: 2px 7px; min-width: 34px; font-variant-numeric: tabular-nums; }
.ids button.partial { border-color: var(--h3k); color: var(--h3k); }
main { display: grid; gap: 14px; padding-block: 14px 40px; }
figure { margin: 0; scroll-margin-top: 190px; }
figure img { display: block; width: 100%; max-width: 100%; height: auto; aspect-ratio: 16 / 9; background: var(--panel); }
</style>
<div class="top"><div class="bar">
  <h1>__TITLE__</h1>
  <p class="note">__COUNT__ One slide per cell, in cell-id order. Partial cells are marked in red. PRELIMINARY: no figure review of this deck yet.__MORE__</p>
  <div class="row" id="tabs"></div>
  <div class="row ids" id="ids"></div>
</div></div>
<main id="slides"></main>
<script>
const DATA = __DATA__;
const tabs = document.getElementById("tabs"), ids = document.getElementById("ids"), slides = document.getElementById("slides");
const slug = s => s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
let current = 0, partialOnly = false;

function draw() {
  tabs.replaceChildren();
  DATA.forEach((c, i) => {
    const b = document.createElement("button");
    b.textContent = `${c.label} (${c.cells.length})`;
    b.setAttribute("aria-pressed", i === current);
    b.onclick = () => { location.hash = slug(c.label); show(i); };
    tabs.append(b);
  });
  const t = document.createElement("button");
  const nPartial = DATA[current].cells.filter(c => c.partial).length;
  t.textContent = `Partial only (${nPartial})`;
  t.setAttribute("aria-pressed", partialOnly);
  t.onclick = () => { partialOnly = !partialOnly; draw(); };
  tabs.append(t);

  const cond = DATA[current];
  const cells = cond.cells.filter(c => !partialOnly || c.partial);
  ids.replaceChildren();
  slides.replaceChildren();
  cells.forEach(c => {
    const b = document.createElement("button");
    b.textContent = c.id;
    if (c.partial) b.className = "partial";
    b.title = `Go to cell ${c.id}${c.partial ? ", Partial" : ""}`;
    b.onclick = () => document.getElementById("cell-" + c.id).scrollIntoView();
    ids.append(b);
    const f = document.createElement("figure");
    f.id = "cell-" + c.id;
    const img = new Image();
    img.src = c.file;
    img.loading = "lazy";
    img.width = 1500;
    img.height = 844;
    img.alt = `${cond.label}, cell ${c.id}${c.partial ? ", Partial" : ""}`;
    f.append(img);
    slides.append(f);
  });
}
function show(i) { current = i; partialOnly = false; draw(); window.scrollTo(0, 0); }
function fromHash() {
  const i = DATA.findIndex(c => slug(c.label) === location.hash.slice(1));
  show(i < 0 ? 0 : i);
}
window.addEventListener("hashchange", fromHash);
fromHash();
</script>
"""


def to_jpeg(png: Path) -> str:
    """One slide PNG as a JPEG data URI at reading width."""
    with m.Image.open(m._winlong(png)) as im:
        im = im.convert("RGB")
        h = round(im.size[1] * JPEG_W / im.size[0])
        buf = io.BytesIO()
        im.resize((JPEG_W, h), m.Image.LANCZOS).save(buf, "JPEG", quality=JPEG_QUALITY, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


def conditions(name: str) -> list:
    """[{label, cells: [{id, partial, png}]}] in slide order, each checked against the deck."""
    exp = EXPS[name]
    configure(exp)
    root = Path(exp["root"])
    partial = partial_cells()
    titles = [next(sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip())
              for s in Presentation(OUT[name]).slides]
    data, i = [], 0
    for well, label, _, _ in exp["wells"]:
        cells = []
        for cell in side_cells(root / SIDE_SUB, well):
            i += 1
            cid = cell["cell_id"]
            flag = (well, cid) in partial
            want = f"{label}, cell {cid}{', Partial' if flag else ''}"
            if titles[i - 1] != want:
                raise SystemExit(f"slide {i} reads {titles[i - 1]!r}, expected {want!r}")
            cells.append({"id": int(cid), "partial": flag, "png": root / PNG_SUB / f"slide_{i:02d}.png"})
        data.append({"label": label, "cells": cells})
    if i != len(titles):
        raise SystemExit(f"{i} cells for {len(titles)} slides")
    return data


def build(page_name: str, name: str, picks: tuple, title: str, link: tuple, sets: dict, out_root: Path) -> int:
    conds = [sets[name][i] for i in picks]
    cells = [c for d in conds for c in d["cells"]]
    with ThreadPoolExecutor(WORKERS) as pool:
        uris = iter(list(pool.map(to_jpeg, [c["png"] for c in cells])))
    data = [{"label": d["label"], "cells": [{"id": c["id"], "partial": c["partial"], "file": next(uris)}
                                            for c in d["cells"]]} for d in conds]
    more = f' <a href="{link[0]}">Go to the {link[1]}.</a>' if link else ""
    page = (PAGE.replace("__TITLE__", title)
            .replace("__COUNT__", f"{len(cells)} kept cells, {sum(c['partial'] for c in cells)} Partial.")
            .replace("__MORE__", more)
            .replace("__DATA__", json.dumps(data)))
    out = out_root / f"{page_name}.html"
    out.write_text(page, encoding="utf-8")
    mb = out.stat().st_size / 2**20
    print(f"[side view page {page_name}] {len(cells)} slides, {mb:.1f} MB -> {out}")
    return 0 if mb < PAGE_LIMIT_MB else 1


if __name__ == "__main__":
    out_root = Path(sys.argv[1])
    out_root.mkdir(parents=True, exist_ok=True)
    sets = {n: conditions(n) for n in EXPS}
    sys.exit(max(build(p, n, c, t, ln, sets, out_root) for p, n, c, t, ln in PAGES))
