"""One command for the whole Fig 6 metrics deck rebuild, timed per step (FF 2026-11-19: "it should not take so long
to rebuild! ... improve the infra").

Steps: table + contents images -> deck (batch_insert_images.py) -> deterministic deck check -> slide PNGs (PowerPoint
COM, read-only) -> self-contained review page (slide PNGs as base64). Each step prints its time and a warning when it
runs far over its measured cost (2026-11-19 on this PC, 113 slides: tables 15 s, deck 11 s, export 35 s).

Usage:  python rebuild_fig6_deck.py <config.yaml> <page.html> [--from STEP]
        STEP = tables | deck | export | page   (skip the steps before it; the check always runs)
The figure-reviewer pass and the claude.ai publish stay separate (agent and tool steps).
Not publish-figures/gallery.py: it embeds slide PNGs at full size (113 slides would exceed the 16 MB page limit) and
its subtitle uses a middle dot.
"""
import base64
import html
import io
import os
import re
import subprocess
import sys
import time
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import yaml
from PIL import Image
from pptx import Presentation

HERE = Path(__file__).resolve().parent
EXAMPLES = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(EXAMPLES.parent))
import render_fig6_table_slides as tables  # noqa: E402
from ppt_image_inserter.approved_text import deck_strings, load_approved, unapproved  # noqa: E402

FIGS = "Q:/FF/Nucleus_Data/3D_Nucleus/Fixed/fig6_metrics_page/figs_20261118"
APPROVED = HERE / "approved_text.yaml"   # every string a slide may carry, with its source
EXPECTED_S = {"tables": 15, "deck": 11, "check": 5, "export": 35, "measure": 30, "page": 40}
STEPS = ["tables", "deck", "export", "page"]
PAGE_W = 1400


def timed(name, fn, *args):
    t = time.perf_counter()
    out = fn(*args)
    dt = time.perf_counter() - t
    flag = f"  [SLOW: expected ~{EXPECTED_S[name]} s]" if dt > 3 * EXPECTED_S[name] else ""
    print(f"[{name}] {dt:.1f} s{flag}", flush=True)
    return out


def run(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode:
        print(r.stdout[-3000:], r.stderr[-3000:])
        raise SystemExit(f"failed: {' '.join(map(str, cmd))}")
    return r.stdout


def check(cfg):
    """Per slide: one picture, its config title in the slide text, alt text ending in the image's file name; title
    slide text as configured; no forbidden words (below); contents slide ranges match."""
    prs = Presentation(cfg["output_path"])
    slides, ims, bad = list(prs.slides), cfg["images"], []
    if prs.slides[0].shapes.title.text_frame.text.replace("\v", "\n") != re.sub(r"^\*(.+)\*$", r"\1", cfg["title_slide_text"], flags=re.M):
        bad.append("title slide text differs from title_slide_text")
    if len(slides) != len(ims) + 1:
        bad.append(f"{len(slides)} slides for {len(ims)} images + title")
    for n, (s, e) in enumerate(zip(slides[1:], ims), 2):
        pics = [sh for sh in s.shapes if sh.shape_type == 13]
        text = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
        descr = pics[0]._element.nvPicPr.cNvPr.get("descr", "") if pics else ""
        if not (len(pics) == 1 and e["title"] in text and descr.endswith(os.path.basename(e["path"]))):
            bad.append(f"slide {n}")
    # Never on a slide (FF 2026-11-19): "do not use \u00b7 anywere", "do not mention voxels", "do not use README in
    # slides...ever", "no need to mention C, D, E, F". Checked in slide text and in the rendered table/contents HTML.
    texts = [sh.text_frame.text for s in slides for sh in s.shapes if sh.has_text_frame]
    texts += [p.read_text(encoding="utf-8") for p in (Path(cfg["output_path"]).parent / "table_png").glob("*.html")]
    for word in ("\u00b7", "readme", "voxel", "fig. 6c", "fig 6c"):
        hits = sum(t.lower().count(word) for t in texts)
        if hits:
            bad.append(f"{hits} x '{word}' in slide text")
    if any(re.match(r"[A-F]: ", e["title"]) for e in ims):
        bad.append("a title starts with a panel letter")
    # No character limits on contents rows or notes (FF 2026-11-19: "there should not be any character limit")
    try:   # contents ranges are computed from the titles; this fails unless each row is one unbroken block
        tables.contents_rows([e["title"] for e in ims])
    except SystemExit as e:
        bad.append(str(e))
    # unapproved wording is reported, never blocking (FF 2026-11-19: "if I tell you to build it - you build it. period")
    for t in unapproved_text(cfg):
        print(f"[TEXT] {t}")
    print(f"check: {len(slides)} slides, {len(bad)} problems" + (": " + "; ".join(bad) if bad else ""))
    if bad:
        raise SystemExit("deck check failed")


def slide_strings(cfg):
    """(where, text) for every piece of text on a slide: each line of slide text in the deck, plus the text drawn in
    the rendered table / contents images. Skipped: numbers, and arm names in table rows (td class "l", condition
    names from the data)."""
    out = deck_strings(Presentation(cfg["output_path"]))
    for p in sorted((Path(cfg["output_path"]).parent / "table_png").glob("*.html")):
        t = re.sub(r"<style.*?</style>|<title>.*?</title>", "", p.read_text(encoding="utf-8"), flags=re.S)
        for m in re.finditer(r"<(/?)(\w+)([^>]*)>([^<]*)", t):
            s = html.unescape(m.group(4)).replace("\xa0", " ").strip()
            arm = m.group(2) == "td" and not m.group(1) and 'class="l"' in m.group(3)
            if s and not arm and not re.fullmatch(r"[\d.±\s*–]+|n/a", s):
                out.append((p.stem, s))
    return out


def unapproved_text(cfg):
    """Every slide string must be in approved_text.yaml (FF 2026-11-19: "we need to prevent these mistakes...period"),
    via the shared ppt_image_inserter.approved_text check."""
    return unapproved(slide_strings(cfg), load_approved(APPROVED))


def measure_titles(cfg):
    """Every title at the deck's fixed size, measured in PowerPoint; reports any title over 2 lines, running into the
    figure, or a bottom line that wraps or overlaps (FF 2026-11-20: "titles sould never span 3 lines"). Reported,
    never blocking."""
    sys.path.insert(0, str(EXAMPLES))
    from measure_title_lines import measure
    bad = []
    for m in measure(cfg["output_path"]):
        if m["slide"] == 1:
            continue
        if m["lines"] > 2 or m["runs_into_figure"]:
            bad.append(f"slide {m['slide']}: title {m['lines']} lines" + (", into the figure" if m["runs_into_figure"] else ""))
        if m.get("bottom") and (m["bottom_lines"] > 1 or m["bottom_overlaps_figure"] or m["bottom_off_slide"]):
            bad.append(f"slide {m['slide']}: bottom line {m['bottom_lines']} lines"
                       + (", overlaps the figure" if m["bottom_overlaps_figure"] else "")
                       + (", off the slide" if m["bottom_off_slide"] else ""))
    print(f"titles: {len(bad)} problems" + (": " + "; ".join(bad) if bad else f" (all at most 2 lines at {cfg['title_size']} pt)"))


def _embed(args):
    path, quantize = args
    im = Image.open(path).convert("RGB")
    if im.width != PAGE_W:
        im = im.resize((PAGE_W, round(im.height * PAGE_W / im.width)), Image.LANCZOS)
    if quantize:
        im = im.quantize(256)
    buf = io.BytesIO()
    im.save(buf, "PNG", optimize=True)
    return base64.b64encode(buf.getvalue()).decode(), im.width, im.height


def page(cfg, png_dir, out_html):
    """Self-contained page: every slide PNG as base64; table and contents slides unquantized, figures 256 colors."""
    pngs = sorted(Path(png_dir).glob("slide_*.png"), key=lambda p: int(p.stem.split("_")[1]))
    caps = ["Title slide"] + [e["title"] for e in cfg["images"]]
    if len(pngs) != len(caps):
        raise SystemExit(f"{len(pngs)} slide PNGs for {len(caps)} slides")
    exact = {n for n, e in enumerate(cfg["images"], 2) if "table_png" in e["path"]}
    with ProcessPoolExecutor(max_workers=min(12, os.cpu_count() or 1)) as pool:
        imgs = list(pool.map(_embed, [(p, n not in exact) for n, p in enumerate(pngs, 1)]))
    title = cfg["title_slide_text"].split("\n")[0]
    figs = "".join(
        f'<figure id="s{n}"><figcaption><span class="n">{n}</span> {html.escape(c)}</figcaption>'
        f'<img src="data:image/png;base64,{b}" alt="{html.escape(c)}" loading="lazy" width="{w}" height="{h}"></figure>\n'
        for n, (c, (b, w, h)) in enumerate(zip(caps, imgs), 1))
    doc = (PAGE_HEAD.replace("{TITLE}", html.escape(title))
           + f'<p class="sub">{len(pngs)} slides exported from {html.escape(cfg["output_path"])}. '
             "Deck check passed.</p>\n" + figs + "</main></body></html>\n")
    Path(out_html).write_text(doc, encoding="utf-8")
    print(f"page: {out_html} ({len(doc.encode()) / 1e6:.1f} MB)")


PAGE_HEAD = """<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Fig 6 Metrics Deck</title>
<style>
:root{--bg:#faf9f5;--fg:#1f1e1b;--muted:#6b6a65;--card:#ffffff;--line:#e3e1da}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#1c1b19;--fg:#ecebe6;--muted:#a3a29c;--card:#262522;--line:#3a3935}}
:root[data-theme="dark"]{--bg:#1c1b19;--fg:#ecebe6;--muted:#a3a29c;--card:#262522;--line:#3a3935}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,-apple-system,Segoe UI,sans-serif}
main{max-width:1240px;margin:0 auto;padding:20px 16px 48px}
h1{font-size:22px;margin:0 0 4px}
p.sub{color:var(--muted);margin:0 0 20px;overflow-wrap:anywhere}
figure{margin:0 0 22px;background:var(--card);border:1px solid var(--line);border-radius:8px;overflow:hidden}
figcaption{padding:8px 12px;font-size:14px;color:var(--muted);border-bottom:1px solid var(--line)}
.n{display:inline-block;min-width:2.2em;font-weight:600;color:var(--fg)}
img{display:block;width:100%;height:auto;background:#fff}
</style></head><body><main>
<h1>{TITLE}</h1>
"""


def main(cfg_path, out_html, start="tables"):
    cfg = yaml.safe_load(open(cfg_path, encoding="utf-8"))
    deck_dir = Path(cfg["output_path"]).parent
    todo = STEPS[STEPS.index(start):]
    t0 = time.perf_counter()
    if "tables" in todo:
        timed("tables", tables.main, FIGS, str(deck_dir / "table_png"), cfg_path)
    if "deck" in todo:
        log = timed("deck", run, [sys.executable, str(EXAMPLES / "batch_insert_images.py"), cfg_path])
        (deck_dir / "build.log").write_text(log, encoding="utf-8")
    timed("check", check, cfg)
    slides_dir = Path(cfg.get("slides_png_dir", deck_dir / "slides_png"))   # one folder per deck
    if "export" in todo:
        for old in slides_dir.glob("slide_*.png"):   # the exporter's numbering width follows the slide count, so a
            old.unlink()                             # shorter deck would leave the old 3-digit files behind
        timed("export", run, [sys.executable, str(EXAMPLES / "export_deck_slides.py"), cfg["output_path"],
                              str(slides_dir)])
        timed("measure", measure_titles, cfg)
    timed("page", page, cfg, slides_dir, out_html)
    print(f"total {time.perf_counter() - t0:.1f} s")


if __name__ == "__main__":
    a = sys.argv[1:]
    frm = a[a.index("--from") + 1] if "--from" in a else "tables"
    main(a[0], a[1], frm)
