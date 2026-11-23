"""Slides for the Fig 6 decision page, made by the deck's own pipeline (render_fig6_table_slides.py,
batch_insert_images.py, export_deck_slides.py, measure_title_lines.py).

Two decks, each built, exported and measured once:
- "now": the deck as it is (deck config + deck_text.yaml, FF's applied answers included), for the All slides view;
  built with the deck's approved_text check on;
- "options": for every open decision, each option's replacements applied (decision_page/text_map.py, the same
  function apply_answers.py uses) and only the slides whose text changes, one copy per option;
- "dates" (when the spec has an item with `dates_preview: true`): the whole deck with each title's date moved to a
  bottom_text line under the figure (FF 2026-11-20: "dates need not be in title. they can be near bottom"), measured
  in full; the page shows one two-date and one one-date slide and a table of every changed title.
Every slide's text comes from the sources (titles, subtitles, deck_text.yaml lines), never scraped from images.

Writes <OUT>/page/*.png and <OUT>/page/page_data.json, read by decision_page/build_decision_page.py.
Usage:  python fig6_decision_previews.py <decision spec.yaml>
"""
import copy
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
EXAMPLES = HERE.parent
sys.path[:0] = [str(HERE), str(EXAMPLES), str(EXAMPLES / "decision_page"), str(EXAMPLES.parent)]
import render_fig6_table_slides as tables  # noqa: E402
import text_map  # noqa: E402
from measure_title_lines import measure  # noqa: E402

FIGS = "Q:/FF/Nucleus_Data/3D_Nucleus/Fixed/fig6_metrics_page/figs_20261118"
OUT = Path("K:/FF/PPT/PPT_autogeneration/Fig6_metrics_across_experiments/decision_previews")
GEN_CFG = EXAMPLES / "configs/fixed/fig6_metrics"   # generated build configs sit beside the deck config (relative paths)
UNLISTED = {"figure": "Labels inside the figure are not listed.",
            "table": "The numbers, experiment and arm names in the table are not listed.", "contents": "", "title": ""}


DATE = re.compile(r" \(([^()]*\d{4})\)$")   # a title's trailing "(... 2026)"


def move_dates(cfg):
    """Copy of cfg with each title's trailing date in parentheses moved, without the parentheses, to bottom_text."""
    c = copy.deepcopy(cfg)
    for e in c["images"]:
        m = DATE.search(e["title"])
        if m:
            e["title"], e["bottom_text"] = e["title"][:m.start()], m.group(1)
    return c


def step(name, t0):
    print(f"[{name}] {time.perf_counter() - t0:.1f} s", flush=True)
    return time.perf_counter()


def slides(cfg, text, pages, data, cent):
    """[{n, id, kind, title, subtitle, lines, table}] for slides 2..N from the sources."""
    titles = [e["title"] for e in cfg["images"]]
    out = []
    for n, e in enumerate(cfg["images"], 2):
        name = Path(e["path"]).name
        s = {"n": n, "id": Path(name).stem, "title": e["title"], "subtitle": e.get("subtitle", ""),
             "bottom": e.get("bottom_text", ""), "path": e["path"], "kind": "figure", "table": None}
        drawn = []
        if name == "contents.png":
            s["kind"], drawn = "contents", tables.contents_lines(titles, text)
        elif name.startswith("table_"):
            i = int(name.split("_")[1]) - 1
            keys = pages[i][2]
            s["kind"], s["table"] = "table", i
            drawn = tables.header_lines(keys, cent, text) + tables.key_lines(keys, data, cent, text)
        s["lines"] = ([e["title"]] + ([e["subtitle"]] if e.get("subtitle") else []) + drawn
                      + ([e["bottom_text"]] if e.get("bottom_text") else []))
        out.append(s)
    return out


def render_images(sl, cfg, text, pages, data, cent, d):
    """Render the contents/table images these slides need into d; returns {slide n: png path}."""
    d.mkdir(parents=True, exist_ok=True)
    titles = [e["title"] for e in cfg["images"]]
    out = {}
    for s in sl:
        if s["kind"] == "contents":
            tables.render_contents(d, titles, text)
            out[s["n"]] = d / "contents.png"
        elif s["kind"] == "table":
            title, slug, keys = pages[s["table"]]
            page, _, _ = tables.slide_html(title, keys, data, cent, text)
            hp, pp = d / f"table_{s['table'] + 1}_{slug}.html", d / f"table_{s['table'] + 1}_{slug}.png"
            hp.write_text(page, encoding="utf-8")
            subprocess.run([tables.EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                            "--force-device-scale-factor=2", "--window-size=2600,1150", f"--screenshot={pp}",
                            hp.resolve().as_uri()], check=True, capture_output=True, timeout=120)
            out[s["n"]] = pp
    return out


def build(cfg, entries, name, approved):
    """Build one deck from (image path, title, subtitle, bottom text) entries with the deck's own settings; export and measure it.
    Returns (exported PNGs from slide 1, measured titles by slide number, deck strings by slide number)."""
    from pptx import Presentation
    from ppt_image_inserter.approved_text import deck_strings
    c = {k: v for k, v in cfg.items() if k != "images" and (approved or k != "approved_text")}
    deck = OUT / f"{name}.pptx"
    c.update(output_path=str(deck).replace("\\", "/"), backup_dir=str(OUT / "backups").replace("\\", "/"),
             images=[dict({"path": str(p).replace("\\", "/"), "title": t}, **({"subtitle": s} if s else {}),
                          **({"bottom_text": b} if b else {})) for p, t, s, b in entries])
    cp = GEN_CFG / f"config_fig6_decision_{name}_GENERATED.yaml"
    cp.write_text("# GENERATED by fig6_metrics/fig6_decision_previews.py; rewritten on every run.\n"
                  + yaml.safe_dump(c, sort_keys=False, allow_unicode=True, width=200), encoding="utf-8")
    t = time.perf_counter()
    r = subprocess.run([sys.executable, str(EXAMPLES / "batch_insert_images.py"), str(cp)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    (OUT / f"build_{name}.log").write_text(r.stdout + r.stderr, encoding="utf-8")
    if r.returncode or f"{len(entries)}/{len(entries)} slides created" not in r.stdout:
        print(r.stdout[-3000:], r.stderr[-2000:])
        raise SystemExit(f"{name} deck build failed (log: build_{name}.log)")
    t = step(f"{name}: build {len(entries) + 1} slides", t)
    pd = OUT / f"png_{name}"
    shutil.rmtree(pd, ignore_errors=True)
    r = subprocess.run([sys.executable, str(EXAMPLES / "export_deck_slides.py"), str(deck), str(pd)],
                       capture_output=True, text=True)
    if r.returncode:
        raise SystemExit(r.stderr[-2000:])
    pngs = sorted(pd.glob("slide_*.png"), key=lambda p: int(p.stem.split("_")[1]))
    if len(pngs) != len(entries) + 1:
        raise SystemExit(f"{name}: {len(pngs)} exported slides for {len(entries) + 1}")
    t = step(f"{name}: export", t)
    try:
        res = measure(str(deck))
    except Exception as e:   # PowerPoint COM now and then drops the call right after an export ("remote procedure
        print(f"[{name}: measure failed once ({e}); retrying]")   # call failed"); a second try has worked each time
        time.sleep(5)
        res = measure(str(deck))
    meas = {m["slide"]: m for m in res}
    step(f"{name}: measure titles", t)
    strings = {}
    for where, line in deck_strings(Presentation(str(deck))):
        strings.setdefault(int(where.split()[1]), []).append(line)
    log = (OUT / f"build_{name}.log").read_text(encoding="utf-8")   # text_clear, run by batch_insert_images.py
    for m in re.finditer(r"\[TEXT CLEAR\] slide (\d+): figure scale ([\d.]+|None), moved ([-\d.]+|None) in(?: \((.+)\))?", log):
        meas[int(m.group(1))].update(moved=f"scale {m.group(2)}, moved {m.group(3)} in" + (f" ({m.group(4)})" if m.group(4) else ""))
    for m in re.finditer(r"\[TEXT\] slide (\d+): the (title|bottom line) overlaps the data", log):
        meas[int(m.group(1))].setdefault("on_data", []).append(m.group(2))
    return pngs, meas, strings


def fact(m, title, size):
    if m is None or m["title"] != title.strip():
        raise SystemExit(f"measured title {m and m['title']!r} is not {title!r}")
    k = m["lines"]
    on = m.get("on_data", [])
    return {"lines": k, "warn": k >= 3 or m["runs_into_figure"] or bool(on),
            "text": f"Title: {k} line{'s' if k != 1 else ''} at {size} pt"
                    + ("; it runs into the figure" if m["runs_into_figure"] else "")
                    + "".join(f"; the {w} overlaps the data" for w in on)
                    + (f"; figure {m['moved']}" if m.get("moved") else "")}


def main(spec_path):
    t0 = time.perf_counter()
    spec_path = Path(spec_path)
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    cfg = yaml.safe_load((spec_path.parent / spec["deck"]["config"]).read_text(encoding="utf-8"))
    text = yaml.safe_load((spec_path.parent / spec["deck"]["text"]).read_text(encoding="utf-8"))
    ap = spec_path.parent / spec["applied"]
    applied = {str(a["code"]) for a in (yaml.safe_load(ap.read_text(encoding="utf-8")) or []) if "code" in a}
    data, _ = tables.collect(FIGS)
    cent = tables.load_cent(FIGS)
    pages = tables.pages(data, cent)
    size = cfg["title_size"]
    bsize = cfg.get("bottom_text_size", 16)
    page_dir = OUT / "page"
    shutil.rmtree(page_dir, ignore_errors=True)
    page_dir.mkdir(parents=True)
    t = step("load", t0)

    # the deck as it is
    now = slides(cfg, text, pages, data, cent)
    imgs = render_images(now, cfg, text, pages, data, cent, OUT / "src" / "now")
    t = step("now: render contents and tables", t)
    pngs, meas, strings = build(cfg, [(imgs.get(s["n"], s["path"]), s["title"], s["subtitle"], s["bottom"]) for s in now],
                                "now", approved=True)
    title_lines = strings.get(1, [])
    deck = [{"n": 1, "id": "title_slide", "png": "deck_01.png", "lines": title_lines, "unlisted": "", "fact": None}]
    shutil.copyfile(pngs[0], page_dir / "deck_01.png")
    for s in now:
        png = f"deck_{s['n']:02d}.png"
        shutil.copyfile(pngs[s["n"] - 1], page_dir / png)
        deck.append({"n": s["n"], "id": s["id"], "png": png, "lines": s["lines"], "unlisted": UNLISTED[s["kind"]],
                     "fact": fact(meas.get(s["n"]), s["title"], size)})
    if len({x["id"] for x in deck}) != len(deck):
        raise SystemExit("two deck slides share an image name (the page keys slide edits by it)")

    # one options deck: each open option's replacements, only the slides whose text changes; dates already moved to
    # the bottom line (FF 2026-11-20), so each option's titles are measured as they would be
    t = time.perf_counter()
    entries, where, decisions = [], [], {}
    base = slides(move_dates(cfg), text, pages, data, cent)
    for it in spec["items"]:
        code = str(it["code"])
        if code in applied or not it.get("options"):
            continue
        dd = decisions[code] = {"applies": [], "options": {}}
        for o in it["options"]:
            if not o.get("replace"):
                continue
            cfg2, text2, _ = text_map.apply(cfg, text, o["replace"])
            cfg2 = move_dates(cfg2)
            new = slides(cfg2, text2, pages, data, cent)
            changed = [s for s, b in zip(new, base) if s["lines"] != b["lines"]]
            if not changed:
                raise SystemExit(f"decision {code} option {o['choice']}: changes no slide")
            imgs2 = render_images(changed, cfg2, text2, pages, data, cent, OUT / "src" / f"{code}_{o['choice']}")
            dd["options"][str(o["choice"])] = []
            for s in changed:
                entries.append((imgs2.get(s["n"], s["path"]), s["title"], s["subtitle"], s["bottom"]))
                where.append((code, str(o["choice"]), s))
                dd["applies"] = sorted(set(dd["applies"]) | {s["n"]})
    t = step(f"options: render {len(entries)} slides' images", t)
    if entries:
        pngs, meas, _ = build(cfg, entries, "options", approved=False)
        for k, (code, choice, s) in enumerate(where, 2):
            png = f"opt_{code}_{choice}_{s['n']:02d}.png"
            shutil.copyfile(pngs[k - 1], page_dir / png)
            decisions[code]["options"][choice].append(
                {"n": s["n"], "png": png, "lines": s["lines"], "fact": fact(meas.get(k), s["title"], size)})
    # the dates deck: every slide, dates moved to the bottom line
    di = next((it for it in spec["items"] if it.get("dates_preview") and str(it["code"]) not in applied), None)
    if di:
        t = time.perf_counter()
        cfg3 = move_dates(cfg)
        new = slides(cfg3, text, pages, data, cent)
        imgs3 = render_images(new, cfg3, text, pages, data, cent, OUT / "src" / "dates")
        t = step("dates: render contents and tables", t)
        pngs, meas, _ = build(cfg3, [(imgs3.get(s["n"], s["path"]), s["title"], s["subtitle"], s["bottom"]) for s in new],
                              "dates", approved=False)
        moved = [s for s in new if s["bottom"]]
        show = [next(s for s in moved if s["bottom"].startswith("N1")), next(s for s in moved if not s["bottom"].startswith("N1"))]
        code = str(di["code"])
        rows = []
        for s in moved:
            m = meas.get(s["n"])
            f = fact(m, s["title"], size)
            bad = m["bottom_lines"] != 1 or m["bottom_overlaps_figure"] or m["bottom_off_slide"]
            rows.append({"n": s["n"], "cells": [str(s["n"]), s["title"], s["bottom"], str(m["lines"]), str(m["bottom_lines"])
                                                + ("; overlaps the figure" if m["bottom_overlaps_figure"] else "")
                                                + ("; off the slide" if m["bottom_off_slide"] else ""),
                                                m.get("moved", "")], "warn": f["warn"] or bad})
        decisions[code] = {"applies": [s["n"] for s in show], "options": {str(di["options"][0]["choice"]): [
            {"n": s["n"], "png": f"opt_{code}_{s['n']:02d}.png", "lines": s["lines"],
             "fact": dict(fact(meas.get(s["n"]), s["title"], size),
                          text=fact(meas.get(s["n"]), s["title"], size)["text"]
                          + f"; date line: {meas[s['n']]['bottom_lines']} line at {bsize} pt"
                          + ("; it touches the figure" if meas[s["n"]]["bottom_overlaps_figure"] else ""))} for s in show]},
            "table": {"head": ["Slide", "Title", "Date line", f"Title lines ({size} pt)", f"Date lines ({bsize} pt)",
                               "Figure moved to clear the text"], "rows": rows}}
        for s in show:
            shutil.copyfile(pngs[s["n"] - 1], page_dir / f"opt_{code}_{s['n']:02d}.png")
        nbad = sum(r["warn"] for r in rows)
        step(f"dates: {len(moved)} titles moved, {nbad} with a title over 2 lines or a bad bottom line", t)
    out = {"deck": deck, "decisions": decisions}
    (page_dir / "page_data.json").write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8")
    warn = [f"slide {x['n']}: {x['fact']['text']}" for x in deck if x["fact"] and x["fact"]["warn"]]
    print(f"{len(deck)} deck slides, {len(entries)} option slides for {len(decisions)} open decisions -> "
          f"{page_dir / 'page_data.json'}; {len(warn)} deck titles with 3+ lines or into the figure")
    step("total", t0)


if __name__ == "__main__":
    main(sys.argv[1])
