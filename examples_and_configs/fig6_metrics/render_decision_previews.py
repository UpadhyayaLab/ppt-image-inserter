"""Full-slide previews of every option on the Fig 6 deck decision page (FF 2026-11-19: "I need to see everything in
context. you need to show full slide for each item!").

Every preview is made by the deck's own pipeline:
- contents slide: render_fig6_table_slides.render_contents with CONTENTS / CONTENTS_NOTES set per option;
- table slide 1 (H3K27me3, Jurkat, 1 of 2, the one with near-centrosome columns): slide_html with KEY / CENT_KEY set
  per option, screenshotted with the same Edge call as render_fig6_table_slides.main;
- everything is placed on the deck's template by batch_insert_images.py (preview config, no approved_text) and
  exported by export_deck_slides.py; each exported slide is then copied to <out>/named/<id>.png;
- the text on each preview slide goes to <out>/named/slide_text.json for the decision page's per-slide change box:
  the slide's own text boxes (read from the built deck) plus the text drawn into its image where the image is
  rendered from HTML (contents, table), with what is not listed (table numbers, labels inside figures);
- each title's line count and whether it runs into the figure, measured in PowerPoint by measure_title_lines.py, go
  to <out>/named/slide_facts.json (titles use the deck's own title settings: title_size 36 since 2026-11-19).

Option strings are the code-checked ones on the decision page (Histone mark data, 2026-11-19). Assumption stated on
the page: the new near-centrosome slides sit at 11-13 (violins) and 35-37 (profiles), as the decision items say, so
the other contents ranges move up by 3 and 6.

Usage:  python render_decision_previews.py [--from render|deck|export|name|full]
(full: only the whole-deck preview for the All slides view; it runs last in a full run.)
"""
import html
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
sys.path.insert(0, str(HERE))
import render_fig6_table_slides as tables  # noqa: E402

FIGS = "Q:/FF/Nucleus_Data/3D_Nucleus/Fixed/fig6_metrics_page/figs_20261118"
DECK_CFG = EXAMPLES / "configs/fixed/fig6_metrics/config_fig6_metrics_by_metric_20261118.yaml"
PREVIEW_CFG = EXAMPLES / "configs/fixed/fig6_metrics/config_fig6_decision_previews_20261120.yaml"
OUT = Path("K:/FF/PPT/PPT_autogeneration/Fig6_metrics_across_experiments/decision_previews")
DECK = OUT / "Fig 6 decision previews.pptx"

# ---- option strings, verbatim from the decision page --------------------------------------------------------------
ROW1 = {"1A": "Relative level around the nadir of the centrosome-proximal region",
        "1B": "Relative level within 25% of the mean surface distance from the nadir of the centrosome-proximal region"}
ROW2 = {"2A": "Relative level vs distance from the nadir of the centrosome-proximal region",
        "2B": "Relative level vs distance from the nadir of the centrosome-proximal region, along the surface"}
NOTE3A = "Level: mean in the 0.25 μm layer inside the nuclear surface, weighted by surface area, ÷ the same over the whole nuclear surface"
NOTE3B = "Near nadir: surface connected to the nadir and deeper than half its depth, within 25% of the mean surface distance from it"
NOTE3C = "Around nadir (near centrosome): surface within 25% of the mean surface distance from the nadir of the centrosome-proximal region"
TITLE4A = "Regions around the nadir of the centrosome-proximal region and the centrosome projection, CD3 DMSO cell {cell}"
KEY6_NOW = "Near centrosome: disk around the deepest point of the centrosome-proximal patch; no test; blank = not analyzed"
KEY6 = "Near centrosome: surface within 25% of the mean surface distance from the nadir of the centrosome-proximal region ÷ whole surface; no test; blank = centrosome not imaged"
KEY7_NOW = "Mean in the 0.25 μm layer near the nadir ÷ the same layer, whole nucleus (1 = no enrichment); mean ± SEM over cells"
KEY7 = "Mean in the 0.25 μm layer near the nadir, weighted by surface area, ÷ the same over the whole nuclear surface (1 = no enrichment); mean ± SEM over cells"
TITLE8A = "Relative {ch} intensity near the deepest invagination: {label}"
ROW1S = "Relative level near the centrosome-proximal nadir"              # item 11a: FF's wording and pick, 2026-11-19
ROW2S = "Relative level vs distance from the centrosome-proximal nadir"  # item 11b shorter (Histone mark data)
DATE_NOW = "(N1 Jun 12, N2 Jun 13, 2026)"                                # item 11c current
DATE_S = "(Jun 12 and 13, 2026)"                                         # item 11c shorter
CENT_KEYS = ["noco_oct", "cilio_oct", "cilio_jun"]
DECK_CONTENTS = list(tables.CONTENTS)
DB_EXPORT = OUT / "db_export"   # the decision page's db, exported with ArtifactData list (out_dir) before each run

# ---- full deck preview for the page's "All slides" view (Histone mark data agent presentation, 2026-11-19) -----------
# Every deck slide in deck order, numbered as the deck will be: the deck config's 72 slides with the 3 near-centrosome
# violins inserted at 11-13 (FF's 11a title) and the 3 near-centrosome profiles at 35-37 (FF's pick 2A), each block in
# the deck's own experiment order. Slide 2 is rendered from the Contents text FF saved on the page and the table keys
# with his decision 6 change, both read from the db export (latest saved state is the base for every slide).
# Deck slides that carry a saved slide change keep the id the change was saved under, so the page shows that change on
# them (db slides collection): image name -> slide id. Slide 2 (contents.png) is always contents_base.
SAVED_IDS = {"Ccent_noco_oct_0.png": "cent_violin_1A_noco_oct"}   # FF's note on the nocodazole near-centrosome violin
FULL_CFG = EXAMPLES / "configs/fixed/fig6_metrics/config_fig6_deck_preview_20261120.yaml"
FULL_DECK = OUT / "Fig 6 deck preview, all slides.pptx"
RANGES = ["3–4", "5–10", "11–13", "14–34", "35–37", "38–58", "59–79"]
# Kept after FF's decision 6 text until he says whether it replaces the whole key line (Histone mark data agent
# presentation, 2026-11-19: render "his text plus '; no test; blank = not analyzed' kept").
KEY6_KEPT_TAIL = "; no test; blank = not analyzed"


def ff_saved():
    """FF's latest saved state from the db export: (contents rows, contents notes, decision 6 key line)."""
    doc = json.loads((DB_EXPORT / "slides" / "contents_base.json").read_text(encoding="utf-8"))
    if doc.get("status") != "change":
        raise SystemExit("db export: no saved change on the contents slide")
    lines = [x for x in doc["text"].split("\n") if x.strip()]
    if lines[0].strip() != "Contents":
        raise SystemExit("db export: saved contents text does not start with 'Contents'")
    rows, notes = [], []
    for x in lines[1:]:
        m = re.match(r"^(\d+–\d+)\s{2,}(.+)$", x)
        if m:
            rows.append((m.group(1), m.group(2).strip(), ""))
        else:
            notes.append(x.strip())
    if [r[0] for r in rows] != RANGES:
        raise SystemExit(f"db export: saved contents rows are {[r[0] for r in rows]}, expected {RANGES}")
    d6 = json.loads((DB_EXPORT / "decisions" / "6.json").read_text(encoding="utf-8"))
    if d6.get("status") != "change" or not d6.get("note", "").strip():
        raise SystemExit("db export: decision 6 has no saved change")
    return rows, notes, d6["note"].strip() + KEY6_KEPT_TAIL


FF_CONTENTS, FF_NOTES, FF_KEY6 = ff_saved()
# Options shown on FF's saved slide 2, each applied alone to his text. code: (kind, key, new string)
#   row: the row whose text is key; note: the note that is key; all: replace key in every row;
#   range: the row with that range; note_prefix: the note starting with key.
# item 12: Histone mark data's code check of FF's slide 2 (relayed 2026-11-19); items 1-3, 11b: the page's options.
FIX12 = {"12a": ("row", "Nadir: surface point farthest from the hull boundary",
                 "Nadir: surface point farthest from the hull boundary (centrosome-proximal: within that region)"),
         "12b": ("note", "Level: mean in the 0.25 μm layer inside the nuclear surface ÷ mean in that layer over the whole nucleus",
                 "Level: mean in the 0.25 μm layer inside the nuclear surface ÷ the area-weighted mean over the whole surface"),
         "12c": ("all", "full nucleus", "whole nucleus"),
         "12d": ("row", "Relative level vs distance from the nadir of the centrosome-proximal region",
                 "Relative level vs distance from the centrosome-proximal nadir, along the surface")}
FF_VARIANTS = dict(FIX12, **{
    "1A": ("range", "11–13", ROW1["1A"]), "1B": ("range", "11–13", ROW1["1B"]),
    "2B": ("range", "35–37", ROW2["2B"]), "11b": ("range", "35–37", ROW2S),
    "3a": ("note_prefix", "Level:", NOTE3A), "3b": ("note_prefix", "Near nadir (whole nucleus):", NOTE3B),
    "3c": ("note_prefix", "Near the centrosome-proximal nadir:", NOTE3C)})


def ff_variant(code):
    """FF's saved contents rows and notes with one option applied; fails unless it hits exactly one place."""
    kind, key, new = FF_VARIANTS[code]
    rows, notes = [list(r) for r in FF_CONTENTS], list(FF_NOTES)
    if kind in ("row", "range"):
        hit = [r for r in rows if r[1 if kind == "row" else 0] == key]
        if len(hit) != 1:
            raise SystemExit(f"{code}: row {key!r} not found once in FF's contents")
        hit[0][1] = new
    elif kind in ("note", "note_prefix"):
        hit = [i for i, n in enumerate(notes) if (n == key if kind == "note" else n.startswith(key))]
        if len(hit) != 1:
            raise SystemExit(f"{code}: note {key!r} not found once in FF's contents")
        notes[hit[0]] = new
    else:
        if not any(key in r[1] for r in rows):
            raise SystemExit(f"{code}: {key!r} in no row of FF's contents")
        for r in rows:
            r[1] = r[1].replace(key, new)
    return [tuple(r) for r in rows], notes


def edge_shot(html_path, png_path, size):
    subprocess.run([tables.EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--force-device-scale-factor=2", f"--window-size={size}", f"--screenshot={png_path}",
                    html_path.resolve().as_uri()], check=True, capture_output=True, timeout=120)
    if not png_path.exists():
        raise RuntimeError(f"no screenshot written: {png_path}")


def table_pages():
    data, _ = tables.collect(FIGS)
    cent = tables.load_cent(FIGS)
    return data, cent, tables.pages(data, cent)


def render():
    """Images rendered from HTML, all on FF's latest saved state -> <OUT>/src/<id>/(contents|table).png|html:
    contents_<code> for every FF_VARIANTS option; the deck's 6 table slides with FF's decision 6 key (deck_tables);
    table6_p<n> (item 6 option) on the pages that carry the near-centrosome key; table7_p<n> (item 7 option) on all 6."""
    saved = (tables.CONTENTS, tables.CONTENTS_NOTES, tables.KEY, tables.CENT_KEY)
    src = {}
    try:
        for code in FF_VARIANTS:
            cid = f"contents_{code}"
            d = OUT / "src" / cid
            d.mkdir(parents=True, exist_ok=True)
            tables.CONTENTS, tables.CONTENTS_NOTES = ff_variant(code)
            tables.render_contents(d)
            src[cid] = d / "contents.png"
        if KEY7_NOW not in saved[2]:
            raise SystemExit("current KEY text differs from the decision page's 'Now' line")
        if KEY6_NOW not in saved[3]:
            raise SystemExit("current CENT_KEY text differs from the decision page's 'Now' line")
        cent_ff = saved[3].replace(KEY6_NOW, html.escape(FF_KEY6))
        data, cent, pgs = table_pages()
        dt = OUT / "src" / "deck_tables"
        dt.mkdir(parents=True, exist_ok=True)
        for n, (title, slug, keys) in enumerate(pgs, 1):
            has_cent = any(k in cent for k in keys)
            versions = [(dt, f"table_{n}_{slug}", saved[2], cent_ff),
                        (OUT / "src" / f"table7_p{n}", "table", saved[2].replace(KEY7_NOW, html.escape(KEY7)), cent_ff)]
            if has_cent:
                versions.append((OUT / "src" / f"table6_p{n}", "table", saved[2],
                                 saved[3].replace(KEY6_NOW, html.escape(KEY6))))
            for d, stem, key, cent_key in versions:
                d.mkdir(parents=True, exist_ok=True)
                tables.KEY, tables.CENT_KEY = key, cent_key
                page, _, _ = tables.slide_html(title, keys, data, cent)
                hp, pp = d / f"{stem}.html", d / f"{stem}.png"
                hp.write_text(page, encoding="utf-8")
                edge_shot(hp, pp, "2600,1150")
                if d != dt:
                    src[d.name] = pp
    finally:
        tables.CONTENTS, tables.CONTENTS_NOTES, tables.KEY, tables.CENT_KEY = saved
    from PIL import Image
    for cid, png in src.items():   # the contents window is fixed (1100 x 600 at 2x): a cut-off bottom row fails
        if cid.startswith("contents_"):
            im = Image.open(png).convert("L")
            if min(im.crop((0, im.height - 6, im.width, im.height)).getdata()) < 200:
                raise SystemExit(f"{cid}: text reaches the bottom edge of the contents image (cut off)")
    return src


def entries(src):
    """(id, image path, title, subtitle) for every preview slide. Each is one option applied to a slide of the deck
    preview (its base on the page) or to another preview."""
    cfg = yaml.safe_load(open(DECK_CFG, encoding="utf-8"))
    by_path = {Path(e["path"]).name: e for e in cfg["images"]}
    label = {k: by_path[f"C_{k}_0.png"]["title"].split("Relative level near nadir: ", 1)[1] for k in CENT_KEYS}
    out = [(cid, png, "Contents", None) for cid, png in src.items() if cid.startswith("contents_")]
    for n, cell in ((49, 10), (50, 34)):
        e = by_path[f"noco_20261029_deck_slide{n}.png"]
        out.append((f"regions_4A_cell{cell}", e["path"], TITLE4A.format(cell=cell), e.get("subtitle")))
    table_titles = [e["title"] for e in cfg["images"] if Path(e["path"]).name.startswith("table_")]
    for cid, png in src.items():
        if cid.startswith("table"):
            out.append((cid, png, table_titles[int(cid.split("_p")[1]) - 1], None))
    for code, row in list(ROW1.items()) + [("11s", ROW1S)]:
        for k in CENT_KEYS:
            out.append((f"cent_violin_{code}_{k}", f"{FIGS}/Ccent_{k}_0.png", f"{row}: {label[k]}", None))
    for code, row in list(ROW2.items()) + [("11s", ROW2S), ("12d", FIX12["12d"][2])]:
        for k in CENT_KEYS:
            out.append((f"cent_profile_{code}_{k}", f"{FIGS}/Dcent_{k}_0.png", f"{row}: {label[k]}", None))
    e_now = yaml.safe_load(open(E_ENTRIES, encoding="utf-8"))[0]
    lab = e_now["title"].split(": ", 1)[1]
    ch = re.match(r"Relative (.+?) intensity near nadir: ", e_now["title"]).group(1)
    out.append(("image_deck_now", f"{FIGS}/{e_now['path']}", e_now["title"], None))
    out.append(("image_deck_8A", f"{FIGS}/{e_now['path']}", TITLE8A.format(ch=ch, label=lab), None))
    for letter, kind in (("C", "violin_whole"), ("D", "profile_whole"), ("F", "depth_whole")):
        e = by_path[f"{letter}_cilio_jun_0.png"]
        out.append((f"{kind}_cilio_jun", f"{FIGS}/{e['path']}", e["title"], e.get("subtitle")))
    jun = [x for x in out if x[0] in ("cent_violin_11s_cilio_jun", "cent_profile_2A_cilio_jun",
                                      "cent_profile_11s_cilio_jun", "violin_whole_cilio_jun",
                                      "profile_whole_cilio_jun", "depth_whole_cilio_jun")]
    for cid, path, title, sub in jun:   # item 11c: the other date on every slide of that experiment
        if DATE_NOW not in title:
            raise SystemExit(f"{cid}: title lacks {DATE_NOW!r}: {title}")
        out.append((f"{cid}_d2", path, title.replace(DATE_NOW, DATE_S), sub))
    taken = set(SAVED_IDS.values()) | {"contents_base"}   # ids that belong to deck slides with a saved change
    out = [((i + "_preview") if i in taken else i, p, t, sub) for i, p, t, sub in out]
    if len({x[0] for x in out}) != len(out):
        raise SystemExit("previews: two share an id")
    return out


E_ENTRIES = Path("C:/Users/UPADHY~1/AppData/Local/Temp/claude/015e1a9efd7d8135/7f1941ab-fefa-4b67-8656-3b95a15636e0/scratchpad/e_entries.yaml")


def write_config(ents):
    cfg = yaml.safe_load(open(DECK_CFG, encoding="utf-8"))
    imgs = []
    for _, path, title, sub in ents:
        e = {"path": str(path).replace("\\", "/"), "title": title}
        if sub:
            e["subtitle"] = sub
        imgs.append(e)
    new = {k: cfg[k] for k in ("presentation", "template_slide", "preserve_slides", "auto_position", "position",
                               "trim_whitespace", "title_fit", "title_size", "add_label", "remove_template_slide")
           if k in cfg}   # the deck's own title settings (title_size: 36 since 2026-11-19)
    new.update(output_path=str(DECK).replace("\\", "/"), backup_dir=str(OUT / "backups").replace("\\", "/"),
               base_dir=FIGS, images=imgs)
    head = ("# GENERATED by examples_and_configs/fig6_metrics/render_decision_previews.py: full-slide previews of the\n"
            "# Fig 6 decision page options. Preview only: no approved_text (the build prints a note).\n")
    PREVIEW_CFG.write_text(head + yaml.safe_dump(new, sort_keys=False, allow_unicode=True, width=200), encoding="utf-8")
    ids = [{"id": i, "slide": n, "title": t, "subtitle": s or ""} for n, (i, _, t, s) in enumerate(ents, 2)]
    (OUT / "preview_ids.json").write_text(json.dumps(ids, indent=1, ensure_ascii=False), encoding="utf-8")
    return ids


def deck(ids):
    """Build with batch_insert_images.py; any failure (its text check included) stops here."""
    r = subprocess.run([sys.executable, str(EXAMPLES / "batch_insert_images.py"), str(PREVIEW_CFG)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    (OUT / "build.log").write_text(r.stdout + r.stderr, encoding="utf-8")
    if r.returncode or f"{len(ids)}/{len(ids)} slides created" not in r.stdout:
        print(r.stdout[-4000:], r.stderr[-2000:])
        raise SystemExit("deck build failed")
    print(f"deck built: {len(ids)} preview slides")


def export(ids):
    d = OUT / "slides_png"
    for old in d.glob("slide_*.png"):
        old.unlink()
    r = subprocess.run([sys.executable, str(EXAMPLES / "export_deck_slides.py"), str(DECK), str(d)],
                       capture_output=True, text=True)
    print(r.stdout.strip())
    if r.returncode:
        raise SystemExit(r.stderr[-2000:])


STAT_CELL = re.compile(r"^[\d.\s±*<>%()+-]*$")   # numbers, ± SEM, stars: table values, not wording


def html_lines(path):
    """Text lines of a rendered HTML image, in reading order: one line per table row (cells joined by two spaces) or
    block. Cells that are only numbers or stars are dropped; returns (lines, dropped_any)."""
    s = Path(path).read_text(encoding="utf-8")
    s = re.sub(r"<(style|title|script)\b.*?</\1>", "", s, flags=re.S | re.I)
    lines, dropped = [], False
    for block in re.split(r"</(?:tr|div|p|h\d|caption|li|span)>|<br\s*/?>", s, flags=re.I):
        cells = []
        for c in re.split(r"</t[dh]>", block, flags=re.I):
            t = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", c))).strip()
            if not t:
                continue
            if STAT_CELL.match(t):
                dropped = True
                continue
            cells.append(t)
        if cells:
            lines.append("  ".join(cells))
    return lines, dropped


def slide_text(ids):
    """{id: {"lines": [...], "unlisted": "..."}} for every preview slide -> named/slide_text.json"""
    from pptx import Presentation
    sys.path.insert(0, str(EXAMPLES.parent))
    from ppt_image_inserter.approved_text import deck_strings
    by_slide = {}
    for where, line in deck_strings(Presentation(str(DECK))):
        by_slide.setdefault(int(where.split()[1]), []).append(line)
    out = {}
    for x in ids:
        lines = list(by_slide.get(x["slide"], []))
        htmls = sorted((OUT / "src" / x["id"]).glob("*.html"))
        if htmls:
            img, dropped = html_lines(htmls[0])
            lines += img
            unlisted = "The numbers in the table are not listed." if dropped else ""
        else:
            unlisted = "Labels inside the figure are not listed."
        if not lines:
            raise SystemExit(f"no text found for preview slide {x['id']}")
        out[x["id"]] = {"lines": lines, "unlisted": unlisted}
    (OUT / "named" / "slide_text.json").write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"slide text: {len(out)} slides -> {OUT / 'named' / 'slide_text.json'}")


def name(ids):
    d, nd = OUT / "slides_png", OUT / "named"
    nd.mkdir(exist_ok=True)
    pngs = sorted(d.glob("slide_*.png"), key=lambda p: int(p.stem.split("_")[1]))
    if len(pngs) != len(ids) + 1:
        raise SystemExit(f"{len(pngs)} exported slides for {len(ids)} previews + title")
    for x in ids:
        shutil.copyfile(pngs[x["slide"] - 1], nd / f"{x['id']}.png")
    print(f"named: {len(ids)} PNGs in {nd}")
    slide_text(ids)
    title_facts(ids)


def title_facts(ids):
    """Title lines as PowerPoint lays them out (measure_title_lines.py) -> named/slide_facts.json, one line shown
    under each slide on the decision page: "Title: 2 lines at 36 pt", flagged when 3+ lines or into the figure."""
    sys.path.insert(0, str(EXAMPLES))
    from measure_title_lines import measure
    size = yaml.safe_load(open(DECK_CFG, encoding="utf-8")).get("title_size")
    by_slide = {m["slide"]: m for m in measure(str(DECK))}
    facts, long3 = {}, []
    for x in ids:
        m = by_slide.get(x["slide"])
        if m is None:
            raise SystemExit(f"no title measured on preview slide {x['id']}")
        if m["title"] != x["title"].strip():
            raise SystemExit(f"slide {x['slide']}: measured title {m['title']!r} is not {x['title']!r}")
        n = m["lines"]
        text = f"Title: {n} line{'s' if n != 1 else ''}" + (f" at {size} pt" if size else "")
        if m["runs_into_figure"]:
            text += "; it runs into the figure"
        warn = n >= 3 or m["runs_into_figure"]
        facts[x["id"]] = {"text": text, "warn": warn, "lines": n, "runs_into_figure": m["runs_into_figure"]}
        if warn:
            long3.append(f"{x['id']}: {text}")
    (OUT / "named" / "slide_facts.json").write_text(json.dumps(facts, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"title lines measured on {len(facts)} slides; {len(long3)} with 3+ lines or into the figure:")
    for line in long3:
        print("  ", line)


def full_entries():
    """(id, image path, title, subtitle) for every slide of the deck preview from slide 2 on, in deck order."""
    d = OUT / "src" / "deck_contents"
    d.mkdir(parents=True, exist_ok=True)
    saved = (tables.CONTENTS, tables.CONTENTS_NOTES)
    try:
        tables.CONTENTS, tables.CONTENTS_NOTES = FF_CONTENTS, FF_NOTES
        tables.render_contents(d)
    finally:
        tables.CONTENTS, tables.CONTENTS_NOTES = saved
    from PIL import Image
    im = Image.open(d / "contents.png").convert("L")
    if min(im.crop((0, im.height - 6, im.width, im.height)).getdata()) < 200:
        raise SystemExit("FF's contents text reaches the bottom edge of the contents image (cut off)")
    cfg = yaml.safe_load(open(DECK_CFG, encoding="utf-8"))
    imgs = cfg["images"]
    if Path(imgs[0]["path"]).name != "contents.png":
        raise SystemExit("deck config: slide 2 is not the contents slide")
    names = [Path(e["path"]).name for e in imgs]
    order = [n[2:-6] for n in names if n.startswith("C_")]
    cent = [k for k in order if k in CENT_KEYS]
    if sorted(cent) != sorted(CENT_KEYS):
        raise SystemExit(f"near-centrosome experiments not all in the deck: {cent}")
    by_path = dict(zip(names, imgs))
    label = {k: by_path[f"C_{k}_0.png"]["title"].split("Relative level near nadir: ", 1)[1] for k in cent}
    out = [("contents_base", d / "contents.png", "Contents", None)]
    for n, e in zip(names[1:], imgs[1:]):
        if n.startswith("C_") and not any(x[0].startswith("cent_v") for x in out):
            out += [(f"cent_v_{k}", f"{FIGS}/Ccent_{k}_0.png", f"{ROW1S}: {label[k]}", None) for k in cent]
        if n.startswith("D_") and not any(x[0].startswith("cent_p") for x in out):
            out += [(f"cent_p_{k}", f"{FIGS}/Dcent_{k}_0.png", f"{ROW2['2A']}: {label[k]}", None) for k in cent]
        if n.startswith("table_"):   # re-rendered with FF's decision 6 key (render() -> src/deck_tables)
            t = OUT / "src" / "deck_tables" / n
            if not t.exists():
                raise SystemExit(f"{t} missing: run from the render step")
            out.append((n, t, e["title"], e.get("subtitle")))
            continue
        out.append((n, e["path"], e["title"], e.get("subtitle")))
    out = [(("contents_base" if i == 0 else SAVED_IDS.get(Path(str(p)).name, f"deck_{i + 2:02d}")), p, t, sub)
           for i, (_, p, t, sub) in enumerate(out)]
    if len({x[0] for x in out}) != len(out):
        raise SystemExit("deck preview: two slides share an id")
    for want, rng in ((11, "11–13"), (35, "35–37")):   # the contents rows must point at the inserted slides
        if not out[want - 2][1].endswith(f"cent_{cent[0]}_0.png"):
            raise SystemExit(f"slide {want} is not the first near-centrosome slide ({rng})")
    if len(out) + 1 != 79:
        raise SystemExit(f"deck preview has {len(out) + 1} slides, the contents says 79")
    return out


def full():
    """Build, export and measure the whole deck preview -> named/<id>.png, named/deck_slides.json (the All slides
    index), and the deck slides' entries merged into named/slide_text.json and named/slide_facts.json."""
    from pptx import Presentation
    sys.path.insert(0, str(EXAMPLES.parent))
    sys.path.insert(0, str(EXAMPLES))
    from ppt_image_inserter.approved_text import deck_strings
    from measure_title_lines import measure
    ents = full_entries()
    cfg = yaml.safe_load(open(DECK_CFG, encoding="utf-8"))
    cfg.pop("approved_text", None)   # preview: titles FF picked on the page are not on the approved list yet
    cfg.update(output_path=str(FULL_DECK).replace("\\", "/"), backup_dir=str(OUT / "backups").replace("\\", "/"),
               images=[dict({"path": str(p).replace("\\", "/"), "title": t}, **({"subtitle": sub} if sub else {}))
                       for _, p, t, sub in ents])
    head = ("# GENERATED by examples_and_configs/fig6_metrics/render_decision_previews.py --from full: the whole Fig 6\n"
            "# deck as it will be, for the decision page's All slides view. Preview only: no approved_text.\n")
    FULL_CFG.write_text(head + yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True, width=200), encoding="utf-8")
    t = time.perf_counter()
    r = subprocess.run([sys.executable, str(EXAMPLES / "batch_insert_images.py"), str(FULL_CFG)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    (OUT / "build_full.log").write_text(r.stdout + r.stderr, encoding="utf-8")
    if r.returncode or f"{len(ents)}/{len(ents)} slides created" not in r.stdout:
        print(r.stdout[-4000:], r.stderr[-2000:])
        raise SystemExit("deck preview build failed")
    print(f"[full deck] {len(ents) + 1} slides built in {time.perf_counter() - t:.1f} s")
    pd = OUT / "deck_png"
    pd.mkdir(exist_ok=True)
    for old in pd.glob("slide_*.png"):
        old.unlink()
    t = time.perf_counter()
    r = subprocess.run([sys.executable, str(EXAMPLES / "export_deck_slides.py"), str(FULL_DECK), str(pd)],
                       capture_output=True, text=True)
    if r.returncode:
        raise SystemExit(r.stderr[-2000:])
    pngs = sorted(pd.glob("slide_*.png"), key=lambda p: int(p.stem.split("_")[1]))
    if len(pngs) != len(ents) + 1:
        raise SystemExit(f"{len(pngs)} exported slides for a {len(ents) + 1}-slide deck")
    print(f"[full export] {time.perf_counter() - t:.1f} s")
    nd = OUT / "named"
    index = [{"id": "deck_01", "png": "deck_01.png", "slide": 1}]
    shutil.copyfile(pngs[0], nd / "deck_01.png")
    for n, (i, _, _, _) in enumerate(ents, 2):
        png = f"deck_{n:02d}.png"
        shutil.copyfile(pngs[n - 1], nd / png)
        index.append({"id": i, "png": png, "slide": n})
    # slide text: the slide's text boxes plus the text drawn into rendered images (contents, tables)
    text = json.loads((nd / "slide_text.json").read_text(encoding="utf-8"))
    by_slide = {}
    for where, line in deck_strings(Presentation(str(FULL_DECK))):
        by_slide.setdefault(int(where.split()[1]), []).append(line)
    for x in index:
        lines = by_slide.get(x["slide"], [])
        if not lines:
            raise SystemExit(f"deck preview slide {x['slide']}: no text found")
        html_src = None
        if x["slide"] == 2:   # contents: FF's saved text, drawn into the image
            html_src = OUT / "src" / "deck_contents" / "contents.html"
        if 5 <= x["slide"] <= 10:   # tables: their text is drawn into the image
            html_src = next(iter(sorted((OUT / "src" / "deck_tables").glob(f"table_{x['slide'] - 4}_*.html"))), None)
        if html_src:
            img, _ = html_lines(html_src)
            text[x["id"]] = {"lines": lines + img,
                             "unlisted": "The numbers in the table are not listed." if x["slide"] > 2 else ""}
        else:
            text[x["id"]] = {"lines": lines, "unlisted": "" if x["slide"] == 1 else "Labels inside the figure are not listed."}
    (nd / "slide_text.json").write_text(json.dumps(text, indent=1, ensure_ascii=False), encoding="utf-8")
    # title lines at 36 pt, measured in PowerPoint
    t = time.perf_counter()
    facts = json.loads((nd / "slide_facts.json").read_text(encoding="utf-8"))
    size = cfg.get("title_size")
    meas = {m["slide"]: m for m in measure(str(FULL_DECK))}
    titles = {n: tt for n, (_, _, tt, _) in enumerate(ents, 2)}
    long3 = []
    for x in index:
        n = x["slide"]
        if n == 1:
            facts[x["id"]] = {"text": "", "warn": False}   # title slide: not a 36 pt title
            continue
        m = meas.get(n)
        if m is None or m["title"] != titles[n].strip():
            raise SystemExit(f"deck preview slide {n}: measured title {m and m['title']!r} is not {titles[n]!r}")
        k = m["lines"]
        f = {"text": f"Title: {k} line{'s' if k != 1 else ''}" + (f" at {size} pt" if size else "")
                     + ("; it runs into the figure" if m["runs_into_figure"] else ""),
             "warn": k >= 3 or m["runs_into_figure"], "lines": k, "runs_into_figure": m["runs_into_figure"]}
        if x["id"] == "contents_base" and facts.get("contents_base", {}).get("text") not in (None, f["text"]):
            raise SystemExit(f"contents slide: {f['text']} in the deck preview, {facts['contents_base']['text']} in the previews")
        facts[x["id"]] = f
        if f["warn"]:
            long3.append(f"slide {n}: {f['text']}: {titles[n]}")
    (nd / "slide_facts.json").write_text(json.dumps(facts, indent=1, ensure_ascii=False), encoding="utf-8")
    (nd / "deck_slides.json").write_text(json.dumps(index, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"[full measure] {time.perf_counter() - t:.1f} s; {len(long3)} deck slides with 3+ lines or into the figure:")
    for line in long3:
        print("  ", line)


def main(start="render"):
    steps = ["render", "deck", "export", "name", "full"]
    todo = steps[steps.index(start):]
    OUT.mkdir(parents=True, exist_ok=True)
    t0 = time.perf_counter()
    if todo == ["full"]:
        full()
        print(f"total {time.perf_counter() - t0:.1f} s")
        return
    if "render" in todo:
        src = render()
        ids = write_config(entries(src))
    else:
        ids = json.loads((OUT / "preview_ids.json").read_text(encoding="utf-8"))
    for step, fn in (("deck", deck), ("export", export), ("name", name)):
        if step in todo:
            t = time.perf_counter()
            fn(ids)
            print(f"[{step}] {time.perf_counter() - t:.1f} s")
    if "full" in todo:
        full()
    print(f"total {time.perf_counter() - t0:.1f} s")


if __name__ == "__main__":
    a = sys.argv[1:]
    main(a[a.index("--from") + 1] if "--from" in a else "render")
