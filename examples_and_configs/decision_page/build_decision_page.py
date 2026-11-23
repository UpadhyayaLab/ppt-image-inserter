"""Build a decision page (decision_page/README.md): page_template.html filled with one JSON block.

Inputs:
- the decision spec: one item per decision (options with `replace` or `request`), plus info items (no options);
- the applied list (answers already applied to the deck by apply_answers.py): shown as Resolved;
- page_data.json from the deck's previews builder: every deck slide with its text and measured title, and per open
  decision the slides each option changes.
Each PNG is embedded once. The build fails on any check below (all reported together).

Usage:  python build_decision_page.py <spec.yaml> <out.html> [--test-dir <data folder>/tests]
"""
import base64
import html
import io
import json
import re
import sys
from pathlib import Path

import yaml
from PIL import Image

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "page_template.html"
SCOPE = "Applies to every slide like this one."


def load(spec_path):
    spec_path = Path(spec_path)
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    ap = spec_path.parent / spec["applied"]
    applied = (yaml.safe_load(ap.read_text(encoding="utf-8")) or []) if ap.exists() else []
    pd = Path(spec["page_data"])
    return spec, applied, json.loads(pd.read_text(encoding="utf-8")), pd.parent


def page_data(spec, applied, pd, png_dir):
    """The JSON the page draws from, and the list of failed checks."""
    errors, warnings, images = [], [], {}

    def img(png):
        if png not in images:
            im = Image.open(png_dir / png).convert("RGB")
            w = spec.get("png_width", 1100)
            if im.width > w:
                im = im.resize((w, round(im.height * w / im.width)), Image.LANCZOS)
            buf = io.BytesIO()   # 256 colors, as the deck's own review page stores figure slides (size limit 16 MB)
            im.quantize(256, method=Image.Quantize.FASTOCTREE).save(buf, "PNG", optimize=True)
            images[png] = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
        return png

    deck = [dict(s, img=img(s["png"])) for s in pd["deck"]]
    num = {s["id"]: s["n"] for s in deck}
    by_n = {s["n"]: s for s in deck}
    done = {str(a["code"]): a for a in applied if "code" in a}

    def carrying(text):   # deck slides showing this line, as a whole line (a contents row without its slide range)
        return [s["n"] for s in deck  # or a title's prefix or label
                if any(t == text or t.startswith(text + ": ") or t.endswith(": " + text)
                       for t in (re.sub(r"^\d+(?:–\d+)?\s{2,}", "", ln) for ln in s["lines"]))]

    items, codes = [], set()
    in_spec = {str(it["code"]) for it in spec["items"]}
    for a in applied:   # answers applied earlier whose items are no longer in the spec
        if "code" in a and str(a["code"]) not in in_spec:
            items.append({"code": str(a["code"]), "title": a["title"], "options": []})
    items += [dict(it, code=str(it["code"])) for it in spec["items"]]
    out = []
    for it in items:
        code = it["code"]
        if code in codes:
            errors.append(f"item {code} appears twice")
        codes.add(code)
        if "options" not in it:
            out.append({"code": code, "title": it["title"], "about": it.get("about")})
            continue
        res = done.get(code)
        p = (pd["decisions"].get(code) or {})
        if res and res.get("rule"):   # settled by a general rule of FF's, not by a line on a slide
            applies = []
        elif res and res["text"] == "Keep current":   # the lines its options would have changed
            applies = sorted({n for o in it.get("options", []) for r in o.get("replace", []) for n in carrying(r["old"])})
        elif res:
            applies = carrying(res["text"])
        elif p:
            applies = list(p["applies"])
        else:
            applies = []
        for stem in it.get("slides", []):
            if stem not in num:
                errors.append(f"item {code}: no deck slide {stem}")
            else:
                applies.append(num[stem])
        applies = sorted(set(applies))
        opts = []
        olds = {r["old"] for o in it["options"] for r in o.get("replace", [])}
        if it["options"] and it.get("keep", True):
            cur = it.get("current") or (olds.pop() if len(olds) == 1 else None)
            opts.append({"choice": "keep", "button": "Keep current", "text": cur, "label": it.get("keep_label") or (None if cur else "Keep current"),
                         "note": "", "nr": "", "slides": []})
        for o in it["options"]:
            ch = str(o["choice"])
            if o.get("note") and not o.get("note_source"):
                errors.append(f"item {code} option {ch}: note without note_source")
            nr = o.get("not_rendered") or (f"Not rendered: needs a re-render by {o['request']}." if o.get("request") else "")
            sl = [{"n": s["n"], "img": img(s["png"]), "fact": s["fact"]} for s in (p.get("options") or {}).get(ch, [])]
            long3 = [s["n"] for s in sl if s["fact"] and s["fact"]["lines"] >= 3]
            if long3 and not res:   # FF 2026-11-20: "titles sould never span 3 lines"
                warnings.append(f"item {code} option {ch}: title takes 3+ lines on slide(s) {long3}")
            if not res and not sl and not nr:
                errors.append(f"item {code} option {ch}: no preview (run the previews builder) and no request")
            opts.append({"choice": ch, "button": o.get("button", ch), "text": o.get("text"), "label": o.get("label"),
                         "note": o.get("note", ""), "source": o.get("wording_source", ""), "nr": nr, "slides": sl})
        out.append({"code": code, "title": it["title"], "about": it.get("about"), "applies": applies,
                    "now": [{"n": n, "img": by_n[n]["img"], "fact": by_n[n]["fact"]} for n in applies],
                    "table": p.get("table") if not res else None,
                    "options": opts, "resolved": ({k: res.get(k) for k in ("text", "how", "at", "note")} if res else None)})
        if res and not applies and not res.get("rule"):
            errors.append(f"item {code}: resolved text is on no deck slide: {res['text']!r}")
    data = {"heading": spec.get("heading", spec["title"]), "lead": spec.get("lead", ""),
            "all_heading": spec.get("all_heading", "All slides"), "all_lead": spec.get("all_lead", ""),
            "coll": spec.get("collection", "decisions"), "slcoll": spec.get("slide_collection", "slides"),
            "scope": spec.get("change_scope_note", SCOPE), "deck": deck, "items": out, "images": images}
    shown = json.dumps({k: v for k, v in data.items() if k != "images"}, ensure_ascii=False)
    # FF 2026-11-20: "it cannot overlap with any data" (text_clear, run by the deck builder; the facts carry what is left)
    for s_ in deck:
        if s_.get("fact") and "overlaps the data" in s_["fact"]["text"]:
            warnings.append(f"deck slide {s_['n']}: {s_['fact']['text']}")
    for it in out:
        for o in it.get("options") or []:
            for sl in o["slides"]:
                if sl["fact"] and "overlaps the data" in sl["fact"]["text"]:
                    warnings.append(f"item {it['code']} option {o['choice']} slide {sl['n']}: {sl['fact']['text']}")
    data["warnings"] = warnings
    for ch, what in (("·", "middle dot"), ("µ", "micro sign (use Greek mu U+03BC)")):
        if ch in shown:
            errors.append(f"{what} in the page text")
    files = re.findall(r"\b[\w-]+\.(?:m|py|csv|yaml|json|mat)\b", shown)
    if files:
        errors.append(f"file names in the page text: {sorted(set(files))}")
    return data, errors


MAX_EXPORT_AGE_MIN = 15


def db_guard(spec_path, spec, applied):
    """The page is built from the live db (2026-11-20: a stale export left 7 of FF's saved answers off v22): the export
    must be at most MAX_EXPORT_AGE_MIN old, and every saved answer in it must already be applied (apply_answers.py
    --write), so the deck sources and the Resolved list carry all of them."""
    import time
    sys.path.insert(0, str(HERE))
    import apply_answers
    exp = Path(spec["db_export"])
    files = list(exp.glob("*/*.json"))
    if not files and spec.get("first_build"):   # a page not yet published has no db to export (2026-11-22, TGT deck)
        return []
    if not files:
        return [f"no db export in {exp}: export the page's db (ArtifactData list, out_dir) first"]
    age = (time.time() - min(f.stat().st_mtime for f in files)) / 60
    errs = []
    if age > MAX_EXPORT_AGE_MIN:
        errs.append(f"db export is {age:.0f} min old (limit {MAX_EXPORT_AGE_MIN}): export the live db again first")
    pairs, records, _ = apply_answers.answers(spec, applied, exp)
    if pairs or records:
        left = sorted({str(r.get("code") or "slide " + r.get("slide", "")) for r in records})
        errs.append(f"saved answers not applied yet: {left}: run apply_answers.py <spec> <export> --write first")
    return errs


def build(spec_path, test_js=""):
    spec, applied, pd, png_dir = load(spec_path)
    data, errors = page_data(spec, applied, pd, png_dir)
    errors += db_guard(spec_path, spec, applied)
    if errors:
        raise SystemExit("build checks failed:\n  " + "\n  ".join(errors))
    js = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    page = (TEMPLATE.read_text(encoding="utf-8").replace("__TITLE__", html.escape(spec["title"]))
            .replace("__TEST__", f"<script>{test_js}</script>" if test_js else "").replace("__DATA__", js))
    return page, data


def main(argv):
    page, data = build(argv[0])
    Path(argv[1]).write_text(page, encoding="utf-8")
    mb = len(page.encode()) / 1e6
    n_open = sum(1 for it in data["items"] if "options" in it and not it["resolved"])
    print(f"{argv[1]}: {len(data['deck'])} deck slides, {n_open} open decisions before saved picks, "
          f"{len(data['images'])} images, {mb:.1f} MB")
    for w in data["warnings"]:
        print(f"[WARNING] {w}")
    if mb > 15:
        raise SystemExit("page over 15 MB (artifact limit 16 MB): lower png_width")
    if "--test-dir" in argv:
        import test_decision_page
        test_decision_page.run(argv[0], argv[argv.index("--test-dir") + 1])


if __name__ == "__main__":
    sys.path.insert(0, str(HERE))
    main(sys.argv[1:])
