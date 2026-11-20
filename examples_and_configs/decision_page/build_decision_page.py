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
    errors, images = [], {}

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

    def carrying(text):   # deck slides showing this line, as a whole line or a title's prefix or label
        return [s["n"] for s in deck if any(ln == text or ln.startswith(text + ": ") or ln.endswith(": " + text)
                                            for ln in s["lines"])]

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
    for ch, what in (("·", "middle dot"), ("µ", "micro sign (use Greek mu U+03BC)")):
        if ch in shown:
            errors.append(f"{what} in the page text")
    files = re.findall(r"\b[\w-]+\.(?:m|py|csv|yaml|json|mat)\b", shown)
    if files:
        errors.append(f"file names in the page text: {sorted(set(files))}")
    return data, errors


def build(spec_path, test_js=""):
    spec, applied, pd, png_dir = load(spec_path)
    data, errors = page_data(spec, applied, pd, png_dir)
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
    if mb > 15:
        raise SystemExit("page over 15 MB (artifact limit 16 MB): lower png_width")
    if "--test-dir" in argv:
        import test_decision_page
        test_decision_page.run(argv[0], argv[argv.index("--test-dir") + 1])


if __name__ == "__main__":
    sys.path.insert(0, str(HERE))
    main(sys.argv[1:])
