"""Slides for the decision page of the naive 0708 deck "TCell mask vs Cellpose 3" (decision_page/README.md).

Builds one preview deck of 3 slides with the same python-pptx helpers as the deck builder
(figure_creation/lib/python/review_deck.py), on black (FF 2026-11-24: "use black bg"): the two methods slides from
deck_text.yaml as bullets and sub-bullets (FF 2026-11-24: "we need bullets"; formatting as bullet_list in
CART/CART_GzmB_ppt/build_gzmb_vs_ctsb_deck.py), and the first cell slide, two sheets side by side (FF 2026-11-24:
"2 cells per slide I guess"): 1.5 kPa 27hr, cells 21 and 25. Titles 32 pt (repo CLAUDE.md). Runs lint_presentation and the
approved-text check and prints both counts, exports the slides through PowerPoint (export_deck_slides.py) and writes
<deck folder>/decision_page/page/page_data.json for decision_page/build_decision_page.py.

Usage (env PPT_editing):  python cellpose_vs_tcell_decision_previews.py <decision spec.yaml>
"""
import json
import shutil
import subprocess
import sys
import time
from pathlib import Path

import yaml
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
sys.path.insert(0, r"C:\Users\Upadhyaya_Lab\Code\figure_creation\lib\python")
from review_deck import SLIDE_H, SLIDE_W, Deck, add_text, fit_in_box  # noqa: E402
from ppt_image_inserter import lint_presentation  # noqa: E402
from ppt_image_inserter.approved_text import deck_strings, load_approved, unapproved  # noqa: E402

SHEETS = Path(r"L:\FF\fixed_cell\CTL_nucleus\tifsFixedNaiveCTLs_nucleus\cellpose3_seg\qc\cmp32_d60_d150"
              r"\montage_mip_slice")
FIRST_SLIDE = ("f-3p0-101010-27hrFix-0708", (21, 25), "1.5 kPa 27hr")   # condition folder, cells (cell list
# order), 0708 config label
EXPORT = REPO / "examples_and_configs" / "export_deck_slides.py"
BASE_PYTHON = Path(r"C:\Users\Upadhyaya_Lab\miniforge3\python.exe")   # export_deck_slides.py needs pywin32
TITLE_PT, LINE_PT, LABEL_PT = 32, 24, 24
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def black_slide(d):
    s = d.prs.slides.add_slide(d.blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
    return s


def title(s, text):
    add_text(s, 0.4, 0.15, SLIDE_W - 0.8, 0.6, text, TITLE_PT, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def bullets(s, lines, size=LINE_PT):
    """deck_text.yaml lines as real bullets: a string is a level-0 bullet, {text, sub} adds level-1 dash
    sub-bullets. Hanging indent, spacing and sizes as bullet_list in build_gzmb_vs_ctsb_deck.py."""
    items = []
    for ln in lines:
        items.append((ln, 0) if isinstance(ln, str) else (ln["text"], 0))
        if not isinstance(ln, str):
            items += [(t, 1) for t in ln.get("sub", [])]
    tf = s.shapes.add_textbox(Inches(0.9), Inches(1.2), Inches(SLIDE_W - 1.8), Inches(SLIDE_H - 1.5)).text_frame
    tf.word_wrap = True
    for i, (t, level) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(16 if level == 0 else 8)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.4 + 0.5 * level)))
        pPr.set("indent", str(-Inches(0.4)))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•" if level == 0 else "–"}))
        r = p.add_run()
        r.text = t
        r.font.size = Pt(size - 2 * level)
        r.font.color.rgb = WHITE


def main(spec_path):
    sys.stdout.reconfigure(encoding="utf-8")   # slide lines carry μ and en dashes
    t0 = time.perf_counter()
    spec_path = Path(spec_path)
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    text = yaml.safe_load((spec_path.parent / spec["deck"]["text"]).read_text(encoding="utf-8"))
    page_dir = Path(spec["page_data"]).parent
    out = page_dir.parent
    shutil.rmtree(page_dir, ignore_errors=True)   # this script's own output from its last run
    page_dir.mkdir(parents=True)

    d = Deck()
    for m in text["methods"]:
        s = black_slide(d)
        title(s, m["title"])
        bullets(s, m["lines"])
    cond, cells, label = FIRST_SLIDE
    s = black_slide(d)
    title(s, text["cell_title"].format(condition=label))
    half = (SLIDE_W - 0.6) / 2
    for k, cell in enumerate(cells):   # two sheets side by side, each with its cell number above it
        left = 0.3 + k * half
        add_text(s, left, 0.8, half, 0.45, text["cell_label"].format(cell=cell), LABEL_PT, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
        fit_in_box(s, str(SHEETS / f"h3k27ac_0708_naive__{cond}__cell{cell}.png"), left + 0.1, 1.3, half - 0.2,
                   SLIDE_H - 1.4)
    deck = out / "preview.pptx"
    d.save(str(deck))

    lint = lint_presentation(d.prs)
    unap = unapproved(deck_strings(d.prs), load_approved(spec_path.parent / spec["deck"]["approved"]))
    print(f"lint_presentation: {len(lint)} problems; unapproved lines: {len(unap)}")
    for p in lint + unap:
        print("  " + p)

    src = out / "src"
    shutil.rmtree(src, ignore_errors=True)
    subprocess.run([str(BASE_PYTHON), str(EXPORT), str(deck), str(src), "1600"], check=True)
    pngs = sorted(src.glob("slide_*.png"))
    if len(pngs) != len(d.prs.slides):
        raise SystemExit(f"exported {len(pngs)} slides, expected {len(d.prs.slides)}")

    slides = []
    for n, slide in enumerate(d.prs.slides, 1):
        name = f"deck_{n:02d}.png"
        shutil.copyfile(pngs[n - 1], page_dir / name)
        lines = [p.text.strip() for sh in slide.shapes if sh.has_text_frame
                 for p in sh.text_frame.paragraphs if p.text.strip()]
        slides.append({"n": n, "id": f"slide_{n:02d}", "png": name, "lines": lines,
                       "unlisted": "Text inside the sheets is not listed." if n == 3 else "", "fact": None})
    by_n = {s["n"]: s for s in slides}
    decisions = {}
    for it in spec["items"]:
        if "options" not in it:
            continue
        dd = decisions[str(it["code"])] = {"applies": list(it.get("applies", [])), "options": {}}
        for o in it["options"]:   # an option that is not a request or not_rendered shows the slides it applies to
            if o.get("request") or o.get("not_rendered"):
                continue
            dd["options"][str(o["choice"])] = [{"n": n, "png": by_n[n]["png"], "lines": by_n[n]["lines"],
                                                "fact": None} for n in it.get("applies", [])]
    (page_dir / "page_data.json").write_text(json.dumps({"deck": slides, "decisions": decisions}, indent=1,
                                                        ensure_ascii=False), encoding="utf-8")
    print(f"{len(slides)} slides -> {page_dir / 'page_data.json'} ({time.perf_counter() - t0:.1f} s)")


if __name__ == "__main__":
    main(sys.argv[1])
