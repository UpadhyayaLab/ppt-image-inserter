"""Slides for the decision page of the naive 0708 deck "TCell mask vs Cellpose 3" (decision_page/README.md).

Builds one preview deck of 3 slides with the same python-pptx helpers as the deck builder
(figure_creation/lib/python/review_deck.py), on black (FF 2026-11-24: "use black bg"): the two methods slides from
deck_text.yaml as bullets and sub-bullets (FF 2026-11-24: "we need bullets"; formatting as bullet_list in
CART/CART_GzmB_ppt/build_gzmb_vs_ctsb_deck.py), and the first cell slide, two sheets side by side (FF 2026-11-24:
"2 cells per slide I guess"): 1.5 kPa 27hr, cells 41 and 52, from the 160-cell no-text render, with the scale-bar
label top right and the method label of each sheet row at the left. An option with `sub_under`/`sub` in the spec gets its own copy of the slides it applies to, with those
lines as sub-sub-bullets under that line. Titles 32 pt (repo CLAUDE.md). Runs lint_presentation and the
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
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
sys.path.insert(0, r"C:\Users\Upadhyaya_Lab\Code\figure_creation\lib\python")
from review_deck import SLIDE_H, SLIDE_W, Deck, add_text, fit_in_box  # noqa: E402
from ppt_image_inserter import lint_presentation  # noqa: E402
from ppt_image_inserter.approved_text import deck_strings, load_approved, unapproved  # noqa: E402

SHEETS = Path(r"L:\FF\fixed_cell\CTL_nucleus\tifsFixedNaiveCTLs_nucleus\cellpose3_seg\qc\cmp160_d60_d150"
              r"\montage_mip_slice_notext")
FIRST_SLIDE = ("f-3p0-101010-27hrFix-0708", (41, 52), "1.5 kPa 27hr")   # condition folder, first two cells of
# cell_lists/per_cond20_s0.csv, 0708 config label
EXPORT = REPO / "examples_and_configs" / "export_deck_slides.py"
BASE_PYTHON = Path(r"C:\Users\Upadhyaya_Lab\miniforge3\python.exe")   # export_deck_slides.py needs pywin32
TITLE_PT, LINE_PT, ROW_PT, SCALE_PT = 32, 24, 16, 18
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def black_slide(d):
    s = d.prs.slides.add_slide(d.blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = RGBColor(0, 0, 0)
    return s


def title(s, text):
    add_text(s, 0.4, 0.15, SLIDE_W - 0.8, 0.6, text, TITLE_PT, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def flatten(lines, level=0, sub_under=None, sub=()):
    """deck_text.yaml lines as (text, level): a string is a bullet, {text, sub} nests its sub lines one level down.
    The lines `sub` go one level under the line equal to `sub_under` (option previews)."""
    items = []
    for ln in lines:
        t = ln if isinstance(ln, str) else ln["text"]
        items.append((t, level))
        if t == sub_under:
            items += [(x, level + 1) for x in sub]
        if not isinstance(ln, str):
            items += flatten(ln.get("sub", []), level + 1, sub_under, sub)
    return items


def bullets(s, items, size=LINE_PT):
    """(text, level) items as real bullets: level 0 "•", deeper levels "–". Hanging indent, spacing and sizes as
    bullet_list in build_gzmb_vs_ctsb_deck.py."""
    tf = s.shapes.add_textbox(Inches(0.9), Inches(1.2), Inches(SLIDE_W - 1.8), Inches(SLIDE_H - 1.5)).text_frame
    tf.word_wrap = True
    for i, (t, level) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        nxt = items[i + 1][1] if i + 1 < len(items) else 0
        p.space_after = Pt(16 if nxt == 0 else 6)   # gap after each group, not between a line and its sub lines
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
        bullets(s, flatten(m["lines"]))
    cond, cells, label = FIRST_SLIDE
    s = black_slide(d)
    title(s, text["cell_title"].format(condition=label))
    add_text(s, SLIDE_W - 3.4, 0.25, 3.1, 0.45, text["scale_label"], SCALE_PT, color=WHITE, align=PP_ALIGN.RIGHT)
    # method labels in one column at the left (FF 2026-11-24), the two sheets side by side to their right; no cell
    # labels (FF 2026-11-24: "no need to label cells"). Sheet rows are equal fifths of the sheet height.
    col, gap = 2.0, 0.8   # label column; gap between the two cells (with a thin divider line in it)
    half = (SLIDE_W - 0.3 - col - gap) / 2
    pics = [fit_in_box(s, str(SHEETS / f"h3k27ac_0708_naive__{cond}__cell{cell}.png"), 0.1 + col + k * (half + gap),
                       0.9, half, SLIDE_H - 1.0) for k, cell in enumerate(cells)]
    top, h = pics[0].top / 914400, pics[0].height / 914400
    mid = 0.1 + col + half + gap / 2   # divider: the two cells read as two, not as four groups of three
    ln = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(mid), Inches(top), Inches(mid), Inches(top + h))
    ln.line.color.rgb = RGBColor(0x80, 0x80, 0x80)
    ln.line.width = Pt(1.5)
    rows = text["row_labels"]
    for i, lab in enumerate(rows):
        tb = add_text(s, 0.2, top + h * (i + 0.5) / len(rows) - 0.4, col - 0.2, 0.8, lab, ROW_PT, bold=True,
                      color=WHITE, align=PP_ALIGN.RIGHT)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    n_deck = len(d.prs.slides)
    variants = {}   # (code, choice) -> [slide number in preview.pptx] for options with sub_under
    for it in spec["items"]:
        for o in it.get("options", []):
            if not o.get("sub_under"):
                continue
            for n in it["applies"]:
                m = text["methods"][n - 1]
                s = black_slide(d)
                title(s, m["title"])
                items = flatten(m["lines"], sub_under=o["sub_under"], sub=o["sub"])
                if len(items) != len(flatten(m["lines"])) + len(o["sub"]):
                    raise SystemExit(f"item {it['code']} option {o['choice']}: line {o['sub_under']!r} not on slide {n}")
                bullets(s, items)
                variants.setdefault((str(it["code"]), str(o["choice"])), []).append(len(d.prs.slides))
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

    def slide_lines(slide):
        return [p.text.strip() for sh in slide.shapes if sh.has_text_frame
                for p in sh.text_frame.paragraphs if p.text.strip()]

    all_slides = list(d.prs.slides)
    slides = []
    for n, slide in enumerate(all_slides[:n_deck], 1):
        name = f"deck_{n:02d}.png"
        shutil.copyfile(pngs[n - 1], page_dir / name)
        slides.append({"n": n, "id": f"slide_{n:02d}", "png": name, "lines": slide_lines(slide),
                       "unlisted": "", "fact": None})
    by_n = {s["n"]: s for s in slides}
    decisions = {}
    for it in spec["items"]:
        if "options" not in it:
            continue
        dd = decisions[str(it["code"])] = {"applies": list(it.get("applies", [])), "options": {}}
        for o in it["options"]:   # an option that is not a request or not_rendered shows the slides it applies to
            if o.get("request") or o.get("not_rendered"):
                continue
            key = (str(it["code"]), str(o["choice"]))
            if key in variants:   # its own copies, with the added lines
                prev = []
                for n, k in zip(it["applies"], variants[key]):
                    name = f"opt_{key[0]}_{key[1]}_{n:02d}.png"
                    shutil.copyfile(pngs[k - 1], page_dir / name)
                    prev.append({"n": n, "png": name, "lines": slide_lines(all_slides[k - 1]), "fact": None})
                dd["options"][key[1]] = prev
                continue
            dd["options"][str(o["choice"])] = [{"n": n, "png": by_n[n]["png"], "lines": by_n[n]["lines"],
                                                "fact": None} for n in it.get("applies", [])]
    (page_dir / "page_data.json").write_text(json.dumps({"deck": slides, "decisions": decisions}, indent=1,
                                                        ensure_ascii=False), encoding="utf-8")
    print(f"{len(slides)} deck slides, {len(all_slides) - n_deck} option slides -> {page_dir / 'page_data.json'} "
          f"({time.perf_counter() - t0:.1f} s)")


if __name__ == "__main__":
    main(sys.argv[1])
