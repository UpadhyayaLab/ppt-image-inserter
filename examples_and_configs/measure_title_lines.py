"""Measure every slide's title as PowerPoint lays it out: how many lines the title takes and whether its letters run
into the figure below (FF 2026-11-19: at a fixed 36 pt "a 3-line title runs into the figure"). Measured from
PowerPoint itself (COM), not estimated from character counts.

Per slide: the title is the topmost text shape (slide-number boxes ignored); its first paragraph is the title (a
subtitle is a later paragraph of the same box). Lines: TextRange.Lines on that paragraph. Letters vs figure: the slide
is exported twice in memory, once with only the title box visible and once with only the pictures visible (the deck
is opened read-only and never saved); the lowest dark pixel row of the title is compared with the highest dark pixel
row of the pictures. gap_in < 0 means the title's letters reach into the figure.

Usage:  python measure_title_lines.py <deck.pptx> <out.json>
Output: [{"slide": n, "title": "...", "lines": n, "gap_in": x | null, "runs_into_figure": bool}, ...]; a slide with a
bottom_text line (shape "Bottom text") also gets bottom, bottom_lines, bottom_overlaps_figure (a figure pixel inside
the box around the line's letters, from a third in-memory export), bottom_off_slide
(slides without a text shape are skipped). The deck must be closed in PowerPoint.
"""
import json
import os
import sys
import tempfile

import win32com.client as w
from PIL import Image

PICTURE = 13      # msoPicture
WIDTH_PX = 2000   # export width; px per inch = WIDTH_PX / slide width in inches
DARK = 140        # gray level below which a pixel counts as ink
FIG_INK = 245     # in the figure, any pixel darker than this counts (light colors such as cyan included)


def ink_rows(png):
    """(first, last) row holding a dark pixel, or None."""
    im = Image.open(png).convert("L").point(lambda v: 255 if v < DARK else 0)
    box = im.getbbox()
    return (box[1], box[3] - 1) if box else None


def touches(text_png, fig_png):
    """True when a figure pixel lies inside the bounding box of the text's ink."""
    box = Image.open(text_png).convert("L").point(lambda v: 255 if v < DARK else 0).getbbox()
    if not box:
        return False
    return Image.open(fig_png).convert("L").crop(box).point(lambda v: 255 if v < FIG_INK else 0).getbbox() is not None


def measure(deck):
    deck = os.path.abspath(deck)
    app = w.Dispatch("PowerPoint.Application")
    for k in range(1, app.Presentations.Count + 1):
        try:
            same = os.path.normcase(app.Presentations(k).FullName) == os.path.normcase(deck)
        except Exception:
            continue
        if same:
            raise SystemExit(f"{deck} is open in PowerPoint: close it first")
    pres = app.Presentations.Open(deck, True, False, False)  # ReadOnly, Untitled, WithWindow
    tmp = tempfile.mkdtemp(prefix="title_measure_")
    out = []
    try:
        sw, sh = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        px_in = WIDTH_PX / (sw / 72.0)
        hpx = round(WIDTH_PX * sh / sw)
        for n in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(n)
            shapes = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]
            texts = [s for s in shapes if s.Type != PICTURE and s.HasTextFrame and s.TextFrame.HasText
                     and not s.TextFrame.TextRange.Text.strip().isdigit()]
            pics = [s for s in shapes if s.Type == PICTURE]
            if not texts:
                continue
            box = min(texts, key=lambda s: s.Top)
            para = box.TextFrame.TextRange.Paragraphs(1)
            lines = 0
            while lines < 50 and para.Lines(lines + 1, 1).Length > 0:
                lines += 1
            gap = None
            bottom = [s for s in texts if s.Name == "Bottom text"]   # workflow._add_bottom_text
            if pics:
                vis = {id(s): s.Visible for s in shapes}
                rows = {}
                try:
                    for name, keep in (("title", [box]), ("figure", pics)) + ((("bottom", bottom),) if bottom else ()):
                        for s in shapes:
                            s.Visible = -1 if any(s is k for k in keep) else 0
                        png = os.path.join(tmp, f"s{n}_{name}.png")
                        slide.Export(png, "PNG", WIDTH_PX, hpx)
                        rows[name] = ink_rows(png)
                finally:
                    for s in shapes:
                        s.Visible = vis[id(s)]
                if rows["title"] and rows["figure"]:
                    gap = round((rows["figure"][0] - rows["title"][1] - 1) / px_in, 3)
            rec = {"slide": n, "title": para.Text.strip(), "lines": lines, "gap_in": gap,
                   "runs_into_figure": bool(gap is not None and gap < 0)}
            if bottom:
                rng, k = bottom[0].TextFrame.TextRange, 0
                while k < 50 and rng.Lines(k + 1, 1).Length > 0:
                    k += 1
                fig_png = os.path.join(tmp, f"s{n}_figure.png")
                rec.update(bottom=rng.Text.strip(), bottom_lines=k,
                           bottom_overlaps_figure=bool(pics) and touches(os.path.join(tmp, f"s{n}_bottom.png"), fig_png),
                           bottom_off_slide=rng.BoundTop + rng.BoundHeight > sh + 0.5)
            out.append(rec)
    finally:
        pres.Close()
    return out


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    res = measure(sys.argv[1])
    with open(sys.argv[2], "w", encoding="utf-8") as f:
        json.dump(res, f, indent=1, ensure_ascii=False)
    print(f"{len(res)} slides measured -> {sys.argv[2]}")
