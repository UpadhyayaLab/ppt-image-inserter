"""Shows the text-off-data check (ppt_image_inserter.text_clear) failing on a slide built to overlap, and passing after
text_clear.fix: a 2-line title at the deck's title size, a picture of solid data placed over the title, and an 18 pt
italic bottom line under the picture's lower edge (slide 1). Slide 2 has no bottom line and its data reaches down near
the slide's lower edge, so moving the data down to clear the title must not push it off the slide (2026-11-20: Fig 6
slides 3-4 were moved until their data ran off the bottom). Uses the builder's own title and bottom-line code (workflow).

Usage:  python test_text_clear.py [<output dir>]   (default: the Fig 6 decision previews' tests folder; needs
PowerPoint). Exit code 0 only if the check fails before the fix and passes after it.
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from ppt_image_inserter import text_clear  # noqa: E402
from ppt_image_inserter.workflow import _add_bottom_text, _add_title  # noqa: E402

OUT = Path(sys.argv[1] if len(sys.argv) > 1 else
           "K:/FF/PPT/PPT_autogeneration/Fig6_metrics_across_experiments/decision_previews/tests/text_clear")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    img = OUT / "data_block.png"
    im = Image.new("RGB", (1300, 760), "white")      # white margin around the data, which must not count
    ImageDraw.Draw(im).rectangle((60, 20, 1240, 755), fill=(0, 200, 200))
    im.save(img)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(img), Inches(0.3), Inches(0.05), width=Inches(12.73))   # data from top to bottom edge
    _add_title(slide, "Relative level vs distance from the centrosome-proximal nadir: Jurkat H3K27me3, nocodazole",
               size=32)
    _add_bottom_text(slide, "Oct 29, 2026", 13.333, 7.5, {"size": 18, "italic": True, "align": "left", "margin": 0.15})
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    slide2.shapes.add_picture(str(img), Inches(0.3), Inches(0.05), width=Inches(12.73))   # no bottom line
    _add_title(slide2, "Relative level vs distance from the centrosome-proximal nadir: Jurkat H3K27me3, nocodazole",
               size=32)
    deck = OUT / "overlap_test.pptx"
    prs.save(deck)

    before = text_clear.check(str(deck))
    moved = text_clear.fix(str(deck))
    after = text_clear.check(str(deck))
    print("before fix:", "; ".join(text_clear.report(before)) or "no overlap")
    print("fix:", moved)
    print("after fix:", "; ".join(text_clear.report(after)) or "no overlap")
    ok = {(p["slide"], p["text"]) for p in before} == {(1, "title"), (1, "bottom"), (2, "title")} and not after
    print("test_text_clear:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
