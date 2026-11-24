#!/usr/bin/env python
"""Make the template for the naive PCA trajectory deck (config: configs/live/ctl_nucleus/naive_pca_trajectory_deck.yaml).

From the Naive CTL summary template (left unchanged): slide 1 = a title slide (theme's "Title Slide" layout, title
placeholder only, its text set by the config's title_slide_text), slide 2 = its two-picture slide with both boxes
widened to fill the slide (FF 2026-11-24: "widen to fill slide"). Its other slides are dropped.

Usage:
    python examples_and_configs/ctl_nucleus/make_naive_pca_trajectory_template.py
"""
from pptx import Presentation
from pptx.util import Inches

SOURCE = "K:/FF/PPT/PPT_autogeneration/Naive_CTL/NaiveCTL_GzmB_histone_summary_template.pptx"
OUT = "K:/FF/PPT/PPT_autogeneration/CTL_Nucleus/naive_pca_trajectory/naive_pca_trajectory_template.pptx"

# Picture boxes (inches): two side by side under the title band, 0.15 in margins, aspect is kept inside each box
BOX_TOP, BOX_H, BOX_W, MARGIN = 1.2, 6.1, 6.45, 0.15


def main():
    prs = Presentation(SOURCE)
    slide_w = prs.slide_width / 914400

    # Keep only the two-picture slide (source slide 1)
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids)[1:]:
        prs.part.drop_rel(sld_id.rId)
        sld_ids.remove(sld_id)

    two_up = prs.slides[0]
    pics = sorted([s for s in two_up.shapes if s.shape_type == 13], key=lambda s: s.left)
    assert len(pics) == 2, f"expected 2 pictures on the source slide, found {len(pics)}"
    lefts = [MARGIN, slide_w - MARGIN - BOX_W]
    for pic, left in zip(pics, lefts):
        pic.left, pic.top, pic.width, pic.height = Inches(left), Inches(BOX_TOP), Inches(BOX_W), Inches(BOX_H)

    # Title slide in front: title placeholder only (placeholder text replaced by title_slide_text)
    layout = next(l for l in prs.slide_layouts if l.name == "Title Slide")
    title_slide = prs.slides.add_slide(layout)
    for ph in list(title_slide.placeholders):
        if ph.placeholder_format.type != 3:   # 3 = CENTER_TITLE; drop the subtitle box
            ph._element.getparent().remove(ph._element)
    title_slide.shapes.title.text = "Title"
    sld_ids.insert(0, sld_ids[-1])            # move it to slide 1

    prs.save(OUT)
    print(f"Saved {OUT}: slide 1 title, slide 2 two boxes {BOX_W} x {BOX_H} in at left {lefts}")


if __name__ == "__main__":
    main()
