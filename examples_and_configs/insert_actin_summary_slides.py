"""
insert_actin_summary_slides.py

Build the accumulating CART actin summary deck. Each slide shows one
PNG plot from the actin-only timecourse summaries with:
  - bold title derived from the filename
  - the plot, full-width, aspect ratio preserved
  - the full source filepath in a small footer (for traceability)

This script is intended to be extended over time by appending entries
to PNG_FILES. The first batch is the actin_bottom_* timecourse scatter
plots from grid_panels/timecourse_scatter_plots/.

Modeled after insert_actin_qc_synapse_mask_xz_mip_slides.py.

Usage:
    python examples_and_configs/insert_actin_summary_slides.py
"""

import os
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

OUTPUT_PATH = Path(
    "K:/FF/PPT/PPT_autogeneration/CART_actin_only/CART_actin_summary.pptx"
)

SOURCE_DIR_TIMECOURSE_SCATTER = Path(
    "Y:/User_data/Kiet/results_compiled/actin_only/"
    "compiled_20260312_20260414_20260510_20260604/"
    "grid_panels/timecourse_scatter_plots"
)

# Each entry: (source_dir, png_filename). Order = slide order. Append
# entries to extend the deck; the script is idempotent.
PNG_FILES = [
    # --- actin_bottom_* timecourse scatter plots (round 1) -----------------
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_MFI_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_mask_area_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_total_sig_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_slice_inner_outer_ratio_50pct_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_slice_inner_outer_ratio_70pct_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_slice_perimeter_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_slice_circularity_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_slice_solidity_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_slice_eccentricity_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_MFI_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_mask_area_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_total_sig_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_mip_inner_outer_ratio_50pct_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_mip_inner_outer_ratio_70pct_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_perimeter_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_circularity_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_solidity_grid.png"),
    (SOURCE_DIR_TIMECOURSE_SCATTER, "actin_bottom_3slice_eccentricity_grid.png"),
]

# Colors
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)

# Slide layout (inches). 13.333 x 7.5 widescreen.
SLIDE_W = 13.333
SLIDE_H = 7.5

MARGIN = 0.10

TITLE_LEFT = MARGIN
TITLE_TOP = 0.05
TITLE_WIDTH = SLIDE_W - 2 * MARGIN
TITLE_HEIGHT = 0.50
TITLE_FONT_PT = 28

SUBTITLE_LEFT = MARGIN
SUBTITLE_TOP = 0.55
SUBTITLE_WIDTH = SLIDE_W - 2 * MARGIN
SUBTITLE_HEIGHT = 0.32
SUBTITLE_FONT_PT = 14

# Image placement: bigger when there's no subtitle, smaller when there is.
IMG_LEFT = MARGIN
IMG_TOP = 0.60
IMG_BOX_W = SLIDE_W - 2 * MARGIN
IMG_BOX_H = 6.40

IMG_TOP_WITH_SUBTITLE = 0.90
IMG_BOX_H_WITH_SUBTITLE = 6.10

FOOTER_LEFT = MARGIN
FOOTER_TOP = 7.05
FOOTER_WIDTH = SLIDE_W - 2 * MARGIN
FOOTER_HEIGHT = 0.40
FOOTER_FONT_PT = 9

# Per-token substitutions for prettify_metric_name (keys are lowercase).
TOKEN_MAP = {
    "mfi": "MFI",
    "mip": "MIP",
    "sig": "Signal",
    "num": "Number",
    "idx": "Index",
    "rel": "Rel.",
    "3slice": "(3-Slice MIP)",
    "bottom": "Synapse",
}

# Tokens dropped entirely from the title.
SKIP_TOKENS = {"slice"}

# Per-filename overrides — applied before any other title logic.
TITLE_OVERRIDES = {
    "actin_bottom_mask_area_grid.png": "Synapse Area",
    "actin_bottom_3slice_mask_area_grid.png": "Synapse (3-Slice MIP) Area",
}

# Optional per-filename subtitle shown under the title (formula / units /
# interpretation). Missing entries -> no subtitle line. Currently empty.
SUBTITLE_BY_FILENAME = {}

# ---------------------------------------------------------------------------


def prettify_metric_name(filename: str) -> str:
    """Convert a PNG filename to a readable slide title.

    Checks TITLE_OVERRIDES first. Otherwise strips `.png` and `_grid`,
    splits on `_`, applies per-token substitutions (dropping tokens in
    SKIP_TOKENS and collapsing the redundant `3slice_mip` pair into a
    single `(3-Slice MIP)`), and rewrites `Inner Outer` as `Inner/Outer`.
    """
    if filename in TITLE_OVERRIDES:
        return TITLE_OVERRIDES[filename]

    stem = filename
    if stem.lower().endswith(".png"):
        stem = stem[:-4]
    if stem.lower().endswith("_grid"):
        stem = stem[:-5]

    tokens = stem.split("_")
    out = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        low = tok.lower()
        nxt = tokens[i + 1].lower() if i + 1 < len(tokens) else None

        if low == "3slice" and nxt == "mip":
            out.append("(3-Slice MIP)")
            i += 2
            continue
        if low in SKIP_TOKENS:
            i += 1
            continue
        if low in TOKEN_MAP:
            out.append(TOKEN_MAP[low])
        elif low.endswith("pct") and low[:-3].isdigit():
            out.append(f"({low[:-3]}% Eff Rad Thresh)")
        else:
            out.append(tok.capitalize())
        i += 1

    pretty = " ".join(out)
    pretty = pretty.replace("Inner Outer", "Inner/Outer")
    return pretty


def add_textbox(slide, text, left, top, width, height, font_pt, color,
                bold=False, italic=False):
    """Add a centered textbox with given text and styling."""
    box = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = box.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.text = text
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.runs[0]
    run.font.size = Pt(font_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def add_image_in_box(slide, image_path, box_left, box_top, box_w, box_h):
    """Place an image inside the given bounding box, preserving aspect
    ratio and centering on the dimension that is < box."""
    pic = slide.shapes.add_picture(
        image_path,
        Inches(box_left),
        Inches(box_top),
        width=Inches(box_w),
    )
    actual_h_in = pic.height / 914400.0
    if actual_h_in > box_h:
        sp = pic._element
        sp.getparent().remove(sp)
        pic = slide.shapes.add_picture(
            image_path,
            Inches(box_left),
            Inches(box_top),
            height=Inches(box_h),
        )
        actual_w_in = pic.width / 914400.0
        pic.left = Inches(box_left + (box_w - actual_w_in) / 2)
    else:
        pic.top = Inches(box_top + (box_h - actual_h_in) / 2)
    return pic


def set_slide_background(slide, rgb: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def build_slide(prs, title_text: str, image_path: Path, footer_text: str,
                subtitle_text=None):
    """Build one slide: title + optional subtitle + full image + filepath footer.
    Returns (slide, missing_flag)."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, WHITE)

    add_textbox(
        slide, title_text,
        TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
        font_pt=TITLE_FONT_PT, color=BLACK, bold=True,
    )

    if subtitle_text:
        add_textbox(
            slide, subtitle_text,
            SUBTITLE_LEFT, SUBTITLE_TOP, SUBTITLE_WIDTH, SUBTITLE_HEIGHT,
            font_pt=SUBTITLE_FONT_PT, color=BLACK, italic=True,
        )
        img_top, img_h = IMG_TOP_WITH_SUBTITLE, IMG_BOX_H_WITH_SUBTITLE
    else:
        img_top, img_h = IMG_TOP, IMG_BOX_H

    missing = not image_path.exists()
    if not missing:
        add_image_in_box(
            slide, str(image_path), IMG_LEFT, img_top, IMG_BOX_W, img_h
        )
    else:
        add_textbox(
            slide, "(missing)",
            IMG_LEFT, img_top + img_h / 2 - 0.2, IMG_BOX_W, 0.4,
            font_pt=18, color=BLACK,
        )

    add_textbox(
        slide, footer_text,
        FOOTER_LEFT, FOOTER_TOP, FOOTER_WIDTH, FOOTER_HEIGHT,
        font_pt=FOOTER_FONT_PT, color=BLACK, bold=False,
    )
    return slide, missing


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    missing = []
    print(f"Writing deck to: {OUTPUT_PATH}\n")

    for source_dir, png_name in PNG_FILES:
        image_path = source_dir / png_name
        title = prettify_metric_name(png_name)
        footer = image_path.as_posix()
        subtitle = SUBTITLE_BY_FILENAME.get(png_name)
        _, is_missing = build_slide(
            prs, title, image_path, footer, subtitle_text=subtitle
        )
        status = "OK" if not is_missing else "MISSING"
        print(f"[{png_name}]  {status}  -> {title!r}")
        if is_missing:
            missing.append(f"{png_name}  ({image_path})")

    prs.save(str(OUTPUT_PATH))
    print(f"\nDone. {len(PNG_FILES)} slides written to:\n  {OUTPUT_PATH}")

    if missing:
        print(f"\nMissing ({len(missing)}):")
        for m in missing:
            print(f"  - {m}")
        sys.exit(1)
    print("\nAll images found - no missing items.")


if __name__ == "__main__":
    main()
