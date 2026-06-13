"""
insert_actin_qc_synapse_mask_xz_mip_20260607_slides.py

Per-condition single-image actin QC deck SCOPED TO the 20260607 Kiet
CART dataset only (Y:/User_data/Kiet/20260607_pMLC_CART_actin_hoescht).
Same layout as insert_actin_qc_synapse_mask_xz_mip_slides.py but
trimmed to one dataset so it can be run as soon as that dataset's
MATLAB pipeline outputs land.

Result: 1 dataset x 6 conditions x 2 kinds = 12 slides.

Until pipeline results exist, every slide will render as a
`(missing)` placeholder.

Usage:
    python examples_and_configs/insert_actin_qc_synapse_mask_xz_mip_20260607_slides.py
"""

import os
import re
import sys
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

OUTPUT_PATH = (
    "K:/FF/PPT/PPT_autogeneration/CART_actin_only/"
    "CART_actin_QC_synapse_mask_xz_mip_20260607.pptx"
)

KIET_ROOT = "Y:/User_data/Kiet"

# Single dataset scope. YYYYMMDD acquisition date shown in slide titles.
DATASETS = [
    ("20260607", "20260607_pMLC_CART_actin_hoescht"),
]

CONDITIONS = [
    "CAT_5min",  "FMC_5min",
    "CAT_10min", "FMC_10min",
    "CAT_15min", "FMC_15min",
]

CONDITION_SUBPATH = "converted/cropped/split_channels"
PROG_SUBPATH = "prog_fixed_cells_actin_only/actin"

# (kind label, subpath under PROG_SUBPATH/)
KINDS = [
    ("Actin at Synapse", "synapse/inner_outer/bot_combined/montages"),
    ("Actin XZ MIP",     "xz_mip/montages"),
]

CHUNK_GLOB = "montage_cells_*.png"

# Colors
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)

# Slide layout (inches). 13.333 x 7.5 widescreen.
SLIDE_W = 13.333
SLIDE_H = 7.5

TITLE_LEFT = 0.10
TITLE_TOP = 0.05
TITLE_WIDTH = SLIDE_W - 2 * 0.10
TITLE_HEIGHT = 0.50
TITLE_FONT_PT = 28

IMG_LEFT = 0.10
IMG_TOP = 0.60
IMG_BOX_W = SLIDE_W - 2 * 0.10           # 13.13"
IMG_BOX_H = SLIDE_H - IMG_TOP - 0.10     # 6.80"

# ---------------------------------------------------------------------------


def format_condition(folder_name: str) -> str:
    """CAT_5min -> 'CAT 5 min', FMC_10min -> 'FMC 10 min'."""
    time_map = {"5min": "5 min", "10min": "10 min", "15min": "15 min"}
    parts = folder_name.split("_", 1)
    if len(parts) != 2:
        return folder_name
    cell, time = parts
    return f"{cell} {time_map.get(time.lower(), time)}"


def add_textbox(slide, text, left, top, width, height, font_pt, color, bold=False):
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


def _chunk_start_index(p: Path) -> int:
    """Extract the first integer after 'montage_cells_' for natural sort."""
    m = re.match(r"montage_cells_(\d+)", p.name)
    return int(m.group(1)) if m else 0


def find_first_chunk(montages_dir: Path) -> Optional[Path]:
    """Return the lowest-numbered montage_cells_*.png in the dir, or None."""
    if not montages_dir.is_dir():
        return None
    chunks = list(montages_dir.glob(CHUNK_GLOB))
    if not chunks:
        return None
    chunks.sort(key=_chunk_start_index)
    return chunks[0]


def build_slide(prs, title_text: str, image_path: Optional[Path]):
    """Build one full-image slide with title. Returns (slide, missing_flag)."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, BLACK)

    add_textbox(
        slide, title_text,
        TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
        font_pt=TITLE_FONT_PT, color=WHITE, bold=True,
    )

    if image_path is not None and image_path.exists():
        add_image_in_box(slide, str(image_path), IMG_LEFT, IMG_TOP, IMG_BOX_W, IMG_BOX_H)
        return slide, False

    add_textbox(
        slide, "(missing)",
        IMG_LEFT, IMG_TOP + IMG_BOX_H / 2 - 0.2, IMG_BOX_W, 0.4,
        font_pt=18, color=WHITE,
    )
    return slide, True


def main() -> None:
    out_path = Path(OUTPUT_PATH)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    kiet_root = Path(KIET_ROOT)
    missing = []
    slides_added = 0

    print(f"Writing deck to: {OUTPUT_PATH}\n")

    for kind_label, kind_subpath in KINDS:
        for date_tag, dataset_name in DATASETS:
            for cond in CONDITIONS:
                cond_pretty = format_condition(cond)
                montages_dir = (
                    kiet_root
                    / dataset_name
                    / cond
                    / CONDITION_SUBPATH
                    / PROG_SUBPATH
                    / kind_subpath
                )
                chunk = find_first_chunk(montages_dir)
                title = f"{kind_label}: {cond_pretty} ({date_tag})"
                _, is_missing = build_slide(prs, title, chunk)
                slides_added += 1
                status = "OK" if not is_missing else "MISSING"
                print(f"[{kind_label}/{date_tag}/{cond}]  {status}")
                if is_missing:
                    missing.append(f"{kind_label}/{date_tag}/{cond}  ({montages_dir})")

    prs.save(str(out_path))
    print(f"\nDone. {slides_added} slides written to:\n  {out_path}")

    if missing:
        print(f"\nMissing ({len(missing)}):")
        for m in missing:
            print(f"  - {m}")
    else:
        print("\nAll images found - no missing items.")


if __name__ == "__main__":
    main()
