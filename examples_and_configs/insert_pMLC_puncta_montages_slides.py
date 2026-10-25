"""
insert_pMLC_puncta_montages_slides.py

CAR-T pMLC band-pass peak-to-mean puncta-detector overlays.
CAT (left) vs FMC (right) side-by-side, ONE MONTAGE per condition
(the first chunk), one slide per (experiment, tp) group.

Reads only QC-passing cells from `montages_filtered/` (same
`montage_cells_<...>.png` naming as full `montages/`).

Note: the source folder is always `peak_to_mean_bandpass/`. The plain
`pMLC/synapse/peak_to_mean/` detector (without `_bandpass`) is the
broken original method - do NOT insert it into any deck. (For pMLC it
hasn't been renamed to `_ORIGINAL_NOT_WORKING` yet because a run is
still active, but the same "band-pass only" rule applies.)

Enumerates conditions directly from each experiment root (no manifest).

Layout per experiment:
    <exp_root>/<cond>/converted/cropped/split_channels/
        prog_fixed_cells/pMLC/synapse/peak_to_mean_bandpass/montages_filtered/

Conditions used (6 per experiment): CAT_5min, CAT_10min, CAT_15min,
FMC_5min, FMC_10min, FMC_15min. 3 timepoints x 3 experiments =
9 slides total (Mar 12, May 10, Jun 07 2026).

Usage:
    python examples_and_configs/insert_pMLC_puncta_montages_slides.py
"""

import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ppt_image_inserter import backup_presentation  # noqa: E402

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# All experiments that have finished reprocessing.
# Jun 07 was added 2026-09-16 (reprocess finished 2026-09-15 13:14-18:56).
EXPERIMENTS = [
    ("Mar 12, 2026",
     r"\\10.225.56.176\ULab_data\User_data\Kiet\03122026_pMLC_actin_CAR_T"),
    ("May 10, 2026",
     r"\\10.225.56.176\ULab_data\User_data\Kiet\20260510_pMLC_Actin561_nucleus_CAR_Tcell_"),
    ("Jun 07, 2026",
     r"\\10.225.56.176\ULab_data\User_data\Kiet\20260607_pMLC_CART_actin_hoescht"),
]

CHAN_SUB = "converted/cropped/split_channels"
PROGRESS_FOLDER = "prog_fixed_cells"

# (kind label, subpath under <base>/<progress_folder>/).
# Source is always `peak_to_mean_bandpass/`. Never use plain
# `peak_to_mean/` - that is the broken original detector.
KINDS = [
    ("pMLC Puncta", "pMLC/synapse/peak_to_mean_bandpass/montages_filtered"),
]

# Which timepoints in what order (per experiment).
TIMEPOINTS = [5, 10, 15]

OUTPUT_PATH = (
    "K:/FF/PPT/PPT_autogeneration/CART/pMLC/"
    "CART_pMLC_puncta_montages_20260915.pptx"
)

CHUNK_GLOB = "montage_cells_*.png"

# Colors
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)

SLIDE_W = 13.333
SLIDE_H = 7.5

TITLE_LEFT = 0.05
TITLE_TOP = 0.05
TITLE_WIDTH = SLIDE_W - 2 * 0.05
TITLE_HEIGHT = 0.40
TITLE_FONT_PT = 24

GRID_LEFT = 0.05
GRID_TOP = 0.50
CELL_W = 6.60
CELL_H = SLIDE_H - GRID_TOP - 0.05
LABEL_H = 0.22
IMG_H = CELL_H - LABEL_H
LABEL_FONT_PT = 14
COL_GAP = SLIDE_W - 2 * GRID_LEFT - 2 * CELL_W

CELL_POSITIONS = [
    (GRID_LEFT,                    GRID_TOP),
    (GRID_LEFT + CELL_W + COL_GAP, GRID_TOP),
]

PPUM_SOURCE = 30
SCALEBAR_UM = 5
SCALEBAR_PX = PPUM_SOURCE * SCALEBAR_UM

_COND_RE = re.compile(r"^(CAT|FMC63|FMC(?!63))_(\d+)min$", re.IGNORECASE)


def _long_path_str(p: Path) -> str:
    s = str(p)
    if os.name == "nt" and len(s) >= 240 and not s.startswith("\\\\?\\"):
        s = s.replace("/", "\\")
        if not s.startswith("\\\\"):
            return "\\\\?\\" + s
    return s


def parse_cond(name: str) -> Tuple[Optional[str], Optional[int]]:
    m = _COND_RE.match(name)
    if not m:
        return (None, None)
    cell = m.group(1).upper()
    if cell == "FMC63":
        cell = "FMC"
    return (cell, int(m.group(2)))


def resolve_kind_dir(exp_root: str, cond: str, kind_subpath: str) -> Optional[Path]:
    cand = Path(exp_root) / cond / CHAN_SUB / PROGRESS_FOLDER / kind_subpath
    return cand if os.path.isdir(_long_path_str(cand)) else None


def _parse_chunk_range(p: Path) -> Tuple[int, int]:
    m4 = re.match(r"montage_cells_(\d+)_(\d+)_(\d+)_(\d+)\.png$", p.name)
    if m4:
        f_a, c_a, f_b, c_b = (int(x) for x in m4.groups())
        return (f_a * 1000 + c_a, f_b * 1000 + c_b)
    m2 = re.match(r"montage_cells_(\d+)_(\d+)\.png$", p.name)
    if m2:
        return (int(m2.group(1)), int(m2.group(2)))
    return (0, 0)


def _list_chunk_files(montages_dir):
    """Long-path-safe list of montage chunk PNGs (unsorted); [] if dir absent.
    pathlib .is_dir()/.glob() silently return False/empty past Windows MAX_PATH
    (260) even when the dir exists, so enumerate via os.listdir over the
    \\?\-prefixed path."""
    import fnmatch
    long_dir = _long_path_str(montages_dir)
    if not os.path.isdir(long_dir):
        return []
    return [montages_dir / n for n in os.listdir(long_dir)
            if fnmatch.fnmatch(n, CHUNK_GLOB)]


def find_first_chunk(montages_dir: Optional[Path]) -> Optional[Path]:
    chunks = _list_chunk_files(montages_dir)
    if not chunks:
        return None
    parsed = [(p, *_parse_chunk_range(p)) for p in chunks]
    keep = []
    for (p, s, e) in parsed:
        shadowed = any(
            s2 <= s and e <= e2 and (s2 < s or e < e2)
            for (p2, s2, e2) in parsed
            if p2 is not p
        )
        if not shadowed:
            keep.append((p, s, e))
    if not keep:
        return None
    keep.sort(key=lambda x: x[1])
    return keep[0][0]


def _png_dims(path: Path) -> Tuple[int, int]:
    with Image.open(_long_path_str(path)) as im:
        return im.size


def _path_exists(p: Path) -> bool:
    return os.path.exists(_long_path_str(p))


def compute_deck_ppi(image_paths: List[Path], max_w_in: float, max_h_in: float) -> float:
    ppi = 0.0
    for p in image_paths:
        w_px, h_px = _png_dims(p)
        ppi = max(ppi, w_px / max_w_in, h_px / max_h_in)
    return ppi


def add_textbox(slide, text, left, top, width, height, font_pt, color, bold=False):
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


def set_slide_background(slide, rgb: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def add_image_in_cell_at_ppi(slide, image_path: Path, ppi: float,
                             cell_left: float, cell_top: float):
    w_px, h_px = _png_dims(image_path)
    w_in = w_px / ppi
    h_in = h_px / ppi
    img_area_top = cell_top + LABEL_H
    left_in = cell_left + (CELL_W - w_in) / 2
    top_in  = img_area_top + (IMG_H - h_in) / 2
    return slide.shapes.add_picture(
        _long_path_str(image_path),
        Inches(left_in),
        Inches(top_in),
        width=Inches(w_in),
    )


def build_compare_slide(prs, title_text: str,
                        cat_img: Optional[Path], fmc_img: Optional[Path],
                        deck_ppi: float):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, BLACK)

    add_textbox(
        slide, title_text,
        TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
        font_pt=TITLE_FONT_PT, color=WHITE, bold=True,
    )

    cells = [
        ("CAT", cat_img, CELL_POSITIONS[0]),
        ("FMC", fmc_img, CELL_POSITIONS[1]),
    ]
    missing = []
    for label, img_path, (cell_left, cell_top) in cells:
        add_textbox(
            slide, label,
            cell_left, cell_top, CELL_W, LABEL_H,
            font_pt=LABEL_FONT_PT, color=WHITE, bold=True,
        )
        if img_path is not None and _path_exists(img_path):
            add_image_in_cell_at_ppi(slide, img_path, deck_ppi, cell_left, cell_top)
        else:
            add_textbox(
                slide, "(missing)",
                cell_left, cell_top + LABEL_H + IMG_H / 2 - 0.15, CELL_W, 0.3,
                font_pt=14, color=WHITE,
            )
            missing.append(label)
    return slide, missing


def main() -> None:
    out_path = Path(OUTPUT_PATH)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Build slide specs by (experiment, tp), with CAT and FMC slots.
    slide_specs: List[Tuple] = []
    missing_total: List[str] = []

    for exp_label, exp_root in EXPERIMENTS:
        if not os.path.isdir(_long_path_str(Path(exp_root))):
            print(f"ERROR: experiment root missing: {exp_root}")
            continue

        for tp in TIMEPOINTS:
            for kind_label, kind_subpath in KINDS:
                cat_cond = f"CAT_{tp}min"
                fmc_cond = f"FMC_{tp}min"
                cat_dir = resolve_kind_dir(exp_root, cat_cond, kind_subpath)
                fmc_dir = resolve_kind_dir(exp_root, fmc_cond, kind_subpath)
                cat_img = find_first_chunk(cat_dir)
                fmc_img = find_first_chunk(fmc_dir)

                title = f"{kind_label} \u2014 {exp_label} \u2014 {tp} min"
                log_key = f"{exp_label}/{tp}min/{kind_label}"
                slide_specs.append((
                    title,
                    cat_img, fmc_img,
                    cat_dir, fmc_dir,
                    log_key,
                ))

    present_count = sum(
        1 for (_, cat_img, fmc_img, _, _, _) in slide_specs
        for p in (cat_img, fmc_img)
        if p is not None and _path_exists(p)
    )
    print(
        f"Experiments: {len(EXPERIMENTS)}  Timepoints per exp: {len(TIMEPOINTS)}\n"
        f"Kinds: {len(KINDS)}\n"
        f"Slides emitted: {len(slide_specs)}\n"
        f"Present cells: {present_count} / {2 * len(slide_specs)}\n"
        f"Deck PPI: ONE px/in for every slide, pinned to the largest montage in the deck\n"
        f"  (fits every cell box, and every slide shows one physical scale).\n"
        f"Source PPUM = {PPUM_SOURCE} px/\u03bcm (locked).\n"
    )
    print(f"Writing deck to: {OUTPUT_PATH}\n")

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    # One px/in for the whole deck (FF's one-scale rule, montage_scale.py): a per-slide PPI
    # shrank each slide by its own factor, 652-774 px/in across the 9 slides (2026-10-24).
    all_present = [
        p for (_, cat_img, fmc_img, _, _, _) in slide_specs
        for p in (cat_img, fmc_img)
        if p is not None and _path_exists(p)
    ]
    deck_ppi = compute_deck_ppi(all_present, CELL_W, IMG_H) if all_present else 100.0

    slides_added = 0
    for (title, cat_img, fmc_img, cat_dir, fmc_dir, log_key) in slide_specs:
        slide_ppi = deck_ppi
        bar_cm = (SCALEBAR_PX / slide_ppi) * 2.54

        _, missing = build_compare_slide(prs, title, cat_img, fmc_img, slide_ppi)
        slides_added += 1

        status_parts = [
            "CAT:OK" if cat_img and _path_exists(cat_img) else "CAT:MISSING",
            "FMC:OK" if fmc_img and _path_exists(fmc_img) else "FMC:MISSING",
        ]
        print(
            f"[{log_key}]  PPI={slide_ppi:7.2f}  bar={bar_cm:.3f}cm  "
            + "  ".join(status_parts)
        )

        for cell in missing:
            src = cat_dir if cell == "CAT" else fmc_dir
            missing_total.append(f"{log_key}/{cell}  ({src})")

    if Path(out_path).exists():    # the output name is fixed: back up the previous deck first
        backup_presentation(str(out_path), backup_base=str(Path(out_path).parent / "backups"))
    prs.save(str(out_path))
    print(f"\nDone. {slides_added} slides written to:\n  {out_path}")

    if missing_total:
        print(f"\nMissing ({len(missing_total)}):")
        for m in missing_total:
            print(f"  - {m}")
    else:
        print("\nAll images found - no missing items.")


if __name__ == "__main__":
    main()
