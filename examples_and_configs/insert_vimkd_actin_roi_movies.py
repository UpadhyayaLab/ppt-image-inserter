"""Build the Fig 7C actin ROI review deck: one slide per cell, one looping movie per slide.

Vimentin paper, Kiet's siCtrl / siVim live F-tractin movies (2026-10-10). Each movie is the
composite written by Fiji (K:/FF/Code/Scripts/Frank/FIJI_macros/vimkd_actin_roi_movie.ijm via
run_vimkd_actin_roi_movies.ps1): the whole cell with the 4 analyzed ROI boxes, plus the 4 ROIs
zoomed. This script converts each PNG-compressed AVI to a PowerPoint-safe H.264 MP4 (written
beside the AVI) and embeds it with the first frame as poster.

Run in the ``PPT_editing`` environment, with the output deck closed in PowerPoint:

    conda run -n PPT_editing python examples_and_configs/insert_vimkd_actin_roi_movies.py
    conda run -n PPT_editing python examples_and_configs/insert_vimkd_actin_roi_movies.py --list
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import List, Tuple

import tifffile
from PIL import Image

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from ppt_image_inserter import (  # noqa: E402
    MovieSpec,
    SlideSpec,
    TextboxSpec,
    backup_presentation,
    subset_output_path,
    build_movie_deck_via_com,
)
from insert_ctl_nuc_cent_coculture_seg_qc_movies import (  # noqa: E402
    fit_within_box,
    pad_movie_hold_last,
)

DATA_ROOT = Path("N:/VIM_KD_actin_data_for_Ivan/8_9_min_after_dropping")
MOVIE_ROOT = DATA_ROOT / "roi_movies_for_ppt_20261010"
DEFAULT_OUTPUT = Path(
    "K:/FF/Papers/vimentin_paper/Manuscript/figures_PPT/"
    "Fig7_actin_ROI_movies_20261010/Fig7C_actin_ROI_movies_20261010.pptx"
)
CONDITIONS = ("siCtrl", "siVim")

# siVim Cell 11 has no peripheral fit in siVimentin_FitData.mat (set 1 and set 2).
TITLE_NOTES = {("siVim", 11): "no peripheral fit saved; likely not in Fig 7C"}

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
POINTS_PER_INCH = 72.0
FONT_NAME = "Calibri"
TITLE_FONT_SIZE_PT = 24.0
NOTE_FONT_SIZE_PT = 14.0
SIDE_MARGIN_IN = 0.3
TITLE_TOP_IN = 0.08
TITLE_HEIGHT_IN = 0.5
SOURCE_TOP_IN = 0.56
NOTE_HEIGHT_IN = 0.3
MOVIE_TOP_IN = 0.95
MOVIE_HEIGHT_IN = 6.05
LEGEND_TOP_IN = 7.08
LEGEND = (
    "Frames 1\u2013150 of 301 (1 s/frame). Boxes: analyzed ROIs, 39 px (2.5 \u03bcm)."
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument("--movie-root", type=Path, default=MOVIE_ROOT)
    parser.add_argument(
        "--hold-last", type=float, default=0.5,
        help="Seconds the final frame is held (PowerPoint drops the last frame). Default 0.5.",
    )
    parser.add_argument(
        "--only", default="",
        help="Comma-separated subset, e.g. siCtrl:1,siVim:11 (default: every rendered movie).",
    )
    parser.add_argument("--loop", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--list", action="store_true", help="List the movies and exit.")
    return parser.parse_args()


def collect_movies(movie_root: Path) -> List[Tuple[str, int, Path]]:
    """(condition, cell number, avi) for every rendered movie, siCtrl first, by cell number."""
    records = []
    for cond in CONDITIONS:
        found = []
        for avi in (movie_root / cond).glob("Cell*_roi_movie.avi"):
            match = re.fullmatch(r"Cell(\d+)_roi_movie\.avi", avi.name)
            if match:
                found.append((cond, int(match.group(1)), avi))
        records.extend(sorted(found, key=lambda r: r[1]))
    return records


def source_nd2_name(cond: str, cell: int) -> str:
    """Name of the original .nd2, from the 'Series 0 Name' line of the cell TIFF's Info block."""
    with tifffile.TiffFile(DATA_ROOT / cond / f"Cell{cell}.tif") as tif:
        info = (tif.imagej_metadata or {}).get("Info", "")
    match = re.search(r"Series 0 Name = (.+?\.nd2)", info)
    return match.group(1) if match else ""


def _pt(inches: float) -> float:
    return inches * POINTS_PER_INCH


def _textbox(text: str, top_in: float, height_in: float, font_size_pt: float) -> TextboxSpec:
    return TextboxSpec(
        text=text,
        left_pt=_pt(SIDE_MARGIN_IN),
        top_pt=_pt(top_in),
        width_pt=_pt(SLIDE_WIDTH_IN - 2 * SIDE_MARGIN_IN),
        height_pt=_pt(height_in),
        font_size_pt=font_size_pt,
        bold=False,
        italic=True,
        align="center",
        font_name=FONT_NAME,
    )


def build_slide(cond: str, cell: int, avi: Path, hold_seconds: float) -> SlideSpec:
    poster = avi.with_suffix(".png")  # first frame, written by the Fiji macro
    if not poster.exists():
        raise FileNotFoundError(f"missing poster {poster}")
    mp4 = avi.with_suffix(".mp4")
    pad_movie_hold_last(avi, mp4, hold_seconds)

    with Image.open(poster) as im:
        frame_w, frame_h = im.size
    left, top, width, height = fit_within_box(
        frame_w, frame_h,
        SIDE_MARGIN_IN, MOVIE_TOP_IN, SLIDE_WIDTH_IN - 2 * SIDE_MARGIN_IN, MOVIE_HEIGHT_IN,
    )
    title = f"{cond} Cell {cell}"
    if (cond, cell) in TITLE_NOTES:
        title += f" ({TITLE_NOTES[(cond, cell)]})"
    return SlideSpec(
        textboxes=(
            _textbox(title, TITLE_TOP_IN, TITLE_HEIGHT_IN, TITLE_FONT_SIZE_PT),
            _textbox(source_nd2_name(cond, cell), SOURCE_TOP_IN, NOTE_HEIGHT_IN, NOTE_FONT_SIZE_PT),
            _textbox(LEGEND, LEGEND_TOP_IN, NOTE_HEIGHT_IN, NOTE_FONT_SIZE_PT),
        ),
        movies=(
            MovieSpec(
                movie_path=str(mp4.resolve()),
                poster_path=str(poster.resolve()),
                left_pt=_pt(left), top_pt=_pt(top), width_pt=_pt(width), height_pt=_pt(height),
            ),
        ),
    )


def main() -> int:
    args = parse_args()
    records = collect_movies(args.movie_root)
    only = {token.strip() for token in args.only.split(",") if token.strip()}
    if only:
        records = [r for r in records if f"{r[0]}:{r[1]}" in only]
    if not records:
        print(f"No Cell*_roi_movie.avi under {args.movie_root}")
        return 1
    if args.list:
        for cond, cell, avi in records:
            print(f"{cond} Cell {cell}: {avi}")
        return 0

    slides = []
    for cond, cell, avi in records:
        slides.append(build_slide(cond, cell, avi, args.hold_last))
        print(f"  {cond} Cell {cell}: {avi.with_suffix('.mp4').name}")

    output_path = args.output
    if args.output == DEFAULT_OUTPUT:   # a subset run never replaces the full deck
        output_path = subset_output_path(output_path, only=args.only)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        backup_presentation(str(output_path), backup_base=str(output_path.parent / "backups"))
    rewritten = build_movie_deck_via_com(
        slides, str(output_path), _pt(SLIDE_WIDTH_IN), _pt(SLIDE_HEIGHT_IN), loop=args.loop,
    )
    print(f"Wrote {output_path} ({len(slides)} slides, autoplay set on {rewritten})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
