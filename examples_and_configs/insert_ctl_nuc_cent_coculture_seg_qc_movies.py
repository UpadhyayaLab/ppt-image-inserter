#!/usr/bin/env python
"""
Build a PowerPoint deck of CTL nucleus/centrosome coculture nucleus-seg QC movies.

Source data: a confocal, high-resolution B16-OVA + CTL conjugate imaging
experiment (cent + nuc), dated 20260616. For each field of view (FOV) the
nucleus-segmentation QC pipeline renders four movies of the same field, differing
only in contrast/overlay:

- ``masks_linear5-50``       nucleus outlines, dim contrast window [5, 50]
- ``masks_fullrange_gamma``  outlines, full range (gamma; bright ones saturate)
- ``masks_fullrange_log``    outlines, full range (log; bright ones not saturated)
- ``tracks``                 outlines + track IDs + motion tails

This script builds one slide per FOV, laying the four renderings out in a single
horizontal row (1x4) with a caption above each. Movies are inserted via
PowerPoint COM (Windows-only); after saving, the timing-tree trigger on disk is
rewritten so videos autoplay in slideshow mode, and (with --loop, the default)
each movie is marked to loop until stopped.

Both ``.mp4`` and ``.avi`` copies of each clip exist in the source folder; this
script uses the ``.mp4`` files (better PowerPoint embedding).
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

from PIL import Image


REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from ppt_image_inserter import (
    MovieSpec,
    SlideSpec,
    TextboxSpec,
    backup_presentation,
    build_movie_deck_via_com,
)

try:
    import imageio.v3 as iio
except ImportError:
    iio = None

try:
    import imageio_ffmpeg
except ImportError:
    imageio_ffmpeg = None


MOVIE_DIR = Path(
    "F:/FF/CTL_nucleus_centrosome_coculture/20260616/"
    "20260616_B16-OVA_CTLs_High-res_conjugate_imaging_cent_nuc/"
    "Confocal_MIPs/channels/nucleus_seg_qc_20260708/key_movies"
)

EXPERIMENT_DATE = "20260616"
MODALITY = "confocal"

DEFAULT_OUTPUT = Path(
    "K:/FF/PPT/PPT_autogeneration/CTL_nucleus_centrosome_coculture/"
    "CTL nucleus centrosome coculture, nucleus seg QC movies (confocal, 20260616).pptx"
)

# Poster frames and padded movie copies are transient per-run scratch; keep them
# out of the repo and out of the deck output folder (wiped and recreated each run).
DEFAULT_POSTER_DIR = Path(tempfile.gettempdir()) / "ppt_ctl_nuc_cent_coculture_posters"

# PowerPoint can drop the final timepoint of a movie on playback; cloning the
# last frame for a moment fixes it. But that fix (pad_movie_hold_last) re-encodes
# with a keyframe every 2 frames (-g 2), which is right for the 2 fps Jurkat mesh
# movies but badly inflates these 10 fps / 167-frame clips (well past their
# already-compressed source size). Here the movies loop by default, which masks
# the dropped final 100 ms anyway, so hold-last defaults OFF (0) and the small
# source .mp4s are embedded unchanged. Pass --hold-last 1 to force the re-encode.
HOLD_LAST_FRAME_SEC_DEFAULT = 0.0

# The renderings per FOV, in slide column order (left to right). Each entry is
# (filename token, caption shown above the movie). The full-range gamma variant
# is intentionally omitted (it saturates the bright nuclei); the log variant
# covers the full-range case. Dropping it leaves 3 per row, so each movie is
# larger. The gamma .mp4s still exist on disk and are simply not enumerated.
MOVIE_TYPES: Tuple[Tuple[str, str], ...] = (
    ("masks_linear5-50", "Linear scale [5, 50]"),
    ("masks_fullrange_log", "Log scale"),
    ("tracks", "Outlines + track IDs + tails"),
)

TYPE_TOKENS: Tuple[str, ...] = tuple(token for token, _ in MOVIE_TYPES)


# --- LLS (lattice light-sheet) MIP Cellpose-seg movies -----------------------
# A second modality of the same 20260616 experiment, added as extra slides after
# the confocal FOV slides. One slide per region, laid out in the same 3-across
# format as the confocal slides (linear tight | full-range log | seg+tracking).
# The gamma contrast variant is omitted for consistency; the "dualtone_gamma" in
# the seg movie names is the render style, not a contrast variant to drop.
LLS_BASE = Path(
    "F:/FF/CTL_nucleus_centrosome_coculture/20260616/"
    "20260616_B16-OVA_CTLs_High-res_conjugate_imaging_cent_nuc/"
    "LLS_MIPs/nucleus_cellpose"
)

# Whole-FOV first (the LLS analog of the confocal whole-field FOV slides), then
# the ROI zoom-ins.
LLS_REGIONS: Tuple[str, ...] = (
    "WA1_wholeFOV",
    "WA2_wholeFOV",
    "WA1_ROI1",
    "WA1_ROI2",
    "WA1_ROI3",
    "WA1_ROI4",
)

# (movie path relative to a region folder, caption) in slide column order. Chosen
# to mirror the confocal 3-across format. For LLS the middle column uses the
# gamma-corrected display (seg-only, no tracks) instead of the confocal log scale,
# since gamma reads better on the LLS background; linear_wide / full_range_log are
# omitted. The LLS linear window is auto-derived per file, so no fixed numbers in
# the "Linear scale" label (unlike the confocal fixed [5, 50]).
LLS_MOVIE_TYPES: Tuple[Tuple[str, str], ...] = (
    ("dim_nuclei_check/linear_tight.mp4", "Linear scale"),
    ("movies/seg-only_dualtone_gamma.mp4", "Gamma-corrected"),
    ("movies/seg+tracking_dualtone_gamma.mp4", "Outlines + track IDs + tails"),
)

LLS_MODALITY = "LLS"

# Exact per-WA linear display window shown on the "Linear scale" caption of the
# WHOLE-FOV slides only. The renderer's --auto samples 12 frames across the whole
# stack once and sets lo = p1, hi = p92, then maps every frame identically, so a
# single fixed window is meaningful for the whole field: WA1 ≈ [102, 188],
# WA2 ≈ [103, 194].
#
# The ROI conjugate crops (WA1_ROI1-4) are auto-windowed independently ON EACH
# CROP (lo≈p1, tight hi=p99), so their linear windows differ per ROI (e.g. ROI1
# [111, 738], ROI2 [112, 5341] (bright object, ~20x), ROI3 [111, 307], ROI4
# [108, 285]) and are NOT brightness-comparable. Those slides are therefore
# labelled with the method, not a fixed intensity window.
LLS_LINEAR_WINDOW_BY_WA: Dict[int, str] = {1: "[102, 188]", 2: "[103, 194]"}
LLS_ROI_LINEAR_NOTE = "(per-crop auto, p1–p99)"


# --- Low-SNR per-nucleus QC crops --------------------------------------------
# Per-track crop movies (nucleus outline over time) for dim / low-SNR nuclei,
# used to spot-check segmentation on the hardest cases. There are many (~40 per
# FOV); we add a small grid on its own slide(s), placed BEFORE the LLS section.
# Filenames: FOVn_MFI<mfi>_id<trackid>_len<frames>.mp4.
# MOVIE_DIR is .../nucleus_seg_qc_20260708/key_movies, so its parent is the QC
# folder. Two sibling collections with identical filenames exist:
#   low_snr_track_movies          - red outline + yellow track ID/tail
#   low_snr_crop_movies_notracks  - red outline only (no track overlay)  <- used
# Switch LOWSNR_DIR back to low_snr_track_movies to show the tracked version.
LOWSNR_DIR = MOVIE_DIR.parent / "low_snr_crop_movies_notracks"
LOWSNR_FILENAME_RE = re.compile(
    r"^FOV(?P<fov>\d+)_MFI(?P<mfi>\d+)_id(?P<id>\d+)_len(?P<len>\d+)\.mp4$"
)
# Selection: the dimmest (lowest MFI) tracks long enough to show tracking, taken
# in ascending-MFI order. The default is large enough to include ALL eligible
# (len >= LOWSNR_MIN_LEN) crops; lower it with --lowsnr-count (0 disables). At
# 8 per 4x2 grid page, the eligible set spans several slides.
LOWSNR_COUNT_DEFAULT = 1000
LOWSNR_MIN_LEN = 100

# Grid layout for the low-SNR slide(s). The crops carry their own baked-in banner
# (id / per-frame MFI / length), so no per-crop caption is added here.
LOWSNR_COLS = 4
LOWSNR_ROWS = 2  # cells per slide = COLS * ROWS
LOWSNR_BOX_IN = 3.0
LOWSNR_COL_GAP_IN = 0.2
LOWSNR_ROW_GAP_IN = 0.35

# --- LLS low-SNR per-nucleus crops (whole-FOV WA1/WA2, outline only) ----------
# Sibling of the confocal low-SNR crops, from the LLS whole-FOV data. Placed
# AFTER the LLS region slides (so the whole-FOV region slides precede the crops).
# Filenames: WA<n>_MFI<mfi>_id<id>_len<frames>.mp4. Same rule as the confocal
# grid: all crops (both WA) with len >= threshold, dimmest (lowest MFI) first.
LLS_LOWSNR_DIR = LLS_BASE / "low_snr_crop_movies_notracks"
LLS_LOWSNR_FILENAME_RE = re.compile(
    r"^WA(?P<wa>\d+)_MFI(?P<mfi>\d+)_id(?P<id>\d+)_len(?P<len>\d+)\.mp4$"
)
LLS_LOWSNR_MIN_LEN = 100


SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
TITLE_LEFT_IN = 0.45
TITLE_TOP_IN = 0.12
TITLE_WIDTH_IN = 12.4
TITLE_HEIGHT_IN = 0.35
TITLE_FONT_SIZE_PT = 22.0
POINTS_PER_INCH = 72.0

# Single-row layout (one column per rendering). The source movies are square
# (2048x2048), so the box width is the binding constraint; box height is set
# equal so each square movie fills its box exactly. Fewer columns => wider (and
# thus larger) boxes. Margins are kept tight to maximize movie size.
N_COLS = len(MOVIE_TYPES)
SIDE_MARGIN_IN = 0.3
COL_GAP_IN = 0.15
MOVIE_BOX_WIDTH_IN = (
    SLIDE_WIDTH_IN - 2 * SIDE_MARGIN_IN - (N_COLS - 1) * COL_GAP_IN
) / N_COLS
MOVIE_BOX_HEIGHT_IN = MOVIE_BOX_WIDTH_IN
CAPTION_HEIGHT_IN = 0.24
CAPTION_FONT_SIZE_PT = 12.0

# Vertically center the caption+movie block in the region below the title,
# clamped so the movie never runs off the bottom of the slide. Derived from the
# box height so it stays centered no matter how many columns there are.
_CAPTION_GAP_IN = 0.06
_REGION_TOP_IN = TITLE_TOP_IN + TITLE_HEIGHT_IN + 0.15
_REGION_BOTTOM_IN = SLIDE_HEIGHT_IN - 0.1
_BLOCK_HEIGHT_IN = CAPTION_HEIGHT_IN + _CAPTION_GAP_IN + MOVIE_BOX_HEIGHT_IN
CAPTION_TOP_IN = _REGION_TOP_IN + max(
    0.0, (_REGION_BOTTOM_IN - _REGION_TOP_IN - _BLOCK_HEIGHT_IN) / 2
)
MOVIE_BOX_TOP_IN = CAPTION_TOP_IN + CAPTION_HEIGHT_IN + _CAPTION_GAP_IN

# LLS whole-FOV movies are tall/narrow (444x1024), so the shared square row box
# renders them small. Whole-FOV slides get a taller box (caption just under the
# title, movie spanning most of the slide height) so the field is larger. Width
# stays the shared 3-column width; the tall movie is height-limited and centered.
LLS_WHOLEFOV_CAPTION_TOP_IN = 0.58
LLS_WHOLEFOV_MOVIE_TOP_IN = 0.84
LLS_WHOLEFOV_BOX_HEIGHT_IN = SLIDE_HEIGHT_IN - LLS_WHOLEFOV_MOVIE_TOP_IN - 0.12


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Create a PowerPoint deck of CTL nucleus/centrosome coculture "
            "nucleus-segmentation QC movies (confocal, 20260616). One slide per "
            "FOV, with the four renderings (linear [5,50], full-range gamma, "
            "full-range log, tracks) laid out in a 1x4 row."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=f"Output PowerPoint path. Defaults to {DEFAULT_OUTPUT}.",
    )
    parser.add_argument(
        "--fovs",
        type=str,
        default="",
        help="Optional comma-separated list of FOV numbers to include, e.g. 1,3.",
    )
    parser.add_argument(
        "--hold-last",
        type=float,
        default=HOLD_LAST_FRAME_SEC_DEFAULT,
        help=(
            "Seconds to hold (clone) the final frame of each movie so PowerPoint "
            "does not drop the last timepoint on playback (the re-encode also "
            "compresses the clips). 0 inserts the source movies unchanged. "
            f"Default {HOLD_LAST_FRAME_SEC_DEFAULT}."
        ),
    )
    parser.add_argument(
        "--loop",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Loop each movie until stopped in slideshow (default: on; --no-loop to disable).",
    )
    parser.add_argument(
        "--lowsnr-count",
        type=int,
        default=LOWSNR_COUNT_DEFAULT,
        help=(
            "Number of low-SNR per-track QC crops to add on grid slide(s) before "
            f"the LLS section (dimmest first). 0 disables. Default {LOWSNR_COUNT_DEFAULT}."
        ),
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List the FOV -> movie mapping and slide plan, then exit without building.",
    )
    return parser.parse_args()


def resolve_output_path(output_arg: Optional[Path]) -> Path:
    """Pick the output path based on --output."""
    if output_arg is not None:
        return output_arg.resolve()
    return DEFAULT_OUTPUT.resolve()


def parse_requested_fovs(fovs_arg: str) -> Optional[Set[int]]:
    """Parse optional comma-separated FOV numbers."""
    if not fovs_arg.strip():
        return None

    requested: Set[int] = set()
    for token in fovs_arg.split(","):
        token = token.strip()
        if not token:
            continue
        try:
            requested.add(int(token))
        except ValueError as exc:
            raise ValueError(f"Invalid FOV number in --fovs: {token}") from exc

    if not requested:
        raise ValueError("--fovs was provided but no valid FOV numbers were found")

    return requested


def validate_folder(folder: Path, label: str) -> None:
    """Ensure a required movie folder exists."""
    if not folder.exists():
        raise FileNotFoundError(f"{label} folder not found: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"{label} is not a folder: {folder}")


def _type_patterns() -> Dict[str, "re.Pattern[str]"]:
    """Anchored ``FOV<n>_<token>.mp4`` regex per movie type."""
    return {
        token: re.compile(
            rf"^FOV(?P<fov>\d+)_{re.escape(token)}\.mp4$", re.IGNORECASE
        )
        for token in TYPE_TOKENS
    }


def report_missing_fovs(
    type_to_movies: Dict[str, Dict[int, Path]],
    requested_fovs: Optional[Set[int]] = None,
) -> None:
    """Print warnings for FOVs missing one or more of the four renderings."""
    all_fovs: Set[int] = set()
    for movies in type_to_movies.values():
        all_fovs.update(movies)

    if requested_fovs is not None:
        all_fovs &= requested_fovs

    for fov in sorted(all_fovs):
        missing = [token for token in TYPE_TOKENS if fov not in type_to_movies[token]]
        if missing:
            print(
                f"[WARNING] FOV{fov} skipped; missing rendering(s): "
                f"{', '.join(missing)}"
            )


def collect_fov_movies(
    folder: Path,
    requested_fovs: Optional[Set[int]] = None,
) -> Dict[int, Dict[str, Path]]:
    """Return ``{fov -> {type token -> movie path}}`` for FOVs complete in all four types."""
    validate_folder(folder, "key_movies")

    patterns = _type_patterns()
    by_type: Dict[str, Dict[int, Path]] = {token: {} for token in TYPE_TOKENS}
    for movie_path in sorted(folder.glob("*.mp4")):
        for token, pattern in patterns.items():
            match = pattern.match(movie_path.name)
            if not match:
                continue
            fov = int(match.group("fov"))
            existing = by_type[token].get(fov)
            if existing is not None:
                raise ValueError(
                    f"Duplicate FOV{fov} {token} movie: "
                    f"{existing.name} and {movie_path.name}"
                )
            by_type[token][fov] = movie_path
            break

    for token, mapping in by_type.items():
        if not mapping:
            print(
                f"[WARNING] no FOV*.mp4 files matched type {token!r} in {folder}"
            )

    report_missing_fovs(by_type, requested_fovs)

    complete: Set[int] = set(by_type[TYPE_TOKENS[0]])
    for token in TYPE_TOKENS[1:]:
        complete &= set(by_type[token])
    if requested_fovs is not None:
        complete &= requested_fovs

    return {
        fov: {token: by_type[token][fov] for token in TYPE_TOKENS}
        for fov in sorted(complete)
    }


def pad_movie_hold_last(src: Path, dst: Path, hold_seconds: float) -> Path:
    """Re-encode ``src`` to ``dst``, holding the final frame for ``hold_seconds``.

    PowerPoint ends video playback slightly before the media's true end, so the
    final timepoint can go unshown on playback. Cloning the last frame for ~1s
    pushes the real final frame safely before that cutoff.

    Critically, the re-encode must NOT introduce B-frames: PowerPoint's decoder
    drops reordered (B-frame) tails, which makes it swallow *more* than one final
    frame. We use ``-bf 0`` and a matching keyframe interval so playback drops at
    most the true last frame (now a clone). The source file is never modified.
    Requires the bundled ffmpeg from imageio-ffmpeg.
    """
    if imageio_ffmpeg is None:
        raise RuntimeError(
            "imageio-ffmpeg is required to hold the final movie frame. Install "
            "imageio-ffmpeg in the active environment, or pass --hold-last 0 to "
            "disable."
        )
    exe = imageio_ffmpeg.get_ffmpeg_exe()
    cmd = [
        exe, "-y", "-loglevel", "error",
        "-i", str(src),
        "-vf", f"tpad=stop_mode=clone:stop_duration={hold_seconds}",
        "-c:v", "libx264", "-profile:v", "main", "-pix_fmt", "yuv420p",
        "-bf", "0", "-g", "2", "-crf", "18",
        "-movflags", "+faststart",
        str(dst),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0 or not dst.exists():
        raise RuntimeError(f"ffmpeg failed to pad {src.name}:\n{result.stderr}")
    return dst


def extract_first_frame(movie_path: Path, poster_path: Path) -> Tuple[int, int]:
    """Extract the first frame from a movie, save it, and return width/height."""
    if iio is None:
        raise RuntimeError(
            "imageio is required to extract poster frames. "
            "Install imageio and imageio-ffmpeg in the active environment."
        )

    frame = iio.imread(movie_path, index=0)
    poster = Image.fromarray(frame)
    poster.save(poster_path)
    return frame.shape[1], frame.shape[0]


def fit_within_box(
    content_width: int,
    content_height: int,
    box_left: float,
    box_top: float,
    box_width: float,
    box_height: float,
) -> Tuple[float, float, float, float]:
    """Fit content inside a box while preserving aspect ratio."""
    content_aspect = content_width / content_height
    box_aspect = box_width / box_height

    if content_aspect >= box_aspect:
        final_width = box_width
        final_height = box_width / content_aspect
    else:
        final_height = box_height
        final_width = box_height * content_aspect

    final_left = box_left + (box_width - final_width) / 2
    final_top = box_top + (box_height - final_height) / 2
    return final_left, final_top, final_width, final_height


def format_title(fov: int) -> str:
    """Return the slide title for a FOV."""
    return (
        f"FOV {fov} — B16-OVA + CTL conjugate, nucleus+centrosome "
        f"({MODALITY}, {EXPERIMENT_DATE})"
    )


def _to_points(inches: float) -> float:
    """Convert inches to PowerPoint points."""
    return inches * POINTS_PER_INCH


def _column_left_in(col: int) -> float:
    """Left edge (inches) of the movie box in column ``col`` (0-based)."""
    return SIDE_MARGIN_IN + col * (MOVIE_BOX_WIDTH_IN + COL_GAP_IN)


def _make_title_textbox(text: str) -> TextboxSpec:
    return TextboxSpec(
        text=text,
        left_pt=_to_points(TITLE_LEFT_IN),
        top_pt=_to_points(TITLE_TOP_IN),
        width_pt=_to_points(TITLE_WIDTH_IN),
        height_pt=_to_points(TITLE_HEIGHT_IN),
        font_size_pt=TITLE_FONT_SIZE_PT,
        bold=True,
        align="center",
        font_name="Arial",
    )


def _make_caption_box(
    text: str,
    left_in: float,
    top_in: float,
    width_in: float,
    font_size_pt: float = CAPTION_FONT_SIZE_PT,
    height_in: float = CAPTION_HEIGHT_IN,
) -> TextboxSpec:
    """A centered caption at an explicit position (inches)."""
    return TextboxSpec(
        text=text,
        left_pt=_to_points(left_in),
        top_pt=_to_points(top_in),
        width_pt=_to_points(width_in),
        height_pt=_to_points(height_in),
        font_size_pt=font_size_pt,
        bold=True,
        align="center",
        font_name="Arial",
    )


def _make_caption_textbox(text: str, box_left_in: float) -> TextboxSpec:
    """A caption centered over the single-row movie column at ``box_left_in``."""
    return _make_caption_box(
        text, box_left_in, CAPTION_TOP_IN, MOVIE_BOX_WIDTH_IN
    )


def _make_movie_spec_in_box(
    movie_path: Path,
    poster_path: Path,
    box_left_in: float,
    box_top_in: float,
    box_width_in: float,
    box_height_in: float,
    hold_seconds: float,
) -> MovieSpec:
    """Build a MovieSpec fit (aspect-preserving) into an explicit box (inches)."""
    # Insert a padded copy (final frame held) so PowerPoint does not drop the
    # last timepoint; the poster (first frame) is unaffected by the padding.
    movie_to_insert = movie_path
    if hold_seconds and hold_seconds > 0:
        held = poster_path.parent / f"{movie_path.stem}_held.mp4"
        movie_to_insert = pad_movie_hold_last(movie_path, held, hold_seconds)

    frame_width, frame_height = extract_first_frame(movie_to_insert, poster_path)
    left_in, top_in, width_in, height_in = fit_within_box(
        frame_width,
        frame_height,
        box_left_in,
        box_top_in,
        box_width_in,
        box_height_in,
    )
    return MovieSpec(
        movie_path=str(movie_to_insert.resolve()),
        poster_path=str(poster_path.resolve()),
        left_pt=_to_points(left_in),
        top_pt=_to_points(top_in),
        width_pt=_to_points(width_in),
        height_pt=_to_points(height_in),
    )


def _make_movie_spec(
    movie_path: Path,
    poster_path: Path,
    box_left_in: float,
    hold_seconds: float,
) -> MovieSpec:
    """Single-row movie: fit into the shared row box at ``box_left_in``."""
    return _make_movie_spec_in_box(
        movie_path,
        poster_path,
        box_left_in,
        MOVIE_BOX_TOP_IN,
        MOVIE_BOX_WIDTH_IN,
        MOVIE_BOX_HEIGHT_IN,
        hold_seconds,
    )


def _build_fov_slide(
    fov: int,
    type_to_movie: Dict[str, Path],
    poster_dir: Path,
    hold_seconds: float,
) -> SlideSpec:
    """Build one slide: title + a 1x4 row of captioned movies for one FOV."""
    textboxes: List[TextboxSpec] = [_make_title_textbox(format_title(fov))]
    movies: List[MovieSpec] = []
    for col, (token, caption) in enumerate(MOVIE_TYPES):
        box_left_in = _column_left_in(col)
        textboxes.append(_make_caption_textbox(caption, box_left_in))
        poster_path = poster_dir / f"FOV{fov}_{token}_poster.png"
        movies.append(
            _make_movie_spec(
                type_to_movie[token], poster_path, box_left_in, hold_seconds
            )
        )
    return SlideSpec(textboxes=tuple(textboxes), movies=tuple(movies))


def build_slide_specs(
    fov_movies: Dict[int, Dict[str, Path]],
    poster_dir: Path,
    hold_seconds: float,
) -> List[SlideSpec]:
    """Build one SlideSpec per FOV, in numeric order."""
    return [
        _build_fov_slide(fov, fov_movies[fov], poster_dir, hold_seconds)
        for fov in sorted(fov_movies)
    ]


def format_lls_region_label(region: str) -> str:
    """Human-readable region name, e.g. WA1_ROI1 -> 'WA1 ROI1', *_wholeFOV -> 'whole FOV'."""
    return region.replace("_wholeFOV", " whole FOV").replace("_", " ")


def _region_wa(region: str) -> int:
    """Well/acquisition number from a region name, e.g. WA1_ROI1 -> 1."""
    m = re.match(r"WA(\d+)", region)
    return int(m.group(1)) if m else 0


def _lls_caption(region: str, rel: str, caption: str) -> str:
    """Annotate the 'Linear scale' caption with the display window.

    Whole-FOV slides use a single stack-wide window (a fixed [lo, hi]); ROI slides
    are auto-windowed per crop, so they get a method note rather than a fixed
    window (their windows differ per ROI and are not brightness-comparable).
    """
    if "linear_tight" not in rel:
        return caption
    if region.endswith("_wholeFOV"):
        window = LLS_LINEAR_WINDOW_BY_WA.get(_region_wa(region))
        return f"{caption} {window}" if window else caption
    return f"{caption} {LLS_ROI_LINEAR_NOTE}"


def format_lls_title(region: str) -> str:
    """Return the (single-line) slide title for an LLS region."""
    return (
        f"{format_lls_region_label(region)} — nucleus Cellpose seg "
        f"({LLS_MODALITY}, {EXPERIMENT_DATE})"
    )


def collect_lls_regions(base: Path) -> List[Tuple[str, List[Path]]]:
    """Return ``[(region, [movie paths in column order])]`` for complete regions.

    A region is included only if all of ``LLS_MOVIE_TYPES`` are present; otherwise
    it is warned about and skipped so the rest of the deck still builds.
    """
    validate_folder(base, "LLS nucleus_cellpose")

    result: List[Tuple[str, List[Path]]] = []
    for region in LLS_REGIONS:
        paths: List[Path] = []
        missing: List[str] = []
        for rel, _caption in LLS_MOVIE_TYPES:
            movie_path = base / region / rel
            if movie_path.exists():
                paths.append(movie_path)
            else:
                missing.append(rel)
        if missing:
            print(
                f"[WARNING] LLS {region} skipped; missing: {', '.join(missing)}"
            )
            continue
        result.append((region, paths))
    return result


def _build_lls_slide(
    region: str,
    movie_paths: Sequence[Path],
    poster_dir: Path,
    hold_seconds: float,
) -> SlideSpec:
    """Build one slide: title + a 3-across row of captioned movies for one region.

    Whole-FOV regions use a taller movie box so their tall/narrow field fills
    more of the slide; ROI regions keep the shared square row box.
    """
    is_wholefov = region.endswith("_wholeFOV")
    caption_top_in = (
        LLS_WHOLEFOV_CAPTION_TOP_IN if is_wholefov else CAPTION_TOP_IN
    )
    movie_top_in = LLS_WHOLEFOV_MOVIE_TOP_IN if is_wholefov else MOVIE_BOX_TOP_IN
    box_height_in = (
        LLS_WHOLEFOV_BOX_HEIGHT_IN if is_wholefov else MOVIE_BOX_HEIGHT_IN
    )

    textboxes: List[TextboxSpec] = [_make_title_textbox(format_lls_title(region))]
    movies: List[MovieSpec] = []
    for col, ((rel, caption), movie_path) in enumerate(
        zip(LLS_MOVIE_TYPES, movie_paths)
    ):
        box_left_in = _column_left_in(col)
        textboxes.append(
            _make_caption_box(
                _lls_caption(region, rel, caption),
                box_left_in,
                caption_top_in,
                MOVIE_BOX_WIDTH_IN,
            )
        )
        token = rel.replace("/", "_").replace(".mp4", "")
        poster_path = poster_dir / f"{region}_{token}_poster.png"
        movies.append(
            _make_movie_spec_in_box(
                movie_path,
                poster_path,
                box_left_in,
                movie_top_in,
                MOVIE_BOX_WIDTH_IN,
                box_height_in,
                hold_seconds,
            )
        )
    return SlideSpec(textboxes=tuple(textboxes), movies=tuple(movies))


def build_lls_slide_specs(
    lls_regions: Sequence[Tuple[str, List[Path]]],
    poster_dir: Path,
    hold_seconds: float,
) -> List[SlideSpec]:
    """Build one SlideSpec per LLS region, in listed order."""
    return [
        _build_lls_slide(region, paths, poster_dir, hold_seconds)
        for region, paths in lls_regions
    ]


# One selected low-SNR track: (mfi, fov, track_id, length, path). Ordered so a
# plain sort puts the dimmest (lowest MFI) first.
LowSnrTrack = Tuple[int, int, int, int, Path]


def collect_lowsnr_tracks(folder: Path, count: int) -> List[LowSnrTrack]:
    """Return up to ``count`` dimmest low-SNR track crops (len >= LOWSNR_MIN_LEN)."""
    if count <= 0:
        return []
    if not folder.exists():
        print(f"[WARNING] low-SNR track folder not found; skipping ({folder})")
        return []

    records: List[LowSnrTrack] = []
    for p in sorted(folder.glob("*.mp4")):
        m = LOWSNR_FILENAME_RE.match(p.name)
        if not m:
            continue
        records.append(
            (int(m["mfi"]), int(m["fov"]), int(m["id"]), int(m["len"]), p)
        )

    eligible = [r for r in records if r[3] >= LOWSNR_MIN_LEN]
    if not eligible:
        print(
            f"[WARNING] no low-SNR tracks with len>={LOWSNR_MIN_LEN} in {folder}"
        )
        return []
    eligible.sort(key=lambda r: (r[0], r[1], r[2]))  # MFI asc, then FOV, id
    return eligible[:count]


def _lowsnr_grid_geometry() -> Tuple[float, float, float]:
    """Return (grid_left0_in, grid_top0_in, cell_height_in), vertically centered.

    No per-crop caption row: each crop carries its own baked-in banner, so a cell
    is just the movie box.
    """
    grid_w = LOWSNR_COLS * LOWSNR_BOX_IN + (LOWSNR_COLS - 1) * LOWSNR_COL_GAP_IN
    grid_left0 = (SLIDE_WIDTH_IN - grid_w) / 2
    cell_h = LOWSNR_BOX_IN
    grid_h = LOWSNR_ROWS * cell_h + (LOWSNR_ROWS - 1) * LOWSNR_ROW_GAP_IN
    region_top = TITLE_TOP_IN + TITLE_HEIGHT_IN + 0.2
    region_bottom = SLIDE_HEIGHT_IN - 0.1
    grid_top0 = region_top + max(0.0, (region_bottom - region_top - grid_h) / 2)
    return grid_left0, grid_top0, cell_h


# One grid crop to place: (unique poster token, movie path). The crop's own
# baked-in banner (id / per-frame MFI / length) is the only label.
GridCrop = Tuple[str, Path]


def _build_crop_grid_slides(
    title: str,
    crops: Sequence[GridCrop],
    poster_dir: Path,
    hold_seconds: float,
) -> List[SlideSpec]:
    """Build COLS x ROWS grid slide(s) of crop movies (no captions) under ``title``."""
    if not crops:
        return []

    per_page = LOWSNR_COLS * LOWSNR_ROWS
    grid_left0, grid_top0, cell_h = _lowsnr_grid_geometry()
    pages = [crops[i : i + per_page] for i in range(0, len(crops), per_page)]

    slides: List[SlideSpec] = []
    for page_index, page in enumerate(pages):
        page_label = "" if len(pages) == 1 else f" (page {page_index + 1}/{len(pages)})"
        textboxes: List[TextboxSpec] = [_make_title_textbox(title + page_label)]
        movies: List[MovieSpec] = []
        for idx, (token, path) in enumerate(page):
            col = idx % LOWSNR_COLS
            row = idx // LOWSNR_COLS
            left_in = grid_left0 + col * (LOWSNR_BOX_IN + LOWSNR_COL_GAP_IN)
            movie_top_in = grid_top0 + row * (cell_h + LOWSNR_ROW_GAP_IN)
            movies.append(
                _make_movie_spec_in_box(
                    path,
                    poster_dir / f"{token}_poster.png",
                    left_in,
                    movie_top_in,
                    LOWSNR_BOX_IN,
                    LOWSNR_BOX_IN,
                    hold_seconds,
                )
            )
        slides.append(SlideSpec(textboxes=tuple(textboxes), movies=tuple(movies)))
    return slides


def build_lowsnr_slide_specs(
    tracks: Sequence[LowSnrTrack],
    poster_dir: Path,
    hold_seconds: float,
) -> List[SlideSpec]:
    """Build confocal low-SNR QC grid slide(s) from selected tracks."""
    if not tracks:
        return []
    title = (
        f"Low-SNR nucleus crops — dim examples, outline only "
        f"({MODALITY}, {EXPERIMENT_DATE})"
    )
    crops: List[GridCrop] = [(f"lowsnr_FOV{t[1]}_id{t[2]}", t[4]) for t in tracks]
    return _build_crop_grid_slides(title, crops, poster_dir, hold_seconds)


# One selected LLS low-SNR crop: (mfi, wa, track_id, length, path). Ordered so a
# plain sort puts the dimmest (lowest MFI) first, like LowSnrTrack.
LlsLowSnrCrop = Tuple[int, int, int, int, Path]


def collect_lls_lowsnr(folder: Path) -> List[LlsLowSnrCrop]:
    """Return all LLS low-SNR crops (both WA) with len >= threshold, dimmest first."""
    if not folder.exists():
        print(f"[WARNING] LLS low-SNR crop folder not found; skipping ({folder})")
        return []

    records: List[LlsLowSnrCrop] = []
    for p in sorted(folder.glob("*.mp4")):
        m = LLS_LOWSNR_FILENAME_RE.match(p.name)
        if not m:
            continue
        records.append(
            (int(m["mfi"]), int(m["wa"]), int(m["id"]), int(m["len"]), p)
        )

    eligible = [r for r in records if r[3] >= LLS_LOWSNR_MIN_LEN]
    if not eligible:
        print(
            f"[WARNING] no LLS low-SNR crops with len>={LLS_LOWSNR_MIN_LEN} in {folder}"
        )
        return []
    eligible.sort(key=lambda r: (r[0], r[1], r[2]))  # MFI asc, then WA, id
    return eligible


def build_lls_lowsnr_slide_specs(
    crops: Sequence[LlsLowSnrCrop],
    poster_dir: Path,
    hold_seconds: float,
) -> List[SlideSpec]:
    """Build the LLS low-SNR QC grid slide(s) from the eligible WA1/WA2 crops."""
    if not crops:
        return []
    title = (
        f"LLS low-SNR nucleus crops — dim examples, outline only "
        f"({LLS_MODALITY}, {EXPERIMENT_DATE})"
    )
    grid_crops: List[GridCrop] = [
        (f"lls_lowsnr_WA{wa}_id{tid}", path)
        for _mfi, wa, tid, _length, path in crops
    ]
    return _build_crop_grid_slides(title, grid_crops, poster_dir, hold_seconds)


def ensure_parent_dir(output_path: Path) -> None:
    """Create the output directory if needed."""
    if output_path.parent and not output_path.parent.exists():
        output_path.parent.mkdir(parents=True, exist_ok=True)


def prepare_poster_dir(poster_dir: Path) -> Path:
    """Create a clean poster-cache directory."""
    if poster_dir.exists():
        shutil.rmtree(poster_dir, ignore_errors=True)
    poster_dir.mkdir(parents=True, exist_ok=True)
    return poster_dir


def print_plan(
    fov_movies: Dict[int, Dict[str, Path]],
    lowsnr_tracks: Sequence[LowSnrTrack],
    lls_regions: Sequence[Tuple[str, List[Path]]],
    lls_lowsnr: Sequence[LlsLowSnrCrop],
) -> None:
    """Print the FOV/track/region -> movie mapping and slide plan (used by --list)."""
    for fov in sorted(fov_movies):
        print(f"Slide: FOV {fov} (confocal)")
        for token, _caption in MOVIE_TYPES:
            print(f"    {token:24s} {fov_movies[fov][token].name}")
    if lowsnr_tracks:
        print(f"Slide(s): Low-SNR crop QC (confocal) — {len(lowsnr_tracks)} crops")
        for track in lowsnr_tracks:
            print(f"    {track[4].name}")
    for region, paths in lls_regions:
        print(f"Slide: {format_lls_region_label(region)} (LLS)")
        for (rel, _caption), path in zip(LLS_MOVIE_TYPES, paths):
            print(f"    {rel:40s} {path.name}")
    if lls_lowsnr:
        print(f"Slide(s): LLS low-SNR crop QC — {len(lls_lowsnr)} crops")
        for _mfi, _wa, _tid, _length, path in lls_lowsnr:
            print(f"    {path.name}")


def main() -> int:
    """Run the movie deck generation workflow."""
    args = parse_args()

    try:
        requested_fovs = parse_requested_fovs(args.fovs)
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    try:
        fov_movies = collect_fov_movies(MOVIE_DIR, requested_fovs=requested_fovs)
        lowsnr_tracks = collect_lowsnr_tracks(LOWSNR_DIR, args.lowsnr_count)
        lls_regions = collect_lls_regions(LLS_BASE)
        lls_lowsnr = collect_lls_lowsnr(LLS_LOWSNR_DIR)
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1

    if not fov_movies and not lowsnr_tracks and not lls_regions and not lls_lowsnr:
        print("[ERROR] No confocal FOV, low-SNR, or LLS region movies found")
        return 1

    per_page = LOWSNR_COLS * LOWSNR_ROWS
    lowsnr_pages = -(-len(lowsnr_tracks) // per_page)  # ceil
    lls_lowsnr_pages = -(-len(lls_lowsnr) // per_page)
    n_slides = len(fov_movies) + lowsnr_pages + len(lls_regions) + lls_lowsnr_pages
    print(f"Found {len(fov_movies)} confocal FOV set(s): {sorted(fov_movies)}")
    print(f"Found {len(lowsnr_tracks)} confocal low-SNR crop(s) -> {lowsnr_pages} grid slide(s)")
    print(f"Found {len(lls_regions)} LLS region(s): {[r for r, _ in lls_regions]}")
    print(f"Found {len(lls_lowsnr)} LLS low-SNR crop(s) -> {lls_lowsnr_pages} grid slide(s)")
    print(f"Will create {n_slides} slide(s)")

    if args.list:
        print_plan(fov_movies, lowsnr_tracks, lls_regions, lls_lowsnr)
        return 0

    if args.hold_last and args.hold_last > 0:
        print(
            f"Holding final frame {args.hold_last}s per movie "
            "(PowerPoint last-timepoint fix; re-encodes/compresses the clips)"
        )
    if args.loop:
        print("Looping each movie until stopped")

    output_path = resolve_output_path(args.output)
    ensure_parent_dir(output_path)

    if output_path.exists():
        backup_dir = output_path.parent / "backups"
        try:
            backups_created = backup_presentation(str(output_path), backup_base=str(backup_dir))
            print(f"Backed up existing output to: {', '.join(sorted(backups_created))}")
        except Exception as exc:
            print(f"[WARNING] Could not back up existing output: {exc}")

    poster_dir = prepare_poster_dir(DEFAULT_POSTER_DIR)

    try:
        # Modality-grouped: all confocal (FOV overviews then confocal crops),
        # then all LLS (region overviews then LLS crops). Within LLS the region
        # order (whole-FOV first) is set by LLS_REGIONS.
        slide_specs = build_slide_specs(fov_movies, poster_dir, args.hold_last)
        slide_specs += build_lowsnr_slide_specs(lowsnr_tracks, poster_dir, args.hold_last)
        slide_specs += build_lls_slide_specs(lls_regions, poster_dir, args.hold_last)
        slide_specs += build_lls_lowsnr_slide_specs(lls_lowsnr, poster_dir, args.hold_last)
        rewritten = build_movie_deck_via_com(
            slide_specs,
            str(output_path),
            slide_width_pt=_to_points(SLIDE_WIDTH_IN),
            slide_height_pt=_to_points(SLIDE_HEIGHT_IN),
            loop=args.loop,
        )
        print(f"Set {rewritten} slide(s) to autoplay")
    except PermissionError:
        print(
            f"[ERROR] Permission denied when saving {output_path}. "
            "Make sure the PowerPoint file is closed."
        )
        return 1
    except Exception as exc:
        print(f"[ERROR] Failed while building/saving slides: {exc}")
        return 1

    print(f"[SUCCESS] Saved presentation: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
