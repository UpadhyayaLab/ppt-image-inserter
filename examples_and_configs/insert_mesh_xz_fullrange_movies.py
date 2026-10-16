#!/usr/bin/env python
"""
Build a PowerPoint deck of the full-frame-range raw-with-meshes movies.

For each cell, three movie kinds are shown, two slides per kind (an xz pair then
a yz pair), each pairing a direction on top with its reverse on bottom:

1. ``panel_1x3``              -> "Raw, invag depth, min curvature"
2. ``panel_raw_invag_vectors`` -> "Raw + invagination vectors"
3. ``cent_regions``           -> "Centrosome regions"

These are the movies compiled to cover every frame on disk (not just the 12-min
window), re-rendered for both directions of xz and yz. All three kinds are
required: a cell appears only if it has all four views in all three folders.

Scope (``--experiment``, one deck per experiment): the 032022 (NaBu800 Control)
and 04142022 (GFP-Centrin Control) Jurkat datasets, and the 20220614 (OT1 CTLs
antiCD3) CTL dataset. Movies are inserted via PowerPoint COM (Windows-only);
after saving, each movie is set to autoplay and (by default) loop. The Jurkat
movies are 2 fps so their final frame is held ~1s (PowerPoint drops the last
frame otherwise); the CTL movies are 4 fps so the hold is off by default (it
would ~2x the deck and looping masks the dropped 0.25s frame).
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from dataclasses import dataclass
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


JURKAT_BASE = Path("F:/FF/nucleus_live_cell/jurkat_nucleus_centrosome")
CTL_BASE = Path("F:/FF/nucleus_live_cell/ctl_nucleus_centrosome")

PER_EXPT_OUTPUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Live Cells")


@dataclass(frozen=True)
class ExperimentSpec:
    """One live experiment whose xz raw-with-meshes movies feed a deck.

    ``movies_root`` is the ``raw_with_meshes_movies`` folder that holds the
    per-kind subfolders (panel_1x3, panel_raw_invag_vectors, cent_regions).
    ``cell_type`` prefixes the output deck name; ``hold_last_default`` is the
    per-experiment default for the last-frame hold (on for 2 fps Jurkat, off for
    4 fps CTL) and is overridden by an explicit ``--hold-last``.
    """

    key: str
    label: str
    cell_type: str
    movies_root: Path
    hold_last_default: float


EXPERIMENTS: Dict[str, ExperimentSpec] = {
    "032022": ExperimentSpec(
        key="032022",
        label="032022 expt",
        cell_type="Jurkat",
        movies_root=JURKAT_BASE
        / "NaBu800 Experiments/Control_3D_CD3/all_cells_together/prog_live_cells/raw_with_meshes_movies",
        hold_last_default=1.0,
    ),
    "04142022": ExperimentSpec(
        key="04142022",
        label="04142022 expt",
        cell_type="Jurkat",
        movies_root=JURKAT_BASE
        / "GFP-Centrin_SiR-DNA/Control/cells/all_cells_together/prog_live_cells/raw_with_meshes_movies",
        hold_last_default=1.0,
    ),
    "20220614": ExperimentSpec(
        key="20220614",
        label="20220614 CTL antiCD3",
        cell_type="CTL",
        movies_root=CTL_BASE
        / "20220614_OT1_CTLs_Centrin_SiR-DNA/antiCD3/Cells/channels/frames/prog_live_cells/raw_with_meshes_movies",
        hold_last_default=0.0,
    ),
}

# Poster frames and padded movie copies are transient per-run scratch; keep them
# out of the repo and out of the deck output folder (wiped and recreated each run).
DEFAULT_POSTER_DIR = Path(tempfile.gettempdir()) / "ppt_mesh_xz_posters"

# The full-range movies were re-rendered for both directions of xz and yz. Each
# kind becomes two slides per cell: an xz pair and a yz pair (see VIEW_PAIRS).
VIEWS: Tuple[str, ...] = ("xz", "xz_rev", "yz", "yz_rev")
VIEW_PAIRS: Tuple[Tuple[str, str], ...] = (("xz", "xz_rev"), ("yz", "yz_rev"))


def _kind_patterns(token: str, colormap_suffix: str) -> Dict[str, "re.Pattern[str]"]:
    """Per-view (xz, xz_rev) filename regexes for one movie kind.

    Each pattern is fully anchored (``$``) so ``xz`` never matches an ``xz_rev``
    filename. ``token`` is the kind's filename segment after ``rwm_``;
    ``colormap_suffix`` is ``_white_grey_blue`` for panel_1x3 and empty for the
    other two kinds (which carry no colormap token).
    """
    suffix = re.escape(colormap_suffix)
    return {
        view: re.compile(
            rf"^Cell(?P<cell>\d+)_rwm_{re.escape(token)}_{view}{suffix}\.mp4$",
            re.IGNORECASE,
        )
        for view in VIEWS
    }


@dataclass(frozen=True)
class KindSpec:
    """One movie kind shown as its own slide (xz over xz_rev). All required."""

    key: str
    movie_subdir: str
    title: str
    view_patterns: Dict[str, "re.Pattern[str]"]


KINDS: Tuple[KindSpec, ...] = (
    KindSpec(
        key="panel_1x3",
        movie_subdir="panel_1x3",
        title="Raw, invag depth, min curvature",
        view_patterns=_kind_patterns("panel_1x3", "_white_grey_blue"),
    ),
    KindSpec(
        key="invag_vectors",
        movie_subdir="panel_raw_invag_vectors",
        title="Raw + invagination vectors",
        view_patterns=_kind_patterns("panel_raw_invag_vectors", ""),
    ),
    KindSpec(
        key="cent_regions",
        movie_subdir="cent_regions",
        title="Centrosome regions",
        view_patterns=_kind_patterns("cent_regions", ""),
    ),
)


SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
TITLE_LEFT_IN = 0.45
TITLE_TOP_IN = 0.12
TITLE_WIDTH_IN = 12.4
TITLE_HEIGHT_IN = 0.35
TITLE_FONT_SIZE_PT = 22.0
LABEL_LEFT_IN = 0.7
LABEL_WIDTH_IN = 2.0
LABEL_HEIGHT_IN = 0.18
TOP_LABEL_TOP_IN = 0.55
BOTTOM_LABEL_TOP_IN = 3.79
LABEL_FONT_SIZE_PT = 11.0
MOVIE_LEFT_IN = 0.7
MOVIE_BOX_WIDTH_IN = 11.9
TOP_MOVIE_TOP_IN = 0.78
BOTTOM_MOVIE_TOP_IN = 4.02
MOVIE_BOX_HEIGHT_IN = 2.62
POINTS_PER_INCH = 72.0


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Create a PowerPoint deck of Jurkat 1x4 raw-with-meshes movies "
            "(raw | region | depth | min-curv), two slides per cell per variant "
            "pairing each direction with its _rev variant. Builds the raw "
            "variant and, when its movies exist, the centrosome-dot variant, "
            "for the chosen Jurkat experiment (--experiment)."
        )
    )
    parser.add_argument(
        "--experiment",
        type=str,
        required=True,
        choices=sorted(EXPERIMENTS),
        help="Which experiment to build (one deck per experiment).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=(
            "Output PowerPoint path. Defaults to "
            f"{PER_EXPT_OUTPUT_DIR}/<cell type> mesh xz movies, full range (<experiment>).pptx."
        ),
    )
    parser.add_argument(
        "--cells",
        type=str,
        default="",
        help="Optional comma-separated list of cell numbers to include, e.g. 1,2,5,10",
    )
    parser.add_argument(
        "--since",
        type=str,
        default=None,
        help=(
            "Only include movies modified on/after this date (YYYY-MM-DD). Use it "
            "to pick up just a fresh compile, since old movies with the same names "
            "(earlier compiles / 12-min window) can linger in the same folders."
        ),
    )
    parser.add_argument(
        "--hold-last",
        type=float,
        default=None,
        help=(
            "Seconds to hold (clone) the final frame of each movie so PowerPoint "
            "does not drop the last timepoint on playback. 0 disables. If omitted, "
            "uses the experiment default (1.0 for 2 fps Jurkat, 0 for 4 fps CTL)."
        ),
    )
    parser.add_argument(
        "--loop",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Loop each movie until stopped in slideshow (default: on; --no-loop to disable).",
    )
    return parser.parse_args()


def resolve_output_path(output_arg: Optional[Path], experiment: ExperimentSpec) -> Path:
    """Pick the output path based on --output / --experiment."""
    if output_arg is not None:
        return output_arg.resolve()
    return (
        PER_EXPT_OUTPUT_DIR
        / f"{experiment.cell_type} mesh xz movies, full range ({experiment.key}).pptx"
    ).resolve()


def parse_requested_cells(cells_arg: str) -> Optional[Set[int]]:
    """Parse optional comma-separated cell numbers."""
    if not cells_arg.strip():
        return None

    requested: Set[int] = set()
    for token in cells_arg.split(","):
        token = token.strip()
        if not token:
            continue
        try:
            requested.add(int(token))
        except ValueError as exc:
            raise ValueError(f"Invalid cell number in --cells: {token}") from exc

    if not requested:
        raise ValueError("--cells was provided but no valid cell numbers were found")

    return requested


def parse_since(since_arg: Optional[str]) -> Optional[float]:
    """Parse a YYYY-MM-DD ``--since`` date to a min mtime (local start-of-day)."""
    if not since_arg or not since_arg.strip():
        return None
    try:
        return datetime.strptime(since_arg.strip(), "%Y-%m-%d").timestamp()
    except ValueError as exc:
        raise ValueError(f"Invalid --since date (expected YYYY-MM-DD): {since_arg}") from exc


def validate_folder(folder: Path, label: str) -> None:
    """Ensure a required movie folder exists."""
    if not folder.exists():
        raise FileNotFoundError(f"{label} folder not found: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"{label} is not a folder: {folder}")


def collect_movies_by_variant(
    folder: Path,
    variant_patterns: Dict[str, "re.Pattern[str]"],
    label: str,
    min_mtime: Optional[float] = None,
) -> Dict[str, Dict[int, Path]]:
    """Scan a parent folder, classify CellN movies into variant -> cell -> path.

    Files in ``folder`` whose names match any pattern in ``variant_patterns``
    are added to the matching variant bucket. Files that do not match any
    pattern (e.g. alternate color/clim variants) are silently ignored, as are
    files older than ``min_mtime`` (when given). Per-cell duplicates within one
    variant raise ``ValueError``.
    """
    validate_folder(folder, label)

    by_variant: Dict[str, Dict[int, Path]] = {
        variant: {} for variant in variant_patterns
    }
    for movie_path in sorted(folder.glob("*.mp4")):
        if min_mtime is not None and os.path.getmtime(movie_path) < min_mtime:
            continue
        for variant, pattern in variant_patterns.items():
            match = pattern.match(movie_path.name)
            if not match:
                continue
            cell_number = int(match.group("cell"))
            existing = by_variant[variant].get(cell_number)
            if existing is not None:
                raise ValueError(
                    f"Duplicate Cell{cell_number} {variant} movie in {label}: "
                    f"{existing.name} and {movie_path.name}"
                )
            by_variant[variant][cell_number] = movie_path
            break

    for variant, mapping in by_variant.items():
        if not mapping:
            print(
                f"[WARNING] {label}: no CellN .mp4 files matched variant "
                f"{variant!r} in {folder}"
            )

    return by_variant


def report_missing_cells(
    experiment_label: str,
    view_to_movies: Dict[str, Dict[int, Path]],
    requested_cells: Optional[Set[int]] = None,
) -> None:
    """Print warnings for cells missing one or more of a panel's views."""
    all_cells: Set[int] = set()
    for movies in view_to_movies.values():
        all_cells.update(movies)

    if requested_cells is not None:
        all_cells &= requested_cells

    for cell_number in sorted(all_cells):
        missing_from = [
            view for view, movies in view_to_movies.items() if cell_number not in movies
        ]
        if missing_from:
            missing_text = ", ".join(missing_from)
            print(
                f"[WARNING] {experiment_label}: Cell{cell_number} skipped; "
                f"missing view(s): {missing_text}"
            )


def collect_kind_cells(
    kind: KindSpec,
    movies_root: Path,
    experiment_label: str,
    requested_cells: Optional[Set[int]] = None,
    min_mtime: Optional[float] = None,
) -> Dict[int, Dict[str, Path]]:
    """Return ``{cell -> {view -> movie path}}`` for cells complete in a kind.

    Complete = has both xz and xz_rev for this kind. The kind's folder is
    required (raises if missing). Matching by the kind's view patterns naturally
    ignores other views (yz/yz_rev) and other colormaps; ``min_mtime`` drops
    movies older than a ``--since`` cutoff.
    """
    folder = movies_root / kind.movie_subdir
    by_view = collect_movies_by_variant(
        folder, kind.view_patterns, f"{experiment_label} {kind.key}", min_mtime
    )

    report_missing_cells(f"{experiment_label} {kind.key}", by_view, requested_cells)

    complete: Set[int] = set(by_view[VIEWS[0]])
    for view in VIEWS[1:]:
        complete &= set(by_view[view])
    if requested_cells is not None:
        complete &= requested_cells

    return {
        cell_number: {view: by_view[view][cell_number] for view in VIEWS}
        for cell_number in sorted(complete)
    }


def pad_movie_hold_last(src: Path, dst: Path, hold_seconds: float) -> Path:
    """Re-encode ``src`` to ``dst``, holding the final frame for ``hold_seconds``.

    PowerPoint ends video playback slightly before the media's true end, so at
    low frame rates the final timepoint is never shown on playback. Cloning the
    last frame for ~1s pushes the real final frame safely before that cutoff.

    Critically, the re-encode must NOT introduce B-frames: PowerPoint's decoder
    drops reordered (B-frame) tails, which makes it swallow *more* than one final
    frame. The source movies are Main-profile with no B-frames (IPIPIP...); we
    match that with ``-bf 0`` and a matching keyframe interval so playback drops
    at most the true last frame (now a clone). The source file is never
    modified. Requires the bundled ffmpeg from imageio-ffmpeg.
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


def format_title(
    kind: KindSpec, cell_number: int, direction: str, experiment_label: str
) -> str:
    """Return the slide title for a kind/cell/direction (slide holds fwd/rev)."""
    return f"{kind.title} ({direction}): Cell {cell_number} ({experiment_label})"


def _to_points(inches: float) -> float:
    """Convert inches to PowerPoint points."""
    return inches * POINTS_PER_INCH


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


def _make_region_label(text: str, top_in: float) -> TextboxSpec:
    return TextboxSpec(
        text=text,
        left_pt=_to_points(LABEL_LEFT_IN),
        top_pt=_to_points(top_in),
        width_pt=_to_points(LABEL_WIDTH_IN),
        height_pt=_to_points(LABEL_HEIGHT_IN),
        font_size_pt=LABEL_FONT_SIZE_PT,
        bold=True,
        align="left",
        font_name="Arial",
    )


def _make_movie_spec(
    movie_path: Path,
    poster_path: Path,
    box_top_in: float,
    hold_seconds: float,
) -> MovieSpec:
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
        MOVIE_LEFT_IN,
        box_top_in,
        MOVIE_BOX_WIDTH_IN,
        MOVIE_BOX_HEIGHT_IN,
    )
    return MovieSpec(
        movie_path=str(movie_to_insert.resolve()),
        poster_path=str(poster_path.resolve()),
        left_pt=_to_points(left_in),
        top_pt=_to_points(top_in),
        width_pt=_to_points(width_in),
        height_pt=_to_points(height_in),
    )


def _build_movie_pair_slide(
    title: str,
    top_movie: Path,
    bottom_movie: Path,
    top_label: str,
    bottom_label: str,
    poster_dir: Path,
    suffix: str,
    hold_seconds: float,
) -> SlideSpec:
    """Build one slide with title + top/bottom labels + top/bottom movies."""
    top_poster = poster_dir / f"{top_movie.stem}_{suffix}_top_poster.png"
    bottom_poster = poster_dir / f"{bottom_movie.stem}_{suffix}_bottom_poster.png"
    return SlideSpec(
        textboxes=(
            _make_title_textbox(title),
            _make_region_label(top_label, TOP_LABEL_TOP_IN),
            _make_region_label(bottom_label, BOTTOM_LABEL_TOP_IN),
        ),
        movies=(
            _make_movie_spec(top_movie, top_poster, TOP_MOVIE_TOP_IN, hold_seconds),
            _make_movie_spec(bottom_movie, bottom_poster, BOTTOM_MOVIE_TOP_IN, hold_seconds),
        ),
    )


def common_cells_across_kinds(
    kind_cells: Dict[str, Dict[int, Dict[str, Path]]]
) -> List[int]:
    """Cells complete (xz + xz_rev) in EVERY kind, numerically ordered."""
    cell_sets = [set(kind_cells[kind.key]) for kind in KINDS]
    if not cell_sets:
        return []
    return sorted(set.intersection(*cell_sets))


def build_slide_specs(
    kind_cells: Dict[str, Dict[int, Dict[str, Path]]],
    poster_dir: Path,
    hold_seconds: float,
    experiment_label: str,
) -> List[SlideSpec]:
    """Build two SlideSpecs per kind per cell (xz pair, then yz pair).

    Only cells complete in ALL kinds are included; cells are ordered numerically
    and, within each cell, the kinds are emitted in ``KINDS`` order, each
    contributing an xz-pair slide followed by a yz-pair slide.
    """
    slide_specs: List[SlideSpec] = []
    for cell_number in common_cells_across_kinds(kind_cells):
        for kind in KINDS:
            views = kind_cells[kind.key][cell_number]
            for top_view, bottom_view in VIEW_PAIRS:
                direction = "xz" if top_view == "xz" else "yz"
                slide_specs.append(
                    _build_movie_pair_slide(
                        format_title(kind, cell_number, direction, experiment_label),
                        views[top_view],
                        views[bottom_view],
                        top_view,
                        bottom_view,
                        poster_dir,
                        f"cell{cell_number}_{kind.key}_{direction}",
                        hold_seconds,
                    )
                )
    return slide_specs


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


def main() -> int:
    """Run the movie deck generation workflow."""
    args = parse_args()

    try:
        requested_cells = parse_requested_cells(args.cells)
        min_mtime = parse_since(args.since)
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    experiment = EXPERIMENTS[args.experiment]
    hold_seconds = (
        experiment.hold_last_default if args.hold_last is None else args.hold_last
    )
    print(f"Experiment: {experiment.label} ({experiment.movies_root})")
    if args.since:
        print(f"Only movies modified on/after {args.since}")

    try:
        kind_cells: Dict[str, Dict[int, Dict[str, Path]]] = {}
        for kind in KINDS:
            cells = collect_kind_cells(
                kind,
                experiment.movies_root,
                experiment.label,
                requested_cells=requested_cells,
                min_mtime=min_mtime,
            )
            print(f"{kind.key}: found {len(cells)} complete cell set(s)")
            kind_cells[kind.key] = cells
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1

    common_cells = common_cells_across_kinds(kind_cells)
    if not common_cells:
        print("[ERROR] No cells have xz + xz_rev in all three kinds")
        return 1

    union_cells: Set[int] = set()
    for cells in kind_cells.values():
        union_cells.update(cells)
    dropped = sorted(union_cells - set(common_cells))
    if dropped:
        print(
            f"[WARNING] {len(dropped)} cell(s) not complete in all 3 kinds, "
            f"skipped: {dropped}"
        )

    n_slides = len(common_cells) * len(KINDS) * len(VIEW_PAIRS)
    print(f"Found {len(common_cells)} cell(s) complete in all {len(KINDS)} kinds")
    print(f"Will create {n_slides} slide(s)")
    if hold_seconds and hold_seconds > 0:
        print(
            f"Holding final frame {hold_seconds}s per movie "
            "(PowerPoint last-timepoint fix)"
        )
    if args.loop:
        print("Looping each movie until stopped")

    output_path = resolve_output_path(args.output, experiment)
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
        slide_specs = build_slide_specs(
            kind_cells, poster_dir, hold_seconds, experiment.label
        )
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
