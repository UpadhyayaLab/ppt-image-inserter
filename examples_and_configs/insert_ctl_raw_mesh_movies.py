#!/usr/bin/env python
"""
Build a PowerPoint deck of CTL raw-with-meshes movies (four slides per cell).

For each cell number N present in all eight required folders within an
experiment, this script creates:

1. A "panel_1x3 xz" slide pairing xz on top with xz_rev on bottom
2. A "panel_1x3 yz" slide pairing yz on top with yz_rev on bottom
3. A "panel_raw_cent xz" slide pairing xz on top with xz_rev on bottom
4. A "panel_raw_cent yz" slide pairing yz on top with yz_rev on bottom

Three CTL experiments are combined into a single deck:
20210928 OTI CTL Activated, 20211221 OTI CTL Activated, and 20220614 OT1
CTLs antiCD3. Movies are inserted via PowerPoint COM (Windows-only); after
saving, the timing-tree trigger on disk is rewritten so videos autoplay in
slideshow mode.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
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


CTL_BASE = Path("F:/FF/nucleus_live_cell/ctl_nucleus_centrosome")

EXPT_0928_BASE = CTL_BASE / "20210928_OTI_CTL_Activated/combined/frames/prog_live_cells/raw_with_meshes_movies"
EXPT_1221_BASE = CTL_BASE / "20211221_OTI_CTL_Activated_already_cropped/combined/channels/frames/prog_live_cells/raw_with_meshes_movies"
EXPT_0614_BASE = CTL_BASE / "20220614_OT1_CTLs_Centrin_SiR-DNA/antiCD3/Cells/channels/frames/prog_live_cells/raw_with_meshes_movies"

DEFAULT_OUTPUT = Path("K:/FF/PPT/PPT_autogeneration/Live Cells/CTL raw, mesh movies.pptx")
DEFAULT_POSTER_DIR = Path(__file__).resolve().parent / "_movie_posters_tmp_ctl_raw_mesh"

PER_EXPT_OUTPUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Live Cells")

CELL_PATTERN = re.compile(r"^Cell(?P<cell>\d+)_.*\.mp4$", re.IGNORECASE)

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


@dataclass(frozen=True)
class CellMovieSet:
    """All eight movie paths needed to build the four slides for one cell."""

    experiment_label: str
    cell_number: int
    panel_1x3_xz: Path
    panel_1x3_xz_rev: Path
    panel_1x3_yz: Path
    panel_1x3_yz_rev: Path
    raw_cent_xz: Path
    raw_cent_xz_rev: Path
    raw_cent_yz: Path
    raw_cent_yz_rev: Path


@dataclass(frozen=True)
class ExperimentSource:
    """Folder configuration for one experiment to add into the deck."""

    experiment_label: str
    panel_1x3_xz: Path
    panel_1x3_xz_rev: Path
    panel_1x3_yz: Path
    panel_1x3_yz_rev: Path
    raw_cent_xz: Path
    raw_cent_xz_rev: Path
    raw_cent_yz: Path
    raw_cent_yz_rev: Path


DEFAULT_EXPERIMENT_SOURCES: Tuple[ExperimentSource, ...] = (
    ExperimentSource(
        experiment_label="20210928 OTI CTL Activated",
        panel_1x3_xz=EXPT_0928_BASE / "panel_1x3/xz_white_grey_blue",
        panel_1x3_xz_rev=EXPT_0928_BASE / "panel_1x3/xz_rev_white_grey_blue",
        panel_1x3_yz=EXPT_0928_BASE / "panel_1x3/yz_white_grey_blue",
        panel_1x3_yz_rev=EXPT_0928_BASE / "panel_1x3/yz_rev_white_grey_blue",
        raw_cent_xz=EXPT_0928_BASE / "panel_raw_cent/xz",
        raw_cent_xz_rev=EXPT_0928_BASE / "panel_raw_cent/xz_rev",
        raw_cent_yz=EXPT_0928_BASE / "panel_raw_cent/yz",
        raw_cent_yz_rev=EXPT_0928_BASE / "panel_raw_cent/yz_rev",
    ),
    ExperimentSource(
        experiment_label="20211221 OTI CTL Activated",
        panel_1x3_xz=EXPT_1221_BASE / "panel_1x3/xz_white_grey_blue",
        panel_1x3_xz_rev=EXPT_1221_BASE / "panel_1x3/xz_rev_white_grey_blue",
        panel_1x3_yz=EXPT_1221_BASE / "panel_1x3/yz_white_grey_blue",
        panel_1x3_yz_rev=EXPT_1221_BASE / "panel_1x3/yz_rev_white_grey_blue",
        raw_cent_xz=EXPT_1221_BASE / "panel_raw_cent/xz",
        raw_cent_xz_rev=EXPT_1221_BASE / "panel_raw_cent/xz_rev",
        raw_cent_yz=EXPT_1221_BASE / "panel_raw_cent/yz",
        raw_cent_yz_rev=EXPT_1221_BASE / "panel_raw_cent/yz_rev",
    ),
    ExperimentSource(
        experiment_label="20220614 OT1 CTLs antiCD3",
        panel_1x3_xz=EXPT_0614_BASE / "panel_1x3/xz_white_grey_blue",
        panel_1x3_xz_rev=EXPT_0614_BASE / "panel_1x3/xz_rev_white_grey_blue",
        panel_1x3_yz=EXPT_0614_BASE / "panel_1x3/yz_white_grey_blue",
        panel_1x3_yz_rev=EXPT_0614_BASE / "panel_1x3/yz_rev_white_grey_blue",
        raw_cent_xz=EXPT_0614_BASE / "panel_raw_cent/xz",
        raw_cent_xz_rev=EXPT_0614_BASE / "panel_raw_cent/xz_rev",
        raw_cent_yz=EXPT_0614_BASE / "panel_raw_cent/yz",
        raw_cent_yz_rev=EXPT_0614_BASE / "panel_raw_cent/yz_rev",
    ),
)


FOLDER_FIELDS: Tuple[str, ...] = (
    "panel_1x3_xz",
    "panel_1x3_xz_rev",
    "panel_1x3_yz",
    "panel_1x3_yz_rev",
    "raw_cent_xz",
    "raw_cent_xz_rev",
    "raw_cent_yz",
    "raw_cent_yz_rev",
)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Create a PowerPoint deck of CTL raw-with-meshes movies, four "
            "slides per cell pairing each direction with its _rev variant: "
            "panel_1x3 xz pair, panel_1x3 yz pair, panel_raw_cent xz pair, "
            "panel_raw_cent yz pair. Combines 20210928, 20211221, and "
            "20220614 CTL experiments."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=(
            "Output PowerPoint path. If omitted, defaults to the combined deck "
            f"({DEFAULT_OUTPUT}) when --experiment is not used, or to a "
            "per-experiment file in K:/FF/PPT/PPT_autogeneration/Live Cells/ "
            "when --experiment is set."
        ),
    )
    parser.add_argument(
        "--experiment",
        type=str,
        default=None,
        choices=["20210928", "20211221", "20220614"],
        help=(
            "Build a deck for just one experiment (by date prefix). "
            "Useful when the combined deck is too large for PowerPoint COM."
        ),
    )
    parser.add_argument(
        "--cells",
        type=str,
        default="",
        help="Optional comma-separated list of cell numbers to include, e.g. 1,2,5,10",
    )
    return parser.parse_args()


def select_experiment_sources(experiment: Optional[str]) -> Tuple[ExperimentSource, ...]:
    """Filter DEFAULT_EXPERIMENT_SOURCES to a single experiment if requested."""
    if experiment is None:
        return DEFAULT_EXPERIMENT_SOURCES
    matched = tuple(
        source for source in DEFAULT_EXPERIMENT_SOURCES
        if experiment in source.experiment_label
    )
    if not matched:
        raise ValueError(f"No experiment matches: {experiment}")
    return matched


def resolve_output_path(
    output_arg: Optional[Path], experiment: Optional[str]
) -> Path:
    """Pick the output path based on --output / --experiment."""
    if output_arg is not None:
        return output_arg.resolve()
    if experiment is None:
        return DEFAULT_OUTPUT.resolve()
    return (PER_EXPT_OUTPUT_DIR / f"CTL raw, mesh movies ({experiment}).pptx").resolve()


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


def validate_folder(folder: Path, label: str) -> None:
    """Ensure a required movie folder exists."""
    if not folder.exists():
        raise FileNotFoundError(f"{label} folder not found: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"{label} is not a folder: {folder}")


def collect_movies(folder: Path, label: str) -> Dict[int, Path]:
    """Collect CellN movie files from a folder."""
    validate_folder(folder, label)

    movies: Dict[int, Path] = {}
    for movie_path in sorted(folder.glob("*.mp4")):
        match = CELL_PATTERN.match(movie_path.name)
        if not match:
            print(f"[WARNING] Skipping unmatched filename in {label}: {movie_path.name}")
            continue

        cell_number = int(match.group("cell"))
        if cell_number in movies:
            raise ValueError(
                f"Duplicate Cell{cell_number} movie found in {label}: "
                f"{movies[cell_number].name} and {movie_path.name}"
            )
        movies[cell_number] = movie_path

    if not movies:
        raise ValueError(f"No CellN .mp4 files found in {label}: {folder}")

    return movies


def build_complete_cell_sets(
    experiment_label: str,
    folder_maps: Dict[str, Dict[int, Path]],
    requested_cells: Optional[Set[int]] = None,
) -> List[CellMovieSet]:
    """Build complete eight-movie records for cells present in all required folders."""
    available_in_all: Set[int] = set(folder_maps[FOLDER_FIELDS[0]])
    for field in FOLDER_FIELDS[1:]:
        available_in_all &= set(folder_maps[field])

    if requested_cells is not None:
        available_in_all &= requested_cells

    cell_sets = [
        CellMovieSet(
            experiment_label=experiment_label,
            cell_number=cell_number,
            **{field: folder_maps[field][cell_number] for field in FOLDER_FIELDS},
        )
        for cell_number in sorted(available_in_all)
    ]

    return cell_sets


def report_missing_cells(
    experiment_label: str,
    label_to_movies: Dict[str, Dict[int, Path]],
    requested_cells: Optional[Set[int]] = None,
) -> None:
    """Print warnings for cells missing one or more required movie files."""
    all_cells = set()
    for movies in label_to_movies.values():
        all_cells.update(movies)

    if requested_cells is not None:
        all_cells &= requested_cells

    for cell_number in sorted(all_cells):
        missing_from = [
            label for label, movies in label_to_movies.items() if cell_number not in movies
        ]
        if missing_from:
            missing_text = ", ".join(missing_from)
            print(
                f"[WARNING] {experiment_label}: Cell{cell_number} skipped; "
                f"missing movie in: {missing_text}"
            )

    if requested_cells is not None:
        undiscovered = requested_cells - all_cells
        for cell_number in sorted(undiscovered):
            print(
                f"[WARNING] {experiment_label}: Cell{cell_number} skipped; "
                "not found in any source folder"
            )


def collect_cell_sets_for_experiment(
    source: ExperimentSource,
    requested_cells: Optional[Set[int]] = None,
) -> List[CellMovieSet]:
    """Collect complete cell movie sets for one experiment."""
    folder_maps: Dict[str, Dict[int, Path]] = {
        field: collect_movies(
            getattr(source, field),
            f"{source.experiment_label} {field}",
        )
        for field in FOLDER_FIELDS
    }

    report_missing_cells(
        source.experiment_label,
        folder_maps,
        requested_cells=requested_cells,
    )

    return build_complete_cell_sets(
        source.experiment_label,
        folder_maps,
        requested_cells=requested_cells,
    )


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


def format_experiment_suffix(experiment_label: str) -> str:
    """Strip the verbose experiment description, keep the date as `YYYYMMDD expt`."""
    match = re.search(r"\b(\d{8})\b", experiment_label)
    if match:
        return f"{match.group(1)} expt"
    return experiment_label


def format_panel_1x3_title(
    cell_number: int, experiment_label: str, direction: str
) -> str:
    """Return the title for a panel_1x3 slide."""
    return (
        f"Raw, invag depth, min curvature ({direction}): "
        f"Cell {cell_number} ({format_experiment_suffix(experiment_label)})"
    )


def format_raw_cent_title(
    cell_number: int, experiment_label: str, direction: str
) -> str:
    """Return the title for a panel_raw_cent slide."""
    return (
        f"Raw, region near centrosome ({direction}): "
        f"Cell {cell_number} ({format_experiment_suffix(experiment_label)})"
    )


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
) -> MovieSpec:
    frame_width, frame_height = extract_first_frame(movie_path, poster_path)
    left_in, top_in, width_in, height_in = fit_within_box(
        frame_width,
        frame_height,
        MOVIE_LEFT_IN,
        box_top_in,
        MOVIE_BOX_WIDTH_IN,
        MOVIE_BOX_HEIGHT_IN,
    )
    return MovieSpec(
        movie_path=str(movie_path.resolve()),
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
            _make_movie_spec(top_movie, top_poster, TOP_MOVIE_TOP_IN),
            _make_movie_spec(bottom_movie, bottom_poster, BOTTOM_MOVIE_TOP_IN),
        ),
    )


def build_slide_specs(
    cell_sets: Sequence[CellMovieSet],
    poster_dir: Path,
) -> List[SlideSpec]:
    """Build four SlideSpecs per cell, pairing each direction with its _rev variant."""
    slide_specs: List[SlideSpec] = []
    for cell_set in cell_sets:
        slug = f"cell{cell_set.cell_number}_{cell_set.experiment_label.replace(' ', '_')}"

        slide_specs.append(
            _build_movie_pair_slide(
                format_panel_1x3_title(
                    cell_set.cell_number, cell_set.experiment_label, "xz"
                ),
                cell_set.panel_1x3_xz,
                cell_set.panel_1x3_xz_rev,
                "xz",
                "xz_rev",
                poster_dir,
                f"{slug}_panel1x3_xz",
            )
        )
        slide_specs.append(
            _build_movie_pair_slide(
                format_panel_1x3_title(
                    cell_set.cell_number, cell_set.experiment_label, "yz"
                ),
                cell_set.panel_1x3_yz,
                cell_set.panel_1x3_yz_rev,
                "yz",
                "yz_rev",
                poster_dir,
                f"{slug}_panel1x3_yz",
            )
        )
        slide_specs.append(
            _build_movie_pair_slide(
                format_raw_cent_title(
                    cell_set.cell_number, cell_set.experiment_label, "xz"
                ),
                cell_set.raw_cent_xz,
                cell_set.raw_cent_xz_rev,
                "xz",
                "xz_rev",
                poster_dir,
                f"{slug}_rawcent_xz",
            )
        )
        slide_specs.append(
            _build_movie_pair_slide(
                format_raw_cent_title(
                    cell_set.cell_number, cell_set.experiment_label, "yz"
                ),
                cell_set.raw_cent_yz,
                cell_set.raw_cent_yz_rev,
                "yz",
                "yz_rev",
                poster_dir,
                f"{slug}_rawcent_yz",
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
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    try:
        sources = select_experiment_sources(args.experiment)
    except ValueError as exc:
        print(f"[ERROR] {exc}")
        return 1

    try:
        cell_sets: List[CellMovieSet] = []
        for source in sources:
            experiment_cell_sets = collect_cell_sets_for_experiment(
                source,
                requested_cells=requested_cells,
            )
            print(f"{source.experiment_label}: found {len(experiment_cell_sets)} complete cell set(s)")
            cell_sets.extend(experiment_cell_sets)
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1

    if not cell_sets:
        print("[ERROR] No cells have a complete set of all eight required movies")
        return 1

    print(f"Found {len(cell_sets)} complete cell set(s) across all experiments")
    print(f"Will create {len(cell_sets) * 4} slide(s)")

    output_path = resolve_output_path(args.output, args.experiment)
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
        slide_specs = build_slide_specs(cell_sets, poster_dir)
        rewritten = build_movie_deck_via_com(
            slide_specs,
            str(output_path),
            slide_width_pt=_to_points(SLIDE_WIDTH_IN),
            slide_height_pt=_to_points(SLIDE_HEIGHT_IN),
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
