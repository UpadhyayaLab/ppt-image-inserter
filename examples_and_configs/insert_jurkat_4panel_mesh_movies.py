#!/usr/bin/env python
"""
Build a PowerPoint deck of Jurkat 1x4 ("4-panel") raw-with-meshes movies.

The 4-panel montage is ``raw | region | depth | min-curv``. Two variants of it
are supported, each rendered as its own set of slides:

- ``raw`` (``panel_raw_region_depth_curv``): the plain montage.
- ``dot`` (``panel_raw_region_depth_curv_dot``): the same montage with the
  centrosome drawn as a dot on the mesh tiles.

For each cell present (all four views) within a variant, two slides are built,
each pairing a direction with its reverse:

1. An "xz" slide pairing xz on top with xz_rev on bottom
2. A "yz" slide pairing yz on top with yz_rev on bottom

The two variants are independent: the ``raw`` folder is required, but the
``dot`` folder is optional. When the dot movies have not been rendered yet the
script warns and skips them; once those movies exist in
``raw_with_meshes_movies/panel_raw_region_depth_curv_dot/`` a re-run adds the
dot slides automatically (no code change needed).

Scope: the 032022 (NaBu800 Control_3D_CD3) Jurkat dataset. Movies are inserted
via PowerPoint COM (Windows-only); after saving, the timing-tree trigger on disk
is rewritten so videos autoplay in slideshow mode.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
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


JURKAT_BASE = Path("F:/FF/nucleus_live_cell/jurkat_nucleus_centrosome")

# All 4-panel movie folders live directly under this root, one subfolder per
# variant (panel_raw_region_depth_curv, panel_raw_region_depth_curv_dot).
EXPT_032022_MOVIES_ROOT = (
    JURKAT_BASE
    / "NaBu800 Experiments/Control_3D_CD3/all_cells_together/prog_live_cells/raw_with_meshes_movies"
)

EXPERIMENT_LABEL = "032022 expt"

DEFAULT_OUTPUT = Path(
    "K:/FF/PPT/PPT_autogeneration/Live Cells/Jurkat 4-panel mesh movies (032022).pptx"
)

# Poster frames are transient per-run scratch; keep them out of the repo and out
# of the deck output folder (wiped and recreated each run).
DEFAULT_POSTER_DIR = Path(tempfile.gettempdir()) / "ppt_jurkat_4panel_posters"

VIEWS: Tuple[str, ...] = ("xz", "xz_rev", "yz", "yz_rev")


def build_variant_patterns(allow_dot: bool) -> Dict[str, "re.Pattern[str]"]:
    """Build per-view filename regexes for a 4-panel movie folder.

    Each pattern is fully anchored (``$``) so ``xz`` never matches an
    ``xz_rev`` filename and vice versa. ``allow_dot`` makes the ``_dot`` token
    optional so the dot folder matches either the described
    ``..._curv_dot_<view>_...`` naming or a frame-style ``..._curv_<view>_...``
    naming; because the pattern is only ever applied within the dot folder,
    either naming is unambiguous.
    """
    dot = r"(?:_dot)?" if allow_dot else ""
    return {
        view: re.compile(
            rf"^Cell(?P<cell>\d+)_rwm_panel_raw_region_depth_curv{dot}_{view}"
            rf"_white_grey_blue\.mp4$",
            re.IGNORECASE,
        )
        for view in VIEWS
    }


@dataclass(frozen=True)
class PanelSpec:
    """One 4-panel variant to add into the deck.

    ``required`` panels raise if their movie folder is missing; optional panels
    warn and are skipped, so the deck still builds from whatever is present.
    """

    key: str
    movie_subdir: str
    required: bool
    variant_patterns: Dict[str, "re.Pattern[str]"]
    title_prefix: str


PANELS: Tuple[PanelSpec, ...] = (
    PanelSpec(
        key="raw",
        movie_subdir="panel_raw_region_depth_curv",
        required=True,
        variant_patterns=build_variant_patterns(allow_dot=False),
        title_prefix="Raw, region, depth, min curvature",
    ),
    PanelSpec(
        key="dot",
        movie_subdir="panel_raw_region_depth_curv_dot",
        required=False,
        variant_patterns=build_variant_patterns(allow_dot=True),
        title_prefix="Raw, region, depth, min curvature, cent dot",
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
            "for the 032022 (NaBu800 Control) Jurkat experiment."
        )
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=f"Output PowerPoint path. Defaults to {DEFAULT_OUTPUT}.",
    )
    parser.add_argument(
        "--cells",
        type=str,
        default="",
        help="Optional comma-separated list of cell numbers to include, e.g. 1,2,5,10",
    )
    return parser.parse_args()


def resolve_output_path(output_arg: Optional[Path]) -> Path:
    """Pick the output path based on --output."""
    if output_arg is not None:
        return output_arg.resolve()
    return DEFAULT_OUTPUT.resolve()


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


def collect_movies_by_variant(
    folder: Path,
    variant_patterns: Dict[str, "re.Pattern[str]"],
    label: str,
) -> Dict[str, Dict[int, Path]]:
    """Scan a parent folder, classify CellN movies into variant -> cell -> path.

    Files in ``folder`` whose names match any pattern in ``variant_patterns``
    are added to the matching variant bucket. Files that do not match any
    pattern (e.g. alternate color/clim variants) are silently ignored. Per-cell
    duplicates within one variant raise ``ValueError``.
    """
    validate_folder(folder, label)

    by_variant: Dict[str, Dict[int, Path]] = {
        variant: {} for variant in variant_patterns
    }
    for movie_path in sorted(folder.glob("*.mp4")):
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


def collect_panel_cells(
    panel: PanelSpec,
    movies_root: Path,
    requested_cells: Optional[Set[int]] = None,
) -> Dict[int, Dict[str, Path]]:
    """Return ``{cell -> {view -> movie path}}`` for cells complete in a panel.

    A missing folder for a required panel raises; for an optional panel it warns
    and returns an empty mapping (so the deck still builds from other panels,
    and this one is picked up automatically once its movies exist).
    """
    folder = movies_root / panel.movie_subdir
    if not folder.exists():
        if panel.required:
            raise FileNotFoundError(
                f"Required {panel.key} panel folder not found: {folder}"
            )
        print(
            f"[WARNING] {panel.key} panel movie folder not found; skipping "
            f"({folder}). It will be included automatically once those movies "
            "are rendered there."
        )
        return {}

    by_variant = collect_movies_by_variant(
        folder, panel.variant_patterns, f"{EXPERIMENT_LABEL} {panel.key}"
    )

    report_missing_cells(
        f"{EXPERIMENT_LABEL} {panel.key}", by_variant, requested_cells
    )

    complete: Set[int] = set(by_variant[VIEWS[0]])
    for view in VIEWS[1:]:
        complete &= set(by_variant[view])
    if requested_cells is not None:
        complete &= requested_cells

    return {
        cell_number: {view: by_variant[view][cell_number] for view in VIEWS}
        for cell_number in sorted(complete)
    }


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


def format_title(panel: PanelSpec, cell_number: int, direction: str) -> str:
    """Return the slide title for a panel/cell/direction."""
    return (
        f"{panel.title_prefix} ({direction}): "
        f"Cell {cell_number} ({EXPERIMENT_LABEL})"
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
    panel_cells: Dict[str, Dict[int, Dict[str, Path]]],
    poster_dir: Path,
) -> List[SlideSpec]:
    """Build two SlideSpecs per cell per present panel (xz pair, yz pair).

    Cells are ordered numerically; within each cell the panels are emitted in
    ``PANELS`` order (raw, then dot), each contributing its two slides only if
    that cell is complete in that panel.
    """
    all_cells: Set[int] = set()
    for cells in panel_cells.values():
        all_cells.update(cells)

    slide_specs: List[SlideSpec] = []
    for cell_number in sorted(all_cells):
        for panel in PANELS:
            views = panel_cells.get(panel.key, {}).get(cell_number)
            if views is None:
                continue
            slug = f"cell{cell_number}_{panel.key}"
            slide_specs.append(
                _build_movie_pair_slide(
                    format_title(panel, cell_number, "xz"),
                    views["xz"],
                    views["xz_rev"],
                    "xz",
                    "xz_rev",
                    poster_dir,
                    f"{slug}_xz",
                )
            )
            slide_specs.append(
                _build_movie_pair_slide(
                    format_title(panel, cell_number, "yz"),
                    views["yz"],
                    views["yz_rev"],
                    "yz",
                    "yz_rev",
                    poster_dir,
                    f"{slug}_yz",
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
        panel_cells: Dict[str, Dict[int, Dict[str, Path]]] = {}
        for panel in PANELS:
            cells = collect_panel_cells(
                panel, EXPT_032022_MOVIES_ROOT, requested_cells=requested_cells
            )
            print(f"{panel.key} panel: found {len(cells)} complete cell set(s)")
            panel_cells[panel.key] = cells
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1

    all_cells: Set[int] = set()
    for cells in panel_cells.values():
        all_cells.update(cells)
    if not all_cells:
        print("[ERROR] No cells have a complete set of movies in any panel")
        return 1

    n_slides = sum(2 * len(cells) for cells in panel_cells.values())
    print(f"Found {len(all_cells)} cell(s) across {len(PANELS)} panel(s)")
    print(f"Will create {n_slides} slide(s)")

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
        slide_specs = build_slide_specs(panel_cells, poster_dir)
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
