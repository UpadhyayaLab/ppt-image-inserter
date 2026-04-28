#!/usr/bin/env python
"""
Build a PowerPoint deck with one invagination-centrosome panel movie per slide.

This script creates a new presentation from a folder of
`CellN_invag_cent_panels.mp4` movies. Each movie gets its own slide with a
title, an embedded video, and a poster frame generated from frame 1 of the
movie. Movies are inserted via PowerPoint COM (Windows-only); after saving,
``force_autoplay_in_pptx`` rewrites the timing-tree trigger on disk so videos
autoplay in slideshow mode.

The output file is backed up before being overwritten.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence, Set, Tuple

import imageio.v3 as iio
from PIL import Image


REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from ppt_image_inserter import backup_presentation, force_autoplay_in_pptx


DEFAULT_MOVIE_DIR = Path(
    "F:/FF/nucleus_live_cell/ctl_nucleus_centrosome/"
    "20220614_OT1_CTLs_Centrin_SiR-DNA/antiCD3/Cells/channels/frames/"
    "prog_live_cells/invag_avis/cent"
)
DEFAULT_OUTPUT = DEFAULT_MOVIE_DIR / "invag_cent_panels_movies.pptx"
DEFAULT_POSTER_DIR = DEFAULT_MOVIE_DIR / "_movie_posters_tmp"
DEFAULT_EXPERIMENT_LABEL = "20220614 OT1 CTLs antiCD3"
DEFAULT_COMBINED_OUTPUT = Path(
    "K:/FF/PPT/PPT_autogeneration/Live Cells/CTL_invag_cent_panels_movies_combined.pptx"
)
DEFAULT_COMBINED_POSTER_DIR = Path(
    "K:/FF/PPT/PPT_autogeneration/Live Cells/_movie_posters_tmp_invag_cent_panels_combined"
)
DEFAULT_JURKAT_COMBINED_OUTPUT = Path(
    "K:/FF/PPT/PPT_autogeneration/Live Cells/Jurkat_invag_cent_panels_movies_combined.pptx"
)
DEFAULT_JURKAT_COMBINED_POSTER_DIR = Path(
    "K:/FF/PPT/PPT_autogeneration/Live Cells/_movie_posters_tmp_jurkat_invag_cent_panels_combined"
)

MOVIE_PATTERN = re.compile(r"^Cell(?P<cell>\d+)_invag_cent_panels\.mp4$", re.IGNORECASE)

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
TITLE_LEFT_IN = 0.45
TITLE_TOP_IN = 0.12
TITLE_WIDTH_IN = 12.4
TITLE_HEIGHT_IN = 0.45
MOVIE_BOX_LEFT_IN = 0.45
MOVIE_BOX_TOP_IN = 0.78
MOVIE_BOX_WIDTH_IN = 12.43
MOVIE_BOX_HEIGHT_IN = 6.3
POINTS_PER_INCH = 72.0


@dataclass(frozen=True)
class MovieSource:
    """A folder of CellN invag-cent panel movies plus its title label."""

    movie_dir: Path
    experiment_label: str
    exclude_cells: Tuple[int, ...] = ()


@dataclass(frozen=True)
class MovieRecord:
    """One movie to place on one slide."""

    cell_number: int
    movie_path: Path
    experiment_label: str


DEFAULT_COMBINED_SOURCES: Tuple[MovieSource, ...] = (
    MovieSource(
        Path(
            "F:/FF/nucleus_live_cell/ctl_nucleus_centrosome/"
            "20210928_OTI_CTL_Activated/combined/frames/prog_live_cells/invag_avis/cent"
        ),
        "20210928 OTI CTL Activated",
        (2, 3, 4, 5, 7),
    ),
    MovieSource(
        Path(
            "F:/FF/nucleus_live_cell/ctl_nucleus_centrosome/"
            "20211221_OTI_CTL_Activated_already_cropped/combined/channels/frames/"
            "prog_live_cells/invag_avis/cent"
        ),
        "20211221 OTI CTL Activated",
        (7,),
    ),
    MovieSource(
        DEFAULT_MOVIE_DIR,
        DEFAULT_EXPERIMENT_LABEL,
        (5, 8, 18, 19, 20, 23, 25, 27, 28, 43, 44, 48, 50, 51, 53),
    ),
)

DEFAULT_JURKAT_COMBINED_SOURCES: Tuple[MovieSource, ...] = (
    MovieSource(
        Path(
            "F:/FF/nucleus_live_cell/jurkat_nucleus_centrosome/"
            "NaBu800 Experiments/Control_3D_CD3/all_cells_together/"
            "prog_live_cells/invag_avis/cent"
        ),
        "032022",
    ),
    MovieSource(
        Path(
            "F:/FF/nucleus_live_cell/jurkat_nucleus_centrosome/"
            "GFP-Centrin_SiR-DNA/Control/cells/all_cells_together/"
            "prog_live_cells/invag_avis/cent"
        ),
        "04142022",
    ),
)


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Create a PowerPoint deck with one CellN_invag_cent_panels.mp4 "
            "movie per slide."
        )
    )
    parser.add_argument(
        "--movie-dir",
        type=Path,
        default=DEFAULT_MOVIE_DIR,
        help=f"Folder containing CellN_invag_cent_panels.mp4 files. Default: {DEFAULT_MOVIE_DIR}",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Output PowerPoint path. Default: {DEFAULT_OUTPUT}",
    )
    parser.add_argument(
        "--poster-dir",
        type=Path,
        default=DEFAULT_POSTER_DIR,
        help=f"Temporary poster-frame folder. Default: {DEFAULT_POSTER_DIR}",
    )
    parser.add_argument(
        "--experiment-label",
        type=str,
        default=DEFAULT_EXPERIMENT_LABEL,
        help=f"Text appended to slide titles. Default: {DEFAULT_EXPERIMENT_LABEL}",
    )
    parser.add_argument(
        "--cells",
        type=str,
        default="",
        help="Optional comma-separated list of cell numbers to include, e.g. 1,2,5,10",
    )
    parser.add_argument(
        "--combined-defaults",
        action="store_true",
        help=(
            "Ignore --movie-dir and build one deck from the three default CTL "
            "invag-cent experiment folders."
        ),
    )
    parser.add_argument(
        "--jurkat-combined-defaults",
        action="store_true",
        help=(
            "Ignore --movie-dir and build one deck from the two default Jurkat "
            "invag-cent experiment folders."
        ),
    )
    return parser.parse_args()


def parse_requested_cells(cells_arg: str) -> Optional[Set[int]]:
    """Parse an optional comma-separated list of cell numbers."""
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


def validate_movie_dir(movie_dir: Path) -> None:
    """Ensure the input movie folder exists."""
    if not movie_dir.exists():
        raise FileNotFoundError(f"Movie folder not found: {movie_dir}")
    if not movie_dir.is_dir():
        raise NotADirectoryError(f"Movie folder is not a directory: {movie_dir}")


def collect_movies(
    movie_dir: Path,
    requested_cells: Optional[Set[int]] = None,
) -> List[Tuple[int, Path]]:
    """Collect matching CellN_invag_cent_panels.mp4 files in numeric order."""
    validate_movie_dir(movie_dir)

    movies: List[Tuple[int, Path]] = []
    seen_cells: Set[int] = set()

    for movie_path in sorted(movie_dir.glob("Cell*_invag_cent_panels.mp4")):
        match = MOVIE_PATTERN.match(movie_path.name)
        if not match:
            print(f"[WARNING] Skipping unmatched movie filename: {movie_path.name}")
            continue

        cell_number = int(match.group("cell"))
        if requested_cells is not None and cell_number not in requested_cells:
            continue

        if cell_number in seen_cells:
            raise ValueError(f"Duplicate Cell{cell_number} movie found: {movie_path.name}")

        seen_cells.add(cell_number)
        movies.append((cell_number, movie_path))

    if requested_cells is not None:
        missing_cells = sorted(requested_cells - seen_cells)
        for cell_number in missing_cells:
            print(f"[WARNING] Cell{cell_number} requested but no matching movie was found")

    if not movies:
        raise ValueError(f"No matching CellN_invag_cent_panels.mp4 files found in {movie_dir}")

    movies.sort(key=lambda item: item[0])
    return movies


def collect_movie_records(
    sources: Sequence[MovieSource],
    requested_cells: Optional[Set[int]] = None,
) -> List[MovieRecord]:
    """Collect movies from one or more experiment folders."""
    records: List[MovieRecord] = []
    for source in sources:
        source_movies = collect_movies(source.movie_dir, requested_cells=requested_cells)
        excluded = set(source.exclude_cells)
        for cell_number, movie_path in source_movies:
            if cell_number in excluded:
                continue
            records.append(
                MovieRecord(
                    cell_number=cell_number,
                    movie_path=movie_path,
                    experiment_label=source.experiment_label,
                )
            )
    return records


def extract_first_frame(movie_path: Path, poster_path: Path) -> Tuple[int, int]:
    """Extract frame 1 (index 0), save it as PNG, and return width/height."""
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
    """Fit content inside a target box while preserving aspect ratio."""
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
    """Return the parenthetical experiment text for the slide title."""
    match = re.search(r"\b(\d{8})\b", experiment_label)
    if match:
        return f"{match.group(1)} expt"
    match = re.search(r"\b(\d{6})\b", experiment_label)
    if match:
        return f"{match.group(1)} expt"
    return experiment_label


def format_slide_title(cell_number: int, experiment_label: str) -> str:
    """Return the slide title for one movie."""
    return (
        f"Nucleus-Centrosome (single slice): Cell {cell_number} "
        f"({format_experiment_suffix(experiment_label)})"
    )


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


def _to_points(inches: float) -> float:
    """Convert inches to PowerPoint points."""
    return inches * POINTS_PER_INCH


def build_presentation_with_powerpoint_com(
    movies: Sequence[MovieRecord],
    poster_dir: Path,
    output_path: Path,
) -> None:
    """Build the presentation using native PowerPoint media objects via COM.

    Inserts each movie via ``Shapes.AddMediaObject2`` with autoplay+loop flags
    set, and attaches the poster frame via ``MediaFormat.SetDisplayPictureFromFile``.
    Autoplay-trigger correctness is enforced separately by
    :func:`force_autoplay_in_pptx` after the file is saved, since PowerPoint
    COM does not write the required ``<p:cond delay="0"/>`` start condition
    on its own.
    """
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint COM deck generation is only available on Windows")

    manifest = {
        "slide_width_pt": _to_points(SLIDE_WIDTH_IN),
        "slide_height_pt": _to_points(SLIDE_HEIGHT_IN),
        "title_left_pt": _to_points(TITLE_LEFT_IN),
        "title_top_pt": _to_points(TITLE_TOP_IN),
        "title_width_pt": _to_points(TITLE_WIDTH_IN),
        "title_height_pt": _to_points(TITLE_HEIGHT_IN),
        "slides": [],
    }

    for movie in movies:
        poster_name = f"{movie.experiment_label.replace(' ', '_')}_{movie.movie_path.stem}_poster.png"
        poster_path = poster_dir / poster_name
        frame_width, frame_height = extract_first_frame(movie.movie_path, poster_path)
        left_inches, top_inches, width_inches, height_inches = fit_within_box(
            frame_width,
            frame_height,
            MOVIE_BOX_LEFT_IN,
            MOVIE_BOX_TOP_IN,
            MOVIE_BOX_WIDTH_IN,
            MOVIE_BOX_HEIGHT_IN,
        )
        manifest["slides"].append(
            {
                "title": format_slide_title(movie.cell_number, movie.experiment_label),
                "movie_path": str(movie.movie_path.resolve()),
                "poster_path": str(poster_path.resolve()),
                "left_pt": _to_points(left_inches),
                "top_pt": _to_points(top_inches),
                "width_pt": _to_points(width_inches),
                "height_pt": _to_points(height_inches),
            }
        )

    manifest_path = poster_dir / "_ppt_manifest.json"
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)

    normalized_manifest_path = str(manifest_path.resolve()).replace("'", "''")
    normalized_output_path = str(output_path.resolve()).replace("'", "''")
    powershell_script = rf"""
$ErrorActionPreference = 'Stop'
$manifestPath = '{normalized_manifest_path}'
$outputPath = '{normalized_output_path}'
$data = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$presentation = $app.Presentations.Add()
$buildError = $null
try {{
    $presentation.PageSetup.SlideWidth = [int][math]::Round($data.slide_width_pt)
    $presentation.PageSetup.SlideHeight = [int][math]::Round($data.slide_height_pt)
    foreach ($item in $data.slides) {{
        $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
        $titleShape = $slide.Shapes.AddTextbox(
            1,
            [single]$item.title_left_pt,
            [single]$item.title_top_pt,
            [single]$data.title_width_pt,
            [single]$data.title_height_pt
        )
        $titleRange = $titleShape.TextFrame.TextRange
        $titleRange.Text = $item.title
        $titleRange.ParagraphFormat.Alignment = 2
        $titleRange.Font.Size = 22
        $titleRange.Font.Bold = -1
        $titleRange.Font.Name = 'Arial'

        $mediaShape = $slide.Shapes.AddMediaObject2(
            $item.movie_path,
            $false,
            $true,
            [single]$item.left_pt,
            [single]$item.top_pt,
            [single]$item.width_pt,
            [single]$item.height_pt
        )
        $mediaShape.MediaFormat.SetDisplayPictureFromFile($item.poster_path)
        # Add Media Play (effect 83) on the main sequence; force_autoplay_in_pptx
        # then flips the trigger from delay="indefinite" to delay="0" on disk.
        $slide.TimeLine.MainSequence.AddEffect($mediaShape, 83, 0, 2) | Out-Null
    }}
    $presentation.SaveAs($outputPath, 24)
}} catch {{
    $buildError = $_
    throw
}} finally {{
    try {{
        if ($presentation -and -not $buildError) {{ $presentation.Close() }}
    }} catch {{
    }}
    try {{
        if ($app) {{ $app.Quit() }}
    }} catch {{
    }}
}}
"""

    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                powershell_script,
            ],
            capture_output=True,
            text=True,
            check=True,
        )
    except subprocess.CalledProcessError as exc:
        raise RuntimeError(
            "PowerPoint COM build failed.\n"
            f"stdout:\n{exc.stdout}\n"
            f"stderr:\n{exc.stderr}"
        ) from exc


def main() -> int:
    """Run the movie-deck generation workflow."""
    args = parse_args()

    try:
        requested_cells = parse_requested_cells(args.cells)
        if args.combined_defaults:
            movies = collect_movie_records(DEFAULT_COMBINED_SOURCES, requested_cells=requested_cells)
        elif args.jurkat_combined_defaults:
            movies = collect_movie_records(
                DEFAULT_JURKAT_COMBINED_SOURCES,
                requested_cells=requested_cells,
            )
        else:
            movies = [
                MovieRecord(
                    cell_number=cell_number,
                    movie_path=movie_path,
                    experiment_label=args.experiment_label,
                )
                for cell_number, movie_path in collect_movies(
                    args.movie_dir,
                    requested_cells=requested_cells,
                )
            ]
    except Exception as exc:
        print(f"[ERROR] {exc}")
        return 1

    print(f"Found {len(movies)} movie(s)")

    output_path = args.output.resolve()
    ensure_parent_dir(output_path)

    if output_path.exists():
        backup_dir = output_path.parent / "backups"
        try:
            backups_created = backup_presentation(str(output_path), backup_base=str(backup_dir))
            print(f"Backed up existing output to: {', '.join(sorted(backups_created))}")
        except Exception as exc:
            print(f"[WARNING] Could not back up existing output: {exc}")

    poster_dir = prepare_poster_dir(args.poster_dir.resolve())

    try:
        build_presentation_with_powerpoint_com(movies, poster_dir, output_path)
    except PermissionError:
        print(
            f"[ERROR] Permission denied when saving {output_path}. "
            "Make sure the PowerPoint file is closed."
        )
        return 1
    except Exception as exc:
        print(f"[ERROR] Failed while building/saving slides: {exc}")
        return 1

    try:
        rewritten = force_autoplay_in_pptx(str(output_path))
        print(f"Set {rewritten} slide(s) to autoplay")
    except Exception as exc:
        print(f"[WARNING] Could not set autoplay timing: {exc}")

    print(f"[SUCCESS] Saved presentation: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
