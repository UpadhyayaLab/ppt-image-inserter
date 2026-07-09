#!/usr/bin/env python
"""
Build a PowerPoint deck of rotatable 3D models, one mesh per slide.

Point this at one or more mesh files (``.obj`` and friends) and/or folders of
meshes; each mesh becomes its own slide holding a native, interactively
rotatable 3D model with a title above it. In PowerPoint you can grab the model
and orbit it freely with the mouse, or add the ribbon's 3D "Turntable"
animation to spin it (see the note printed at the end of a run).

Meshes are inserted via PowerPoint COM (Windows-only), exactly like the movie
decks in this repo: PowerPoint's ``Add3DModel`` embeds the mesh (converting it
to an internal glTF binary and rendering a fallback preview), so the resulting
deck is self-contained and needs no external files.

Supported input formats: .obj, .glb, .gltf, .fbx, .stl, .ply, .3mf
(``.obj`` geometry-only meshes import fine and render as a solid model.)

Examples
--------
Single mesh to a named deck::

    conda run -n PPT_editing python examples_and_configs/insert_mesh_obj_3d_slides.py \
        "K:/FF/meshes/cell3_nucleus.obj" --output "K:/FF/PPT/PPT_autogeneration/3D models/nucleus.pptx"

Every mesh in a folder, one slide each, opening at a steeper top-down angle::

    conda run -n PPT_editing python examples_and_configs/insert_mesh_obj_3d_slides.py \
        "K:/FF/meshes/nuclei" --rot-x 35 --rot-y -25

Preview what would be built without touching PowerPoint::

    conda run -n PPT_editing python examples_and_configs/insert_mesh_obj_3d_slides.py \
        "K:/FF/meshes/nuclei" --list

IMPORTANT: The output .pptx must be closed in PowerPoint before running, and no
stray ``POWERPNT`` process may be running, or the COM build fails. If a build
fails with an "RPC" / "call was rejected" error, close PowerPoint (and end any
orphaned POWERPNT task) and re-run.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Sequence

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from ppt_image_inserter import (
    Model3DSlideSpec,
    Model3DSpec,
    TextboxSpec,
    SUPPORTED_MODEL_EXTENSIONS,
    backup_presentation,
    build_model3d_deck_via_com,
)

# Deck geometry (16:9 widescreen), all in inches unless noted.
SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
POINTS_PER_INCH = 72.0

TITLE_LEFT_IN = 0.45
TITLE_TOP_IN = 0.15
TITLE_WIDTH_IN = SLIDE_WIDTH_IN - 2 * TITLE_LEFT_IN
TITLE_HEIGHT_IN = 0.5
TITLE_FONT_SIZE_PT = 24.0

# The model's bounding box: a large square centred below the title. PowerPoint
# fits the mesh inside this box preserving the mesh's own proportions.
MODEL_BOX_SIZE_IN = 6.2
MODEL_BOX_TOP_IN = 0.95
MODEL_BOX_LEFT_IN = (SLIDE_WIDTH_IN - MODEL_BOX_SIZE_IN) / 2.0

# Default opening orientation, in degrees. A three-quarter view so a mesh reads
# as 3D immediately instead of appearing flat face-on. Override with --rot-*.
DEFAULT_ROT_X_DEG = 20.0
DEFAULT_ROT_Y_DEG = -30.0
DEFAULT_ROT_Z_DEG = 0.0

DEFAULT_OUTPUT = Path(
    "K:/FF/PPT/PPT_autogeneration/3D models/mesh 3D models.pptx"
)


@dataclass(frozen=True)
class MeshItem:
    """One mesh to place on its own slide."""

    mesh_path: Path
    title: str


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Create a PowerPoint deck of rotatable 3D models, one mesh per "
            "slide. Accepts mesh files and/or folders of meshes "
            f"({', '.join(SUPPORTED_MODEL_EXTENSIONS)})."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        type=Path,
        help="Mesh file(s) and/or folder(s) to include.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=f"Output .pptx path. Defaults to {DEFAULT_OUTPUT}.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recurse into subfolders when an input is a directory.",
    )
    parser.add_argument(
        "--rot-x",
        type=float,
        default=DEFAULT_ROT_X_DEG,
        help=f"Default model pitch (X) in degrees. Default: {DEFAULT_ROT_X_DEG}.",
    )
    parser.add_argument(
        "--rot-y",
        type=float,
        default=DEFAULT_ROT_Y_DEG,
        help=f"Default model yaw (Y) in degrees. Default: {DEFAULT_ROT_Y_DEG}.",
    )
    parser.add_argument(
        "--rot-z",
        type=float,
        default=DEFAULT_ROT_Z_DEG,
        help=f"Default model roll (Z) in degrees. Default: {DEFAULT_ROT_Z_DEG}.",
    )
    parser.add_argument(
        "--fov",
        type=float,
        default=None,
        help=(
            "Optional perspective field of view in degrees. Omit to keep "
            "PowerPoint's default camera."
        ),
    )
    parser.add_argument(
        "--title-prefix",
        type=str,
        default="",
        help="Optional text prepended to every slide title.",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List the meshes that would be inserted and exit (no PowerPoint).",
    )
    return parser.parse_args()


def title_from_stem(stem: str, prefix: str) -> str:
    """Turn a filename stem into a slide title (underscores -> spaces)."""
    cleaned = stem.replace("_", " ").strip()
    if prefix:
        return f"{prefix.strip()} {cleaned}".strip()
    return cleaned


def collect_meshes(
    inputs: Sequence[Path],
    recursive: bool,
    title_prefix: str,
) -> List[MeshItem]:
    """Expand the input paths into an ordered, de-duplicated list of meshes.

    Files with a supported extension are taken as-is. Directories are scanned
    for supported meshes (recursively if requested). Unsupported files are
    skipped with a warning; missing paths raise ``FileNotFoundError``.
    """
    supported = set(SUPPORTED_MODEL_EXTENSIONS)
    seen: set = set()
    items: List[MeshItem] = []

    def add_file(path: Path) -> None:
        resolved = path.resolve()
        if resolved in seen:
            return
        seen.add(resolved)
        items.append(
            MeshItem(
                mesh_path=resolved,
                title=title_from_stem(resolved.stem, title_prefix),
            )
        )

    for raw in inputs:
        if not raw.exists():
            raise FileNotFoundError(f"Input path not found: {raw}")

        if raw.is_dir():
            pattern = "**/*" if recursive else "*"
            matched = sorted(
                p
                for p in raw.glob(pattern)
                if p.is_file() and p.suffix.lower() in supported
            )
            if not matched:
                print(f"[WARNING] No supported meshes found in folder: {raw}")
            for mesh_path in matched:
                add_file(mesh_path)
        elif raw.suffix.lower() in supported:
            add_file(raw)
        else:
            print(
                f"[WARNING] Skipping unsupported file type: {raw.name} "
                f"(supported: {', '.join(SUPPORTED_MODEL_EXTENSIONS)})"
            )

    return items


def _to_points(inches: float) -> float:
    """Convert inches to PowerPoint points."""
    return inches * POINTS_PER_INCH


def build_slide_specs(
    meshes: Sequence[MeshItem],
    rot_x: float,
    rot_y: float,
    rot_z: float,
    fov: float | None,
) -> List[Model3DSlideSpec]:
    """Build one slide spec per mesh: a centred model with a title above it."""
    slides: List[Model3DSlideSpec] = []
    for mesh in meshes:
        title = TextboxSpec(
            text=mesh.title,
            left_pt=_to_points(TITLE_LEFT_IN),
            top_pt=_to_points(TITLE_TOP_IN),
            width_pt=_to_points(TITLE_WIDTH_IN),
            height_pt=_to_points(TITLE_HEIGHT_IN),
            font_size_pt=TITLE_FONT_SIZE_PT,
            bold=True,
            align="center",
            font_name="Arial",
        )
        model = Model3DSpec(
            model_path=str(mesh.mesh_path),
            left_pt=_to_points(MODEL_BOX_LEFT_IN),
            top_pt=_to_points(MODEL_BOX_TOP_IN),
            width_pt=_to_points(MODEL_BOX_SIZE_IN),
            height_pt=_to_points(MODEL_BOX_SIZE_IN),
            rot_x_deg=rot_x,
            rot_y_deg=rot_y,
            rot_z_deg=rot_z,
            field_of_view_deg=fov,
        )
        slides.append(Model3DSlideSpec(textboxes=(title,), models=(model,)))
    return slides


def resolve_output_path(output_arg: Path | None) -> Path:
    """Pick the output path from --output or the default."""
    if output_arg is not None:
        return output_arg.resolve()
    return DEFAULT_OUTPUT.resolve()


def ensure_parent_dir(output_path: Path) -> None:
    """Create the output directory if needed."""
    if output_path.parent and not output_path.parent.exists():
        output_path.parent.mkdir(parents=True, exist_ok=True)


def main() -> int:
    """Run the 3D-model deck generation workflow."""
    args = parse_args()

    try:
        meshes = collect_meshes(args.inputs, args.recursive, args.title_prefix)
    except FileNotFoundError as exc:
        print(f"[ERROR] {exc}")
        return 1

    if not meshes:
        print("[ERROR] No supported mesh files were found in the given inputs.")
        return 1

    print(f"Found {len(meshes)} mesh(es); will create {len(meshes)} slide(s).")
    if args.list:
        for mesh in meshes:
            print(f"  - {mesh.title}  <-  {mesh.mesh_path}")
        print("(--list) No PowerPoint file written.")
        return 0

    output_path = resolve_output_path(args.output)
    ensure_parent_dir(output_path)

    if output_path.exists():
        backup_dir = output_path.parent / "backups"
        try:
            created = backup_presentation(
                str(output_path), backup_base=str(backup_dir)
            )
            print(f"Backed up existing output to: {', '.join(sorted(created))}")
        except Exception as exc:
            print(f"[WARNING] Could not back up existing output: {exc}")

    try:
        slide_specs = build_slide_specs(
            meshes, args.rot_x, args.rot_y, args.rot_z, args.fov
        )
        inserted = build_model3d_deck_via_com(
            slide_specs,
            str(output_path),
            slide_width_pt=_to_points(SLIDE_WIDTH_IN),
            slide_height_pt=_to_points(SLIDE_HEIGHT_IN),
        )
        print(f"Inserted {inserted} rotatable 3D model(s).")
    except PermissionError:
        print(
            f"[ERROR] Permission denied when saving {output_path}. "
            "Make sure the PowerPoint file is closed."
        )
        return 1
    except Exception as exc:
        print(f"[ERROR] Failed while building/saving slides: {exc}")
        print(
            "        If this is an RPC / 'call was rejected' error, close "
            "PowerPoint and end any orphaned POWERPNT task, then re-run."
        )
        return 1

    print(f"[SUCCESS] Saved presentation: {output_path}")
    print(
        "\nEach model is rotatable: click it in PowerPoint and drag the 3D "
        "rotation handle to orbit it.\nTo auto-spin during a slideshow, select "
        "the model -> Animations tab -> Turntable, then\nAnimations -> Timing "
        "-> Repeat: Until End of Slide."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
