#!/usr/bin/env python
"""
Build a PowerPoint deck of COLOR-PRESERVED, rotatable nucleus 3D models.

The plain ``Cell<N>_t0MM_mesh.obj`` exports are geometry only -- PowerPoint
shows them flat grey. The per-frame MATLAB figures
(``Cell<N>/frame<M>/nucleus/mesh/Cell_<N>_<field>.fig``) instead store the mesh
*with* its scalar coloring (``FaceVertexCData`` + colormap + CLim), e.g. the
min-curvature invagination map. This script recovers that color and bakes it
into the 3D model so the rotatable nucleus in PowerPoint looks like the MATLAB
figure.

Pipeline, per timepoint:

1. MATLAB (``extract_fig_mesh_color_batch``) opens each ``.fig``, reproduces the
   displayed RGB, and writes V / F / RGB to a temp ``.mat``. All frames are done
   in a single MATLAB session.
2. :func:`ppt_image_inserter.build_colored_glb` bakes the per-face colors into a
   ``.glb`` (glTF ``COLOR_0``), which PowerPoint's 3D engine renders in color.
3. The ``.glb`` models are inserted via COM (``Add3DModel``), one per slide,
   titled by timepoint -- each natively rotatable in PowerPoint.

Scope: Jurkat live-cell nucleus data (032022 NaBu800 Control_3D_CD3 and
04142022 GFP-Centrin). Windows-only (PowerPoint COM + MATLAB).

Examples
--------
Cell 1, min-curvature coloring (default), all timepoints::

    conda run -n PPT_editing python examples_and_configs/insert_jurkat_nuc_colored_mesh_slides.py --cell 1

A different scalar field, and preview first::

    conda run -n PPT_editing python examples_and_configs/insert_jurkat_nuc_colored_mesh_slides.py \
        --cell 1 --field chull_dist --list

IMPORTANT: Close the output .pptx in PowerPoint and end any orphaned POWERPNT
task before running, or the COM build fails.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import scipy.io as sio

REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT))

from ppt_image_inserter import (
    Model3DSlideSpec,
    Model3DSpec,
    TextboxSpec,
    backup_presentation,
    build_colored_glb,
    build_model3d_deck_via_com,
)

MATLAB_DIR = Path(__file__).resolve().parent / "matlab"

JURKAT_BASE = Path("F:/FF/nucleus_live_cell/jurkat_nucleus_centrosome")

# prog_live_cells root per experiment; per-cell frames live beneath it as
# Cell<N>/frame<M>/nucleus/mesh/Cell_<N>_<field>.fig
EXPERIMENTS: Dict[str, Path] = {
    "032022": JURKAT_BASE
    / "NaBu800 Experiments/Control_3D_CD3/all_cells_together/prog_live_cells",
    "04142022": JURKAT_BASE
    / "GFP-Centrin_SiR-DNA/Control/cells/all_cells_together/prog_live_cells",
}

# Scalar-field figures written per frame (Cell_<N>_<field>.fig). min_curv is the
# canonical invagination map (concave grooves red, convex blue).
DEFAULT_FIELD = "min_curv"

OUTPUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Live Cells")

# Deck geometry (16:9 widescreen), inches.
SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
POINTS_PER_INCH = 72.0
TITLE_LEFT_IN = 0.45
TITLE_TOP_IN = 0.15
TITLE_WIDTH_IN = SLIDE_WIDTH_IN - 2 * TITLE_LEFT_IN
TITLE_HEIGHT_IN = 0.5
TITLE_FONT_SIZE_PT = 24.0
MODEL_BOX_SIZE_IN = 6.2
MODEL_BOX_TOP_IN = 0.95
MODEL_BOX_LEFT_IN = (SLIDE_WIDTH_IN - MODEL_BOX_SIZE_IN) / 2.0
DEFAULT_ROT_X_DEG = 20.0
DEFAULT_ROT_Y_DEG = -30.0

FRAME_DIR_RE = re.compile(r"^frame(?P<n>\d+)$", re.IGNORECASE)


@dataclass(frozen=True)
class FrameFig:
    """One timepoint's source figure."""

    frame: int
    fig_path: Path


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Build a deck of color-preserved rotatable nucleus 3D models by "
            "extracting per-frame MATLAB .fig colors and baking them into glTF."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--cell", type=int, default=1, help="Cell number (default 1).")
    parser.add_argument(
        "--experiment",
        choices=sorted(EXPERIMENTS),
        default="032022",
        help="Experiment (default 032022).",
    )
    parser.add_argument(
        "--field",
        default=DEFAULT_FIELD,
        help=(
            "Scalar-field figure to color by, i.e. the Cell_<N>_<field>.fig "
            f"suffix (default {DEFAULT_FIELD}). e.g. min_curv, mean_curv, "
            "chull_dist, concav."
        ),
    )
    parser.add_argument(
        "--frames",
        default="",
        help="Optional comma-separated frame numbers to include, e.g. 1,5,10.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output .pptx path. Defaults into the Live Cells folder.",
    )
    parser.add_argument(
        "--rot-x", type=float, default=DEFAULT_ROT_X_DEG, help="Default model pitch (deg)."
    )
    parser.add_argument(
        "--rot-y", type=float, default=DEFAULT_ROT_Y_DEG, help="Default model yaw (deg)."
    )
    parser.add_argument(
        "--matlab",
        type=str,
        default=None,
        help="Path to matlab.exe (auto-detected if omitted).",
    )
    parser.add_argument(
        "--list",
        action="store_true",
        help="List the frames/figs that would be used and exit.",
    )
    return parser.parse_args()


def field_label(field: str) -> str:
    """Human-readable field name for titles (min_curv -> 'min curv')."""
    return field.replace("_", " ").strip()


def discover_frames(
    prog_root: Path, cell: int, field: str, wanted: Optional[set]
) -> List[FrameFig]:
    """Find the per-frame .fig files for a cell/field, ordered by frame."""
    cell_dir = prog_root / f"Cell{cell}"
    if not cell_dir.is_dir():
        raise FileNotFoundError(f"Cell folder not found: {cell_dir}")

    found: List[FrameFig] = []
    for frame_dir in cell_dir.iterdir():
        if not frame_dir.is_dir():
            continue
        m = FRAME_DIR_RE.match(frame_dir.name)
        if not m:
            continue
        frame = int(m.group("n"))
        if wanted is not None and frame not in wanted:
            continue
        fig = frame_dir / "nucleus" / "mesh" / f"Cell_{cell}_{field}.fig"
        if fig.is_file():
            found.append(FrameFig(frame=frame, fig_path=fig))
        else:
            print(f"[WARNING] frame {frame}: missing {fig.name} ({fig})")

    found.sort(key=lambda ff: ff.frame)
    return found


def parse_frames_arg(frames_arg: str) -> Optional[set]:
    """Parse an optional comma-separated frame list."""
    if not frames_arg.strip():
        return None
    out = set()
    for tok in frames_arg.split(","):
        tok = tok.strip()
        if tok:
            out.add(int(tok))
    return out or None


def find_matlab(explicit: Optional[str]) -> str:
    """Locate matlab.exe: explicit arg, then PATH, then a Program Files scan."""
    if explicit:
        if not Path(explicit).exists():
            raise FileNotFoundError(f"--matlab not found: {explicit}")
        return explicit
    on_path = shutil.which("matlab")
    if on_path:
        return on_path
    candidates = sorted(
        Path("C:/Program Files/MATLAB").glob("R*/bin/matlab.exe"), reverse=True
    )
    if candidates:
        return str(candidates[0])
    raise FileNotFoundError(
        "Could not find matlab.exe. Pass --matlab C:/path/to/matlab.exe"
    )


def run_matlab_extraction(
    matlab_exe: str, jobs: Sequence[Tuple[Path, Path]]
) -> None:
    """Extract colors for all figs in one MATLAB session via a manifest.

    Args:
        matlab_exe: path to matlab.exe.
        jobs: (fig_path, out_mat) pairs.
    """
    manifest_lines = [
        f"{fig.resolve().as_posix()}\t{mat.resolve().as_posix()}" for fig, mat in jobs
    ]
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".fig_manifest.txt", delete=False, encoding="utf-8"
    ) as tmp:
        tmp.write("\n".join(manifest_lines))
        manifest_path = tmp.name

    try:
        cmd = [
            matlab_exe,
            "-batch",
            f"extract_fig_mesh_color_batch('{Path(manifest_path).as_posix()}')",
        ]
        proc = subprocess.run(
            cmd, cwd=str(MATLAB_DIR), capture_output=True, text=True
        )
        if proc.returncode != 0:
            raise RuntimeError(
                "MATLAB color extraction failed.\n"
                f"stdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
            )
        # Surface any per-figure failures the batch reported.
        for line in proc.stdout.splitlines():
            if line.startswith("FAIL") or line.startswith("DONE"):
                print(f"  [matlab] {line}")
    finally:
        try:
            Path(manifest_path).unlink()
        except OSError:
            pass


def bake_glb_from_mat(mat_path: Path, glb_path: Path) -> None:
    """Load an extracted .mat and write a colored .glb."""
    d = sio.loadmat(str(mat_path))
    V = np.asarray(d["V"], dtype=np.float64)
    F = np.asarray(d["F"], dtype=np.int64) - 1  # MATLAB 1-based -> 0-based
    RGB = np.asarray(d["RGB"], dtype=np.float64)
    per_face = bool(np.ravel(d["perface"])[0])
    build_colored_glb(V, F, RGB, str(glb_path), per_face=per_face)


def _to_points(inches: float) -> float:
    return inches * POINTS_PER_INCH


def build_slide_spec(title_text: str, glb_path: Path, rot_x: float, rot_y: float) -> Model3DSlideSpec:
    """One slide: centered colored model with a title above it."""
    title = TextboxSpec(
        text=title_text,
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
        model_path=str(glb_path),
        left_pt=_to_points(MODEL_BOX_LEFT_IN),
        top_pt=_to_points(MODEL_BOX_TOP_IN),
        width_pt=_to_points(MODEL_BOX_SIZE_IN),
        height_pt=_to_points(MODEL_BOX_SIZE_IN),
        rot_x_deg=rot_x,
        rot_y_deg=rot_y,
    )
    return Model3DSlideSpec(textboxes=(title,), models=(model,))


def resolve_output_path(
    output_arg: Optional[Path], experiment: str, cell: int, field: str
) -> Path:
    """Pick the output path from --output or a Live Cells default."""
    if output_arg is not None:
        return output_arg.resolve()
    name = f"Jurkat nucleus 3D meshes ({experiment} Cell {cell}, {field_label(field)}).pptx"
    return (OUTPUT_DIR / name).resolve()


def main() -> int:
    """Run the colored-nucleus deck workflow."""
    args = parse_args()

    prog_root = EXPERIMENTS[args.experiment]
    if not prog_root.is_dir():
        print(f"[ERROR] Experiment root not found: {prog_root}")
        return 1

    try:
        wanted = parse_frames_arg(args.frames)
        frames = discover_frames(prog_root, args.cell, args.field, wanted)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[ERROR] {exc}")
        return 1

    if not frames:
        print(
            f"[ERROR] No '{args.field}' figures found for Cell {args.cell} "
            f"({args.experiment})."
        )
        return 1

    label = field_label(args.field)
    print(
        f"{args.experiment} Cell {args.cell}: {len(frames)} frame(s) with "
        f"'{args.field}' coloring."
    )
    if args.list:
        for ff in frames:
            print(f"  - t{ff.frame:03d}  <-  {ff.fig_path}")
        print("(--list) Nothing extracted or written.")
        return 0

    output_path = resolve_output_path(args.output, args.experiment, args.cell, args.field)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        matlab_exe = find_matlab(args.matlab)
    except FileNotFoundError as exc:
        print(f"[ERROR] {exc}")
        return 1
    print(f"Using MATLAB: {matlab_exe}")

    work_dir = Path(tempfile.mkdtemp(prefix="ppt_jurkat_colored_"))
    try:
        jobs: List[Tuple[Path, Path]] = []
        mat_for_frame: Dict[int, Path] = {}
        for ff in frames:
            mat_path = work_dir / f"cell{args.cell}_t{ff.frame:03d}_{args.field}.mat"
            jobs.append((ff.fig_path, mat_path))
            mat_for_frame[ff.frame] = mat_path

        print(f"Extracting colors from {len(jobs)} figure(s) via MATLAB ...")
        run_matlab_extraction(matlab_exe, jobs)

        slide_specs: List[Model3DSlideSpec] = []
        for ff in frames:
            mat_path = mat_for_frame[ff.frame]
            if not mat_path.exists():
                print(f"[WARNING] t{ff.frame:03d}: extraction produced no .mat; skipping")
                continue
            glb_path = work_dir / f"cell{args.cell}_t{ff.frame:03d}_{args.field}.glb"
            try:
                bake_glb_from_mat(mat_path, glb_path)
            except Exception as exc:
                print(f"[WARNING] t{ff.frame:03d}: could not bake glb ({exc}); skipping")
                continue
            title = f"{args.experiment} Jurkat Cell {args.cell} t{ff.frame:03d} {label}"
            slide_specs.append(build_slide_spec(title, glb_path, args.rot_x, args.rot_y))

        if not slide_specs:
            print("[ERROR] No colored models were produced.")
            return 1

        if output_path.exists():
            backup_dir = output_path.parent / "backups"
            try:
                created = backup_presentation(str(output_path), backup_base=str(backup_dir))
                print(f"Backed up existing output to: {', '.join(sorted(created))}")
            except Exception as exc:
                print(f"[WARNING] Could not back up existing output: {exc}")

        print(f"Building deck with {len(slide_specs)} colored model(s) ...")
        inserted = build_model3d_deck_via_com(
            slide_specs,
            str(output_path),
            slide_width_pt=_to_points(SLIDE_WIDTH_IN),
            slide_height_pt=_to_points(SLIDE_HEIGHT_IN),
        )
        print(f"Inserted {inserted} colored rotatable 3D model(s).")
    except PermissionError:
        print(
            f"[ERROR] Permission denied saving {output_path}. "
            "Close the file in PowerPoint and re-run."
        )
        return 1
    except Exception as exc:
        print(f"[ERROR] {exc}")
        print(
            "        If this is an RPC / 'call was rejected' error, close "
            "PowerPoint and end any orphaned POWERPNT task, then re-run."
        )
        return 1
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

    print(f"[SUCCESS] Saved presentation: {output_path}")
    print(
        "\nEach model keeps the MATLAB face colors and is rotatable: click it "
        "and drag the 3D handle to orbit."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
