"""3D-model helpers for PowerPoint decks.

Insert native, interactively-rotatable 3D models into a deck using PowerPoint
COM (Windows-only), mirroring :mod:`ppt_image_inserter.movies`.

``python-pptx`` has no API for 3D models, so -- exactly as with embedded movies
-- the deck is built by driving PowerPoint through COM:

- :func:`build_model3d_deck_via_com` writes a manifest and runs a PowerShell
  script that adds each slide's textboxes and 3D models.
- ``Shapes.Add3DModel`` inserts the model. PowerPoint itself converts the
  source mesh to an embedded glTF binary (``ppt/media/model3dN.glb``), renders
  a PNG fallback for older viewers, and wires up the
  ``.../2017/06/relationships/model3d`` relationship. The source file may be
  any format PowerPoint accepts (see :data:`SUPPORTED_MODEL_EXTENSIONS`).

The result is a fully native 3D model. In PowerPoint the user can grab the
on-slide rotation handle and orbit the mesh freely, or apply the ribbon's 3D
"Turntable" animation. An optional default orientation
(``rot_x_deg`` / ``rot_y_deg`` / ``rot_z_deg``) is baked in via
``Model3D.RotationX/Y/Z`` so a mesh opens at an informative three-quarter angle
instead of flat face-on.

Note:
    There is deliberately no auto-spin helper here. Unlike movie autoplay
    (:func:`ppt_image_inserter.movies.force_autoplay_in_pptx`), the 3D
    "Turntable" emphasis animation is not exposed through the COM ``AddEffect``
    enumeration (``MsoAnimEffect`` has no 3D members), and hand-authoring its
    timing OOXML is corruption-prone. Auto-spin is therefore left as a
    two-click manual step in PowerPoint: select the model, Animations >
    Turntable, then Timing > Repeat: Until End of Slide.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from typing import Optional, Sequence

from .movies import TextboxSpec

__all__ = [
    "Model3DSpec",
    "Model3DSlideSpec",
    "build_model3d_deck_via_com",
    "SUPPORTED_MODEL_EXTENSIONS",
]


#: 3D-model file formats PowerPoint's ``Add3DModel`` accepts. Lower-case, with
#: the leading dot. ``.obj`` typically references a companion ``.mtl`` for
#: materials; geometry-only meshes import fine and render as a solid model.
SUPPORTED_MODEL_EXTENSIONS = (
    ".obj",
    ".glb",
    ".gltf",
    ".fbx",
    ".stl",
    ".ply",
    ".3mf",
)

_ALIGN_MAP = {"left": 1, "center": 2, "right": 3, "justify": 4}


@dataclass(frozen=True)
class Model3DSpec:
    """One 3D model to place on a slide. Measurements are in points/degrees.

    Attributes:
        model_path: Path to the source mesh (any of
            :data:`SUPPORTED_MODEL_EXTENSIONS`). Resolved to an absolute path
            before being handed to PowerPoint.
        left_pt: Left edge of the model's bounding box, in points from the
            slide's left edge.
        top_pt: Top edge of the model's bounding box, in points from the top.
        width_pt: Bounding-box width in points. The mesh is fit into the box
            preserving its own proportions.
        height_pt: Bounding-box height in points.
        rot_x_deg: Default model rotation about X (pitch), in degrees.
        rot_y_deg: Default model rotation about Y (yaw), in degrees.
        rot_z_deg: Default model rotation about Z (roll), in degrees.
        field_of_view_deg: Optional perspective field of view in degrees. When
            ``None`` (default), PowerPoint's default camera is left untouched.
    """

    model_path: str
    left_pt: float
    top_pt: float
    width_pt: float
    height_pt: float
    rot_x_deg: float = 0.0
    rot_y_deg: float = 0.0
    rot_z_deg: float = 0.0
    field_of_view_deg: Optional[float] = None


@dataclass(frozen=True)
class Model3DSlideSpec:
    """One slide's textboxes and 3D models."""

    textboxes: Sequence[TextboxSpec] = field(default_factory=tuple)
    models: Sequence[Model3DSpec] = field(default_factory=tuple)


def build_model3d_deck_via_com(
    slides: Sequence[Model3DSlideSpec],
    output_path: str,
    slide_width_pt: float,
    slide_height_pt: float,
    background_rgb: Optional[int] = None,
) -> int:
    """Build a PowerPoint deck of textboxes and native 3D models via COM.

    Each slide gets its textboxes added via ``Shapes.AddTextbox`` and each 3D
    model inserted via ``Shapes.Add3DModel`` (embedded in the deck, not linked).
    A default orientation is applied per model via ``Model3D.RotationX/Y/Z``.
    The models are natively rotatable in PowerPoint after opening.

    Args:
        slides: The slides to build, in order.
        output_path: Where to save the .pptx. Overwritten if it exists. Must be
            closed in PowerPoint first, or the save fails with a permission
            error.
        slide_width_pt: Slide width in points (e.g. 13.333 in * 72 = 960).
        slide_height_pt: Slide height in points (e.g. 7.5 in * 72 = 540).
        background_rgb: Optional solid slide-background color as a BGR integer
            (as VBA ``RGB()``; e.g. ``0`` for black). ``None`` keeps the default
            (white) background. When using a dark background, give textboxes a
            light ``color_rgb`` so their text stays visible.

    Returns:
        The total number of 3D models inserted across all slides.

    Raises:
        RuntimeError: If not running on Windows, or the PowerPoint COM build
            fails (the PowerShell stdout/stderr is included in the message).
    """
    if sys.platform != "win32":
        raise RuntimeError(
            "PowerPoint COM deck generation is only available on Windows"
        )

    manifest = {
        "slide_width_pt": slide_width_pt,
        "slide_height_pt": slide_height_pt,
        "background_rgb": background_rgb,
        "slides": [_slide_to_manifest(slide) for slide in slides],
    }
    total_models = sum(len(slide.models) for slide in slides)

    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".manifest.json", delete=False, encoding="utf-8"
    ) as tmp:
        json.dump(manifest, tmp)
        manifest_path = tmp.name

    try:
        normalized_manifest_path = manifest_path.replace("'", "''")
        normalized_output_path = os.path.abspath(output_path).replace("'", "''")
        powershell_script = _POWERSHELL_TEMPLATE.format(
            manifest_path=normalized_manifest_path,
            output_path=normalized_output_path,
        )

        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", powershell_script],
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
    finally:
        try:
            os.unlink(manifest_path)
        except OSError:
            pass

    return total_models


def _slide_to_manifest(slide: Model3DSlideSpec) -> dict:
    return {
        "textboxes": [
            {
                "text": tb.text,
                "left_pt": tb.left_pt,
                "top_pt": tb.top_pt,
                "width_pt": tb.width_pt,
                "height_pt": tb.height_pt,
                "font_size_pt": tb.font_size_pt,
                "bold": bool(tb.bold),
                "align": _ALIGN_MAP[tb.align.lower()],
                "font_name": tb.font_name,
                "color_rgb": tb.color_rgb,
            }
            for tb in slide.textboxes
        ],
        "models": [
            {
                "model_path": os.path.abspath(md.model_path),
                "left_pt": md.left_pt,
                "top_pt": md.top_pt,
                "width_pt": md.width_pt,
                "height_pt": md.height_pt,
                "rot_x_deg": md.rot_x_deg,
                "rot_y_deg": md.rot_y_deg,
                "rot_z_deg": md.rot_z_deg,
                "field_of_view_deg": md.field_of_view_deg,
            }
            for md in slide.models
        ],
    }


# msoFalse = 0 (do not link), msoTrue = -1 (embed in document). SaveWithDocument
# must be msoTrue when LinkToFile is msoFalse, so the deck is self-contained.
_POWERSHELL_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$manifestPath = '{manifest_path}'
$outputPath = '{output_path}'
$data = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$presentation = $app.Presentations.Add()
$buildError = $null
try {{
    $presentation.PageSetup.SlideWidth = [int][math]::Round($data.slide_width_pt)
    $presentation.PageSetup.SlideHeight = [int][math]::Round($data.slide_height_pt)
    foreach ($slideSpec in $data.slides) {{
        $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, 12)
        if ($data.background_rgb -ne $null) {{
            $slide.FollowMasterBackground = 0
            $slide.Background.Fill.Solid()
            $slide.Background.Fill.ForeColor.RGB = [int]$data.background_rgb
        }}
        foreach ($tb in $slideSpec.textboxes) {{
            $textShape = $slide.Shapes.AddTextbox(
                1,
                [single]$tb.left_pt,
                [single]$tb.top_pt,
                [single]$tb.width_pt,
                [single]$tb.height_pt
            )
            $textRange = $textShape.TextFrame.TextRange
            $textRange.Text = $tb.text
            $textRange.ParagraphFormat.Alignment = [int]$tb.align
            $textRange.Font.Size = [single]$tb.font_size_pt
            if ($tb.bold) {{ $textRange.Font.Bold = -1 }} else {{ $textRange.Font.Bold = 0 }}
            $textRange.Font.Name = $tb.font_name
            if ($tb.color_rgb -ne $null) {{
                $textRange.Font.Color.RGB = [int]$tb.color_rgb
            }}
        }}
        foreach ($md in $slideSpec.models) {{
            $modelShape = $slide.Shapes.Add3DModel(
                $md.model_path,
                0,
                -1,
                [single]$md.left_pt,
                [single]$md.top_pt,
                [single]$md.width_pt,
                [single]$md.height_pt
            )
            $model3D = $modelShape.Model3D
            $model3D.RotationX = [single]$md.rot_x_deg
            $model3D.RotationY = [single]$md.rot_y_deg
            $model3D.RotationZ = [single]$md.rot_z_deg
            if ($md.field_of_view_deg -ne $null) {{
                $model3D.FieldOfView = [single]$md.field_of_view_deg
            }}
        }}
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
