"""Movie-related helpers for PowerPoint decks.

This module contains:

- :class:`TextboxSpec`, :class:`MovieSpec`, :class:`SlideSpec`: simple
  dataclasses describing a slide's textboxes and embedded movies in points.
- :func:`build_movie_deck_via_com`: builds a deck from those specs using
  PowerPoint COM (Windows-only) and rewrites the timing tree on disk so each
  movie autoplays in slideshow mode.
- :func:`force_autoplay_in_pptx`: low-level OOXML rewrite that fixes the
  well-known "video does not autoplay" issue produced by both
  ``python-pptx``'s ``shapes.add_movie`` and PowerPoint COM's
  ``Shapes.AddMediaObject2``. Both pipelines write a ``<p:cTn>`` start
  condition of ``delay="indefinite"`` (click-to-start) on the trigger node
  above the media-play effect; this helper rewrites it to ``delay="0"``.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Sequence, Tuple

from lxml import etree


P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main"
_CTN_TAG = f"{{{P_NAMESPACE}}}cTn"
_STCONDLST_TAG = f"{{{P_NAMESPACE}}}stCondLst"
_COND_TAG = f"{{{P_NAMESPACE}}}cond"
_CMEDIANODE_TAG = f"{{{P_NAMESPACE}}}cMediaNode"

_ALIGN_MAP = {"left": 1, "center": 2, "right": 3, "justify": 4}


@dataclass(frozen=True)
class TextboxSpec:
    """One textbox to add to a slide. All measurements are in points."""

    text: str
    left_pt: float
    top_pt: float
    width_pt: float
    height_pt: float
    font_size_pt: float = 22.0
    bold: bool = True
    align: str = "center"
    font_name: str = "Arial"


@dataclass(frozen=True)
class MovieSpec:
    """One embedded movie on a slide. All measurements are in points."""

    movie_path: str
    poster_path: str
    left_pt: float
    top_pt: float
    width_pt: float
    height_pt: float


@dataclass(frozen=True)
class SlideSpec:
    """One slide's textboxes and movies."""

    textboxes: Sequence[TextboxSpec] = field(default_factory=tuple)
    movies: Sequence[MovieSpec] = field(default_factory=tuple)


def build_movie_deck_via_com(
    slides: Sequence[SlideSpec],
    output_path: str,
    slide_width_pt: float,
    slide_height_pt: float,
    loop: bool = False,
) -> int:
    """Build a PowerPoint deck of textboxes and embedded movies via COM.

    Each slide gets its textboxes added via ``Shapes.AddTextbox`` and each
    movie inserted via ``Shapes.AddMediaObject2`` with a poster frame from
    ``MediaFormat.SetDisplayPictureFromFile``. A media-play effect (effect
    id 83) is attached on the slide's main timeline sequence; the trigger
    condition is then rewritten on disk by :func:`force_autoplay_in_pptx`
    so the movie autoplays in slideshow mode.

    Args:
        slides: The slides to build, in order.
        output_path: Where to save the .pptx. Overwritten if it exists.
        slide_width_pt: Slide width in points (e.g. 13.333 in * 72 = 960).
        slide_height_pt: Slide height in points (e.g. 7.5 in * 72 = 540).
        loop: If True, also mark every embedded movie to loop until stopped via
            :func:`force_loop_in_pptx`. Defaults to False (unchanged behavior).

    Returns:
        The number of slides whose autoplay trigger was rewritten on disk.
    """
    if sys.platform != "win32":
        raise RuntimeError(
            "PowerPoint COM deck generation is only available on Windows"
        )

    manifest = {
        "slide_width_pt": slide_width_pt,
        "slide_height_pt": slide_height_pt,
        "slides": [_slide_to_manifest(slide) for slide in slides],
    }

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

    rewritten = force_autoplay_in_pptx(output_path)
    if loop:
        force_loop_in_pptx(output_path)
    return rewritten


def _slide_to_manifest(slide: SlideSpec) -> dict:
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
            }
            for tb in slide.textboxes
        ],
        "movies": [
            {
                "movie_path": os.path.abspath(mv.movie_path),
                "poster_path": os.path.abspath(mv.poster_path),
                "left_pt": mv.left_pt,
                "top_pt": mv.top_pt,
                "width_pt": mv.width_pt,
                "height_pt": mv.height_pt,
            }
            for mv in slide.movies
        ],
    }


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
        }}
        foreach ($mv in $slideSpec.movies) {{
            $mediaShape = $slide.Shapes.AddMediaObject2(
                $mv.movie_path,
                $false,
                $true,
                [single]$mv.left_pt,
                [single]$mv.top_pt,
                [single]$mv.width_pt,
                [single]$mv.height_pt
            )
            $mediaShape.MediaFormat.SetDisplayPictureFromFile($mv.poster_path)
            # Add Media Play (effect 83) on the main sequence; force_autoplay_in_pptx
            # then flips the trigger from delay="indefinite" to delay="0" on disk.
            $slide.TimeLine.MainSequence.AddEffect($mediaShape, 83, 0, 2) | Out-Null
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


def force_autoplay_in_pptx(pptx_path: str) -> int:
    """Rewrite click-to-start triggers on media-play effects to autoplay.

    Opens the .pptx as a zip, walks every ``ppt/slides/slideN.xml``, and for
    each ``<p:cTn>`` that wraps a ``presetClass="mediacall" presetID="1"``
    media-play effect rewrites the trigger's ``<p:cond delay="indefinite"/>``
    start condition to ``<p:cond delay="0"/>``. This is the canonical OOXML
    that PowerPoint UI generates for "Start: Automatically", and PowerPoint
    will not normalize it back on the next save.

    The fallback ``<p:cond evt="onBegin" ...>`` condition is left untouched.

    Args:
        pptx_path: Path to the .pptx file. The file is rewritten in place.

    Returns:
        The number of slides whose timing tree was modified.
    """
    tmp_path = pptx_path + ".autoplay.tmp"
    rewritten = 0
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(
        tmp_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _is_slide_xml(item.filename):
                new_data, changed = _patch_slide_xml(data)
                if changed:
                    rewritten += 1
                data = new_data
            zout.writestr(item, data)
    os.replace(tmp_path, pptx_path)
    return rewritten


def force_loop_in_pptx(pptx_path: str) -> int:
    """Mark every embedded movie to loop until stopped.

    Opens the .pptx as a zip, walks every ``ppt/slides/slideN.xml``, and for
    each ``<p:cMediaNode>`` sets ``repeatCount="indefinite"`` on its child
    ``<p:cTn>``. This is the exact OOXML the PowerPoint UI generates for the
    video "Loop until Stopped" playback option (verified via COM), so the movie
    restarts automatically when it reaches the end.

    Args:
        pptx_path: Path to the .pptx file. The file is rewritten in place.

    Returns:
        The number of slides whose timing tree was modified.
    """
    tmp_path = pptx_path + ".loop.tmp"
    rewritten = 0
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(
        tmp_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _is_slide_xml(item.filename):
                new_data, changed = _patch_slide_loop_xml(data)
                if changed:
                    rewritten += 1
                data = new_data
            zout.writestr(item, data)
    os.replace(tmp_path, pptx_path)
    return rewritten


def _patch_slide_loop_xml(xml_bytes: bytes) -> Tuple[bytes, bool]:
    """Return slide XML with every media node set to loop, and a changed flag."""
    root = etree.fromstring(xml_bytes)
    changed = False

    for media_node in root.iter(_CMEDIANODE_TAG):
        media_ctn = media_node.find(_CTN_TAG)
        if media_ctn is None:
            continue
        if media_ctn.get("repeatCount") == "indefinite":
            continue
        media_ctn.set("repeatCount", "indefinite")
        changed = True

    if not changed:
        return xml_bytes, False

    out = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return out, True


def _is_slide_xml(name: str) -> bool:
    """Return True for ``ppt/slides/slideN.xml`` entries (not slideLayouts)."""
    if not name.startswith("ppt/slides/slide"):
        return False
    if not name.endswith(".xml"):
        return False
    return "/" not in name[len("ppt/slides/") :]


def _patch_slide_xml(xml_bytes: bytes) -> Tuple[bytes, bool]:
    """Return possibly-rewritten slide XML and a changed flag."""
    root = etree.fromstring(xml_bytes)
    changed = False

    for play_ctn in root.iter(_CTN_TAG):
        if play_ctn.get("presetClass") != "mediacall":
            continue
        if play_ctn.get("presetID") != "1":
            continue

        trigger_ctn = _find_trigger_ctn(play_ctn)
        if trigger_ctn is None:
            continue

        cond_list = trigger_ctn.find(_STCONDLST_TAG)
        if cond_list is None:
            continue

        for cond in cond_list.findall(_COND_TAG):
            if cond.get("evt") is not None:
                continue
            if cond.get("delay") != "indefinite":
                continue
            cond.set("delay", "0")
            changed = True

    if not changed:
        return xml_bytes, False

    out = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    return out, True


def _find_trigger_ctn(play_ctn: etree._Element):
    """Walk up the cTn chain to the trigger cTn directly under mainSeq.

    The OOXML structure is::

        <p:cTn id="2" nodeType="mainSeq">
          <p:childTnLst><p:par>
            <p:cTn id="3" ...>          # trigger (target of this walk)
              <p:childTnLst><p:par>
                <p:cTn id="4" ...>
                  <p:childTnLst><p:par>
                    <p:cTn id="5" presetClass="mediacall" .../>  # play_ctn

    Each step up walks one cTn level: cTn -> par -> childTnLst -> cTn.
    Stops when the next cTn ancestor is the mainSeq node.
    """
    current = play_ctn
    while current is not None:
        par = current.getparent()
        if par is None:
            return None
        child_tn_lst = par.getparent()
        if child_tn_lst is None:
            return None
        ancestor_ctn = child_tn_lst.getparent()
        if ancestor_ctn is None or ancestor_ctn.tag != _CTN_TAG:
            return None
        if ancestor_ctn.get("nodeType") == "mainSeq":
            return current
        current = ancestor_ctn
    return None
