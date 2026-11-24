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
from typing import Optional, Sequence, Tuple

from lxml import etree


P_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main"
_CTN_TAG = f"{{{P_NAMESPACE}}}cTn"
_STCONDLST_TAG = f"{{{P_NAMESPACE}}}stCondLst"
_COND_TAG = f"{{{P_NAMESPACE}}}cond"
_CMEDIANODE_TAG = f"{{{P_NAMESPACE}}}cMediaNode"

_ALIGN_MAP = {"left": 1, "center": 2, "right": 3, "justify": 4}
_ANCHOR_MAP = {"top": 1, "middle": 3, "bottom": 4}   # MsoVerticalAnchor


def make_run_scratch_dir(base_dir, max_age_s: float = 86400.0) -> str:
    """Create a fresh scratch folder for ONE deck build (posters, held movie copies)
    under ``base_dir`` and return its path; the caller removes it when done.

    Per run, not shared: scripts that wiped and reused one fixed folder let two runs at
    once overwrite each other's files (movie names repeat across experiments), so a deck
    could embed another experiment's movie (bug hunt 2026-10-24). Run folders older than
    ``max_age_s`` (left by crashed runs) are pruned.
    """
    import shutil
    import time

    os.makedirs(base_dir, exist_ok=True)
    cutoff = time.time() - max_age_s
    for name in os.listdir(base_dir):
        path = os.path.join(base_dir, name)
        marker = os.path.join(path, _SCRATCH_MARKER)
        try:
            # Only folders this function made (marker inside): base_dir can be a user-given
            # --poster-dir, where any 'run_*' folder was deleted (bug hunt 2026-10-24, 2nd pass).
            if (name.startswith("run_") and os.path.isfile(marker)
                    and os.path.getmtime(marker) < cutoff):
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass
    run_dir = tempfile.mkdtemp(prefix=f"run_{os.getpid()}_", dir=base_dir)
    with open(os.path.join(run_dir, _SCRATCH_MARKER), "w", encoding="utf-8") as fh:
        fh.write("ppt_image_inserter per-run scratch; safe to delete\n")
    return run_dir


_SCRATCH_MARKER = ".ppt_image_inserter_scratch"


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
    # PowerPoint font color as a BGR integer (as VBA RGB()), or None for the
    # theme default (dark text). Use e.g. 0xFFFFFF (white) on dark slides.
    color_rgb: Optional[int] = None
    italic: bool = False
    # Optional, default None = unchanged behavior (PowerPoint's auto-fit box, top-anchored).
    # anchor "top" | "middle" | "bottom": the box keeps its given height (no auto-fit) and the text sits
    # at that edge, as python-pptx vertical_anchor. space_after_pt: space after each paragraph
    # (separate paragraphs in ``text`` with "\r").
    anchor: Optional[str] = None
    space_after_pt: Optional[float] = None
    # Optional, default None = unchanged. Paragraphs that start with a typed bullet ("•") get a hanging indent of
    # this many points, so a wrapped line starts under the text, not under the mark; other paragraphs stay flush left.
    hanging_indent_pt: Optional[float] = None
    # Optional sub-bullets (used only with hanging_indent_pt). para_levels: one level per paragraph of ``text``, 0 or
    # 1; a level-1 paragraph is a sub-bullet, indented under the bullet text, with its own hanging indent
    # (sub_indent_pt) and font size (sub_font_size_pt); a paragraph followed by a sub-bullet gets sub_space_after_pt
    # after it. main_bold: when set, level-0 bullet paragraphs are bold (True) or not, and sub-bullets never bold
    # (FF 2026-11-23: "bold stuff in main bllets -but not in subbelets"). Sub-bullets keep a bullet mark, never a dash
    # (FF 2026-11-23: "use subbullets not -").
    para_levels: Optional[Sequence[int]] = None
    sub_indent_pt: Optional[float] = None
    sub_font_size_pt: Optional[float] = None
    sub_space_after_pt: Optional[float] = None
    main_bold: Optional[bool] = None
    # Optional, default None = PowerPoint's own inner margins (7.2 pt left/right). A number sets the left and right
    # inner margins of the box, e.g. 0 for narrow number columns.
    side_margin_pt: Optional[float] = None
    # Optional, default None = PowerPoint's own inner margins (3.6 pt top/bottom). A number sets the top and bottom
    # inner margins, e.g. 0 so a box is exactly its lines' height and stacked boxes can sit close without overlapping.
    vert_margin_pt: Optional[float] = None
    # Optional, default None = unchanged. subscripts: (start, length) character ranges of ``text`` (1-based, a "\r"
    # paragraph break counts as one character) set as subscript, e.g. the "under" of I_under. line_spacing_pt: exact
    # line pitch in points, so lines of side-by-side boxes stay level (subscripts can otherwise change a line's height).
    subscripts: Optional[Sequence[Tuple[int, int]]] = None
    line_spacing_pt: Optional[float] = None


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
class PictureSpec:
    """One still image (PNG/JPEG) on a slide, embedded in the deck. All measurements are in points."""

    image_path: str
    left_pt: float
    top_pt: float
    width_pt: float
    height_pt: float


@dataclass(frozen=True)
class SlideSpec:
    """One slide's textboxes, movies and still pictures."""

    textboxes: Sequence[TextboxSpec] = field(default_factory=tuple)
    movies: Sequence[MovieSpec] = field(default_factory=tuple)
    pictures: Sequence[PictureSpec] = field(default_factory=tuple)


def build_movie_deck_via_com(
    slides: Sequence[SlideSpec],
    output_path: str,
    slide_width_pt: float,
    slide_height_pt: float,
    loop: bool = False,
    restore_pictures: bool = False,
    timer=None,
) -> int:
    """Build a PowerPoint deck of textboxes and embedded movies via COM.

    To change a few slides of a deck built before, see :func:`ppt_image_inserter.incremental.update_deck`: it builds
    only the changed slides with this function and splices them into the existing deck.

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
        restore_pictures: If True, put each still picture's source file bytes back into the deck after the save
            (in the same single rewrite as the autoplay patch; PowerPoint downsamples some pictures). Defaults to False (unchanged).
        timer: Optional :class:`ppt_image_inserter.stage_timer.StageTimer`; times the PowerPoint build and save and
            each on-disk rewrite as stages.

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

    from contextlib import nullcontext
    stage = timer.stage if timer else (lambda name: nullcontext())
    try:
        normalized_manifest_path = manifest_path.replace("'", "''")
        normalized_output_path = os.path.abspath(output_path).replace("'", "''")
        powershell_script = _POWERSHELL_TEMPLATE.format(
            manifest_path=normalized_manifest_path,
            output_path=normalized_output_path,
        )

        try:
            with stage(f"PowerPoint build and save ({sum(len(s.movies) for s in slides)} movies, "
                       f"{len(slides)} slides)"):
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

    restore = restore_pictures and any(s.pictures for s in slides)
    with stage("autoplay" + (", loop" if loop else "") + (", source pictures" if restore else "")
               + " (one rewrite of the file)"):
        replace = _source_picture_members(output_path, slides) if restore else {}
        rewritten = _rewrite_pptx_once(output_path, loop, replace)
    if restore:
        with stage("check source pictures"):
            _check_source_pictures(output_path, slides)
    return rewritten


_STORED_MEDIA_EXTS = (".mp4", ".m4v", ".mov", ".wmv", ".avi", ".mp3", ".m4a", ".jpeg", ".jpg", ".png", ".gif")


def _rewrite_pptx_once(pptx_path: str, loop: bool, replace: dict) -> int:
    """force_autoplay_in_pptx, force_loop_in_pptx (if loop) and replacing zip members by given bytes, in ONE rewrite
    of the file; already-compressed media are stored, not deflated again (2026-11-23, 481 MB deck: three separate
    rewrites took 71 s). Returns the number of slides whose autoplay trigger was rewritten."""
    tmp_path = pptx_path + ".rewrite.tmp"
    rewritten = 0
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = replace.get(item.filename)
            if data is None:
                data = zin.read(item.filename)
            if _is_slide_xml(item.filename):
                data, changed = _patch_slide_xml(data)
                rewritten += bool(changed)
                if loop:
                    data, _ = _patch_slide_loop_xml(data)
            stored = item.filename.lower().endswith(_STORED_MEDIA_EXTS)
            zout.writestr(item.filename, data, compress_type=zipfile.ZIP_STORED if stored else zipfile.ZIP_DEFLATED)
    os.replace(tmp_path, pptx_path)
    return rewritten


def _source_picture_members(pptx_path: str, slides: Sequence["SlideSpec"]) -> dict:
    """{zip member name of each still picture's image part: its source file's bytes} (slide k = slides[k], pictures
    matched in shape order), for _rewrite_pptx_once: PowerPoint downsamples some pictures on save."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(pptx_path)
    out = {}
    for k, spec in enumerate(slides):
        pics = [s for s in prs.slides[k].shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
        for shape, pic in zip(pics, spec.pictures):
            with open(pic.image_path, "rb") as f:
                out[prs.slides[k].part.related_part(shape._element.blip_rId).partname.membername] = f.read()
    return out


def _check_source_pictures(pptx_path: str, slides: Sequence["SlideSpec"]) -> None:
    """Raise RuntimeError unless every still picture is byte-identical to its source file."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(pptx_path)
    for k, spec in enumerate(slides):
        pics = [s for s in prs.slides[k].shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
        for shape, pic in zip(pics, spec.pictures):
            with open(pic.image_path, "rb") as f:
                if shape.image.blob != f.read():
                    raise RuntimeError(f"slide {k + 1} picture is not byte-identical to {pic.image_path}")


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
                "italic": bool(tb.italic),
                "align": _ALIGN_MAP[tb.align.lower()],
                "font_name": tb.font_name,
                "color_rgb": tb.color_rgb,
                "anchor": _ANCHOR_MAP[tb.anchor.lower()] if tb.anchor else None,
                "space_after_pt": tb.space_after_pt,
                "hanging_indent_pt": tb.hanging_indent_pt,
                "sub_indent_pt": tb.sub_indent_pt,
                "sub_font_size_pt": tb.sub_font_size_pt,
                "sub_space_after_pt": tb.sub_space_after_pt,
                "para_levels": list(tb.para_levels) if tb.para_levels is not None else None,
                "main_bold": tb.main_bold,
                "side_margin_pt": tb.side_margin_pt,
                "vert_margin_pt": tb.vert_margin_pt,
                "subscripts": [list(s) for s in tb.subscripts] if tb.subscripts else None,
                "line_spacing_pt": tb.line_spacing_pt,
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
        "pictures": [
            {
                "image_path": os.path.abspath(pic.image_path),
                "left_pt": pic.left_pt,
                "top_pt": pic.top_pt,
                "width_pt": pic.width_pt,
                "height_pt": pic.height_pt,
            }
            for pic in slide.pictures
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
            if ($null -ne $tb.side_margin_pt) {{
                $textShape.TextFrame.MarginLeft = [single]$tb.side_margin_pt
                $textShape.TextFrame.MarginRight = [single]$tb.side_margin_pt
            }}
            if ($null -ne $tb.vert_margin_pt) {{
                $textShape.TextFrame.MarginTop = [single]$tb.vert_margin_pt
                $textShape.TextFrame.MarginBottom = [single]$tb.vert_margin_pt
            }}
            $textRange = $textShape.TextFrame.TextRange
            $textRange.Text = $tb.text
            $textRange.ParagraphFormat.Alignment = [int]$tb.align
            $textRange.Font.Size = [single]$tb.font_size_pt
            if ($tb.bold) {{ $textRange.Font.Bold = -1 }} else {{ $textRange.Font.Bold = 0 }}
            if ($tb.italic) {{ $textRange.Font.Italic = -1 }} else {{ $textRange.Font.Italic = 0 }}
            $textRange.Font.Name = $tb.font_name
            if ($null -ne $tb.color_rgb) {{ $textRange.Font.Color.RGB = [int]$tb.color_rgb }}
            if ($null -ne $tb.line_spacing_pt) {{
                $textRange.ParagraphFormat.LineRuleWithin = 0
                $textRange.ParagraphFormat.SpaceWithin = [single]$tb.line_spacing_pt
            }}
            if ($null -ne $tb.subscripts) {{
                foreach ($sub in $tb.subscripts) {{
                    $textRange.Characters([int]$sub[0], [int]$sub[1]).Font.Subscript = -1
                }}
            }}
            if ($null -ne $tb.space_after_pt) {{
                $textRange.ParagraphFormat.LineRuleAfter = 0
                $textRange.ParagraphFormat.SpaceAfter = [single]$tb.space_after_pt
            }}
            if ($null -ne $tb.hanging_indent_pt) {{
                $tr2 = $textShape.TextFrame2.TextRange
                $nPara = $tr2.Paragraphs().Count
                $levels = $tb.para_levels
                if (($null -ne $levels) -and ($levels.Count -ne $nPara)) {{
                    throw "para_levels has $($levels.Count) entries, text has $nPara paragraphs"
                }}
                for ($ip = 1; $ip -le $nPara; $ip++) {{
                    $para = $tr2.Paragraphs($ip, 1)
                    $lvl = 0; if ($null -ne $levels) {{ $lvl = [int]$levels[$ip - 1] }}
                    if (($lvl -eq 1) -and ($null -ne $tb.sub_indent_pt)) {{
                        $para.ParagraphFormat.LeftIndent = [single]$tb.hanging_indent_pt + [single]$tb.sub_indent_pt
                        $para.ParagraphFormat.FirstLineIndent = -[single]$tb.sub_indent_pt
                        if ($null -ne $tb.sub_font_size_pt) {{ $para.Font.Size = [single]$tb.sub_font_size_pt }}
                        if ($null -ne $tb.main_bold) {{ $para.Font.Bold = 0 }}
                    }} elseif ($para.Text.StartsWith([string][char]0x2022)) {{
                        $para.ParagraphFormat.LeftIndent = [single]$tb.hanging_indent_pt
                        $para.ParagraphFormat.FirstLineIndent = -[single]$tb.hanging_indent_pt
                        if ($null -ne $tb.main_bold) {{ if ($tb.main_bold) {{ $para.Font.Bold = -1 }} else {{ $para.Font.Bold = 0 }} }}
                    }}
                    if (($null -ne $tb.sub_space_after_pt) -and ($null -ne $levels) -and ($ip -lt $nPara) -and
                        ([int]$levels[$ip] -eq 1)) {{
                        $para.ParagraphFormat.SpaceAfter = [single]$tb.sub_space_after_pt
                    }}
                }}
            }}
            if ($null -ne $tb.anchor) {{
                $textShape.TextFrame.AutoSize = 0
                $textShape.TextFrame.WordWrap = -1
                $textShape.Top = [single]$tb.top_pt
                $textShape.Height = [single]$tb.height_pt
                $textShape.TextFrame.VerticalAnchor = [int]$tb.anchor
            }}
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
        foreach ($pic in $slideSpec.pictures) {{
            # LinkToFile = msoFalse (0), SaveWithDocument = msoTrue (-1): the image is embedded in the deck
            $picShape = $slide.Shapes.AddPicture(
                $pic.image_path, 0, -1,
                [single]$pic.left_pt, [single]$pic.top_pt, [single]$pic.width_pt, [single]$pic.height_pt
            )
            # msoSendToBack (1): text boxes placed over a picture (e.g. legend labels) stay visible
            $picShape.ZOrder(1) | Out-Null
        }}
    }}
    $presentation.SaveAs($outputPath, 24)
}} catch {{
    $buildError = $_
    throw
}} finally {{
    # Close our own presentation, discarding it on a build error: left open, it lingered
    # in the user's PowerPoint and kept the Quit guard below from ever firing.
    try {{
        if ($presentation) {{
            if ($buildError) {{ $presentation.Saved = -1 }}
            $presentation.Close()
        }}
    }} catch {{
    }}
    # PowerPoint is single-instance: New-Object attached to the user's running PowerPoint
    # if there was one. Quit only when no presentation is left open (the user's decks or a
    # concurrent build's); an unconditional Quit closed them (bug hunt 2026-10-24).
    try {{
        if ($app -and $app.Presentations.Count -eq 0) {{ $app.Quit() }}
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
