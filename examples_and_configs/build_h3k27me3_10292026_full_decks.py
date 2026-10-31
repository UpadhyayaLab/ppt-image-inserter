#!/usr/bin/env python
"""
One montage deck per 10/29/2026 H3K27me3 set (FF 2026-10-31: "separate full deck into
montages + results"): the comparison montages, then the mesh overview and pair slides. The results
stay in their own deck. --with-results instead builds one deck with the results after the
montages, the results deck being the base.

Assembles the finished decks and draws nothing itself, so it never reads the compile:
rebuild a source deck, then rerun this. Per-cell mesh rows and side views stay in their own
decks, as on 10/26. The source decks are not changed.

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_full_decks.py [noco|cilio|all] [--with-results]
"""

import io
import sys
from copy import deepcopy
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_h3k27me3_10292026_compare_decks import EXPS, OUT_ROOT, m  # noqa: E402
from build_h3k27me3_10292026_mesh_decks import OUT as MESH_OUT  # noqa: E402

RESULTS = {"noco": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Effects of Noco, H3K27me3, 10292026.pptx",
           "cilio": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/Effects of CilioD, H3K27me3, 10292026.pptx"}
OUT = {"noco": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Noco, Jurkats, H3K27me3, 20261029.pptx",
       "cilio": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/CilioD, Jurkats, H3K27me3, 20261029.pptx"}
MONTAGES = {"noco": f"{OUT_ROOT}/Noco/H3K27me3_20261029/Noco, Jurkats, H3K27me3, montages, 20261029.pptx",
            "cilio": f"{OUT_ROOT}/CilioD/H3K27me3_20261029/CilioD, Jurkats, H3K27me3, montages, 20261029.pptx"}
TITLE_SLIDES = {"noco": 1, "cilio": 0}   # slides kept ahead of the montages; CilioD has no title slide yet
ROWS_MARK = "mesh views"                 # per-cell mesh slides, left out


def title_of(slide) -> str:
    """The slide's title: its first text box with text."""
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip():
            return shape.text_frame.text
    return ""


def copy_slide(src, prs) -> None:
    """Append a black copy of src: pictures through the package, other shapes as they are."""
    dst = prs.slides.add_slide(prs.slide_layouts[6])
    m.set_slide_background(dst, m.BLACK)
    for shape in src.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            pic = dst.shapes.add_picture(io.BytesIO(shape.image.blob), shape.left, shape.top,
                                         shape.width, shape.height)
            pic.crop_left, pic.crop_right = shape.crop_left, shape.crop_right
            pic.crop_top, pic.crop_bottom = shape.crop_top, shape.crop_bottom
        else:
            el = deepcopy(shape._element)
            dst.shapes._spTree.append(el)
            ids = [int(i) for i in dst.shapes._spTree.xpath(".//p:cNvPr/@id")]
            el.xpath(".//p:cNvPr")[0].set("id", str(max(ids) + 1))


def montage_order(compare, meshes) -> list:
    """The comparison montages, then the mesh slides without the per-cell rows."""
    return compare + [s for s in meshes if ROWS_MARK not in title_of(s)]


def build(name: str, with_results: bool) -> int:
    out = Path((OUT if with_results else MONTAGES)[name])
    lock = out.parent / f"~${out.name}"
    if lock.exists():
        print(f"[full {name}] {out.name} is open in PowerPoint: not built")
        return 0
    if with_results:
        prs = Presentation(RESULTS[name])
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(m.SLIDE_W), Inches(m.SLIDE_H)
    n_results = len(prs.slides)
    compare = list(Presentation(EXPS[name]["out"]).slides)
    if Path(MESH_OUT[name]).exists():
        meshes = list(Presentation(MESH_OUT[name]).slides)
    else:                       # CilioD renders land last; rerun then
        print(f"[full {name}] no mesh deck yet: {MESH_OUT[name]}")
        meshes = []
    added = montage_order(compare, meshes)
    for slide in added:
        copy_slide(slide, prs)
    # Added at the end; move them to just after the title slide.
    lst = prs.slides._sldIdLst
    els = list(lst)
    for el in els:
        lst.remove(el)
    k = TITLE_SLIDES[name] if with_results else 0
    for el in els[:k] + els[n_results:] + els[k:n_results]:
        lst.append(el)
    if out.exists():
        m.backup_presentation(str(out))
    prs.save(str(out))
    print(f"[{name}] {len(prs.slides)} slides: {k} title, {len(added)} montage and mesh, "
          f"{n_results - k} results -> {out}")
    return len(prs.slides)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--with-results"]
    which = args[0] if args else "all"
    names = list(OUT) if which == "all" else [which]
    sys.exit(0 if all(build(n, "--with-results" in sys.argv) for n in names) else 1)
