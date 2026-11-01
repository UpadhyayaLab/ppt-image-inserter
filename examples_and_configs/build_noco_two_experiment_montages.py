#!/usr/bin/env python
"""
One Noco montage deck with both experiments (FF 2026-11-01): 10/26/2026 (MT) and 10/29/2026
(H3K27me3), never pooled, each slide names its date. Slide order of the 10/26 montage deck
(COMBINED_ORDER in insert_noco_10262026_montage_slides.py), H3K27me3 at the end. Families in
both experiments: the Oct 26 slide, then the Oct 29 slide. MT families exist on Oct 26 only.
One page per family and date, 12 kept cells per condition (the 10/26 format); Hull Surf. Dist.
and the other mesh families show every cell, Partial cells left out on Oct 29.

Reuses the three builders; rebuild after a compile change.

Usage (base miniforge python):
    python examples_and_configs/build_noco_two_experiment_montages.py [--smoke]
--smoke builds the Oct 26 slides only, into the 10/26 experiment's tests folder.
"""

import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_h3k27me3_10292026_compare_decks as c  # noqa: E402
import build_h3k27me3_10292026_mesh_decks as mesh  # noqa: E402

from ppt_image_inserter.slide_utils import fast_image_lookup  # noqa: E402

m = c.m
OUT = Path(f"{c.OUT_ROOT}/Noco/Noco, Jurkats, montages, two experiments, 10262026 and 10292026.pptx")
OCT26, OCT29 = "Oct 26", "Oct 29"
COMPILE_1026 = "results_compilation_20261031"
HULL_FOOT = mesh.VIEWS[0][3]    # the 10/29 wording, on both dates
# (date, builder tag) in slide order.
ORDER = [(OCT26, "MT_xz_nolines_fixedclim"), (OCT26, "MT_bz_fixedclim"),
         (OCT26, "actin_cent_xz_nolines"), (OCT29, "actin_cent_xz_nolines"),
         (OCT26, "MT_cent_xz_nolines_fixedclim"),
         (OCT26, "cent_nuc_xz"), (OCT29, "cent_nuc_xz"),
         (OCT26, "actin_nuc_xz_nolines"), (OCT29, "actin_nuc_xz_nolines"),
         (OCT26, "hull"), (OCT29, "hull"),
         (OCT26, "synapse"), (OCT29, "synapse"),
         (OCT29, "H3K27me3_nuc_xz"), (OCT29, "H3K27me3_nuc_bz"), (OCT29, "h3k_side"),
         (OCT29, "h3k27me3"), (OCT29, "cent_proximal"), (OCT29, "deepest_by_cent"), (OCT29, "pairs")]

date_now = OCT26
plain_slide = m.new_slide


def dated_slide(prs, title: str, channels):
    """m.new_slide with the date in the title, before any k/n page number."""
    head, page = re.fullmatch(r"(.*?)( \d+/\d+)?", title).groups()
    return plain_slide(prs, f"{head}, {date_now}{page or ''}", channels)


def add_1026(sink: dict) -> None:
    m.COMPILE_TRACING = m.EXP_ROOT / COMPILE_1026 / "cell_tracing.csv"
    m.COMPARE_ALL_CELLS = set()     # one page per family
    m.HULL_FOOTER = HULL_FOOT
    m.build_compare_decks(sink)
    m.STRIP_FOOT_FONT_PT = 15
    m.build_hull_deck(None, sink, label_pt=mesh.MESH_LABEL_PT, bar_thick=mesh.TILE_BAR_THICK)


def add_1029(sink: dict) -> None:
    global date_now
    date_now = OCT29
    start = len(sink["tags"])
    c.add_slides("noco", sink, one_page=True)
    m.COMPARE_SOURCES = {w: c.PROG for w, _, _, _ in c.EXPS["noco"]["wells"]}
    m.build_synapse_deck(0, sink)
    depth = mesh.cent_depths("noco")
    m.STRIP_FOOT_FONT_PT = 15
    with ThreadPoolExecutor(mesh.WORKERS) as pool:
        mesh.overview("noco", sink, depth, pool)
        mesh.pairs("noco", sink, depth, pool)
    sink["tags"][start:] = [(OCT29, t) for t in sink["tags"][start:]]


def main() -> int:
    smoke = "--smoke" in sys.argv
    prs = Presentation()
    fast_image_lookup(prs)
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    m.new_slide = dated_slide
    add_1026(sink)
    sink["tags"][:] = [(OCT26, t) for t in sink["tags"]]
    if not smoke:
        add_1029(sink)
    # Slides were added in builder order; stable within a family.
    lst = prs.slides._sldIdLst
    els = list(lst)
    for el in els:
        lst.remove(el)
    for k in sorted(range(len(els)), key=lambda k: ORDER.index(sink["tags"][k])):
        lst.append(els[k])
    out = m.EXP_ROOT / "tests" / "two_experiment_montages_smoke.pptx" if smoke else OUT
    if out.exists() and not smoke:
        m.backup_presentation(str(out))
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"[two experiments] {len(prs.slides)} slides -> {out}")
    if sink.get("empty"):
        print(f"[MISSING] {len(sink['empty'])} block(s): " + "; ".join(sink["empty"]))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
