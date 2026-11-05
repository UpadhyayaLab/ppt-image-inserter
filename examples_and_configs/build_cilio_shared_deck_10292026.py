#!/usr/bin/env python
"""
The CilioD deck in the form of the unified Noco deck (FF 2026-11-01), one experiment, Oct 29: FF's
Noco deck gives the opening slides and the master, then the same order: montages, side-by-side
mesh pairs, synapse, results plots, H3K27me3 montages, the region slides, deepest region, and the
H3K27me3 plots, the slices after the Region Nadir 25% plot.

The Noco results text states Noco findings, so it stays out: the results slides carry a plain title
and the plot. The MT slides have no CilioD counterpart. The region slides show cells of the Noco experiment, as in
the ordered CilioD deck: no CilioD tiles are rendered.

Rebuild the montages (--montages) only when cells or tiles change.

Usage (base miniforge python):
    python examples_and_configs/build_cilio_shared_deck_10292026.py [--montages]
"""

import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import build_noco_shared_deck_two_experiments as n  # noqa: E402
import build_noco_two_experiment_montages as b  # noqa: E402
from build_h3k27me3_10292026_full_decks import copy_slide  # noqa: E402
from build_h3k27me3_10292026_mesh_decks import PAIR_FOOT, PAIR_TITLE  # noqa: E402
from ppt_image_inserter.backup import backup_presentation  # noqa: E402
from ppt_image_inserter.slide_utils import fast_image_lookup  # noqa: E402

m = b.m
CILIO = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/CilioD")
MONTAGES = CILIO / "CilioD, Jurkats, montages, 20261029.pptx"
OUT = CILIO / "CilioD, Jurkats, 20261029.pptx"
CONFIG = n.CONFIGS / "config_CilioD_H3K27me3_10292026_fixed_cells.yaml"
ORDERED = n.CONFIGS / "config_CilioD_H3K27me3_ordered_10292026.yaml"

TITLE_OLD = "Effects of High Dose Nocodazole Treatment on Jurkat T Cell Nuclear Morphology and Association with the Centrosome"
TITLE_NEW = "Effects of Ciliobrevin D Treatment on Jurkat T Cell Nuclear Morphology and Association with the Centrosome"
AUTHOR_NEW = "Based on experiment performed by Matthew Connell: October 29, 2026"
HULL_TITLE = "Hull Surf. Dist."
SIDE_VIEW_TITLE = "H3K27me3 and DNA, side view"
REGION_SLIDES = ("The regions measured, CD3 DMSO cell 10", "The regions measured, CD3 DMSO cell 34")
# (slide title, config title), or (slide title, panel in the compile) where the config has no such slide.
RESULTS = [("Centrosome-synapse distance", "Centrosome-synapse distance"),
           ("Fraction polarized (< 2 μm)", "Fraction polarized (< 2 μm)"),
           ("Invagination depth near centrosome", "Invagination depth near centrosome"),
           ("Min curvature near vs away from centrosome", "Min curvature near vs away from centrosome"),
           ("Mean Hull Surf. Dist.: Near / Away", "Mean Hull Surf. Dist.: Near / Away"),
           ("Synapse Area", "Synapse Area"),
           ("Actin inner/outer synapse ratio", "Actin inner/outer synapse ratio"),
           ("Nuclear aspect ratio", "Nuclear aspect ratio"),
           ("Nuclear volume", "Nuclear volume"),
           ("Nuclear surface area", "Nuclear surface area"),
           ("Invagination depth over the full nucleus", "grid_panels/chull_max_D_per_experiment_grid.png"),
           ("Invagination ratio", "grid_panels/invag_ratio_per_experiment_grid.png"),
           ("Concave Ratio", "grid_panels/concave_ratio_per_experiment_grid.png"),
           ("Nucleus-centrosome closest distance", "Nucleus-centrosome closest distance")]


def build_montages() -> int:
    """The Oct 29 half of the two-experiment montage deck, for CilioD."""
    prs = Presentation()
    fast_image_lookup(prs)
    prs.slide_width = Inches(m.SLIDE_W)
    prs.slide_height = Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    m.new_slide = b.dated_slide
    m.COMPARE_ALL_CELLS = set()
    m.STRIP_FOOT_FONT_PT = 15
    b.add_1029(sink, "cilio")
    lst = prs.slides._sldIdLst
    els = list(lst)
    for el in els:
        lst.remove(el)
    for k in sorted(range(len(els)), key=lambda k: b.ORDER.index(sink["tags"][k])):
        lst.append(els[k])
    if MONTAGES.exists():
        m.backup_presentation(str(MONTAGES))
    prs.save(str(MONTAGES))
    print(f"[CilioD montages] {len(prs.slides)} slides -> {MONTAGES}")
    if sink.get("empty"):
        print(f"[MISSING] {len(sink['empty'])} block(s): " + "; ".join(sink["empty"]))
        return 1
    return 0


def rework_overview(slide) -> None:
    """The Noco overview for this experiment. What only the Noco plan says stays out: the question on MT
    organization and the drug in the imaging chambers."""
    n.rework_overview(slide)
    paras = slide.placeholders[1].text_frame.paragraphs
    paras[0].runs[2].text = "H3K27me3"
    for run, text in zip(paras[3].runs[2:], ("50 ", "μM", " ciliobrevin D", "", "", " treatment. ")):
        run.text = text
    paras[4].runs[0].text = "30 min pre-incubation."
    paras[5].runs[0].text = "Does ciliobrevin D treatment:"
    paras[7].runs[0].text = "Lead to less prominent nuclear "
    paras[8].runs[2].text = "Does dynein activity influence "
    paras[6]._p.getparent().remove(paras[6]._p)


def build() -> int:
    if (OUT.parent / f"~${OUT.name}").exists():
        print(f"{OUT.name} is open in PowerPoint: not built")
        return 1
    config = yaml.safe_load(CONFIG.read_text(encoding="utf-8"))
    src = n.sources(config)
    ordered = yaml.safe_load(ORDERED.read_text(encoding="utf-8"))
    h3k_src, h3k_sub = n.sources(ordered), n.subtitles(ordered)
    h3k_src[None] = [f"{ordered['base_dir']}/{n.MFI_PANEL}"]
    panels = [src.get(key) or ([f"{config['base_dir']}/{key}"] if key.endswith(".png") else None) for _, key in RESULTS]
    missing = [t for t, p in zip(RESULTS, panels) if not p]
    missing += [t for t in REGION_SLIDES if not h3k_src.get(t)]
    missing += [t for _, titles in n.H3K_PLOTS for t in titles or () if not h3k_src.get(t)]
    if missing:
        print("no plots in the config for:", missing)
        return 1

    prs = Presentation(str(n.SHARED))
    fast_image_lookup(prs)
    lst = prs.slides._sldIdLst
    own = list(lst)
    slides = list(prs.slides)
    assert n.swap_text(slides[0], TITLE_OLD, TITLE_NEW) == 1
    assert n.swap_text(slides[0], n.AUTHOR_OLD, AUTHOR_NEW) == 1
    rework_overview(slides[1])

    source = list(Presentation(str(MONTAGES)).slides)
    n_regions = next(k for k, s in enumerate(source) if n.has_title(s, n.MESH_LEFT_OUT))
    source = [s for s in source if not n.has_title(s, n.MESH_LEFT_OUT)]
    pair_slides = [s for s in source if n.has_title(s, PAIR_TITLE)]
    assert pair_slides, "no pair slides in the montage deck"
    source = [s for s in source if s not in pair_slides]
    n_hull = 1 + max(k for k, s in enumerate(source) if n.has_title(s, HULL_TITLE))
    n_h3k = next(k for k, s in enumerate(source) if n.has_title(s, SIDE_VIEW_TITLE))
    start = len(lst)
    swapped = 0
    for slide in source:
        copy_slide(slide, prs)
        swapped += n.swap_text(prs.slides[len(prs.slides) - 1], n.HULL_FOOT_OLD, n.HULL_FOOT_NEW)
    assert swapped, "hull footer not found"
    added = list(lst)[start:]
    slices = [el for s, el in zip(source, added) if n.has_title(s, n.SLICE_TITLE)]
    assert len(slices) == 2, "slice slides not found"
    for slide in pair_slides:
        copy_slide(slide, prs)
        assert n.swap_text(prs.slides[len(prs.slides) - 1], PAIR_FOOT, n.PAIR_FOOT_NEW) == 1
    pairs = list(lst)[start + len(added):]

    results = [n.plot_slide(prs, title, None, paths) for (title, _), paths in zip(RESULTS, panels)]
    regions = [n.plot_slide(prs, t, h3k_sub[t], h3k_src[t]) for t in REGION_SLIDES]
    h3k = [n.plot_slide(prs, f"{title}, Oct 29", h3k_sub.get((titles or (None,))[0]),
                        [p for t in titles or (None,) for p in h3k_src[t]]) for title, titles in n.H3K_PLOTS]

    for el in own[2:]:                              # Noco's montages and results
        prs.part.drop_rel(el.rId)
    for el in list(lst):
        lst.remove(el)
    order = (own[:2] + added[:n_hull] + pairs + added[n_hull:n_h3k] + results
             + [el for el in added[n_h3k:n_regions] if el not in slices] + regions + added[n_regions:]
             + h3k[:n.N_THROUGH_NADIR_25] + slices + h3k[n.N_THROUGH_NADIR_25:])
    for el in order:
        lst.append(el)
    if OUT.exists():
        backup_presentation(str(OUT))
    n.drop_sections(prs)
    prs.save(str(OUT))
    print(f"{len(prs.slides)} slides: 2 opening, {n_h3k} montages, {len(pairs)} mesh pairs, {len(results)} results, "
          f"{len(added) - n_h3k} H3K27me3 montages, {len(regions)} region, {len(h3k)} H3K27me3 plots, "
          f"{swapped} hull footers -> {OUT}")
    return 0


if __name__ == "__main__":
    code = build_montages() if "--montages" in sys.argv else 0
    sys.exit(code or build())
