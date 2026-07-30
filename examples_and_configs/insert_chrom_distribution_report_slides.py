"""
insert_chrom_distribution_report_slides.py

Build the CHROMATIN DISTRIBUTION report deck from the hand-curated report set:

  chromatin-analysis-figures/cross_experiment/nucleus_morphology_3expt/report_violins/
      1_dna_distribution/
      2_mark_distribution/
      3_mark_dna_codistribution/
      4_mark_intensity/
      5_nuclear_size_shape/

This is a STRAIGHT MIRROR of that folder: sections follow the numbered folder
order and metrics follow file order within each folder. Unlike the other chromark
decks, the shared classifier's family grouping, metric ordering and drop-lists are
deliberately NOT applied here -- the folder IS the curation. `classify()` is used
only to look up a nice display title, with a filename fallback.

Every PNG is a 3-column grid (one column per acquisition: H3K27me3 2024-03-21 |
H3K9me3 2024-04-24 | H3K27ac 2024-05-31). Each metric has three views:
    grid_<stem>.png / _stiffness.png / _timepoint.png
One full-size plot per slide.

Usage:
    conda run -n PPT_editing python examples_and_configs/insert_chrom_distribution_report_slides.py
    ... --list     # dry run: sections/metrics/titles, build nothing
"""

import os
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import insert_chromark_h3k27me3_summary_slides as M  # noqa: E402

SRC = Path(
    "J:/FF/fixed_cell/CTL_nucleus/tifsFixed3SIactivatedCTLs_nucleus/"
    "chromatin-analysis-figures/cross_experiment/nucleus_morphology_3expt/report_violins"
)
OUT = Path("K:/FF/PPT/PPT_autogeneration/Naive_CTL/chromark/"
           "CTL_chrom_distribution_report_3expt.pptx")

DECK_TITLE = "Chromatin distribution — cross-experiment report"
DECK_SUBTITLE = (
    "Curated report set  -  every panel: 3 columns = H3K27me3 (2024-03-21) | "
    "H3K9me3 (2024-04-24) | H3K27ac (2024-05-31)"
)

# Folder name -> section divider title. Folder order (1..5) is the deck order.
SECTION_TITLES = {
    "1_dna_distribution": "DNA distribution",
    "2_mark_distribution": "Chromatin mark distribution",
    "3_mark_dna_codistribution": "Mark–DNA co-distribution",
    "4_mark_intensity": "Chromatin mark intensity",
    "5_nuclear_size_shape": "Nuclear size & shape",
    "6_peripheral_chromatin": "Peripheral chromatin",
}

PREFIX = "grid_"
VIEWS = [
    ("", "all conditions"),
    ("_stiffness", "stiffness comparison"),
    ("_timepoint", "timepoint comparison"),
]

MARK_LABEL = "Chromatin mark"


def section_title(folder_name):
    if folder_name in SECTION_TITLES:
        return SECTION_TITLES[folder_name]
    # fall back to the folder name minus its numeric prefix
    name = folder_name.split("_", 1)[-1] if folder_name[:1].isdigit() else folder_name
    return name.replace("_", " ").capitalize()


def metric_title(stem):
    rec = M.classify(stem)
    if rec:
        return rec["title"]
    return stem.replace("_", " ")


def _natkey(text):
    """Natural sort: '..._3' before '..._10'. Ints and strings are wrapped so
    mixed tuples stay comparable."""
    return tuple((0, int(t)) if t.isdigit() else (1, t)
                 for t in re.split(r"(\d+)", text))


def _order_key(stem):
    """Within a section: 3D first, then max-area slice, then 2D/MIP; within a
    scope, group by the metric itself so a DNA panel is immediately followed by
    its chromatin-mark counterpart."""
    if "_2dmid" in stem:
        scope = 1
    elif "_2d_" in stem or stem.endswith("_2d"):
        scope = 2
    elif "_3d" in stem:
        scope = 0
    else:
        scope = 3                      # no scope token (e.g. mark_dna_ratio)
    channel = 1 if "mark" in stem else 0   # DNA before mark
    base = stem
    for tok in ("hoechst_", "mark_", "_hoechst", "_mark"):
        base = base.replace(tok, "_")
    for tok in ("_2dmid_", "_3d_", "_2d_"):
        base = base.replace(tok, "_")
    base = base.replace("chan_", "").replace("_int", "").strip("_")
    return (scope, _natkey(base), channel)


def discover(folder):
    """[(stem, {available view suffixes})] in file order (no re-sorting)."""
    known = [s for s, _ in VIEWS if s]
    order, views = [], {}
    for p in sorted(folder.glob(PREFIX + "*.png")):
        stem = p.stem[len(PREFIX):]
        for suf in known:
            if stem.endswith(suf):
                stem, suf_found = stem[: -len(suf)], suf
                break
        else:
            suf_found = ""
        if stem not in views:
            views[stem] = set()
            order.append(stem)
        views[stem].add(suf_found)
    return [(s, views[s]) for s in sorted(order, key=_order_key)]


def main():
    list_only = "--list" in sys.argv[1:]
    M.MARK_LABEL = MARK_LABEL
    # Specific antibody tokens keep their own names; the generic "mark" token
    # (used by this report set) still renders as MARK_LABEL.
    M.FORCE_MARK_LABEL = False

    sections = []
    for folder in sorted(d for d in SRC.iterdir() if d.is_dir()):
        items = discover(folder)
        if items:
            sections.append((folder, section_title(folder.name), items))

    n_metrics = sum(len(i) for _, _, i in sections)
    n_plots = sum(len(v) for _, _, items in sections for _, v in items)
    print("source: {}".format(SRC))
    print("{} sections, {} metrics, {} plot slides\n".format(
        len(sections), n_metrics, n_plots))

    for _, title, items in sections:
        print("=== {} ({}) ===".format(title, len(items)))
        for stem, v in items:
            print("   {:<40} {:<52} views={}".format(
                stem, metric_title(stem), len(v)))
        print("")

    if list_only:
        return

    prs = Presentation()
    prs.slide_width = Inches(M.SLIDE_W)
    prs.slide_height = Inches(M.SLIDE_H)
    M.build_title_slide(prs, DECK_TITLE, DECK_SUBTITLE)

    missing = []
    for folder, title, items in sections:
        M.build_divider_slide(prs, title)
        for stem, avail in items:
            for suf, vlabel in VIEWS:
                if suf not in avail:
                    continue
                p = folder / "{}{}{}.png".format(PREFIX, stem, suf)
                if not p.exists():
                    missing.append(p.name)
                    continue
                M.build_slide(prs, "{} — {}".format(metric_title(stem), vlabel), p,
                              "{}/{}".format(folder.name, p.name))

    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        M.backup_presentation(str(OUT), backup_base=str(OUT.parent / "backups"))
    prs.save(str(OUT))
    print("Done. {} slides -> {}".format(len(prs.slides._sldIdLst), OUT))
    if missing:
        print("\nMissing {}: {}".format(len(missing), missing))
    else:
        print("All panels found.")


if __name__ == "__main__":
    main()
