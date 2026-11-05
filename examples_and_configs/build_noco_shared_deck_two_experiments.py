#!/usr/bin/env python
"""
The unified Noco deck for the two experiments (FF 2026-11-01), built on the deck FF shared with
colleagues, "Noco, Jurkats, 20261026.pptx", which is read and left as it is. Its order and its
wording stay: slides 1-2, then the montages, one page per experiment (from the two-experiment
montage deck, build_noco_two_experiment_montages.py), then its results slides 19-32 with their
text untouched and both experiments' plots in place of the one. Slide 20 becomes two slides,
distance then fraction polarized. The H3K27me3 montages close the deck.

FF's later edits (2026-11-01): slide 2 names both stains and asks the H3K27me3 question, and the
hull / H3K27me3 pair slides of the Oct 29 mesh deck follow the Oct 29 hull pages, the two region
slides replace the Cent.-Prox. Reg. mesh pages, and FF's choice of the H3K27me3 plots closes the
deck, H3K27me3 and DNA on one slide, in the wording of the ordered H3K27me3 config.

Plots come from the two-experiment results config, so a rebuilt panel lands here on the next run.
Rebuild the montage deck first, and only when cells or tiles change; a text or plot edit here takes
one run of this script.

Usage (base miniforge python):
    python examples_and_configs/build_noco_shared_deck_two_experiments.py
"""

import copy
import sys
from pathlib import Path

import yaml
from PIL import Image
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from build_h3k27me3_10292026_full_decks import copy_slide  # noqa: E402
from build_h3k27me3_10292026_mesh_decks import PAIR_FOOT, PAIR_TITLE, VIEWS  # noqa: E402
from ppt_image_inserter.backup import backup_presentation  # noqa: E402
from ppt_image_inserter.slide_utils import duplicate_slide, fast_image_lookup, remove_pictures_from_slide  # noqa: E402
from ppt_image_inserter.workflow import _add_title, _content_crop, _to_open_path  # noqa: E402

NOCO = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Noco")
SHARED = NOCO / "Noco, Jurkats, 20261026.pptx"
MONTAGES = NOCO / "Noco, Jurkats, montages, two experiments, 10262026 and 10292026.pptx"
MESH_LEFT_OUT = "Cent.-Prox. Reg., Oct 29"
OUT = NOCO / "Noco, Jurkats, 20261026 and 20261029.pptx"
CONFIGS = Path(__file__).resolve().parent / "configs/fixed/Jurkats"
CONFIG = CONFIGS / "config_Noco_results_two_experiments_20261101.yaml"
ORDERED = CONFIGS / "config_Noco_H3K27me3_ordered_10292026.yaml"

N_SHARED_MONTAGE = 20       # montage deck slides before the H3K27me3 ones
N_THROUGH_HULL = 18         # montage deck slides through "Hull Surf. Dist., Oct 29 3/3"
# FF 2026-11-01: the slices print each cell's Region Nadir 25% value, so they follow that plot.
SLICE_TITLE = "H3K27me3, slice through Cent.-Prox. Reg."
N_THROUGH_NADIR_25 = 3      # H3K_PLOTS through "H3K27me3 and DNA around the Region Nadir, 25% region"
AUTHOR_OLD = "Based on experiment performed by Matthew Connell: October 26, 2026"
AUTHOR_NEW = "Based on experiments performed by Matthew Connell: October 26, 2026 and October 29, 2026"
HULL_FOOT_OLD = VIEWS[0][3]
HULL_FOOT_NEW = ("Number = invagination depth near centrosome (μm); each nucleus rotated to best show it. "
                 "Dot = centrosome. Bars = 5 μm.")
PAIR_FOOT_NEW = "Number = invagination depth near centrosome (μm). Dot = centrosome. Bars = 5 μm."
OVERVIEW_BOX = {"left": 0.59, "top": 1.2, "width": 12.47, "height": 5.9}     # room for the added lines
OVERVIEW_PT = 26

OCT26, OCT29 = "Oct 26, 2026", "Oct 29, 2026"
# (shared deck slide, config titles of the plots or None to keep the slide as it is, a date over each
# of 2 plots, or None where the plot prints its own). One title = a two-experiment grid with its dates.
RESULTS = [(19, None, None),
           (20, ("Centrosome-synapse distance",), None),
           (20, ("Fraction polarized (< 2 μm)",), (OCT26, OCT29)),
           (21, ("Invagination depth near centrosome",), None),
           (22, ("Min curvature near vs away from centrosome, Oct 26",
                 "Min curvature near vs away from centrosome, Oct 29"), (OCT26, OCT29)),
           (23, ("Mean Hull Surf. Dist.: Near / Away, Oct 26", "Mean Hull Surf. Dist.: Near / Away, Oct 29"),
            (OCT26, OCT29)),
           (24, None, None),
           (25, ("Synapse Area",), None),
           (26, ("Actin inner/outer synapse ratio",), None),
           (27, ("Nuclear aspect ratio",), None),
           (28, ("Nuclear volume",), None),
           (29, ("Nuclear surface area",), None),
           (30, ("Invagination depth over the full nucleus",), None),
           (31, ("Invagination ratio (concave + saddle faces), Oct 26",
                 "Invagination ratio (concave + saddle faces), Oct 29"), (None, OCT29)),
           (32, ("Concave Ratio, Oct 26", "Concave Ratio, Oct 29"), (None, OCT29))]
DISTANCE_TITLE = "Nucleus-centrosome closest distance"     # FF 2026-11-01: not in the shared deck, added after it
# FF's two representative cells, in place of the left-out mesh pages.
REGION_SLIDES = ("The regions measured, CD3 DMSO cell 10", "The regions measured, CD3 DMSO cell 34")
SIZES = ("10", "25", "50")
POINTS = ("Region Nadir", "Centrosome Projection")
# (slide title, config titles of the plots), FF 2026-11-01: H3K27me3 with DNA beside it, the ratio
# after them, then near vs away and the profiles. The MFI panel is not in the ordered config.
MFI_PANEL = "tests/seed_depth_20261101/seed_depth/violins/MFI_CHANPANEL_violin_Oct_29,_2026.png"
H3K_PLOTS = ([("H3K27me3 and DNA MFI", None)]
             + [(t, (t,)) for t in (f"H3K27me3 and DNA around the {p}, {z}% region" for p in POINTS for z in SIZES)]
             + [(t, (t,)) for t in (f"H3K27me3 / DNA around the {p}, {z}% region" for p in POINTS for z in SIZES)]
             + [(t, (t,)) for t in ("H3K27me3 and DNA: near / away from centrosome",
                                    "H3K27me3 and DNA: near centrosome / whole surface",
                                    "H3K27me3 and DNA vs distance from the Region Nadir",
                                    "H3K27me3 and DNA vs distance from the Centrosome Projection",
                                    "H3K27me3 and DNA density vs hull distance",
                                    "H3K27me3 and DNA density vs min curvature")])
FULL_BOX = {"left": 0.3, "width": 12.73}
PAIR_TOP = 2.1
PAIR_BOXES = ({"left": 0.1, "width": 6.52}, {"left": 6.71, "width": 6.52})
BOTTOM = 7.4
DATE_H = 0.46
DATE_PT = 24
# Text boxes that sat over the one plot, moved under the title; the plot then starts below them.
MOVED_TEXT = {23: "invaginated", 27: "flattened"}
TEXT_BOX = {"left": 1.2, "top": 0.68, "width": 10.9, "height": 1.0}
PLOT_TOP = {20: 1.8, 27: 1.65}
H3K_TOP = 1.2
STACK_ASPECT = 2.0          # two plots wider than this go one above the other: side by side they come out tiny
STACK_BOX = {"left": 1.6, "width": 10.13}


def sources(config: dict) -> dict:
    """{config title: image paths}; a title used twice keeps its first slide, the near-the-top one."""
    out = {}
    for entry in config["images"]:
        paths = [entry["path"]] if "path" in entry else entry.get("images", [])
        out.setdefault(entry.get("title"), [p if Path(p).is_absolute() else f"{config['base_dir']}/{p}" for p in paths])
    return out


def subtitles(config: dict) -> dict:
    """{config title: subtitle}, without the note on left-out cells (FF: mention it nowhere)."""
    out = {}
    for entry in config["images"]:
        parts = (entry.get("subtitle") or "").split(" · ")
        out.setdefault(entry.get("title"), " · ".join(x for x in parts if "left out" not in x) or None)
    return out


def plot_slide(prs, title: str, subtitle, paths: list):
    """A new white slide with its title and one plot, or two side by side; returns its slide id."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if len(paths) == 1:
        place(slide, paths[0], FULL_BOX, H3K_TOP)
    else:
        if all(aspect(path) > STACK_ASPECT for path in paths):
            mid = (H3K_TOP + BOTTOM) / 2
            place(slide, paths[0], STACK_BOX, H3K_TOP, mid)
            place(slide, paths[1], STACK_BOX, mid, BOTTOM)
        else:
            for path, box in zip(paths, PAIR_BOXES):
                place(slide, path, box, H3K_TOP)
    _add_title(slide, title, subtitle, fit=True)
    return list(prs.slides._sldIdLst)[-1]


def aspect(path: str) -> float:
    """Width / height of a picture, margins trimmed."""
    with Image.open(_to_open_path(path)) as img:
        crop = _content_crop(img)
        return img.width * (1 - crop[0] - crop[2]) / (img.height * (1 - crop[1] - crop[3]))


def place(slide, path: str, box: dict, top: float, bottom: float = BOTTOM) -> float:
    """Margin-trimmed picture, fitted and centered in the box, behind the slide's text."""
    with Image.open(_to_open_path(path)) as img:
        crop = _content_crop(img)
        w, h = img.size
    w *= 1 - crop[0] - crop[2]
    h *= 1 - crop[1] - crop[3]
    box_h = bottom - top
    scale = min(box["width"] / w, box_h / h)
    pw, ph = w * scale, h * scale
    pic = slide.shapes.add_picture(_to_open_path(path), Inches(box["left"] + (box["width"] - pw) / 2),
                                   Inches(top + (box_h - ph) / 2), Inches(pw), Inches(ph))
    pic.crop_left, pic.crop_top, pic.crop_right, pic.crop_bottom = crop
    tree = slide.shapes._spTree
    tree.remove(pic._element)
    tree.insert(2, pic._element)
    return top + (box_h - ph) / 2


def add_date(slide, text: str, box: dict, plot_top: float) -> None:
    """The experiment's date over a plot that does not print it."""
    frame = slide.shapes.add_textbox(Inches(box["left"]), Inches(plot_top - DATE_H - 0.02), Inches(box["width"]),
                                     Inches(DATE_H)).text_frame
    frame.text = text
    para = frame.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    para.runs[0].font.size, para.runs[0].font.bold = Pt(DATE_PT), True


def fill_results(slide, number: int, titles, dates, src: dict) -> None:
    """Both experiments' plots where the slide's own plots were."""
    top = PLOT_TOP.get(number, min(s.top for s in slide.shapes if s.shape_type == 13) / 914400)
    remove_pictures_from_slide(slide)
    paths = [p for t in titles for p in src[t]]
    if len(paths) == 1:
        place(slide, paths[0], FULL_BOX, top)
        return
    for path, box, date in zip(paths, PAIR_BOXES, dates):
        plot_top = place(slide, path, box, PAIR_TOP)
        if date:
            add_date(slide, date, box, plot_top)


def swap_text(slide, old: str, new: str) -> int:
    """Replace one run's whole text; returns how many runs changed."""
    n = 0
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text == old:
                        run.text = new
                        n += 1
    return n


def rework_overview(slide) -> None:
    """Slide 2 for both experiments: each one's stain, and the H3K27me3 question (FF's words)."""
    body = slide.placeholders[1]
    paras = body.text_frame.paragraphs
    first, last = paras[0], paras[-1]
    bold_run, plain_run = copy.deepcopy(first.runs[0]._r), copy.deepcopy(first.runs[1]._r)
    first.runs[0].text, first.runs[1].text = "The nucleus, centrosome", ", actin, and "
    first.runs[1]._r.addnext(bold_run)
    bold_run.addnext(plain_run)
    first.runs[2].text = "MT (Oct 26) or H3K27me3 (Oct 29)"
    first.runs[3].text = " imaged in confocal mode. Jurkat T cells fixed at 7 min post-drop"

    new_p = copy.deepcopy(last._p)
    last._p.addnext(new_p)
    new = body.text_frame.paragraphs[-1]
    new.runs[1]._r.addnext(copy.deepcopy(new.runs[0]._r))
    new.runs[2]._r.addnext(copy.deepcopy(new.runs[1]._r))
    for run, text in zip(new.runs, ("Influence ", "H3K27me3 distribution near these invaginations? ",
                                    "Does MT activity influence ", "chromatin organization?")):
        run.text = text

    for para in body.text_frame.paragraphs:
        for run in para.runs:
            run.font.size = Pt(OVERVIEW_PT)
    for key, inches in OVERVIEW_BOX.items():
        setattr(body, key, Inches(inches))


def drop_sections(prs) -> None:
    """The source deck's section list names slides this deck drops and misses the new ones."""
    for sections in prs.part._element.xpath(".//*[local-name()='sectionLst']"):
        ext = sections.getparent()
        ext.getparent().remove(ext)


def drop_paragraph(slide, text: str) -> int:
    """Remove the paragraphs that read exactly text; returns how many went."""
    hits = [para for shape in slide.shapes if shape.has_text_frame for para in shape.text_frame.paragraphs
            if para.text.strip() == text]
    for para in hits:
        para._p.getparent().remove(para._p)
    return len(hits)


def has_title(slide, start: str) -> bool:
    """Any text box on the slide begins with start."""
    return any(s.has_text_frame and s.text_frame.text.startswith(start) for s in slide.shapes)


def build() -> int:
    if (OUT.parent / f"~${OUT.name}").exists():
        print(f"{OUT.name} is open in PowerPoint: not built")
        return 1
    src = sources(yaml.safe_load(CONFIG.read_text(encoding="utf-8")))
    ordered = yaml.safe_load(ORDERED.read_text(encoding="utf-8"))
    h3k_src, h3k_sub = sources(ordered), subtitles(ordered)
    h3k_src[None] = [f"{ordered['base_dir']}/{MFI_PANEL}"]
    missing = [t for _, titles, _ in RESULTS for t in titles or () if not src.get(t)]
    missing += [t for t in (DISTANCE_TITLE,) if not src.get(t)]
    missing += [t for t in REGION_SLIDES if not h3k_src.get(t)]
    missing += [t for _, titles in H3K_PLOTS for t in titles or () if not h3k_src.get(t)]
    if missing:
        print("no plots in the config for:", missing)
        return 1
    prs = Presentation(str(SHARED))
    fast_image_lookup(prs)
    lst = prs.slides._sldIdLst
    own = list(lst)
    slides = list(prs.slides)
    assert swap_text(slides[0], AUTHOR_OLD, AUTHOR_NEW) == 1
    rework_overview(slides[1])

    results, extra = [], 0
    for number, titles, dates in RESULTS:
        if results and results[-1][0] == number:    # the second slide made from slide 20
            duplicate_slide(prs, number - 1)
            slide, el = prs.slides[len(prs.slides) - 1], list(lst)[-1]
            extra += 1
        else:
            slide, el = slides[number - 1], own[number - 1]
        if number in MOVED_TEXT:
            for shape in slide.placeholders:
                if shape.has_text_frame and MOVED_TEXT[number] in shape.text_frame.text:
                    for key, inches in TEXT_BOX.items():
                        setattr(shape, key, Inches(inches))
        results.append((number, el, slide, titles, dates))
    for number, _, slide, titles, dates in results:
        if titles:
            fill_results(slide, number, titles, dates, src)
    assert sum(drop_paragraph(r[2], "Artifacts excluded") for r in results) == 1   # FF: mention it nowhere
    assert swap_text(results[-1][2], "Concave surface fraction", "Concave Ratio") == 1      # FF 2026-11-01

    source = list(Presentation(str(MONTAGES)).slides)
    n_before_regions = next(k for k, s in enumerate(source) if has_title(s, MESH_LEFT_OUT))
    source = [s for s in source if not has_title(s, MESH_LEFT_OUT)]
    pair_slides = [s for s in source if has_title(s, PAIR_TITLE)]    # they follow the Oct 29 hull pages
    assert pair_slides, "no pair slides in the montage deck"
    source = [s for s in source if s not in pair_slides]
    start = len(lst)
    swapped = 0
    for slide in source:
        copy_slide(slide, prs)
        swapped += swap_text(prs.slides[len(prs.slides) - 1], HULL_FOOT_OLD, HULL_FOOT_NEW)
    assert swapped, "hull footer not found"
    added = list(lst)[start:]

    for slide in pair_slides:
        copy_slide(slide, prs)
        assert swap_text(prs.slides[len(prs.slides) - 1], PAIR_FOOT, PAIR_FOOT_NEW) == 1
    pairs = list(lst)[start + len(added):]

    distance = plot_slide(prs, DISTANCE_TITLE, None, src[DISTANCE_TITLE])
    regions = [plot_slide(prs, t, h3k_sub[t], h3k_src[t]) for t in REGION_SLIDES]
    h3k = [plot_slide(prs, f"{title}, Oct 29", h3k_sub.get((titles or (None,))[0]),
                      [p for t in titles or (None,) for p in h3k_src[t]]) for title, titles in H3K_PLOTS]

    for number in range(3, 19):                     # the shared deck's montages, replaced
        el = own[number - 1]
        prs.part.drop_rel(el.rId)
        lst.remove(el)
    for el in list(lst):
        lst.remove(el)
    slices = [el for s, el in zip(source, added) if has_title(s, SLICE_TITLE)]
    assert len(slices) == 2, "slice slides not found"
    order = (own[:2] + added[:N_THROUGH_HULL] + pairs + added[N_THROUGH_HULL:N_SHARED_MONTAGE]
             + [r[1] for r in results] + [distance]
             + [el for el in added[N_SHARED_MONTAGE:n_before_regions] if el not in slices] + regions
             + added[n_before_regions:] + h3k[:N_THROUGH_NADIR_25] + slices + h3k[N_THROUGH_NADIR_25:])
    for el in order:
        lst.append(el)
    if OUT.exists():
        backup_presentation(str(OUT))
    drop_sections(prs)
    prs.save(str(OUT))
    print(f"{len(prs.slides)} slides: 2 opening, {N_SHARED_MONTAGE} montages, {len(pairs)} mesh pairs, "
          f"{len(results)} results, {len(added) - N_SHARED_MONTAGE} H3K27me3 montages, {len(regions)} region, {len(h3k)} H3K27me3 plots, "
          f"{swapped} hull footers -> {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(build())
