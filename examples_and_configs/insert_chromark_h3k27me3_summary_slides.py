"""
insert_chromark_h3k27me3_summary_slides.py

Build a condition-comparison summary deck for the fixed activated-CTL nucleus
"chromark" analysis (tifsCTLsFixed101010aCD3aCD28ICAM_H3K27me3_01292024,
compiled 2026-06-22). Every metric in the curated grid_panels_curated/ folder
is included; each is shown across three comparison views:

  1. all conditions          - 8-condition overview (wide, no stat brackets)
  2. stiffness comparison     - stiffness_27hr + stiffness_51hr side by side
  3. timecourses              - timepoint_1p5kPa + _12kPa + _glass

Metrics are discovered from disk and auto-classified into families
(intensity, GLCM texture, GLCM mid-slice, chromatin organization, 3D/2D
morphology). Ordering: 3D before 2D; DNA (Hoechst) before H3K27me3; where both
channels have the SAME metric they are a "pair" whose slides interleave per view
(DNA all-conditions, H3K27me3 all-conditions, DNA stiffness, ...). Titles use
"DNA" for Hoechst, tag (2D)/(3D) unless obvious (e.g. volume), and give GLCM
distance in pixels. Unrecognized stems land in an "Other metrics" family.

Self-contained: builds a blank deck (no template .pptx). Missing/excluded panels
are skipped (no placeholder, no failure). Previous decks are backed up.

Usage:
    conda run -n PPT_editing python examples_and_configs/insert_chromark_h3k27me3_summary_slides.py
    # dry run (print the planned families/titles, build nothing):
    conda run -n PPT_editing python examples_and_configs/insert_chromark_h3k27me3_summary_slides.py --list
"""

import os
import re
import sys
from collections import defaultdict
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ppt_image_inserter import backup_presentation  # noqa: E402

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
ROOT = Path(
    "J:/FF/fixed_cell/CTL_nucleus/tifsFixed3SIactivatedCTLs_nucleus/"
    "tifsCTLsFixed101010aCD3aCD28ICAM_H3K27me3_01292024/chromark/compiled_all_20260622"
)
GRID_DIR = ROOT / "grid_panels_curated"   # the curated metric set (~126 features)

OUTPUT_PATH = Path(
    "K:/FF/PPT/PPT_autogeneration/Naive_CTL/chromark_H3K27me3/"
    "NaiveCTL_chromark_H3K27me3_summary.pptx"
)

DECK_TITLE = "Activated CTLs on stiffness substrates"
DECK_SUBTITLE = (
    "H3K27me3 / DNA nuclear chromark summary (curated metrics)  -  fixed 01/29/2024  -  "
    "compiled 2026-06-22"
)

GRID_SUFFIX = "_grid.png"

# Comparison views (folder under grid_panels/ -> caption).
ALL_COND_VIEW = "all_conditions"
STIFFNESS_VIEWS = [("stiffness_27hr", "27 h"), ("stiffness_51hr", "51 h")]
TIMECOURSE_VIEWS = [
    ("timepoint_1p5kPa", "1.5 kPa"),
    ("timepoint_12kPa", "12 kPa"),
    ("timepoint_glass", "glass"),
]
VIEW_ORDER = ("all", "stiffness", "timecourse")

EXCLUDE_PANELS_RAW = set()
EXCLUDE_PANELS = {
    p.split("grid_panels_curated/", 1)[-1].split("grid_panels/", 1)[-1].lstrip("/")
    for p in EXCLUDE_PANELS_RAW
}
DROP_METRICS = set()  # "include everything" -- nothing dropped

# Family order (index -> divider title).
FAMILIES_ORDER = [
    "Nuclear intensity",
    "Texture (GLCM)",
    "Texture (GLCM, max area slice)",
    "Chromatin organization",
    "Nuclear morphology (3D)",
    "Nuclear morphology (2D)",
    "Other metrics",
]
FAM = {name: i for i, name in enumerate(FAMILIES_ORDER)}

# ---------------------------------------------------------------------------
# Auto-classification / titling
# ---------------------------------------------------------------------------
CH_TITLE = {"hoechst": "DNA", "h3k27me3": "H3K27me3"}

INT_STAT_RANK = {"mean": 0, "std": 1, "sd": 1, "skewness": 2, "kurtosis": 3,
                 "median": 4, "q25": 5, "q75": 6, "min": 7, "max": 8,
                 "mode": 9, "rel": 10, "d25": 11, "d75": 12}
INT_STAT_TITLE = {"mean": "mean", "std": "SD", "sd": "SD", "skewness": "skewness",
                  "kurtosis": "excess kurtosis", "median": "median", "q25": "Q25",
                  "q75": "Q75", "min": "min", "max": "max", "mode": "mode",
                  "rel": "relative", "d25": "d25", "d75": "d75"}

GLCM_RANK = {"contrast": 0, "correlation": 1, "dissimilarity": 2, "energy": 3,
             "homogeneity": 4, "asm": 5}
GLCM_TITLE = {"asm": "ASM", "contrast": "contrast", "correlation": "correlation",
              "dissimilarity": "dissimilarity", "energy": "energy",
              "homogeneity": "homogeneity"}

# Chromatin 2D names -> (sub-rank, title-without-channel). Channel prepended.
CHROM2D = {
    "i80_i20": (0, "I80/I20 ratio (2D)"),
    "hc_area_ec_area": (1, "heterochromatin/euchromatin area (2D)"),
    "hc_area_nuc_area": (2, "heterochromatin/nucleus area (2D)"),
    "hc_content_ec_content": (3, "HC/EC content (2D)"),
    "hc_content_dna_content": (4, "HC/DNA content (2D)"),
    "nhigh_nlow": (9, "N_high / N_low (2D)"),
}

# Morphology titles (channel-independent). Insertion order = display order.
MORPH3D_TITLES = {
    "morph3d_nuclear_volume": "Nuclear volume",
    "morph3d_surface_area": "Surface area (3D)",
    "morph3d_convex_hull_vol": "Convex hull volume",
    "morph3d_equivalent_diameter": "Equivalent diameter (3D)",
    "morph3d_extent": "Nuclear extent (vol/bbox)",
    "morph3d_solidity": "Solidity (3D)",
    "morph3d_concavity_3d": "Concavity (3D)",
    "morph3d_major_axis_length": "Major axis length (3D)",
    "morph3d_minor_axis_length": "Minor axis length (3D)",
}
MORPH2D_TITLES = {
    # size / global shape
    "morph2d_area": "Area (2D)",
    "morph2d_perimeter": "Perimeter (2D)",
    "morph2d_convex_area": "Convex area (2D)",
    "morph2d_bbox_area": "Bounding-box area (2D)",
    "morph2d_area_bbarea": "Area / bbox area (2D)",
    "morph2d_equivalent_diameter": "Equivalent diameter (2D)",
    "morph2d_eccentricity": "Eccentricity (2D)",
    "morph2d_solidity": "Solidity (2D)",
    "morph2d_shape_factor": "Shape factor (2D)",
    "morph2d_a_r": "Aspect ratio (2D)",
    "morph2d_orientation": "Orientation (2D)",
    "morph2d_major_axis_length": "Major axis length (2D)",
    "morph2d_minor_axis_length": "Minor axis length (2D)",
    "morph2d_feret_max": "Max Feret (2D)",
    "morph2d_max_calliper": "Max calliper (2D)",
    "morph2d_min_calliper": "Min calliper (2D)",
    "morph2d_smallest_largest_calliper": "Smallest/largest calliper (2D)",
    # radius
    "morph2d_avg_radius": "Avg radius (2D)",
    "morph2d_med_radius": "Median radius (2D)",
    "morph2d_std_radius": "Radius SD (2D)",
    "morph2d_min_radius": "Min radius (2D)",
    "morph2d_max_radius": "Max radius (2D)",
    "morph2d_mode_radius": "Mode radius (2D)",
    "morph2d_d25_radius": "Radius d25 (2D)",
    "morph2d_d75_radius": "Radius d75 (2D)",
    # curvature
    "morph2d_avg_curvature": "Avg curvature (2D)",
    "morph2d_std_curvature": "Curvature SD (2D)",
    "morph2d_avg_posi_curv": "Avg positive curvature (2D)",
    "morph2d_avg_neg_curv": "Avg negative curvature (2D)",
    "morph2d_std_posi_curv": "Positive-curvature SD (2D)",
    "morph2d_std_neg_curv": "Negative-curvature SD (2D)",
    "morph2d_max_posi_curv": "Max positive curvature (2D)",
    "morph2d_max_neg_curv": "Max negative curvature (2D)",
    "morph2d_med_posi_curv": "Median positive curvature (2D)",
    "morph2d_med_neg_curv": "Median negative curvature (2D)",
    "morph2d_sum_posi_curv": "Sum positive curvature (2D)",
    "morph2d_sum_neg_curv": "Sum negative curvature (2D)",
    "morph2d_len_posi_curv": "Length of positive curvature (2D)",
    "morph2d_len_neg_curv": "Length of negative curvature (2D)",
    "morph2d_frac_peri_w_posi_curvature": "Fraction of perimeter w/ positive curvature (2D)",
    "morph2d_frac_peri_w_neg_curvature": "Fraction of perimeter w/ negative curvature (2D)",
    "morph2d_concavity": "Concavity (2D)",
    # prominence / polarity
    "morph2d_prominant_pos_curv": "Prominent positive curvature (2D)",
    "morph2d_prominant_neg_curv": "Prominent negative curvature (2D)",
    "morph2d_num_prominant_pos_curv": "Number of prominent positive-curvature points (2D)",
    "morph2d_num_prominant_neg_curv": "Number of prominent negative-curvature points (2D)",
    "morph2d_prominance_prominant_pos_curv": "Prominence of prominent positive curvature (2D)",
    "morph2d_prominance_prominant_neg_curv": "Prominence of prominent negative curvature (2D)",
    "morph2d_width_prominant_pos_curv": "Width of prominent positive curvature (2D)",
    "morph2d_width_prominant_neg_curv": "Width of prominent negative curvature (2D)",
    "morph2d_npolarity_changes": "Number of curvature sign changes (2D)",
    "morph2d_frac_peri_w_polarity_changes": "Fraction of perimeter w/ curvature sign changes (2D)",
}
MORPH_ORDER = {stem: i for i, stem in enumerate(
    list(MORPH3D_TITLES) + list(MORPH2D_TITLES))}


def _ch(tok):
    return CH_TITLE[tok]


def _periph_dist(tok):
    """Parse a peripheral chromatin/enrichment distance token -> (display, sortkey).
    Handles pixel (``10``), micron (``0p5um``, ``1um``, ``2um``) and radial-percent
    (``r10pct``) forms. sortkey groups px < um < pct, then by value."""
    m = re.match(r"^(\d+)$", tok)
    if m:
        return ("{} px".format(tok), (0, int(tok)))
    m = re.match(r"^(\d+)(?:p(\d+))?um$", tok)
    if m:
        whole, frac = m.group(1), m.group(2)
        disp = ("{}.{}".format(whole, frac) if frac else whole) + " µm"
        num = float("{}.{}".format(whole, frac)) if frac else float(whole)
        return (disp, (1, num))
    m = re.match(r"^r(\d+)pct$", tok)
    if m:
        return ("r{}%".format(m.group(1)), (2, int(m.group(1))))
    return (tok, (3, 0))


def classify(stem):
    """Return a record dict or None. Record keys: fam, sort, channel, pair, title.
    `fam` is a FAMILIES_ORDER name; `sort` orders within the family; `pair` is a
    channel-agnostic key (records sharing it across DNA/H3K27me3 form a pair);
    `channel` in {'DNA','H3K27me3', None}."""

    # --- intensity: chan_<stat>_<ch>_3d_int ---
    m = re.match(r"^chan_(.+)_(hoechst|h3k27me3)_3d_int$", stem)
    if m:
        stat, ch = m.group(1), m.group(2)
        return dict(fam="Nuclear intensity",
                    sort=(0, INT_STAT_RANK.get(stat, 50), 1, stem),
                    channel=_ch(ch), pair=("int_chan", stat),
                    title="{} {} intensity (3D)".format(
                        _ch(ch), INT_STAT_TITLE.get(stat, stat)))

    # --- intensity: <ch>_3d_nuclear_<stat>_int ---
    m = re.match(r"^(hoechst|h3k27me3)_3d_nuclear_(.+)_int$", stem)
    if m:
        ch, stat = m.group(1), m.group(2)
        return dict(fam="Nuclear intensity",
                    sort=(0, INT_STAT_RANK.get(stat, 50), 0, stem),
                    channel=_ch(ch), pair=("int_nuclear", stat),
                    title="{} nuclear {} intensity (3D)".format(
                        _ch(ch), INT_STAT_TITLE.get(stat, stat)))

    # --- intensity: <ch>_2d_int_<stat> ---
    m = re.match(r"^(hoechst|h3k27me3)_2d_int_(.+)$", stem)
    if m:
        ch, stat = m.group(1), m.group(2)
        return dict(fam="Nuclear intensity",
                    sort=(1, INT_STAT_RANK.get(stat, 50), 0, stem),
                    channel=_ch(ch), pair=("int_2d", stat),
                    title="{} {} intensity (2D)".format(
                        _ch(ch), INT_STAT_TITLE.get(stat, stat)))

    # --- intensity: <ch>_2d_skewness | _2d_kurtosis ---
    m = re.match(r"^(hoechst|h3k27me3)_2d_(skewness|kurtosis)$", stem)
    if m:
        ch, stat = m.group(1), m.group(2)
        return dict(fam="Nuclear intensity",
                    sort=(1, INT_STAT_RANK[stat], 0, stem),
                    channel=_ch(ch), pair=("int_2d", stat),
                    title="{} {} intensity (2D)".format(_ch(ch), INT_STAT_TITLE[stat]))

    # --- GLCM: <ch>_(2d|2dmid)_<type>_<dist> ---
    m = re.match(r"^(hoechst|h3k27me3)_(2d|2dmid)_"
                 r"(asm|contrast|correlation|dissimilarity|energy|homogeneity)_(\d+)$", stem)
    if m:
        ch, variant, gt, dist = m.groups()
        fam = "Texture (GLCM)" if variant == "2d" else "Texture (GLCM, max area slice)"
        suffix = "2D" if variant == "2d" else "max area slice"
        return dict(fam=fam,
                    sort=(GLCM_RANK[gt], int(dist), stem),
                    channel=_ch(ch), pair=("glcm", variant, gt, dist),
                    title="{} GLCM {} ({} px, {})".format(
                        _ch(ch), GLCM_TITLE[gt], dist, suffix))

    # --- texture: <ch>_2d_entropy ---
    m = re.match(r"^(hoechst|h3k27me3)_2d_entropy$", stem)
    if m:
        ch = m.group(1)
        return dict(fam="Texture (GLCM)", sort=(90, 0, stem),
                    channel=_ch(ch), pair=("entropy",),
                    title="{} entropy (2D)".format(_ch(ch)))

    # --- chromatin: <ch>_3d_rdp_<n> ---
    m = re.match(r"^(hoechst|h3k27me3)_3d_rdp_(\d+)$", stem)
    if m:
        ch, n = m.group(1), m.group(2)
        return dict(fam="Chromatin organization",
                    sort=(0, 0, int(n), stem), channel=_ch(ch), pair=("rdp", n),
                    title="{} radial density profile, shell {} (3D)".format(_ch(ch), n))

    # --- chromatin: <ch>_3d_rel_(hc|ec)_volume ---
    m = re.match(r"^(hoechst|h3k27me3)_3d_rel_(hc|ec)_volume$", stem)
    if m:
        ch, which = m.group(1), m.group(2)
        full = "heterochromatin" if which == "hc" else "euchromatin"
        return dict(fam="Chromatin organization",
                    sort=(0, 1, 0 if which == "hc" else 1, stem),
                    channel=_ch(ch), pair=("rel_vol", which),
                    title="{} relative {} volume".format(_ch(ch), full))

    # --- chromatin: <ch>_3d_hc_ec_ratio_3d ---
    m = re.match(r"^(hoechst|h3k27me3)_3d_hc_ec_ratio_3d$", stem)
    if m:
        ch = m.group(1)
        return dict(fam="Chromatin organization", sort=(0, 2, 0, stem),
                    channel=_ch(ch), pair=("hc_ec_ratio",),
                    title="{} HC/EC volume ratio (3D)".format(_ch(ch)))

    # --- chromatin: <ch>_(2d|3d)_peripheral_(chromatin|enrichment)_<dist> ---
    # dist may be pixels (10), microns (0p5um/1um/2um) or radial-percent (r10pct).
    m = re.match(r"^(hoechst|h3k27me3)_(2d|3d)_peripheral_(chromatin|enrichment)_(.+)$", stem)
    if m:
        ch, dim, kind, tok = m.groups()
        disp, dsort = _periph_dist(tok)
        sub = 6 if kind == "chromatin" else 7
        return dict(fam="Chromatin organization",
                    sort=(0 if dim == "3d" else 1, sub, dsort, stem), channel=_ch(ch),
                    pair=("peripheral", dim, kind, tok),
                    title="{} peripheral {}, {} ({})".format(
                        _ch(ch), kind, disp, "3D" if dim == "3d" else "2D"))

    # --- chromatin: named 2D ratios/contents ---
    m = re.match(r"^(hoechst|h3k27me3)_2d_(.+)$", stem)
    if m and m.group(2) in CHROM2D:
        ch, key = m.group(1), m.group(2)
        sub, ttl = CHROM2D[key]
        return dict(fam="Chromatin organization", sort=(1, sub, 0, stem),
                    channel=_ch(ch), pair=("chrom2d", key),
                    title="{} {}".format(_ch(ch), ttl))

    # --- morphology (channel-independent) ---
    if stem in MORPH3D_TITLES:
        return dict(fam="Nuclear morphology (3D)", sort=(MORPH_ORDER[stem], stem),
                    channel=None, pair=("morph", stem), title=MORPH3D_TITLES[stem])
    if stem in MORPH2D_TITLES:
        return dict(fam="Nuclear morphology (2D)", sort=(MORPH_ORDER[stem], stem),
                    channel=None, pair=("morph", stem), title=MORPH2D_TITLES[stem])
    if stem.startswith(("morph3d_", "morph2d_")):
        fam = "Nuclear morphology (3D)" if stem.startswith("morph3d_") else "Nuclear morphology (2D)"
        return dict(fam=fam, sort=(999, stem), channel=None, pair=("morph", stem),
                    title=stem.replace("morph2d_", "").replace("morph3d_", "").replace("_", " ").capitalize())

    return None


def fallback_title(stem):
    t = stem.replace("hoechst", "DNA").replace("h3k27me3", "H3K27me3")
    return t.replace("_", " ")


# ---------------------------------------------------------------------------
# Colors / layout
# ---------------------------------------------------------------------------
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
DIVIDER_BG = RGBColor(0xF0, 0xF0, 0xF0)

SLIDE_W = 13.333
SLIDE_H = 7.5
MARGIN = 0.10

TITLE_LEFT = MARGIN
TITLE_TOP = 0.05
TITLE_WIDTH = SLIDE_W - 2 * MARGIN
TITLE_HEIGHT = 0.55
TITLE_FONT_PT = 28

IMG_LEFT = MARGIN
IMG_TOP = 0.66
IMG_BOX_W = SLIDE_W - 2 * MARGIN
IMG_BOX_H = 6.36

FOOTER_LEFT = MARGIN
FOOTER_TOP = 7.06
FOOTER_WIDTH = SLIDE_W - 2 * MARGIN
FOOTER_HEIGHT = 0.40
FOOTER_FONT_PT = 9

COL_GAP = 0.20
MULTI_LABEL_TOP = 0.64
MULTI_LABEL_HEIGHT = 0.42
MULTI_IMG_TOP = 1.10
MULTI_IMG_H = 5.90
MULTI_FOOTER_FONT_PT = 8


# ---------------------------------------------------------------------------
# Generic slide helpers
# ---------------------------------------------------------------------------
def add_textbox(slide, text, left, top, width, height, font_pt, color,
                bold=False, italic=False, align=PP_ALIGN.CENTER):
    box = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.text = text
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.runs[0]
    run.font.size = Pt(font_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def title_font_for(text):
    n = len(text)
    if n <= 52:
        return TITLE_FONT_PT
    if n <= 70:
        return 24
    if n <= 90:
        return 20
    return 18


def add_image_in_box(slide, image_path, box_left, box_top, box_w, box_h):
    pic = slide.shapes.add_picture(
        image_path, Inches(box_left), Inches(box_top), width=Inches(box_w))
    actual_h_in = pic.height / 914400.0
    if actual_h_in > box_h:
        sp = pic._element
        sp.getparent().remove(sp)
        pic = slide.shapes.add_picture(
            image_path, Inches(box_left), Inches(box_top), height=Inches(box_h))
        actual_w_in = pic.width / 914400.0
        pic.left = Inches(box_left + (box_w - actual_w_in) / 2)
    else:
        pic.top = Inches(box_top + (box_h - actual_h_in) / 2)
    return pic


def set_slide_background(slide, rgb):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def _new_slide(prs, bg=WHITE):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, bg)
    return slide


def build_slide(prs, title_text, image_path, footer_text):
    slide = _new_slide(prs)
    add_textbox(slide, title_text, TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
                font_pt=title_font_for(title_text), color=BLACK, bold=True)
    add_image_in_box(slide, str(image_path), IMG_LEFT, IMG_TOP, IMG_BOX_W, IMG_BOX_H)
    add_textbox(slide, footer_text, FOOTER_LEFT, FOOTER_TOP, FOOTER_WIDTH,
                FOOTER_HEIGHT, font_pt=FOOTER_FONT_PT, color=BLACK)
    return slide


def build_multi_slide(prs, title, image_paths, labels, footers):
    n = len(image_paths)
    slide = _new_slide(prs)
    add_textbox(slide, title, TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
                font_pt=title_font_for(title), color=BLACK, bold=True)
    col_w = (SLIDE_W - 2 * MARGIN - (n - 1) * COL_GAP) / n
    label_font = 26 if n <= 2 else 20
    for i in range(n):
        left = MARGIN + i * (col_w + COL_GAP)
        add_textbox(slide, labels[i], left, MULTI_LABEL_TOP, col_w, MULTI_LABEL_HEIGHT,
                    font_pt=label_font, color=BLACK, bold=True)
        add_image_in_box(slide, str(image_paths[i]), left, MULTI_IMG_TOP, col_w, MULTI_IMG_H)
    box = slide.shapes.add_textbox(
        Inches(FOOTER_LEFT), Inches(FOOTER_TOP), Inches(FOOTER_WIDTH), Inches(FOOTER_HEIGHT))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.text = footers[0]
    for idx, footer in enumerate(footers):
        para = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx != 0:
            para.text = footer
        para.alignment = PP_ALIGN.CENTER
        para.runs[0].font.size = Pt(MULTI_FOOTER_FONT_PT)
        para.runs[0].font.color.rgb = BLACK
    return slide


def build_title_slide(prs, title, subtitle):
    slide = _new_slide(prs)
    add_textbox(slide, title, MARGIN, 2.7, SLIDE_W - 2 * MARGIN, 1.3,
                font_pt=40, color=BLACK, bold=True)
    add_textbox(slide, subtitle, MARGIN, 4.1, SLIDE_W - 2 * MARGIN, 1.0,
                font_pt=18, color=BLACK, italic=True)


def build_divider_slide(prs, family_name):
    slide = _new_slide(prs, bg=DIVIDER_BG)
    add_textbox(slide, family_name, MARGIN, 3.1, SLIDE_W - 2 * MARGIN, 1.3,
                font_pt=44, color=BLACK, bold=True)


# ---------------------------------------------------------------------------
# Deck assembly
# ---------------------------------------------------------------------------
def rel_footer(path):
    try:
        return path.relative_to(ROOT).as_posix()
    except ValueError:
        return path.as_posix()


def _excluded(view, fname):
    return "{}/{}".format(view, fname) in EXCLUDE_PANELS


def _emit_multi(prs, pretty, kind, views, fname, missing, omitted, sep):
    kept = []
    for v, lab in views:
        if _excluded(v, fname):
            omitted.append("{}/{}".format(v, fname))
        elif not (GRID_DIR / v / fname).exists():
            missing.append("{}/{}".format(v, fname))
        else:
            kept.append((v, lab))
    if not kept:
        return
    paths = [GRID_DIR / v / fname for v, _ in kept]
    labels = [lab for _, lab in kept]
    title = "{} — {} ({})".format(pretty, kind, sep.join(labels))
    build_multi_slide(prs, title, paths, labels, [rel_footer(pp) for pp in paths])


def emit_view(prs, view_kind, stem, pretty, missing, omitted):
    fname = stem + GRID_SUFFIX
    if view_kind == "all":
        if _excluded(ALL_COND_VIEW, fname):
            omitted.append("{}/{}".format(ALL_COND_VIEW, fname))
            return
        p = GRID_DIR / ALL_COND_VIEW / fname
        if p.exists():
            build_slide(prs, "{} — all conditions".format(pretty), p, rel_footer(p))
        else:
            missing.append("{}/{}".format(ALL_COND_VIEW, fname))
    elif view_kind == "stiffness":
        _emit_multi(prs, pretty, "stiffness comparison", STIFFNESS_VIEWS, fname,
                    missing, omitted, " / ")
    elif view_kind == "timecourse":
        _emit_multi(prs, pretty, "timecourses", TIMECOURSE_VIEWS, fname,
                    missing, omitted, ", ")


def build_item(prs, item, missing, omitted):
    if item[0] == "pair":
        _, dna_stem, dna_title, h3_stem, h3_title = item
        for view in VIEW_ORDER:
            emit_view(prs, view, dna_stem, dna_title, missing, omitted)
            emit_view(prs, view, h3_stem, h3_title, missing, omitted)
    else:  # solo
        _, stem, title = item
        for view in VIEW_ORDER:
            emit_view(prs, view, stem, title, missing, omitted)


def build_item_allcond(prs, item, missing):
    """All-conditions-only variant: one slide per item. A pair shows DNA and
    H3K27me3 all-conditions panels side by side under a shared metric title; a
    solo shows its single all-conditions panel."""
    if item[0] == "pair":
        _, dna_stem, dna_title, h3_stem, h3_title = item
        base = dna_title[len("DNA "):] if dna_title.startswith("DNA ") else dna_title
        cols = []
        for stem, lab in ((dna_stem, "DNA"), (h3_stem, "H3K27me3")):
            fn = stem + GRID_SUFFIX
            p = GRID_DIR / ALL_COND_VIEW / fn
            if p.exists() and not _excluded(ALL_COND_VIEW, fn):
                cols.append((p, lab))
            else:
                missing.append("{}/{}".format(ALL_COND_VIEW, fn))
        if not cols:
            return
        title = "{} — all conditions".format(base)
        if len(cols) == 1:
            build_slide(prs, title, cols[0][0], rel_footer(cols[0][0]))
        else:
            build_multi_slide(prs, title, [c[0] for c in cols],
                               [c[1] for c in cols], [rel_footer(c[0]) for c in cols])
    else:  # solo
        _, stem, title = item
        fn = stem + GRID_SUFFIX
        p = GRID_DIR / ALL_COND_VIEW / fn
        if p.exists() and not _excluded(ALL_COND_VIEW, fn):
            build_slide(prs, "{} — all conditions".format(title), p, rel_footer(p))
        else:
            missing.append("{}/{}".format(ALL_COND_VIEW, fn))


def _item_stems(item):
    return [item[1], item[3]] if item[0] == "pair" else [item[1]]


def _item_log(item):
    return "{} | {}".format(item[2], item[4]) if item[0] == "pair" else item[2]


def discover_stems():
    """Union of metric stems across all views (so view-specific metrics are kept)."""
    stems = set()
    if not GRID_DIR.is_dir():
        return []
    for d in GRID_DIR.iterdir():
        if d.is_dir():
            for p in d.glob("*" + GRID_SUFFIX):
                stems.add(p.name[: -len(GRID_SUFFIX)])
    return sorted(stems)


def build_families():
    """Discover + classify all metrics into ordered (family_name, [items])."""
    stems = [s for s in discover_stems() if s not in DROP_METRICS]
    records, unknown = [], []
    for s in stems:
        rec = classify(s)
        if rec is None:
            unknown.append(s)
        else:
            rec["stem"] = s
            records.append(rec)

    # Group records into pair/solo items by (family, pair-key).
    groups = defaultdict(list)
    for r in records:
        groups[(r["fam"], r["pair"])].append(r)

    fam_items = defaultdict(list)  # fam -> list of (sort_key, item)
    for (fam, _pair), recs in groups.items():
        chans = {r["channel"] for r in recs}
        sort_key = min(r["sort"] for r in recs)
        if len(recs) == 2 and chans == {"DNA", "H3K27me3"}:
            dna = next(r for r in recs if r["channel"] == "DNA")
            h3 = next(r for r in recs if r["channel"] == "H3K27me3")
            fam_items[fam].append((sort_key, ("pair", dna["stem"], dna["title"],
                                              h3["stem"], h3["title"])))
        else:
            for r in recs:  # solos (incl. any unpaired channel metric)
                fam_items[fam].append((r["sort"], ("solo", r["stem"], r["title"])))

    if unknown:
        fam_items["Other metrics"].extend(
            ((s,), ("solo", s, fallback_title(s))) for s in sorted(unknown))

    families = []
    for fam in FAMILIES_ORDER:
        if fam in fam_items:
            ordered = [it for _, it in sorted(fam_items[fam], key=lambda x: x[0])]
            families.append((fam, ordered))
    return families, unknown


def main():
    list_only = "--list" in sys.argv
    allcond = "--allcond" in sys.argv  # all-conditions-only deck (DNA vs H3K side by side)
    families, unknown = build_families()

    n_metrics = sum(len(_item_stems(it)) for _, items in families for it in items)
    n_items = sum(len(items) for _, items in families)
    n_pairs = sum(1 for _, items in families for it in items if it[0] == "pair")
    est_slides = 1 + len(families) + (n_items if allcond else n_metrics * 3)

    print("Source: {}".format(GRID_DIR))
    print("{} metrics ({} pairs), est. {} slides across {} families{}\n".format(
        n_metrics, n_pairs, est_slides, len(families),
        "  [all-conditions only]" if allcond else ""))

    if list_only:
        for fam, items in families:
            print("=== {} ({} items) ===".format(fam, len(items)))
            for it in items:
                print("  {}".format(_item_log(it)))
            print("")
        if unknown:
            print("UNKNOWN (-> Other): {}".format(unknown))
        return

    if allcond:
        out_path = OUTPUT_PATH.with_name(OUTPUT_PATH.stem + "_all_conditions.pptx")
        subtitle = DECK_SUBTITLE + "  -  all conditions only (DNA vs H3K27me3)"
    else:
        out_path = OUTPUT_PATH
        subtitle = DECK_SUBTITLE

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    build_title_slide(prs, DECK_TITLE, subtitle)

    missing, omitted = [], []
    for fam, items in families:
        build_divider_slide(prs, fam)
        print("=== {} ===".format(fam))
        for it in items:
            if allcond:
                build_item_allcond(prs, it, missing)
            else:
                build_item(prs, it, missing, omitted)
            print("  {}".format(_item_log(it)))
        print("")

    if out_path.exists():
        backup_dir = out_path.parent / "backups"
        created = backup_presentation(str(out_path), backup_base=str(backup_dir))
        if created:
            print("Backed up previous deck to: {}\n".format(backup_dir))

    prs.save(str(out_path))
    total = len(prs.slides._sldIdLst)
    print("Done. {} metrics, {} slides written to:\n  {}".format(
        n_metrics, total, out_path))
    if unknown:
        print("\n{} unrecognized stem(s) under 'Other metrics': {}".format(
            len(unknown), unknown))
    if missing:
        print("\nSkipped {} missing panel(s) (not on disk).".format(len(missing)))


if __name__ == "__main__":
    main()
