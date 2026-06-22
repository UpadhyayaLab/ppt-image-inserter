"""
insert_actin_summary_slides.py

Build one CART actin summary deck per dataset listed in DATASETS.
Each slide shows one PNG plot with:
  - bold title derived from the filename (with overrides)
  - the plot, full-width, aspect ratio preserved
  - the full source filepath in a small footer (for traceability)

To add a new dataset, append an entry to DATASETS with its grid_panels/
root, output path, and PNG list. To add slides, extend the relevant
*_METRICS list. Each run regenerates every deck (idempotent).

Modeled after insert_actin_qc_synapse_mask_xz_mip_slides.py.

Usage:
    python examples_and_configs/insert_actin_summary_slides.py
"""

import os
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ppt_image_inserter import backup_presentation  # noqa: E402

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Subfolder names — identical across every dataset's grid_panels/ tree.
SUB_SCATTER    = "timecourse_scatter_plots"
SUB_CAT_VS_FMC = "CAT_vs_FMC_by_timepoint"

# Slide entries: (subfolder, png_filename). Joined with each dataset's
# compiled_root at deck-build time. Missing files render "(missing)".

SLICE_METRICS = [
    # area + spatial concentration
    (SUB_SCATTER, "actin_bottom_mask_area_grid.png"),
    (SUB_SCATTER, "actin_bottom_slice_inner_outer_ratio_50pct_grid.png"),
    (SUB_SCATTER, "actin_bottom_slice_inner_outer_ratio_70pct_grid.png"),
    # intensity
    (SUB_SCATTER, "actin_bottom_MFI_grid.png"),
    (SUB_SCATTER, "actin_bottom_total_sig_grid.png"),
    # shape
    (SUB_SCATTER, "actin_bottom_slice_perimeter_grid.png"),
    (SUB_SCATTER, "actin_bottom_slice_circularity_grid.png"),
    (SUB_SCATTER, "actin_bottom_slice_solidity_grid.png"),
    (SUB_SCATTER, "actin_bottom_slice_eccentricity_grid.png"),
]
RAD_PROFILE_SLICE = (
    SUB_CAT_VS_FMC,
    "actin_bottom_slice_rad_profile_auc1_all_cells_with_average_grid.png",
)
# Slice block with rad_profile spliced in after the two IOR slides
# (slot 4 in 1-based) per the agreed order.
SLICE_METRICS_WITH_RAD = (
    SLICE_METRICS[:3] + [RAD_PROFILE_SLICE] + SLICE_METRICS[3:]
)

THREE_SLICE_METRICS = [
    # area + spatial concentration
    (SUB_SCATTER, "actin_bottom_3slice_mask_area_grid.png"),
    (SUB_SCATTER, "actin_bottom_3slice_mip_inner_outer_ratio_50pct_grid.png"),
    (SUB_SCATTER, "actin_bottom_3slice_mip_inner_outer_ratio_70pct_grid.png"),
    # intensity
    (SUB_SCATTER, "actin_bottom_3slice_MFI_grid.png"),
    (SUB_SCATTER, "actin_bottom_3slice_total_sig_grid.png"),
    # shape
    (SUB_SCATTER, "actin_bottom_3slice_perimeter_grid.png"),
    (SUB_SCATTER, "actin_bottom_3slice_circularity_grid.png"),
    (SUB_SCATTER, "actin_bottom_3slice_solidity_grid.png"),
    (SUB_SCATTER, "actin_bottom_3slice_eccentricity_grid.png"),
]
RAD_PROFILE_3SLICE = (
    SUB_CAT_VS_FMC,
    "actin_bottom_3slice_mip_rad_profile_auc1_all_cells_with_average_grid.png",
)
THREE_SLICE_METRICS_WITH_RAD = (
    THREE_SLICE_METRICS[:3] + [RAD_PROFILE_3SLICE] + THREE_SLICE_METRICS[3:]
)

# by_timepoint compiles are flat: every PNG lives under
# <by_timepoint_root>/all_<tp>_<date>/all_conditions/, with no _grid suffix.
BY_TIMEPOINT_ROOT = Path(
    "J:/FF/fixed_cell/CAR_TCell/results_compiled/actin/by_timepoint"
)


def _by_timepoint(tp_short: str, date: str = "20260618") -> Path:
    """Path to one timepoint's flat all_conditions/ dir."""
    return BY_TIMEPOINT_ROOT / f"all_{tp_short}_{date}" / "all_conditions"


# Flat png_files list usable with strip_grid_suffix=True datasets (the
# subfolder slot is "" because the PNGs sit directly under compiled_root).
ALL_CONDITIONS_PNG_FILES = [
    ("", fn) for _, fn in (SLICE_METRICS_WITH_RAD + THREE_SLICE_METRICS_WITH_RAD)
]

# ---------------------------------------------------------------------------
# MT/CatB compile (CART_summary_MT_CatB deck)

MT_CATB_ROOT = Path(
    "J:/FF/fixed_cell/CAR_TCell/results_compiled/"
    "MT_CATB_20231127_20240620_20240624/compiled_20260618/grid_panels"
)
OUTPUT_DIR_MT_CATB = Path("K:/FF/PPT/PPT_autogeneration/CART/MT_CatB")

# 1 centrosome slide (user picked the z-distance variant; skipped the
# center_slice_from_MT variant).
CENTROSOME_METRICS = [
    (SUB_SCATTER, "centrosome_center_z_cell_bottom_distance_grid.png"),
]

# All 48 CathepsinB PNGs, grouped by analysis category.
CATB_METRICS = [
    # --- synapse intensity & spatial (16) ----------------------------------
    (SUB_SCATTER, "CathepsinB_synapse_MFI_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_MFI_3mip_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_total_sig_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_total_sig_3mip_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_outer_ratio_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_outer_ratio_3mip_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_outer_ratio_50pct_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_outer_ratio_70pct_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_mask_MFI_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_mask_MFI_3mip_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_outer_mask_MFI_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_outer_mask_MFI_3mip_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_g_ave_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_g_ave_3mip_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_r_eff_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_r_eff_3mip_grid.png"),
    # --- synapse autocorr (2) ----------------------------------------------
    (SUB_SCATTER, "CathepsinB_synapse_autocorr_rmax_um_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_autocorr_smoothSig_grid.png"),
    # --- total + peak signal (2) -------------------------------------------
    (SUB_SCATTER, "CathepsinB_total_sig_grid.png"),
    (SUB_SCATTER, "CathepsinB_peak_sig_grid.png"),
    # --- granule morphology (6) --------------------------------------------
    (SUB_SCATTER, "CathepsinB_granule_area_um2_grid.png"),
    (SUB_SCATTER, "CathepsinB_granule_area_fraction_grid.png"),
    (SUB_SCATTER, "CathepsinB_granule_mean_intensity_segmented_grid.png"),
    (SUB_SCATTER, "CathepsinB_granule_mean_intensity_nonsegmented_grid.png"),
    (SUB_SCATTER, "CathepsinB_granule_total_intensity_segmented_grid.png"),
    (SUB_SCATTER, "CathepsinB_granule_total_intensity_nonsegmented_grid.png"),
    # --- centrosome proximity (6) ------------------------------------------
    (SUB_SCATTER, "CathepsinB_MFI_around_cent_1um_grid.png"),
    (SUB_SCATTER, "CathepsinB_MFI_around_cent_2um_grid.png"),
    (SUB_SCATTER, "CathepsinB_MFI_around_cent_3um_grid.png"),
    (SUB_SCATTER, "CathepsinB_frac_around_cent_1um_grid.png"),
    (SUB_SCATTER, "CathepsinB_frac_around_cent_2um_grid.png"),
    (SUB_SCATTER, "CathepsinB_frac_around_cent_3um_grid.png"),
    # --- Z-distribution (10) -----------------------------------------------
    (SUB_SCATTER, "CathepsinB_z50_rel_cell_bottom_grid.png"),
    (SUB_SCATTER, "CathepsinB_z50_norm_cell_height_grid.png"),
    (SUB_SCATTER, "CathepsinB_z75_rel_cell_bottom_grid.png"),
    (SUB_SCATTER, "CathepsinB_z75_norm_cell_height_grid.png"),
    (SUB_SCATTER, "CathepsinB_z90_rel_cell_bottom_grid.png"),
    (SUB_SCATTER, "CathepsinB_z90_norm_cell_height_grid.png"),
    (SUB_SCATTER, "CathepsinB_zCOF_grid.png"),
    (SUB_SCATTER, "CathepsinB_zCOF_actin_scale_grid.png"),
    (SUB_SCATTER, "CathepsinB_zCOF_cell_bottom_distance_grid.png"),
    (SUB_SCATTER, "CathepsinB_zCOF_cell_bottom_distance_norm_cell_height_grid.png"),
    # --- FDD / structural (6) ----------------------------------------------
    (SUB_SCATTER, "CathepsinB_FDD_3D_grid.png"),
    (SUB_SCATTER, "CathepsinB_FDD_3D_RMS_grid.png"),
    (SUB_SCATTER, "CathepsinB_FDD_3D_rel_cent_grid.png"),
    (SUB_SCATTER, "CathepsinB_z_FDD_grid.png"),
    (SUB_SCATTER, "CathepsinB_z_FDD_RMS_grid.png"),
    (SUB_SCATTER, "CathepsinB_z_FDD_rel_cent_grid.png"),
]

CATB_Z50_METRICS = [
    (SUB_SCATTER, "CathepsinB_z50_rel_cell_bottom_grid.png"),
    (SUB_SCATTER, "CathepsinB_z50_norm_cell_height_grid.png"),
]
CATB_Z75_METRICS = [
    (SUB_SCATTER, "CathepsinB_z75_rel_cell_bottom_grid.png"),
    (SUB_SCATTER, "CathepsinB_z75_norm_cell_height_grid.png"),
]

CATB_METRICS_CONCISE = [
    # spatial concentration (granule_area_um2 dropped per user)
    (SUB_SCATTER, "CathepsinB_synapse_inner_outer_ratio_50pct_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_inner_outer_ratio_70pct_grid.png"),
    # intensity
    (SUB_SCATTER, "CathepsinB_synapse_MFI_grid.png"),
    (SUB_SCATTER, "CathepsinB_synapse_total_sig_grid.png"),
    (SUB_SCATTER, "CathepsinB_total_sig_grid.png"),
    # centrosome-proximity (fraction of CatB signal within N µm of centrosome)
    (SUB_SCATTER, "CathepsinB_frac_around_cent_1um_grid.png"),
    (SUB_SCATTER, "CathepsinB_frac_around_cent_2um_grid.png"),
    (SUB_SCATTER, "CathepsinB_frac_around_cent_3um_grid.png"),
]

# Polarization grid PNGs: one per target field; each PNG embeds all
# configured thresholds in a panel grid. Output by
# assemble_polarization_grids.m at `<compiled_root>/<target>_polarization_grid.png`
# (directly under grid_panels/, NOT in a sub-folder). Filenames will be
# MISSING until the MATLAB compile is rerun with the polarization flag
# turned on for MT_CatB (config: d20231127_20240620_20240624_compilation_config.m).
POLARIZATION_METRICS = [
    ("", "centrosome_center_z_cell_bottom_distance_polarization_grid.png"),
    ("", "CathepsinB_z50_rel_cell_bottom_polarization_grid.png"),
    ("", "CathepsinB_z75_rel_cell_bottom_polarization_grid.png"),
]

# All 17 foci PNGs: actin foci first, then foci x CatB colocalization.
FOCI_METRICS = [
    # --- actin foci (8) ----------------------------------------------------
    (SUB_SCATTER, "actin_foci_count_grid.png"),
    (SUB_SCATTER, "actin_foci_area_um2_grid.png"),
    (SUB_SCATTER, "actin_foci_area_fraction_grid.png"),
    (SUB_SCATTER, "actin_foci_mean_intensity_grid.png"),
    (SUB_SCATTER, "actin_foci_total_intensity_grid.png"),
    (SUB_SCATTER, "actin_foci_mean_norm_score_grid.png"),
    (SUB_SCATTER, "actin_foci_max_norm_score_grid.png"),
    # --- foci x CatB colocalization (9) ------------------------------------
    (SUB_SCATTER, "foci_CathepsinB_mean_intensity_synapse_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_enrichment_ratio_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_foci_overlap_fraction_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_granule_overlap_fraction_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_mean_granule_intensity_in_foci_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_mean_granule_intensity_out_foci_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_m1_granule_intensity_in_foci_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_m2_foci_intensity_in_granules_grid.png"),
    (SUB_SCATTER, "foci_CathepsinB_pearsons_coeff_grid.png"),
]

# Each dataset -> one deck. compiled_root points at the grid_panels/ folder.
OUTPUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/CART/actin_only")

DATASETS = [
    {
        "label": "Kiet 2026",
        "compiled_root": Path(
            "Y:/User_data/Kiet/results_compiled/actin_only/"
            "compiled_20260312_20260414_20260510_20260604/grid_panels"
        ),
        "output_path": OUTPUT_DIR / "CART_actin_summary.pptx",
        # Kiet's CAT_vs_FMC dir has no 3slice rad_profile PNG, so we
        # leave both rad_profile entries off this deck (18 slides).
        "png_files": SLICE_METRICS + THREE_SLICE_METRICS,
    },
    {
        "label": "2024-06 data",
        "compiled_root": Path(
            "J:/FF/fixed_cell/CAR_TCell/results_compiled/actin_only/"
            "compiled_20240620_20240624_20260605/grid_panels"
        ),
        "output_path": OUTPUT_DIR / "CART_actin_summary_202406_data.pptx",
        "png_files": SLICE_METRICS_WITH_RAD + THREE_SLICE_METRICS_WITH_RAD,
    },
    {
        "label": "Kiet 20260607",
        # Compile name: <dataset_folder>_<compile_date> (no `_test_` segment).
        "compiled_root": Path(
            "Y:/User_data/Kiet/results_compiled/actin_only/"
            "compiled_20260607_pMLC_CART_actin_hoescht_20260616/grid_panels"
        ),
        "output_path": OUTPUT_DIR / "CART_actin_summary_20260607.pptx",
        "png_files": SLICE_METRICS_WITH_RAD + THREE_SLICE_METRICS_WITH_RAD,
    },
    {
        "label": "5min by_timepoint",
        # Flat layout: PNGs live directly under compiled_root with no
        # _grid suffix. strip_grid_suffix=True tells build_deck to drop
        # `_grid` when resolving disk paths while keeping the canonical
        # names for prettify / TITLE_OVERRIDES lookup.
        "compiled_root": _by_timepoint("5min"),
        "output_path": OUTPUT_DIR / "CART_actin_summary_5min.pptx",
        "png_files": ALL_CONDITIONS_PNG_FILES,
        "strip_grid_suffix": True,
    },
    {
        "label": "All timepoints",
        # 60 slides: 20 metrics x 3 timepoints, interleaved per-metric.
        # Each metric gets 3 consecutive slides (5/10/15 min).
        "timepoints": [
            ("5 min",  _by_timepoint("5min")),
            ("10 min", _by_timepoint("10min")),
            ("15 min", _by_timepoint("15min")),
        ],
        "output_path": OUTPUT_DIR / "CART_actin_summary_all_timepoints.pptx",
        "png_files": ALL_CONDITIONS_PNG_FILES,
        "strip_grid_suffix": True,
    },
    {
        "label": "MT_CatB",
        # 30 slides: 1 centrosome + 1 cent-polarization + 9 CatB +
        # 2 CatB-z-polarization + 17 foci. Polarization slides will
        # render "(missing)" until the MATLAB compile is rerun with the
        # newly-enabled polarization flag.
        "compiled_root": MT_CATB_ROOT,
        "output_path": OUTPUT_DIR_MT_CATB / "CART_summary_MT_CatB.pptx",
        "png_files": (
            CENTROSOME_METRICS                              # 1: distance
            + [POLARIZATION_METRICS[0]]                     # 1: centrosome polarized
            + CATB_METRICS_CONCISE                          # 8: CatB analogs
            + CATB_Z50_METRICS + [POLARIZATION_METRICS[1]]  # 3: Z50 raw + polarized
            + CATB_Z75_METRICS + [POLARIZATION_METRICS[2]]  # 3: Z75 raw + polarized
            + FOCI_METRICS                                  # 16 foci
        ),
    },
]

# Colors
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)

# Slide layout (inches). 13.333 x 7.5 widescreen.
SLIDE_W = 13.333
SLIDE_H = 7.5

MARGIN = 0.10

TITLE_LEFT = MARGIN
TITLE_TOP = 0.05
TITLE_WIDTH = SLIDE_W - 2 * MARGIN
TITLE_HEIGHT = 0.50
TITLE_FONT_PT = 28

SUBTITLE_LEFT = MARGIN
SUBTITLE_TOP = 0.55
SUBTITLE_WIDTH = SLIDE_W - 2 * MARGIN
SUBTITLE_HEIGHT = 0.32
SUBTITLE_FONT_PT = 14

# Image placement: bigger when there's no subtitle, smaller when there is.
IMG_LEFT = MARGIN
IMG_TOP = 0.60
IMG_BOX_W = SLIDE_W - 2 * MARGIN
IMG_BOX_H = 6.40

IMG_TOP_WITH_SUBTITLE = 0.90
IMG_BOX_H_WITH_SUBTITLE = 6.10

FOOTER_LEFT = MARGIN
FOOTER_TOP = 7.05
FOOTER_WIDTH = SLIDE_W - 2 * MARGIN
FOOTER_HEIGHT = 0.40
FOOTER_FONT_PT = 9

# Per-token substitutions for prettify_metric_name (keys are lowercase).
TOKEN_MAP = {
    "mfi": "MFI",
    "mip": "MIP",
    "sig": "Signal",
    "num": "Number",
    "idx": "Index",
    "rel": "Rel.",
    "3slice": "(3-Slice MIP)",
    "bottom": "Synapse",
    # MT_CatB additions
    "cathepsinb": "CatB",
    "3mip": "(3-Slice MIP)",
    "um": "µm",
    "um2": "µm²",
    "cent": "Centrosome",
    "frac": "Fraction",
    "fdd": "FDD",
    "3d": "3D",
    "rms": "RMS",
    "zcof": "Z-COF",
}

# Tokens dropped entirely from the title.
SKIP_TOKENS = {"slice"}

# Per-filename overrides — applied before any other title logic.
TITLE_OVERRIDES = {
    "actin_bottom_mask_area_grid.png": "Synapse Area",
    "actin_bottom_3slice_mask_area_grid.png": "Synapse (3-Slice MIP) Area",
    "actin_bottom_slice_rad_profile_auc1_all_cells_with_average_grid.png":
        "Actin Synapse Radial Profile AUC₁ (all cells + avg)",
    "actin_bottom_3slice_mip_rad_profile_auc1_all_cells_with_average_grid.png":
        "Actin Synapse (3-Slice MIP) Radial Profile AUC₁ (all cells + avg)",
    # --- MT_CatB compile overrides ----------------------------------------
    "centrosome_center_z_cell_bottom_distance_grid.png":
        "Centrosome-Synapse Distance",
    "centrosome_center_z_cell_bottom_distance_polarization_grid.png":
        "Centrosome-Synapse Distance: Fraction Polarized",
    "CathepsinB_z50_rel_cell_bottom_grid.png":
        "CatB Z₅₀ (µm from cell bottom)",
    "CathepsinB_z50_norm_cell_height_grid.png":
        "CatB Z₅₀ (fraction of cell height)",
    "CathepsinB_z50_rel_cell_bottom_polarization_grid.png":
        "CatB Z₅₀: Fraction Polarized",
    "CathepsinB_z75_rel_cell_bottom_grid.png":
        "CatB Z₇₅ (µm from cell bottom)",
    "CathepsinB_z75_norm_cell_height_grid.png":
        "CatB Z₇₅ (fraction of cell height)",
    "CathepsinB_z75_rel_cell_bottom_polarization_grid.png":
        "CatB Z₇₅: Fraction Polarized",
    "CathepsinB_granule_area_um2_grid.png": "CatB Synapse Area",
    "CathepsinB_synapse_g_ave_grid.png": "CatB Synapse g(r) Average",
    "CathepsinB_synapse_g_ave_3mip_grid.png":
        "CatB Synapse g(r) Average (3-Slice MIP)",
    "CathepsinB_synapse_r_eff_grid.png": "CatB Synapse Effective Radius",
    "CathepsinB_synapse_r_eff_3mip_grid.png":
        "CatB Synapse Effective Radius (3-Slice MIP)",
    "CathepsinB_synapse_autocorr_rmax_um_grid.png":
        "CatB Synapse Autocorrelation: r_max (µm)",
    "CathepsinB_synapse_autocorr_smoothSig_grid.png":
        "CatB Synapse Autocorrelation: smoothed σ",
    "CathepsinB_zCOF_grid.png": "CatB Z-Center of Fluorescence",
    "CathepsinB_zCOF_actin_scale_grid.png":
        "CatB Z-COF (actin scale)",
    "CathepsinB_zCOF_cell_bottom_distance_grid.png":
        "CatB Z-COF: Distance from Cell Bottom",
    "CathepsinB_zCOF_cell_bottom_distance_norm_cell_height_grid.png":
        "CatB Z-COF: Distance from Cell Bottom (norm. cell height)",
}

# Optional per-filename subtitle shown under the title (formula / units /
# interpretation). Missing entries -> no subtitle line. Currently empty.
SUBTITLE_BY_FILENAME = {}

# ---------------------------------------------------------------------------


def prettify_metric_name(filename: str) -> str:
    """Convert a PNG filename to a readable slide title.

    Checks TITLE_OVERRIDES first. Otherwise strips `.png` and `_grid`,
    splits on `_`, applies per-token substitutions (dropping tokens in
    SKIP_TOKENS and collapsing the redundant `3slice_mip` pair into a
    single `(3-Slice MIP)`), and rewrites `Inner Outer` as `Inner/Outer`.
    """
    if filename in TITLE_OVERRIDES:
        return TITLE_OVERRIDES[filename]

    stem = filename
    if stem.lower().endswith(".png"):
        stem = stem[:-4]
    if stem.lower().endswith("_grid"):
        stem = stem[:-5]

    tokens = stem.split("_")
    out = []
    i = 0
    while i < len(tokens):
        tok = tokens[i]
        low = tok.lower()
        nxt = tokens[i + 1].lower() if i + 1 < len(tokens) else None

        if low == "3slice" and nxt == "mip":
            out.append("(3-Slice MIP)")
            i += 2
            continue
        if low in SKIP_TOKENS:
            i += 1
            continue
        if low in TOKEN_MAP:
            out.append(TOKEN_MAP[low])
        elif low.endswith("pct") and low[:-3].isdigit():
            out.append(f"({low[:-3]}% Eff Rad Thresh)")
        elif low.endswith("um2") and low[:-3].isdigit():
            out.append(f"{low[:-3]} µm²")
        elif low.endswith("um") and low[:-2].isdigit():
            out.append(f"{low[:-2]} µm")
        else:
            out.append(tok.capitalize())
        i += 1

    pretty = " ".join(out)
    pretty = pretty.replace("Inner Outer", "Inner/Outer")
    return pretty


def add_textbox(slide, text, left, top, width, height, font_pt, color,
                bold=False, italic=False):
    """Add a centered textbox with given text and styling."""
    box = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = box.text_frame
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.text = text
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.runs[0]
    run.font.size = Pt(font_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def add_image_in_box(slide, image_path, box_left, box_top, box_w, box_h):
    """Place an image inside the given bounding box, preserving aspect
    ratio and centering on the dimension that is < box."""
    pic = slide.shapes.add_picture(
        image_path,
        Inches(box_left),
        Inches(box_top),
        width=Inches(box_w),
    )
    actual_h_in = pic.height / 914400.0
    if actual_h_in > box_h:
        sp = pic._element
        sp.getparent().remove(sp)
        pic = slide.shapes.add_picture(
            image_path,
            Inches(box_left),
            Inches(box_top),
            height=Inches(box_h),
        )
        actual_w_in = pic.width / 914400.0
        pic.left = Inches(box_left + (box_w - actual_w_in) / 2)
    else:
        pic.top = Inches(box_top + (box_h - actual_h_in) / 2)
    return pic


def set_slide_background(slide, rgb: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def build_slide(prs, title_text: str, image_path: Path, footer_text: str,
                subtitle_text=None):
    """Build one slide: title + optional subtitle + full image + filepath footer.
    Returns (slide, missing_flag)."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, WHITE)

    add_textbox(
        slide, title_text,
        TITLE_LEFT, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT,
        font_pt=TITLE_FONT_PT, color=BLACK, bold=True,
    )

    if subtitle_text:
        add_textbox(
            slide, subtitle_text,
            SUBTITLE_LEFT, SUBTITLE_TOP, SUBTITLE_WIDTH, SUBTITLE_HEIGHT,
            font_pt=SUBTITLE_FONT_PT, color=BLACK, italic=True,
        )
        img_top, img_h = IMG_TOP_WITH_SUBTITLE, IMG_BOX_H_WITH_SUBTITLE
    else:
        img_top, img_h = IMG_TOP, IMG_BOX_H

    missing = not image_path.exists()
    if not missing:
        add_image_in_box(
            slide, str(image_path), IMG_LEFT, img_top, IMG_BOX_W, img_h
        )
    else:
        add_textbox(
            slide, "(missing)",
            IMG_LEFT, img_top + img_h / 2 - 0.2, IMG_BOX_W, 0.4,
            font_pt=18, color=BLACK,
        )

    add_textbox(
        slide, footer_text,
        FOOTER_LEFT, FOOTER_TOP, FOOTER_WIDTH, FOOTER_HEIGHT,
        font_pt=FOOTER_FONT_PT, color=BLACK, bold=False,
    )
    return slide, missing


def build_deck(dataset) -> int:
    """Build one deck for the given dataset entry. Returns missing-file count."""
    label = dataset["label"]
    output_path = dataset["output_path"]
    png_files = dataset["png_files"]
    strip_grid = dataset.get("strip_grid_suffix", False)
    # Either a list of (timepoint_label, root) pairs (combined deck),
    # or fall back to the single compiled_root field. The single-root
    # case is modeled as a 1-element list with empty timepoint label.
    timepoints = dataset.get("timepoints") or [("", dataset["compiled_root"])]

    output_path.parent.mkdir(parents=True, exist_ok=True)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    print(f"\n=== Dataset: {label} ===")
    if len(timepoints) == 1 and not timepoints[0][0]:
        print(f"Source root: {timepoints[0][1]}")
    else:
        print(f"Timepoints: {', '.join(tp for tp, _ in timepoints)}")
    print(f"Writing deck to: {output_path}\n")

    missing = []
    slides_written = 0
    for sub, png_name in png_files:
        for tp_label, tp_root in timepoints:
            disk_name = png_name.replace("_grid.png", ".png") if strip_grid else png_name
            image_path = tp_root / sub / disk_name
            base_title = prettify_metric_name(png_name)
            title = f"{base_title} ({tp_label})" if tp_label else base_title
            footer = image_path.as_posix()
            subtitle = SUBTITLE_BY_FILENAME.get(png_name)
            _, is_missing = build_slide(
                prs, title, image_path, footer, subtitle_text=subtitle
            )
            slides_written += 1
            status = "OK" if not is_missing else "MISSING"
            tag = f"{png_name} @ {tp_label}" if tp_label else png_name
            print(f"[{tag}]  {status}  -> {title!r}")
            if is_missing:
                missing.append(f"{png_name}  ({image_path})")

    # Snapshot the previous deck (if any) into backups/ before overwriting.
    if output_path.exists():
        backup_dir = output_path.parent / "backups"
        created = backup_presentation(str(output_path), backup_base=str(backup_dir))
        if created:
            print(f"\nBacked up previous deck to: {backup_dir}")
            for cat, path in created.items():
                print(f"  [{cat:6}] {path}")

    prs.save(str(output_path))
    print(f"\nDone. {slides_written} slides written to:\n  {output_path}")

    if missing:
        print(f"\nMissing ({len(missing)}):")
        for m in missing:
            print(f"  - {m}")
    else:
        print("\nAll images found - no missing items.")

    return len(missing)


def main() -> None:
    total_missing = 0
    for dataset in DATASETS:
        total_missing += build_deck(dataset)
    if total_missing:
        sys.exit(1)


if __name__ == "__main__":
    main()
