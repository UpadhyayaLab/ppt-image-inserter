#!/usr/bin/env python
"""
Generate the Naive CTL GzmB/beta-Tubulin 12-condition summary config.

This emits a static YAML config (reactivation-style) for
``batch_insert_images.py``. It scans the compiled-results directory and writes
one slide entry per (metric x comparison) grid panel, plus profile and
polarization sections -- **but only for files that actually exist**.

Why a generator? ``batch_insert_images.py`` pre-validates every listed image and
aborts the whole run (sys.exit(1)) if a single file is missing. Some
cross-condition metrics are absent from the 2-condition comparison folders, so
hand-authoring the ~175 entries is fragile. Re-run this whenever the data is
reprocessed (the results dir is date-stamped).

Usage:
    conda run -n PPT_editing python examples_and_configs/gen_naiveCTL_GzmB_bTub_summary_config.py
"""

import os

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
# Compiled-results root. base_dir is the ROOT (not grid_panels) because the
# profiles/ and polarization/ folders are siblings of grid_panels/.
RESULTS_ROOT = (
    "L:/FF/Naive_CTLs/GzB_bTub/06172025_naiveCTLs_grzmB_bTub_/compiled_results/"
    "naiveCTLs_grzmB_bTub_06172025_12conditions_20260611"
)

PRESENTATION = (
    "K:/FF/PPT/PPT_autogeneration/Naive_CTL/NaiveCTL_GzmB_bTub_summary_template.pptx"
)
OUTPUT_PATH = (
    "K:/FF/PPT/PPT_autogeneration/Naive_CTL/NaiveCTL_GzmB_bTub_summary.pptx"
)

# Where to write the generated config (beside the reactivation configs).
CONFIG_OUT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "configs", "fixed", "CTLs", "config_naiveCTL_GzmB_bTub_summary.yaml",
)

# Template layout slides:
#   1 = 1-image square panel   (default; activation + profiles)
#   2 = 1-image wide panel      (all_conditions + polarization)
#   3 = 3-up square panels      (substrate trio + timecourse trio on one slide)
SQUARE_SLIDE = 1
WIDE_SLIDE = 2
MULTI_SLIDE = 3

# ---------------------------------------------------------------------------
# Curated scalar metrics (ordered) -> (raw filename stem, pretty title)
# ---------------------------------------------------------------------------
METRICS = [
    # GzmB amount at synapse / overall
    ("GzmB_synapse_total_sig", "Granzyme B: total signal at synapse"),
    ("GzmB_synapse_total_sig_3mip", "Granzyme B: total signal at synapse (3-slice MIP)"),
    ("GzmB_total_sig", "Granzyme B: total signal"),
    # GzmB axial position (Z50 = median signal height above synapse).
    # Concise label matches the analysis ylabel convention "<Channel> Z_{50} (um)".
    ("GzmB_z50_cell_bottom_distance", "Granzyme B Z₅₀"),
    ("GzmB_z50_cell_bottom_distance_norm_cell_height", "Granzyme B Z₅₀ (norm. cell height)"),
    ("GzmB_z75_cell_bottom_distance", "Granzyme B Z₇₅"),
    ("GzmB_z75_cell_bottom_distance_norm_cell_height", "Granzyme B Z₇₅ (norm. cell height)"),
    # Actin synapse footprint
    ("actin_bottom_mask_area", "Synapse area"),
]

# ---------------------------------------------------------------------------
# Per-metric grid layout (ordered). Each panel is one of:
#   ("single", folder, label, is_wide)          -> 1 image (wide or square)
#   ("multi",  [folder, folder, ...], label)     -> several images on one slide
# Related comparisons are grouped onto single multi-image slides to keep the
# deck compact (substrate 24/48/72 together; the three timecourses together).
# ---------------------------------------------------------------------------
GRID_LAYOUT = [
    ("single", "all_conditions", "all conditions", True),
    ("multi", ["substrate_24h", "substrate_48h", "substrate_72h"],
        "substrate comparison (24 / 48 / 72 h)"),
    ("multi", ["gel_1p5kPa_timecourse", "gel_12kPa_timecourse", "glass_timecourse"],
        "timecourses: 1.5 kPa, 12 kPa, glass"),
    ("single", "activation_2h", "effects of activation: aCD3 vs PLL (2 h)", False),
]

# ---------------------------------------------------------------------------
# Profiles (axial + autocorrelation) and polarization shown ACROSS comparisons
# (same GRID_LAYOUT as the metrics: all conditions + substrate trio + timecourse
# trio + activation). The all-conditions panel falls back to the FLAT profiles/
# or polarization/ file when grid_panels/all_conditions/<raw>.png is absent; the
# per-comparison panels are read from grid_panels/<comparison>/<raw>.png and are
# skipped until they exist (some empty at 2 h due to low Granzyme B signal).
# Each entry: (raw filename stem, pretty title).
# ---------------------------------------------------------------------------
# profiles -> all-conditions square (live in profiles/)
PROFILE_EXTRAS = [
    ("GzmB_axial_profile_auc1", "Granzyme B: axial intensity profile"),
    ("GzmB_axial_profile_auc1_all_cells_with_average", "Granzyme B: axial intensity profile (all cells)"),
    ("GzmB_synapse_autocorr_c", "Granzyme B: pair autocorrelation at synapse"),
    ("GzmB_synapse_autocorr_c_all_cells_with_average", "Granzyme B: pair autocorrelation at synapse (all cells)"),
    ("GzmB_synapse_autocorr_c_3mip", "Granzyme B: pair autocorrelation at synapse, 3-slice MIP"),
    ("GzmB_synapse_autocorr_c_3mip_all_cells_with_average", "Granzyme B: pair autocorrelation at synapse, 3-slice MIP (all cells)"),
]
# SPATIAL polarization -> all-conditions wide (lives in polarization/). These are
# true spatial-polarization plots (granule height / centrosome distance) and keep
# the *_polarization suffix.
POLARIZATION_EXTRAS = [
    ("GzmB_z50_cell_bottom_distance_polarization", "Granzyme B Z₅₀: fraction polarized"),
    ("GzmB_z50_cell_bottom_distance_3um_polarization", "Granzyme B Z₅₀: fraction polarized (3 µm)"),
    ("GzmB_z75_cell_bottom_distance_polarization", "Granzyme B Z₇₅: fraction polarized"),
    ("GzmB_z75_cell_bottom_distance_3um_polarization", "Granzyme B Z₇₅: fraction polarized (3 µm)"),
]
# ABUNDANCE (% of cells; NOT spatial) -> all-conditions wide (lives in polarization/).
# Stacked "% of cells" bars: positive vs negative / high vs low total. These use the
# *_fraction suffix (renamed from the old *_sig_polarization files, now removed).
ABUNDANCE_EXTRAS = [
    ("GzmB_peak_sig_fraction", "Granzyme B: % positive cells (peak > 20)"),
    ("GzmB_total_sig_fraction", "Granzyme B: % cells with high total signal (> 1M)"),
]


def exists(rel_path):
    """True if rel_path (relative to RESULTS_ROOT) exists on disk."""
    return os.path.exists(os.path.join(RESULTS_ROOT, rel_path))


def yq(text):
    """Quote a string for a double-quoted YAML scalar."""
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def emit_entry(lines, rel_path, title, wide):
    """Append one single-image slide entry block to lines."""
    lines.append('  - images: [{}]'.format(yq(rel_path)))
    if wide:
        lines.append('    template_slide: {}'.format(WIDE_SLIDE))
    lines.append('    title: {}'.format(yq(title)))
    lines.append('')


def emit_multi(lines, rel_paths, title):
    """Append one multi-image slide entry block (uses the 3-up square layout)."""
    flow = ", ".join(yq(r) for r in rel_paths)
    lines.append('  - images: [{}]'.format(flow))
    lines.append('    template_slide: {}'.format(MULTI_SLIDE))
    lines.append('    title: {}'.format(yq(title)))
    lines.append('')


def emit_across_layout(lines, skipped, raw, pretty, allcond_dir=None, allcond_wide=True):
    """Emit one metric across GRID_LAYOUT (all conditions + substrate trio +
    timecourse trio + activation). Per-comparison panels come from
    grid_panels/<comparison>/<raw>.png; for all_conditions, fall back to
    <allcond_dir>/<raw>.png (flat) when the grid_panels copy is absent. Missing
    panels are skipped. Returns the number of slide entries emitted."""
    emitted = 0
    for panel in GRID_LAYOUT:
        if panel[0] == "single":
            _, folder, label, layout_wide = panel
            if folder == "all_conditions":
                rel = "grid_panels/all_conditions/{}.png".format(raw)
                if not exists(rel) and allcond_dir:
                    rel = "{}/{}.png".format(allcond_dir, raw)
                wide = allcond_wide
            else:
                rel = "grid_panels/{}/{}.png".format(folder, raw)
                wide = layout_wide
            if exists(rel):
                emit_entry(lines, rel, "{} — {}".format(pretty, label), wide)
                emitted += 1
            else:
                skipped.append(rel)
        else:  # "multi"
            _, folders, label = panel
            rels = ["grid_panels/{}/{}.png".format(f, raw) for f in folders]
            if all(exists(r) for r in rels):
                emit_multi(lines, rels, "{} — {}".format(pretty, label))
                emitted += 1
            else:
                skipped.extend(r for r in rels if not exists(r))
    return emitted


def main():
    lines = []
    skipped = []
    n = 0

    # ----- header -----
    lines += [
        "# Naive CTL GzmB / beta-Tubulin -- 12-condition summary deck",
        "#",
        "# AUTO-GENERATED by examples_and_configs/gen_naiveCTL_GzmB_bTub_summary_config.py",
        "# Edit the generator (metric list / titles / comparison order), not this file.",
        "#",
        "# Layout: metric -> grid panels (related comparisons grouped on one slide),",
        "# then profiles, then polarization.",
        "# Template slides: {} = square 1-image, {} = wide 1-image, {} = 3-up square.".format(
            SQUARE_SLIDE, WIDE_SLIDE, MULTI_SLIDE),
        "",
        "presentation: {}".format(yq(PRESENTATION)),
        "output_path: {}".format(yq(OUTPUT_PATH)),
        "template_slide: {}".format(SQUARE_SLIDE),
        "preserve_slides: [0, 1, 2, 3]",
        "base_dir: {}".format(yq(RESULTS_ROOT)),
        "auto_position: true",
        "add_label: false",
        "",
        "images:",
        "",
    ]

    # ----- section 1: grid panels (metric -> grouped comparisons) -----
    lines.append("  # =========================================================================")
    lines.append("  # GRID PANELS  (metric -> grouped comparisons)")
    lines.append("  # =========================================================================")
    lines.append("")
    for raw, pretty in METRICS:
        lines.append("  # --- {} ---".format(raw))
        lines.append("")
        n += emit_across_layout(lines, skipped, raw, pretty, allcond_dir=None, allcond_wide=True)

    # ----- section 2: profiles (axial + autocorrelation, across comparisons) -----
    lines.append("  # =========================================================================")
    lines.append("  # PROFILES  (axial + autocorrelation; across comparisons, square)")
    lines.append("  # =========================================================================")
    lines.append("")
    for raw, pretty in PROFILE_EXTRAS:
        lines.append("  # --- {} ---".format(raw))
        lines.append("")
        n += emit_across_layout(lines, skipped, raw, pretty,
                                allcond_dir="profiles", allcond_wide=False)

    # ----- section 3: spatial polarization (across comparisons) -----
    lines.append("  # =========================================================================")
    lines.append("  # POLARIZATION  (spatial; across comparisons; all-conditions panel is wide)")
    lines.append("  # =========================================================================")
    lines.append("")
    for raw, pretty in POLARIZATION_EXTRAS:
        lines.append("  # --- {} ---".format(raw))
        lines.append("")
        n += emit_across_layout(lines, skipped, raw, pretty,
                                allcond_dir="polarization", allcond_wide=True)

    # ----- section 4: abundance (% of cells; across comparisons) -----
    lines.append("  # =========================================================================")
    lines.append("  # ABUNDANCE  (% of cells; across comparisons; all-conditions panel is wide)")
    lines.append("  # =========================================================================")
    lines.append("")
    for raw, pretty in ABUNDANCE_EXTRAS:
        lines.append("  # --- {} ---".format(raw))
        lines.append("")
        n += emit_across_layout(lines, skipped, raw, pretty,
                                allcond_dir="polarization", allcond_wide=True)

    os.makedirs(os.path.dirname(CONFIG_OUT), exist_ok=True)
    with open(CONFIG_OUT, "w", encoding="utf-8") as f:
        f.write("\n".join(lines).rstrip() + "\n")

    print("Wrote {} slide entries to:".format(n))
    print("  {}".format(CONFIG_OUT))
    if skipped:
        print("\nSkipped {} not-yet-generated file(s) (will appear automatically on a "
              "later regenerate once the pipeline writes them):".format(len(skipped)))
        for s in skipped:
            print("  - {}".format(s))


if __name__ == "__main__":
    main()
