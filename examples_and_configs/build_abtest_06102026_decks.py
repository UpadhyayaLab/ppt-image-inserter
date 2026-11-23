#!/usr/bin/env python
"""
Results and montage decks for the 06/10/2026 Jurkat antibody test (CD3+ cells on PLL, H3K27me3, H3K27ac, H3),
in the format of the 10/29/2026 Noco H3K27me3 decks (FF 2026-11-21).

Results deck: the slide list of config_Noco_H3K27me3_10292026_fixed_cells.yaml ("Effects of Noco, H3K27me3,
10292026.pptx"), same order and template, run once per compile. The experiment has three compiles, one per mark,
each comparing the wells that carry that mark; wells are never pooled:
    H3K27me3: A1 (H3K27me3 + H3K27ac) vs A2 (H3K27me3)
    H3K27ac:  A1 (H3K27me3 + H3K27ac) vs B1 (H3K27ac)
    H3:       B2 alone
Left out of the 10/29 list: every centrosome slide (no centrin channel in this experiment), the near / away from
centrosome profiles, and the statistics-table placeholder. "Partial cells left out" is dropped from subtitles: no
mesh review marked Partial cells here (no partial_exclusions_applied.csv in the compiles).
Each slide's subtitle names the comparison with the compile's own condition labels.

Montage deck: the build_h3k27me3_10292026_montage_decks.py format (one slide per pipeline montage chunk), the
H3K27me3 + H3K27ac well (A1) first, then every A1 merge option (FF had not picked colors, contrast or panels; one
section per panel layout), then the single-mark wells, then side views. Every montage carries a readable cell number
above each tile (MONTAGE_ROOT, the 2026-11-21 re-montage; the run's own montages are unchanged).

Both decks: the generated config is written next to the reusable configs and run through batch_insert_images.py
(title_size 32, lint and approved-text check on every build).

Usage (base miniforge python, from PowerShell):
    python examples_and_configs/build_abtest_06102026_decks.py [results|montages|all] [--compile-date YYYYMMDD]
"""

import argparse
import os
import re
import subprocess
import sys
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
CFG_DIR = HERE / "configs" / "fixed" / "Jurkats" / "histone_mods"
APPROVED = CFG_DIR / "abtest_06102026_approved_text.yaml"

EXP = "L:/FF/Nucleus_H3K27me3/06102026_jurkatsCD3plus_onPLL_testingAntibodiesH3K27me3H3K27ac_"
OUT_DIR = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Histone Mods_RNAPII/H3K27_AbTest_20260610"
TEMPLATE = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Noco/Effects of Noco, Fixed Cells_MT_template.pptx"
# Proposed to FF 2026-11-21; not approved yet (listed by the approved-text check until it is).
TITLE_SLIDE = "Jurkat antibody test, H3K27me3, H3K27ac and H3, June 10, 2026"
# The compiles were rerun 2026-11-21 without the actin-below-FOV filter (FF's choice, session "H3K27me3 06102026:
# Cellpose crop + QC deck"); 20261120 is the earlier run.
COMPILE_DATE = "20261121"

# (mark, compile folder stem, grid_panels subfolder, comparison shown in the subtitle = the compile's labels)
COMPILES = [
    ("H3K27me3", "Jurkats_H3K27_AbTest_H3K27me3", "A1_vs_A2", "H3K27me3 + H3K27ac vs H3K27me3"),
    ("H3K27ac", "Jurkats_H3K27_AbTest_H3K27ac", "A1_vs_B1", "H3K27me3 + H3K27ac vs H3K27ac"),
    ("H3", "Jurkats_H3K27_AbTest_H3", "B2_H3", "H3"),
]

SURF = "Within 0.25 μm of the nuclear surface"
# The 10/29 results slides kept here, in its order: (path under the compile, title, extra subtitle line or None).
# {M} = mark; G = grid_panels/<pair>; C = <mark>_loc_wrto_invag (the curated set).
RESULTS = [
    ("C/09a_cell_counts_filter_breakdown.png", "Cell counts per QC filter", None),
    ("G/deepest_invag_volume.png", "Deepest invagination volume", None),
    ("G/deepest_invag_fraction_chull_volume.png", "Deepest invag: frac of convex hull volume", None),
    ("G/{M}_MFI.png", "{M} MFI", None),
    ("G/{M}_deepest_region_periph_ratio_025um.png", "{M} levels near deepest invagination", None),
    # No prefix = the DNA channel (generate_geom_density_seed_depth_study.m local_stored_fields).
    ("G/deepest_region_periph_ratio_025um.png", "DNA levels near deepest invagination", None),
    ("C/02_chan_hulldist_{M}.png", "Density vs hull-boundary distance", None),
    ("C/04a_{M}_correlation_hulldist.png", "Per-cell correlation vs hull distance", SURF),
    ("C/02_chan_mincurv_{M}.png", "Density vs min curvature", None),
    ("G/nuc_aspect_ratio.png", "Nuclear aspect ratio", None),
    ("G/nuc_solidity.png", "Nuclear solidity", None),
    ("G/nuc_volume_mesh.png", "Nuclear volume", None),
    ("G/nuc_SA_mesh.png", "Nuclear surface area", None),
    ("G/actin_bottom_mask_area.png", "Synapse Area", None),
    ("G/actin_bottom_MFI.png", "Actin MFI at synapse", None),
    ("G/actin_bottom_inner_outer_ratio.png", "Actin inner/outer synapse ratio", None),
    ("C/02_chan_meancurv_{M}.png", "Density vs mean curvature", None),
    ("C/03a_{M}_enrichment_hulldist_gt0.5um.png", "Enrichment in deep invaginations (hull distance > 0.5 μm)", SURF),
    ("C/03b_{M}_enrichment_hulldist_gt1.0um.png", "Enrichment in deep invaginations (hull distance > 1.0 μm)", SURF),
    ("C/03c_{M}_enrichment_minCurvature_lt0.png", "Enrichment on concave surface (min curv < 0)", SURF),
    ("C/03d_{M}_enrichment_minCurvature_ltm0.25.png", "Enrichment on strongly concave surface (min curv < −0.25)", SURF),
    ("C/03e_{M}_enrichment_meanCurvature.png", "Enrichment on concave surface (mean curv < 0)", SURF),
    ("C/04b_{M}_correlation_minCurvature.png", "Per-cell correlation vs min curvature", SURF),
    ("C/04c_{M}_correlation_meanCurvature.png", "Per-cell correlation vs mean curvature", SURF),
    ("C/06_chan_normalangle_{M}.png", "Density vs surface orientation", None),
]

# Montages: (family, legend). Legends follow the 10/29 wording; colors checked on the A1 montages 2026-11-21.
WELLS = {
    "A1": ("H3K27me3 + H3K27ac", "A1_H3K27me3_H3K27ac_"),
    "A2": ("H3K27me3", "A2_H3K27me3_"),
    "B1": ("H3K27ac", "B1_H3K27ac_"),
    "B2": ("H3", "B2_H3_"),
}
TOP = "Top view, broadest nuclear slice: {m} (red), DNA (cyan)"
# Montages with a readable cell number above each tile, and the A1 merge families (2026-11-21 overnight run,
# TCell-3D-Morphodynamics dev/runs/run_abtest_montages_20261121.m: relabel_abtest_montages + render_abtest_merges).
MONTAGE_ROOT = f"{EXP}/montage_render_20261121"
# A1 H3K27me3 + H3K27ac merge options (decision page https://claude.ai/artifact/LBwr6KzxdzLpj26bocwDp6 had no answer at
# render time, so every option is in the deck, one section per panel layout). Wording from that page's buttons.
MERGE_COLORS = [("mg", "magenta / green", "H3K27me3 magenta, H3K27ac green"),
                ("gm", "green / magenta", "H3K27me3 green, H3K27ac magenta"),
                ("rc", "red / cyan", "H3K27me3 red, H3K27ac cyan")]
MERGE_CONTRAST = [("cell", "per cell", "Contrast per cell"),
                  ("fixed", "fixed per mark", "Contrast fixed per mark across the well")]
MERGE_LAYOUTS = [("merge", "Merge only", None),
                 ("three", "3 panels", "Panels: H3K27me3, H3K27ac, merge"),
                 ("four", "4 panels with DNA", "Panels: H3K27me3, H3K27ac, DNA, merge with DNA blue")]
def merge_plan(colors=None, contrast=None, layouts=None):
    """(well, family, title stem, legend) per merge option, in slide order: layout, then colors, then contrast.
    colors / contrast / layouts: FF's picks (keys as above); None = every option."""
    plan = []
    for lay, lay_txt, panels in MERGE_LAYOUTS:
        if layouts and lay not in layouts:
            continue
        for ck, c_txt, c_key in MERGE_COLORS:
            if colors and ck not in colors:
                continue
            for k, k_txt, k_key in MERGE_CONTRAST:
                if contrast and k not in contrast:
                    continue
                legend = ". ".join(x for x in (panels, c_key, k_key) if x)
                plan.append(("A1", f"{lay}_{ck}_{k}", f"{lay_txt}, {c_txt}, {k_txt}", legend))
    return plan


MERGES = merge_plan()
MONTAGES = [  # (well, family, legend), in slide order; "MERGES" marks where the merge sections go
    ("A1", "H3K27me3_nuc_bz", TOP.format(m="H3K27me3")),
    ("A1", "H3K27ac_nuc_bz", TOP.format(m="H3K27ac")),
    ("MERGES", None, None),
    ("A2", "H3K27me3_nuc_bz", TOP.format(m="H3K27me3")),
    ("B1", "H3K27ac_nuc_bz", TOP.format(m="H3K27ac")),
    ("B2", "H3_nuc_bz", TOP.format(m="H3")),
    ("A1", "actin_nuc_xz_nolines", "Side view: actin (orange), DNA (cyan)"),
    ("A2", "actin_nuc_xz_nolines", "Side view: actin (orange), DNA (cyan)"),
    ("B1", "actin_nuc_xz_nolines", "Side view: actin (orange), DNA (cyan)"),
    ("B2", "actin_nuc_xz_nolines", "Side view: actin (orange), DNA (cyan)"),
]

COMMON = {
    "presentation": TEMPLATE,
    "template_slide": 1,
    "backup_dir": f"{OUT_DIR}/backups",
    "preserve_slides": [0, 1],
    "title_slide_text": TITLE_SLIDE,
    "auto_position": False,
    "position": {"left": 0.3, "top": 1.17, "width": 12.73, "height": 6.15},
    "title_size": 32,
    "label_side": "left",
    "add_label": False,
    "remove_template_slide": True,
    "approved_text": str(APPROVED).replace("\\", "/"),
}


def long_path(p) -> str:
    """Extended-length form: the well folders run past Windows' 260-character limit."""
    s = str(p).replace("/", "\\")
    return "\\\\?\\" + s if sys.platform == "win32" and not s.startswith("\\\\?\\") else s


def results_images(date: str):
    images, missing = [], []
    for mark, stem, pair, comparison in COMPILES:
        root = Path(f"{EXP}/results_compilation/{stem}_{date}")
        for rel, title, extra in RESULTS:
            rel = rel.replace("{M}", mark)
            rel = re.sub(r"^G/", f"grid_panels/{pair}/", rel)
            rel = re.sub(r"^C/", f"{mark}_loc_wrto_invag/", rel)
            path = root / rel
            if not path.exists():
                missing.append(str(path))
                continue
            sub = comparison if extra is None else f"{comparison}. {extra}"
            images.append({"path": str(path).replace("\\", "/"), "title": title.replace("{M}", mark),
                           "subtitle": sub})
    return images, missing


def chunks(montage_dir):
    """Montage chunk paths sorted by first cell id, or [] when the folder is missing."""
    try:
        names = os.listdir(long_path(montage_dir))
    except OSError:
        return []
    found = []
    for n in names:
        m = re.fullmatch(r"montage_cells_(\d+)_(\d+)\.png", n)
        if m:
            found.append((int(m.group(1)), int(m.group(2)), Path(montage_dir) / n))
    return sorted(found)


def montage_images():
    """One slide per montage chunk, every chunk of one view (top / side) at ONE displayed scale: the pipeline
    renders all chunks at one px/μm (5 μm bar = 105 px in every chunk, checked 2026-11-21 with
    check_scalebar_pixel_widths.py), so each picture is sized source px / the view's px/in, the px/in that fits the
    view's largest chunk in the box. Without this a 2-cell chunk fills the box and its nuclei look 3x larger."""
    from PIL import Image
    box = COMMON["position"]
    rows, missing = [], []
    plan = []   # (well, family, title stem or None, legend)
    for well, fam, legend in MONTAGES:
        if well == "MERGES":
            plan.extend(MERGES)
        else:
            plan.append((well, fam, None, legend))
    for well, fam, stem, legend in plan:
        cond, folder = WELLS[well]
        mdir = Path(f"{MONTAGE_ROOT}/{folder}/{fam}/montages")
        found = chunks(mdir)
        if not found:
            missing.append(str(mdir))
            continue
        # one displayed scale per view: top views (single families and merge only), each strip layout, side views
        view = "side" if "_xz" in fam else fam.split("_")[0] if fam.startswith(("three_", "four_")) else "top"
        # cell ids that have a tile (the run's cells folder, or the merge family's); a gap in a sheet's range is named
        src = (Path(f"{MONTAGE_ROOT}/{folder}/{fam}/cells") if stem else
               Path(f"{EXP}/prog_fixed_cells/{folder}/physical_scale_images/{fam}/cells"))
        ids = {int(m.group(1)) for n in os.listdir(long_path(src)) if (m := re.match(r"Cell_(\d+)_", n))}
        # a two-mark well's single-mark sheets name the mark shown
        mark = fam.split("_")[0] if not stem and "+" in cond and fam.endswith("_nuc_bz") else None
        for a, b, p in found:
            head = stem or (f"{cond}, {mark} channel" if mark else cond)
            title = f"{head}, cells {a}–{b}"
            gap = [str(i) for i in range(a, b + 1) if i not in ids]
            note = f" Cell{'s' if len(gap) > 1 else ''} {', '.join(gap)} not shown." if gap else ""
            with Image.open(long_path(p)) as im:
                rows.append((view, title, legend + "." + note, p, im.size))
    ppi = {}
    for view, *_, (w, h) in rows:
        ppi[view] = max(ppi.get(view, 0), w / box["width"], h / box["height"])
    images = []
    for view, title, legend, p, (w, h) in rows:
        wi, hi = w / ppi[view], h / ppi[view]
        pos = {"left": round(box["left"] + (box["width"] - wi) / 2, 4),
               "top": round(box["top"] + (box["height"] - hi) / 2, 4),
               "width": round(wi, 4), "height": round(hi, 4)}
        images.append({"images": [long_path(p)], "positions": [pos], "title": title,
                       "subtitle": f"{legend} Bar = 5 μm."})
    for view, v in ppi.items():
        print(f"[scale] {view} view: {v:.1f} source px per displayed inch on every slide")
    return images, missing


def build(deck: str, date: str) -> int:
    if deck == "results":
        images, missing = results_images(date)
        cfg = dict(COMMON, output_path=f"{OUT_DIR}/H3K27 antibody test, results, 06102026.pptx",
                   trim_whitespace=True, images=images)
    else:
        images, missing = montage_images()
        cfg = dict(COMMON, output_path=f"{OUT_DIR}/H3K27 antibody test, montages, 06102026.pptx",
                   trim_whitespace=False, images=images)
    lock = Path(cfg["output_path"]).with_name("~$" + Path(cfg["output_path"]).name)
    if lock.exists():
        print(f"[STOP] {lock.name} exists: the deck is open in PowerPoint. Close it and rerun.")
        return 2
    cfg_path = CFG_DIR / f"config_abtest_06102026_{deck}_GENERATED.yaml"
    head = (f"# GENERATED by examples_and_configs/build_abtest_06102026_decks.py ({deck}); edit the builder, "
            f"not this file.\n")
    cfg_path.write_text(head + yaml.safe_dump(cfg, allow_unicode=True, sort_keys=False), encoding="utf-8")
    Path(OUT_DIR).mkdir(parents=True, exist_ok=True)
    rc = subprocess.run([sys.executable, str(HERE / "batch_insert_images.py"), str(cfg_path)], check=False).returncode
    print(f"[{deck}] {len(images)} slides requested")
    if missing:
        print(f"[MISSING] {len(missing)} panel(s):")
        for m in missing:
            print(f"  {m}")
    return rc or (1 if missing else 0)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("deck", nargs="?", default="all", choices=["results", "montages", "all"])
    ap.add_argument("--compile-date", default=COMPILE_DATE)
    # FF's merge picks (comma-separated keys); omitted = every option
    ap.add_argument("--colors", help="mg,gm,rc")
    ap.add_argument("--contrast", help="cell,fixed")
    ap.add_argument("--layouts", help="merge,three,four")
    args = ap.parse_args()
    MERGES[:] = merge_plan(*(s.split(",") if s else None for s in (args.colors, args.contrast, args.layouts)))
    decks = ["results", "montages"] if args.deck == "all" else [args.deck]
    sys.exit(max(build(d, args.compile_date) for d in decks))
