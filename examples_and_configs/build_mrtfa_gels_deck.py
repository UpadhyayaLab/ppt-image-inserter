#!/usr/bin/env python
"""
MRTF-A deck for CTLs on gels of three stiffnesses (FF 2026-11-04), the analog of the Noco decks
(docs/mrtfa_20261105_deck_plan.md). Three experiments, never pooled: one page per experiment for
every montage family, the three gels side by side (1.5 | 6 | 12 kPa), the first 12 kept cells
by id per gel, all tiles of a family at one scale across the three experiments. Then one results
slide per metric panel of the config's report_extra, in its order and with its captions, with the
config's per-experiment geom-density sections after the first (nadir) section: BYEXP profiles one per
slide, enrichment and Pearson r violins with the three experiments side by side. Before those, the
geodesic (pocket / deepest-region) figures the report does not curate, in the Noco deck's order (GEODESIC).

Everything comes from the compilation config (experiments, wells, gel labels, panels), and the
kept cells from the compile's seg_review_cells.csv (kept_by_compile), so after FF's exclusions
and the recompile, one run of this script rebuilds the deck.

No centrosome stain: no centrin side views, no centrosome-ordered hulls (see the plan).

Usage (base miniforge python):
    python examples_and_configs/build_mrtfa_gels_deck.py [--compile <dir>]
    python examples_and_configs/build_mrtfa_gels_deck.py --smoke
--smoke: Nov 3 only, every cell with a bottom slice (no compile needed), montage slides only,
into the MRTF-A data folder's tests\\ subfolder.
"""

import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path

import numpy as np
import scipy.io as sio
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import insert_noco_10262026_montage_slides as m  # noqa: E402
import build_noco_shared_deck_two_experiments as ns  # noqa: E402
from build_noco_shared_deck_two_experiments import plot_slide  # noqa: E402
from ppt_image_inserter.backup import backup_presentation  # noqa: E402

CONFIG = Path("C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/config/compilation/fixed/CTL/MRTF/"
              "CTL_Gels_MRTFA_compilation.m")
DATA_ROOT = Path("J:/FF/fixed_cell/CTL_nucleus/MRTF-A_Hoechst")
COMPILE_ROOT = DATA_ROOT / "compiled_results"
COMPILE_PREFIX = "CTL_Gels_MRTFA_"
OUT = Path("K:/FF/PPT/PPT_autogeneration/Transcription_factors/"
           "MRTF-A vs gel stiffness, CTLs (20231103, 20240205, 20240222).pptx")
SMOKE_OUT = DATA_ROOT / "tests" / "mrtfa_deck_smoke" / "mrtfa_deck_smoke.pptx"

PER_PAGE = 12               # FF 2026-11-04: first 12 kept cells by id per gel
PER_SHELF = 3               # 4 rows of 3 in each gel's column
SPLIT_PER_PAGE = 4          # 3-panel tiles are 3x wide: the same 12 cells over 3 pages
# Cell id on every tile (Noco turned them off, FF 2026-10-28; FF 2026-11-04: keep them here for now).
SHOW_IDS = True
ID_BOX = (60, 45)           # (w, h) px of the burned-in id, bottom right of the XY tiles
SPLIT_ID_BOX = (10 ** 5, 45)  # 3-panel tiles burn an id and small bars under each panel: blank the strip
BLOB_LEVEL = 30             # XY crop to the largest blob above this, so a neighbor's fragment stays out

MRTFA = ("MRTF-A", m.RED)   # pipeline render colors (save_channel_nuc_combos_physical.m)
DNA = ("DNA", m.CYAN)
ACTIN = ("Actin", m.ORANGE)
# (tag, physical-scale family or None for the synapse mask, tile kind, title, channel legend,
#  cells per page, footer), in deck order.
MONTAGES = [
    ("bz", "MRTFA_nuc_bz", "xy", "MRTF-A, broadest slice of nucleus", [MRTFA, DNA], PER_PAGE, ""),
    ("deep", "MRTFA_nuc", "xy", "MRTF-A, slice through deepest invagination", [MRTFA, DNA], PER_PAGE, ""),
    ("split", "MRTFA_nuc_3p_bz", "xy", "MRTF-A, broadest slice of nucleus",
     [("DNA", m.WHITE), ("MRTF-A", m.WHITE), ("Merge", m.WHITE)], SPLIT_PER_PAGE,
     "Left to right: DNA, MRTF-A, merge."),
    ("xz", "actin_nuc_xz_nolines", "xz", "Actin and DNA, side view", [ACTIN, DNA], PER_PAGE,
     "Each cell cut at its coverslip (actin bottom slice); all cells aligned on it."),
    ("synapse", None, "synapse", "Synapse, actin bottom slice", [("Actin", m.GRAY), ("Mask outline", m.RED)],
     PER_PAGE, "Red = synapse mask outline."),
]


def read_config(path: Path) -> dict:
    """Experiments, gels and report panels from the MATLAB compilation config."""
    text = path.read_text(encoding="utf-8")
    bases = dict(re.findall(r"(base_path_\w+) = '([^']+)';", text))
    subpath = re.search(r"subpath = '([^']+)';", text).group(1)
    exps = {}
    for k, name in re.findall(r"config\.experiments\((\d+)\)\.name = '([^']+)';", text):
        exps[int(k)] = {"name": name, "wells": {}}
    for k, f in re.findall(r"config\.experiments\((\d+)\)\.results_filename = '([^']+)';", text):
        exps[int(k)]["results"] = f
    for k, cond, base, well in re.findall(
            r"config\.experiments\((\d+)\)\.conditions\.(\w+) = \[(base_path_\w+), '([^']+)', subpath\];", text):
        exps[int(k)]["wells"][cond] = Path(bases[base] + well + subpath)

    def cell_list(field: str) -> list:
        return re.findall(r"'([^']+)'", re.search(rf"config\.{field} = \{{([^}}]*)\}};", text).group(1))

    # report_extra: x_ = { 'Section', { gp_('field'), 'caption'; ... } ... };
    block = re.search(r"\n\s*x_ = \{\n(.*?)\n\s*for k_ = ", text, re.DOTALL).group(1)
    panels = []
    section = None
    for line in block.splitlines():
        head = re.match(r"\s*'((?:[^']|'')+)', \{\s*$", line)
        if head:
            section = head.group(1).replace("''", "'")
        for field, caption in re.findall(r"gp_\('(\w+)'\),\s*'((?:[^']|'')*)'", line):
            panels.append((section, field, caption.replace("''", "'").replace("µ", "μ")))
    return {"experiments": [exps[k] for k in sorted(exps)],
            "conditions": list(zip(cell_list("condition_names"), cell_list("condition_labels"))),
            "panels": panels, "geom": read_geom_sections(text)}


def read_geom_sections(text: str) -> list:
    """The config's per-experiment geom-density report sections, mirroring its loops (lines 'ex_ =' to
    'its_ ='): [(section title, lead, [(caption, [png relative to the compile, one per experiment])])],
    MRTF-A then DNA within each section."""
    def block(name: str) -> str:
        return re.search(rf"\n\s*{name} = \{{(.*?)\}};", text, re.DOTALL).group(1)

    def pairs(name: str) -> list:
        s = re.findall(r"'([^']*)'", block(name))
        return list(zip(s[0::2], s[1::2]))

    def mu(s: str) -> str:
        return s.replace("µ", "μ")  # the config types the micro sign; slides use Greek mu

    dates = re.findall(r"'([^']+)'", block("ex_"))
    prof, enr, cor = pairs("prof_"), pairs("enr_"), pairs("cor_")
    keys, words = re.findall(r"containers\.Map\(\{([^}]*)\}, \{([^}]*)\}\)", text)[0]
    shell_words = dict(zip(re.findall(r"'([^']+)'", keys), re.findall(r"'([^']+)'", words)))
    sec = re.findall(r"'((?:[^']|'')*)'", block("sec_"))
    titles = [s for s in sec if s.endswith("by experiment")]
    leads = [" ".join(sec[sec.index(t) + 1:sec.index(titles[k + 1]) if k + 1 < len(titles) else None]).replace("''", "'")
             for k, t in enumerate(titles)]
    shells = {"MRTFA": ["inperinuc025", "perinuc05"], "DNA": ["inperinuc025"]}  # the config's sh_
    tags = [d.replace(" ", "_") for d in dates]   # sanitize_filename('Nov 3, 2023') = 'Nov_3,_2023'
    out = [(titles[q], mu(leads[q]), []) for q in range(3)]
    for p, label in (("MRTFA", "MRTF-A"), ("DNA", "DNA")):
        for s in shells[p]:
            for geom, word in prof:
                out[0][2].append((mu(f"{label} density vs {word}, {shell_words[s]}"),
                                  [f"geom_density/profiles/singles/{p}_geomdens_{geom}_{s}_BYEXP_line.png"]))
        for geom, word in enr:
            out[1][2].append((mu(f"{label} enrichment where {word}"),
                              [f"geom_density/enrichment/{p}_{geom}_shells_{t}.png" for t in tags]))
        for geom, word in cor:
            out[2][2].append((mu(f"{label} Pearson r vs {word}"),
                              [f"geom_density/enrichment/{p}_{geom}_shells_{t}.png" for t in tags]))
    return out


# Panels FF has not cleared, flagged in the subtitle (MRTF dataset configs session, 2026-11-05): with no 3D cell
# mask the cytoplasm is the actin MIP footprint through every plane; the MFI ratio built on it is dropped from the
# config, these summed-signal metrics are less affected but still pending.
PENDING = {f: " · PENDING REVIEW: cytoplasm = actin footprint through all planes (no 3D cell mask)"
           for f in ("MRTFA_nuc_fraction", "MRTFA_nuc_cyto_ratio")}

THREE_BOXES = ({"left": 0.1, "width": 4.3}, {"left": 4.52, "width": 4.3}, {"left": 8.93, "width": 4.3})


def three_slide(prs, title: str, subtitle: str, paths: list, dates=None) -> None:
    """A new white slide with its title and three plots side by side (one per experiment); dates given =
    the plots do not print theirs, so each gets its date above it."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for k, (path, box) in enumerate(zip(paths, THREE_BOXES)):
        top = ns.place(slide, path, box, ns.H3K_TOP + (ns.DATE_H if dates else 0))
        if dates:
            ns.add_date(slide, dates[k], box, top)
    ns._add_title(slide, title, subtitle, fit=True)


GEO_SECTION = "MRTF-A and DNA around the deepest invagination, geodesic"
GEO_LEAD = ("Pocket = the deepest invagination's geodesic region; distances run along the nuclear surface from its "
            "center. Density relative to the cell's mean in the shell; 1 = no enrichment. Outer 0.25 μm = inside the "
            "nucleus, Perinuclear = 0.5 μm outside. One figure or panel per experiment, never pooled.")
# The geodesic figures the compile draws but its report does not curate, in the Noco deck's order (region profile,
# region violins, then per-cell region metrics; FF 2026-11-05 "more metrics and profiles from the geodesic stuff,
# to more closely match noco"). (layout, title, path under the compile; {} = the experiment tag). Layouts:
# each = one slide per experiment, date in the title; three = the experiments side by side, a date over each
# (the figures do not print theirs); grid = one per-experiment grid.
GEODESIC = [
    ("each", "MRTF-A and DNA vs geodesic distance from the pocket center",
     "geom_density/profiles/pocket/MRTFA_x_DNA_geomdens_pocketdist_CHANPANEL_line_{}.png"),
    ("three", "MRTF-A and DNA within 2 μm of the pocket center, inside, outer 0.25 μm",
     "geom_density/pocket_violins/MRTFA_x_DNA_geomdens_pocketdisk_lt2um_inperinuc025_CHANPANEL_violin_{}.png"),
    ("three", "MRTF-A and DNA within 4 μm of the pocket center, inside, outer 0.25 μm",
     "geom_density/pocket_violins/MRTFA_x_DNA_geomdens_pocketdisk_lt4um_inperinuc025_CHANPANEL_violin_{}.png"),
    ("three", "MRTF-A within 2 and 4 μm of the pocket center, outside, 0.5 μm perinuclear",
     "geom_density/pocket_violins/MRTFA_geomdens_pocketdisk_perinuc05_violin_{}.png"),
    ("each", "MRTF-A pocket profile fits, inside, outer 0.25 μm",
     "geom_density/profiles/pocket/MRTFA_geomdens_pocketdist_fits_inperinuc025_violin_{}.png"),
    ("each", "MRTF-A pocket profile fits, outside, 0.5 μm perinuclear",
     "geom_density/profiles/pocket/MRTFA_geomdens_pocketdist_fits_perinuc05_violin_{}.png"),
    ("each", "DNA pocket profile fits, inside, outer 0.25 μm",
     "geom_density/profiles/pocket/DNA_geomdens_pocketdist_fits_inperinuc025_violin_{}.png"),
    ("grid", "MRTF-A near the deepest invagination, 0.5 μm shell",
     "grid_panels/MRTFA_deepest_region_geo025_periph_ratio_05um_per_experiment_grid.png"),
    ("grid", "DNA near the deepest invagination, 0.5 μm shell",
     "grid_panels/nuc_deepest_region_geo025_periph_ratio_05um_per_experiment_grid.png"),
    ("grid", "MRTF-A near the deepest invagination, 1 μm shell",
     "grid_panels/MRTFA_deepest_region_geo025_periph_ratio_1um_per_experiment_grid.png"),
    ("grid", "DNA near the deepest invagination, 1 μm shell",
     "grid_panels/nuc_deepest_region_geo025_periph_ratio_1um_per_experiment_grid.png"),
    ("grid", "MRTF-A intensity slope at the deepest invagination",
     "grid_panels/MRTFA_deepest_invag_geo025_gradient_slope_per_experiment_grid.png"),
    ("grid", "DNA intensity slope at the deepest invagination",
     "grid_panels/nuc_deepest_invag_geo025_gradient_slope_per_experiment_grid.png"),
    ("grid", "MRTF-A slope at the deepest invagination / elsewhere",
     "grid_panels/MRTFA_deepest_invag_geo025_gradient_ratio_per_experiment_grid.png"),
    ("grid", "DNA slope at the deepest invagination / elsewhere",
     "grid_panels/nuc_deepest_invag_geo025_gradient_ratio_per_experiment_grid.png"),
    ("grid", "Radius of the deepest invagination's geodesic region",
     "grid_panels/deepest_region_geo025_radius_um_per_experiment_grid.png"),
]


# --- Showcase: the cells with the most MRTF-A near the deepest invagination, 3-color and large (FF 2026-11-05) ---
SHOW_METRIC = "MRTFA_deepest_region_geo025_periph_ratio_025um"   # the deck's headline nadir metric
SHOW_FIELDS = ["cell_id", SHOW_METRIC, "nuc_deepest_region_geo025_periph_ratio_025um", "chull_max_D",
               "actin_bottom_slice_num", "deepest_invag_z_syn"]
SHOW_TOP = 4                # per experiment, any gel; one cell a slide
# Channel files per experiment (config/datasets/fixed/CTL/transcription_factors/CTL_Gels_MRTFA_<date>_dataset.m).
SHOW_CHANNELS = {"Nov 3, 2023": {"MRTFA": "C1", "Actin": "C3", "DNA": "C4"},
                 "Feb 5, 2024": {"MRTFA": "C3", "Actin": "C2", "DNA": "C4"},
                 "Feb 22, 2024": {"MRTFA": "C3", "Actin": "C2", "DNA": "C4"}}
SHOW_PSIZE = 0.065          # μm per raw xy pixel (the dataset configs' params.psize)
SHOW_UP = 3                 # raw pixels upsampled for the large tiles
SHOW_SEARCH = 5             # planes searched either side of the estimated deepest-invagination plane
# (name, RGB weight, color for the legend): MRTF-A and DNA in the deck's colors, actin dimmed gray under them.
SHOW_RGB = [("MRTF-A", (1.0, 0.0, 0.0), m.RED), ("DNA", (0.0, 1.0, 1.0), m.CYAN), ("Actin", (0.55, 0.55, 0.55), m.GRAY)]
# Panels left to right (FF 2026-11-05: "3 color, not just merge"): each channel alone in grayscale, then the merge.
SHOW_PANELS = [("MRTF-A", None, m.RED), ("DNA", None, m.CYAN), ("Actin", None, m.GRAY),
               ("Merge", None, RGBColor(0, 0, 0))]
SHOW_BOXES = tuple({"left": 0.1 + k * 3.29, "width": 3.2} for k in range(4))
SHOW_IMG_TOP, SHOW_IMG_BOTTOM = 1.75, 5.95


def show_values(prog: Path, results: str) -> dict:
    """{cell id: {field: value}} of the showcase fields from the well's results .mat."""
    mat = sio.loadmat(m._winlong(prog / results), squeeze_me=True, variable_names=SHOW_FIELDS)
    cols = {f: np.asarray(mat[f], float).ravel() for f in SHOW_FIELDS}
    ok = np.isfinite(cols["cell_id"])
    return {int(c): {f: cols[f][k] for f in SHOW_FIELDS} for k, c in enumerate(cols["cell_id"]) if ok[k]}


def deepest_plane(dna: np.ndarray, tile: Path, estimate) -> tuple:
    """(plane, match score): the raw plane whose DNA best matches the DNA (green) of the cell's MRTFA_nuc tile,
    the pipeline's render at the deepest-invagination plane, searched around the estimate (None = every plane)."""
    from skimage.feature import match_template
    from skimage.transform import rescale
    with Image.open(m._winlong(tile)) as im:
        g = np.asarray(im.convert("RGB"), float)[..., 1]
    g[-ID_BOX[1]:, -ID_BOX[0]:] = 0           # the burned-in id
    ys, xs = np.nonzero(g > 20)
    # Matched at raw resolution: the tile shrunk once to raw pixels, not every plane enlarged to the tile's.
    tmpl = rescale(g[ys.min():ys.max() + 1, xs.min():xs.max() + 1], 1 / (m.XZ_PPU * SHOW_PSIZE),
                   preserve_range=True, anti_aliasing=True)
    planes = range(dna.shape[0]) if estimate is None else \
        range(max(0, estimate - SHOW_SEARCH), min(dna.shape[0], estimate + SHOW_SEARCH + 1))
    best = (estimate, -1.0)
    for z in planes:
        img = dna[z].astype(float)
        if img.shape[0] < tmpl.shape[0] or img.shape[1] < tmpl.shape[1]:
            continue
        score = float(match_template(img, tmpl).max())
        if score > best[1]:
            best = (z, score)
    return best


def raw_to_syn(raw: float) -> float:
    """1-based raw slice -> fractional 1-based synthetic (isotropic, psize-spaced) index: TCell-3D's
    original_slice_to_resized, (i - 0.5) * zstep / psize + 0.5 (imresize3). nuc_*_slice_num and the _z<NNN> of the
    physical_scale_images renders are in this frame."""
    return (raw - 0.5) * m.XZ_ZSTEP_UM / SHOW_PSIZE + 0.5


def syn_to_raw(syn: float) -> float:
    """Inverse of raw_to_syn: fractional 1-based raw slice of a synthetic index."""
    return (syn - 0.5) * SHOW_PSIZE / m.XZ_ZSTEP_UM + 0.5


def render_plane(dna: np.ndarray, tile: Path, estimate) -> tuple:
    """(plane 0-based, z_syn, match plane 0-based, match score) for a cell's MRTFA_nuc render. The plane is the
    render's own z (Cell_<id>_z<NNN>.png, synthetic) taken to the raw grid; the DNA template match is kept as a
    cross-check, searched around it. A render without a z in its name falls back to the match (z_syn None)."""
    hit = re.search(r"_z(\d+)\.png$", tile.name, re.IGNORECASE)
    if hit is None:
        zm, score = deepest_plane(dna, tile, estimate)
        return zm, None, zm, score
    z_syn = int(hit.group(1))
    z = min(max(int(round(syn_to_raw(z_syn))) - 1, 0), dna.shape[0] - 1)
    zm, score = deepest_plane(dna, tile, z)
    return z, z_syn, zm, score


def three_color(stacks: dict, z: int, extra_dna: tuple = (), info: dict = None, extra_merges: tuple = (),
                crop: tuple = None, lo_override: dict = None, channel_rgb: list = None) -> list:
    """[MRTF-A, DNA, actin (grayscale), 3-color merge] of one plane, one crop and scale, then a grayscale DNA
    panel per plane in extra_dna (same crop and DNA contrast; planes outside the stack are skipped). Each channel
    is stretched from its median (the background: the cell fills well under half the crop) to its 99.8th
    percentile in that plane (per-cell contrast). A dict passed as info receives the crop box (raw pixel rows and
    columns, 0-based, end exclusive) and each channel's contrast limits. Each tuple of SHOW_RGB names in
    extra_merges adds, last, a merge of only those channels (same crop, contrast and colors). crop = (row0, row1,
    col0, col1), raw pixels, 0-based, end exclusive, replaces the automatic crop (contrast is unchanged).
    lo_override = {SHOW_RGB name: value} replaces that channel's display minimum (its maximum is unchanged).
    channel_rgb replaces SHOW_RGB (same (name, RGB weight, legend color) rows; stacks keyed by those names, as
    for another transcription factor in MRTF-A's slot: render_tf_picker_tiles.py)."""
    spec = channel_rgb or SHOW_RGB
    norm, dna_lohi, lohi = [], None, {}
    for name, _, _ in spec:
        plane = stacks[name.replace("MRTF-A", "MRTFA")][z].astype(float)
        lo, hi = np.percentile(plane, [50, 99.8])
        lo = (lo_override or {}).get(name, lo)
        norm.append(np.clip((plane - lo) / max(hi - lo, 1e-9), 0, 1))
        lohi[name] = (lo, hi)
        if name == "DNA":
            dna_lohi = (lo, hi)
    rgb = np.clip(sum(c[..., None] * np.asarray(w) for c, (_, w, _) in zip(norm, spec)), 0, 1)
    # Crop to the cell: the largest blob of the smoothed merge (camera noise defeats a plain non-black crop).
    from scipy import ndimage
    lab, nlab = ndimage.label(ndimage.gaussian_filter(rgb.max(axis=-1), 3) > 0.12)
    box = (slice(None), slice(None))
    if nlab:
        ys, xs = np.nonzero(lab == 1 + int(np.argmax(ndimage.sum(np.ones_like(lab), lab, range(1, nlab + 1)))))
        pad = 10
        box = (slice(max(0, ys.min() - pad), ys.max() + pad + 1), slice(max(0, xs.min() - pad), xs.max() + pad + 1))
    if crop is not None:
        box = (slice(crop[0], crop[1]), slice(crop[2], crop[3]))
    panels = [np.repeat(c[box][..., None], 3, axis=-1) for c in norm] + [rgb[box]]
    if info is not None:
        info["box"] = tuple((s.indices(n)[0], s.indices(n)[1]) for s, n in zip(box, rgb.shape[:2]))
        info["contrast"] = lohi
        info["norm"] = {name: c for (name, _, _), c in zip(spec, norm)}   # full-frame, contrast applied
    lo, hi = dna_lohi
    for ze in extra_dna:
        if 0 <= ze < stacks["DNA"].shape[0]:
            c = np.clip((stacks["DNA"][ze].astype(float) - lo) / max(hi - lo, 1e-9), 0, 1)
            panels.append(np.repeat(c[box][..., None], 3, axis=-1))
    for keep in extra_merges:
        sub = sum(c[..., None] * np.asarray(w) for c, (name, w, _) in zip(norm, spec) if name in keep)
        panels.append(np.clip(sub, 0, 1)[box])
    out = []
    for a in panels:
        im = Image.fromarray((a * 255).astype(np.uint8))
        out.append(im.resize((im.width * SHOW_UP, im.height * SHOW_UP), Image.LANCZOS))
    return out


def showcase_slides(prs, experiments: list, kept: dict, labels: dict, out_dir: Path) -> int:
    """SHOW_TOP kept cells per experiment with the highest SHOW_METRIC, SHOW_PER_SLIDE large 3-color tiles a
    slide, each cell's values printed under it. Tiles are written to out_dir (next to the data)."""
    import tifffile
    out_dir.mkdir(parents=True, exist_ok=True)
    # The matched plane per cell, kept beside the tiles: the whole-stack match took 79 of 135 s (profile 2026-11-05).
    cache = out_dir / "planes.json"
    planes = json.loads(cache.read_text(encoding="utf-8")) if cache.is_file() else {}
    n = 0
    for exp in experiments:
        rows = []
        for cond, prog in exp["wells"].items():
            for cid, v in show_values(prog, exp["results"]).items():
                if cid in kept.get((exp["name"], cond), set()) and np.isfinite(v[SHOW_METRIC]):
                    rows.append((v[SHOW_METRIC], cond, cid, v))
        rows.sort(key=lambda r: -r[0])
        picks = []
        for value, cond, cid, v in rows[:SHOW_TOP]:
            prog = exp["wells"][cond]
            chans = SHOW_CHANNELS[exp["name"]]
            stacks = {k: tifffile.imread(m._winlong(prog.parent / f"{c}-cell{cid}.tif")) for k, c in chans.items()}
            est = v["actin_bottom_slice_num"] - 1 + v["deepest_invag_z_syn"] / m.XZ_ZSTEP_UM
            est = int(round(est)) if np.isfinite(est) else None   # no synapse plane: search the whole stack
            tile = m.percell_files(prog / "physical_scale_images" / "MRTFA_nuc" / "cells").get(cid)
            if tile is None:
                print(f"WARNING: no MRTFA_nuc tile for {exp['name']} {cond} cell {cid}; left out")
                continue
            key = f"{exp['name']}|{cond}|{cid}"
            stamp = os.path.getmtime(m._winlong(tile))
            # same render as last build: reuse its plane ("z_match" marks the filename-z cache; older entries redo)
            if planes.get(key, {}).get("tile_mtime") == stamp and "z_match" in planes[key]:
                z, zm, score = planes[key]["z"], planes[key]["z_match"], planes[key]["score"]
            else:
                z, z_syn, zm, score = render_plane(stacks["DNA"], tile, est)
                planes[key] = {"z": z, "z_syn": z_syn, "z_match": zm, "score": score, "tile_mtime": stamp}
            print(f"showcase {exp['name']} {cond} cell {cid}: {SHOW_METRIC} {value:.2f}, plane {z + 1} "
                  f"(render z_syn {planes[key]['z_syn']}, DNA match plane {zm + 1}, score {score:.2f})")
            if abs(zm - z) > 1:
                print(f"WARNING: DNA match plane {zm + 1} is {abs(zm - z)} planes from the render's plane {z + 1} "
                      f"for {exp['name']} {cond} cell {cid}; check it")
            if score < 0.6:
                print(f"WARNING: weak DNA match ({score:.2f}) for {exp['name']} {cond} cell {cid}; check its plane")
            stem = f"{exp['name'].replace(',', '').replace(' ', '_')}_{cond}_cell{cid}"
            paths = []
            for (name, _, _), im in zip(SHOW_PANELS, three_color(stacks, z)):
                paths.append(out_dir / f"{stem}_{name.replace('-', '')}.png")
                m.with_scalebar(im, SHOW_UP / SHOW_PSIZE).save(m._winlong(paths[-1]))
            dna = v["nuc_deepest_region_geo025_periph_ratio_025um"]
            picks.append((paths, f"{labels[cond]}, cell {cid}  ·  MRTF-A near deepest invagination "
                                 f"({SHOW_METRIC}): {value:.2f}",
                          f"DNA, same region (nuc_deepest_region_geo025_periph_ratio_025um): {dna:.2f}   "
                          f"MRTF-A / DNA: {value / dna:.2f}   Invag. depth (chull_max_D): {v['chull_max_D']:.2f} μm"))
        for k, (paths, line1, line2) in enumerate(picks):   # one cell a slide: its four panels in a row
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            for (name, _, color), path, box in zip(SHOW_PANELS, paths, SHOW_BOXES):
                top = ns.place(slide, str(path), box, SHOW_IMG_TOP, SHOW_IMG_BOTTOM)
                frame = slide.shapes.add_textbox(Inches(box["left"]), Inches(top - 0.45), Inches(box["width"]),
                                                 Inches(0.42)).text_frame
                para = frame.paragraphs[0]
                para.text, para.alignment = name, PP_ALIGN.CENTER
                para.runs[0].font.size, para.runs[0].font.bold, para.runs[0].font.color.rgb = Pt(20), True, color
            frame = slide.shapes.add_textbox(Inches(0.1), Inches(SHOW_IMG_BOTTOM + 0.02), Inches(13.13),
                                             Inches(1.4)).text_frame
            frame.word_wrap = True
            for j, (text, size, bold) in enumerate(((line1, 18, True), (line2, 14, False))):
                para = frame.paragraphs[0] if j == 0 else frame.add_paragraph()
                para.text, para.alignment = text, PP_ALIGN.CENTER
                para.runs[0].font.size, para.runs[0].font.bold = Pt(size), bold
            ns._add_title(slide, f"Most MRTF-A near the deepest invagination, {short_date(exp['name'])}, "
                          f"{k + 1}/{len(picks)}",
                          f"Top {SHOW_TOP} kept cells, any gel · plane through the deepest invagination · per-cell "
                          "contrast · ratios in the 0.25 μm shell, below 1 = depleted · bars = 5 μm", fit=True)
            n += 1
    cache.write_text(json.dumps(planes, indent=1), encoding="utf-8")
    return n


def geodesic_slides(prs, compile_dir: Path, experiments: list) -> int:
    """The GEODESIC slides; the section lead goes in the notes."""
    dates = [e["name"] for e in experiments]
    n = 0
    for layout, title, rel in GEODESIC:
        paths = [compile_dir / rel] if layout == "grid" else [compile_dir / rel.format(d.replace(" ", "_")) for d in dates]
        missing = [p for p in paths if not p.is_file()]
        if missing:
            print(f"WARNING: missing panel {missing[0]}" + (f" (+{len(missing) - 1})" if len(missing) > 1 else ""))
            continue
        paths = [str(p) for p in paths]
        if layout == "grid":
            plot_slide(prs, title, GEO_SECTION, paths)
        elif layout == "three":
            three_slide(prs, title, GEO_SECTION, paths, dates)
        else:   # each: stacked three to a slide their text came out too small (2026-11-05)
            for path, date in zip(paths, dates):
                plot_slide(prs, f"{title}, {short_date(date)}", GEO_SECTION, [path])
                prs.slides[-1].notes_slide.notes_text_frame.text = GEO_LEAD
                n += 1
            continue
        prs.slides[-1].notes_slide.notes_text_frame.text = GEO_LEAD
        n += 1
    return n


def geom_slides(prs, compile_dir: Path, sections: list) -> int:
    """One slide per curated geom-density figure; the section lead goes in the notes."""
    n = 0
    for section, lead, items in sections:
        for caption, rel in items:
            paths = [compile_dir / r for r in rel]
            missing = [p for p in paths if not p.is_file()]
            if missing:
                print(f"WARNING: missing panel {missing[0]}" + (f" (+{len(missing) - 1})" if len(missing) > 1 else ""))
                continue
            if len(paths) == 1:
                plot_slide(prs, caption, section, [str(paths[0])])
            else:
                three_slide(prs, caption, section, [str(p) for p in paths])
            prs.slides[-1].notes_slide.notes_text_frame.text = lead
            n += 1
    return n


def latest_compile() -> Path:
    """Newest dated compile of the config that has the seg review table."""
    found = sorted(d for d in COMPILE_ROOT.glob(COMPILE_PREFIX + "2*")
                   if re.fullmatch(COMPILE_PREFIX + r"\d{8}", d.name) and (d / "seg_review" / "seg_review_cells.csv").is_file())
    if not found:
        raise SystemExit(f"No {COMPILE_PREFIX}<date> compile with seg_review\\seg_review_cells.csv in {COMPILE_ROOT}")
    return found[-1]


def kept_cells(compile_dir: Path) -> dict:
    """{(experiment name, condition): {cell id}} the compile kept."""
    kept, unread = {}, 0
    with open(compile_dir / "seg_review" / "seg_review_cells.csv", encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            flag = (r["kept_by_compile"] or "").strip().lower()
            if flag in ("", "nan"):
                unread += 1
            elif float(flag) == 1:
                kept.setdefault((r["experiment"], r["condition"]), set()).add(int(float(r["cell_id"])))
    if unread:
        raise SystemExit(f"{unread} rows of seg_review_cells.csv have no kept_by_compile (no compile read); "
                         "rerun the seg review export against the compile")
    return kept


def bottom_slices(prog: Path, results: str) -> dict:
    """{cell id: actin bottom slice} from the well's results .mat."""
    mat = sio.loadmat(m._winlong(prog / results), squeeze_me=True,
                      variable_names=["cell_id", "actin_bottom_slice_num"])
    ids = np.asarray(mat["cell_id"], float).ravel()
    bottom = np.asarray(mat["actin_bottom_slice_num"], float).ravel()
    ok = np.isfinite(ids)  # rows of dropped cells carry a NaN id
    return dict(zip(ids[ok].astype(int).tolist(), bottom[ok].tolist()))


def gel_tiles(exp: dict, cond: str, family, kind: str, keep, n: int) -> list:
    """[(id label, tile)] for the first n cells by id of one gel; keep None = every cell."""
    prog = exp["wells"][cond]
    if kind == "synapse":
        files = m.percell_files(prog / "actin" / "bottom_slice_seg")
    else:
        files = m.percell_files(prog / "physical_scale_images" / family / "cells")
    bottoms = bottom_slices(prog, exp["results"]) if kind == "xz" else {}
    out = []
    for cid in sorted(files):
        if keep is not None and cid not in keep:
            continue
        if kind == "xz":
            if np.isnan(bottoms.get(cid, np.nan)):
                print(f"WARNING: {exp['name']} {cond} cell {cid}: no actin bottom slice; side view left out")
                continue
            t = m.xz_tile(files[cid], bottoms[cid])
        elif kind == "xy":
            # 3-panel tiles are three blobs side by side: no largest-blob crop there.
            split = "_3p" in family
            t = m.xy_tile(files[cid], id_box=SPLIT_ID_BOX if split else ID_BOX, blob_level=0 if split else BLOB_LEVEL)
        else:
            with Image.open(m._winlong(files[cid])) as im:
                t = m.content_crop(im.convert("RGB"))
        if t is None:
            continue
        out.append((str(cid) if SHOW_IDS else "", m.with_scalebar(t, m.XZ_PPU)))
        if len(out) == n:
            break
    if keep is not None and len(out) < min(n, len(keep)):
        print(f"WARNING: {exp['name']} {cond} {family or kind}: {len(out)} tiles for {len(keep)} kept cells")
    return out


def short_date(name: str) -> str:
    """'Nov 3, 2023' -> 'Nov 3 2023' for slide titles."""
    return name.replace(",", "")


def build(cfg: dict, compile_dir, out: Path, experiments: list) -> int:
    m.CONDITIONS = cfg["conditions"]
    labels = dict(cfg["conditions"])
    kept = kept_cells(compile_dir) if compile_dir else None
    for exp in experiments if kept is not None else []:
        for c in exp["wells"]:
            if not kept.get((exp["name"], c)):
                print(f"WARNING: no kept cells for {exp['name']} {c} in seg_review_cells.csv")
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(m.SLIDE_W), Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    n = 0
    for tag, family, kind, title, channels, per_page, footer in MONTAGES:
        pages, titles = [], []
        for exp in experiments:
            by_gel = {labels[c]: gel_tiles(exp, c, family, kind,
                                           None if kept is None else kept.get((exp["name"], c), set()), PER_PAGE)
                      for c in exp["wells"]}
            n_pages = -(-PER_PAGE // per_page)
            for k in range(n_pages):
                pages.append({lab: t[k * per_page:(k + 1) * per_page] for lab, t in by_gel.items()})
                titles.append(f"{title}, {short_date(exp['name'])}" + (f" {k + 1}/{n_pages}" if n_pages > 1 else ""))
        bars = "Bars = 5 μm."
        n += m.build_packed_deck(out, tag, titles, pages, {}, m.XZ_PPU, f"{footer} {bars}".strip(),
                                 channels=channels, slide_bar=False, sink=sink,
                                 per_shelf=1 if per_page == SPLIT_PER_PAGE else PER_SHELF, one_row=True)
    if compile_dir:
        n += showcase_slides(prs, experiments, kept, labels, DATA_ROOT / "deck_tiles" / compile_dir.name / "showcase")
        # The config's report order: the first (nadir) section, the geom-density sections, the rest.
        first = cfg["panels"][0][0]
        geom_done = False
        for section, field, caption in cfg["panels"]:
            if section != first and not geom_done:
                n += geodesic_slides(prs, compile_dir, experiments)
                n += geom_slides(prs, compile_dir, cfg["geom"])
                geom_done = True
            path = compile_dir / "grid_panels" / f"{field}_per_experiment_grid.png"
            if not path.is_file():
                print(f"WARNING: missing panel {path}")
                continue
            plot_slide(prs, caption, section + PENDING.get(field, ""), [str(path)])
            n += 1
    if out.exists():
        backup_presentation(str(out))
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"{n} slides -> {out}")
    return n


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--compile", type=Path, help="compile folder (default: newest with a seg review table)")
    ap.add_argument("--smoke", action="store_true", help="Nov 3 only, every cell, montage slides only")
    ap.add_argument("--out", type=Path, default=OUT, help="deck to write (default: the deck folder)")
    args = ap.parse_args()
    cfg = read_config(CONFIG)
    if args.smoke:
        build(cfg, None, SMOKE_OUT, cfg["experiments"][:1])
    else:
        compile_dir = args.compile or latest_compile()
        print(f"compile: {compile_dir}")
        build(cfg, compile_dir, args.out, cfg["experiments"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
