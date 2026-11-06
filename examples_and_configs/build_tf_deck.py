#!/usr/bin/env python
"""
NFkB / NFAT decks for 3SI CTLs on 1010, 10 min vs 3 hr (FF 2026-11-05, via "MRTF dataset configs": built the same
way as the MRTF-A deck, build_mrtfa_gels_deck.py). One experiment, two conditions side by side, never pooled:
the MRTF-A montage families (the first 12 kept cells by id per condition), the showcase (the kept cells with the
most TF near the deepest invagination, 3-color), the geodesic figures, the geom-density sections, then the MRTF-A
deck's results panels in its order with the TF in MRTF-A's place. Each MRTF-A per-experiment grid becomes the
compile's single 10 min vs 3 hr comparison plot (grid_panels/across_time).

Conditions, wells and results file come from the TF's compilation config; the kept cells from the compile's
cell_tracing.csv (the cells the compile kept, by condition and cell id). No seg review table for these compiles.
Left out, as for MRTF-A: the nuc/cyto MFI ratio (cytoplasm MFI is NaN, no 3D cell mask), and MRTF-A / DNA (no TF
counterpart in the compile). No centrosome stain.

Usage (base miniforge python, from PowerShell):
    python examples_and_configs/build_tf_deck.py --tf nfkb [--compile <dir>]
    python examples_and_configs/build_tf_deck.py --tf nfat --smoke
--smoke: montage slides only, into the data folder's tests\\ subfolder.
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
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_mrtfa_gels_deck as b  # noqa: E402
import render_tf_picker_tiles as rt  # noqa: E402

m, ns = b.m, b.ns
plot_slide = b.plot_slide
CONFIG_DIR = Path("C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/config/compilation/fixed/CTL/"
                  "transcription_factors")
COMPILE_ROOT = Path("L:/FF/transcription_factors/compiled_results")
DATA_ROOT = rt.ROOT
OUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/Transcription_factors")
TF = {"nfkb": ("NFkB", "NFκB"), "nfat": ("NFAT", "NFAT")}   # (field prefix, name on the slides)


def read_config(path: Path) -> dict:
    """Wells, conditions and results file from the TF's compilation config."""
    text = path.read_text(encoding="utf-8")

    def value(field: str) -> str:
        return re.search(rf"config\.{field}\s*=\s*'([^']+)';", text).group(1)

    def cell_list(field: str) -> list:
        return re.findall(r"'([^']+)'", re.search(rf"config\.{field}\s*=\s*\{{([^}}]*)\}};", text).group(1))

    # folders: one per line, commented-out arms (PLL) skipped
    block = re.search(r"config\.conditions\.folders\s*=\s*\{(.*?)\};", text, re.DOTALL).group(1)
    folders = [re.search(r"'([^']+)'", ln).group(1) for ln in block.splitlines()
               if re.match(r"\s*'", ln)]
    names, labels = cell_list("conditions.names"), cell_list("conditions.labels")
    base, sub = value("base_path"), value("subpath")
    return {"results": value("results_filename"), "conditions": list(zip(names, labels)),
            "wells": {n: Path(base) / f / sub for n, f in zip(names, folders)}}


def latest_compile(prefix: str) -> Path:
    """Newest dated compile of the TF that has its cell tracing table."""
    found = sorted(d for d in COMPILE_ROOT.glob(prefix + "2*")
                   if re.fullmatch(prefix + r"\d{8}", d.name) and (d / "cell_tracing.csv").is_file())
    if not found:
        raise SystemExit(f"No {prefix}<date> compile with cell_tracing.csv in {COMPILE_ROOT}")
    return found[-1]


def kept_cells(compile_dir: Path) -> dict:
    """{condition: {cell id}} the compile kept (cell_tracing.csv: one row per kept cell)."""
    kept = {}
    with open(compile_dir / "cell_tracing.csv", encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            if int(float(r["cell_found"])) == 1:
                kept.setdefault(r["condition_name"], set()).add(int(float(r["local_cell_idx"])))
    return kept


def panels(p: str, t: str) -> list:
    """The MRTF-A deck's results panels with the TF in MRTF-A's place; the two with no counterpart dropped."""
    drop = {"MRTFA_nuc_cyto_ratio", "MRTFA_over_DNA_deepest_region_geo025_periph_ratio_025um"}
    return [(s.replace("MRTF-A", t), f.replace("MRTFA", p), c.replace("MRTF-A", t))
            for s, f, c in b.read_config(b.CONFIG)["panels"] if f not in drop]


GEOM_LEADS = [
    "Density on each surface face relative to that cell's own mean in the shell; Mean±SEM across cells, 10 min and "
    "3 hr overlaid. Outer 0.25 μm = inside the nucleus, Perinuclear = 0.5 μm outside ({t} only; DNA is inside only).",
    "One dot per cell, 10 min and 3 hr in one figure (never pooled). Enrichment = density on the faces past the "
    "threshold / the cell's mean in the shell; 1 = no enrichment. n counts cells with enough surface past the "
    "threshold. Outer 0.25 μm = inside the nucleus, Perinuclear = 0.5 μm outside.",
    "One dot per cell, 10 min and 3 hr in one figure. Area-weighted Pearson r between face density and the geometry "
    "axis; 0 = no relation. Outer 0.25 μm = inside the nucleus, Perinuclear = 0.5 μm outside.",
]


def geom_sections(p: str, t: str) -> list:
    """The MRTF-A deck's geom-density sections on this compile: each BYEXP profile -> the 10 min vs 3 hr overlay,
    each set of per-experiment enrichment figures -> the one figure holding both conditions."""
    out = []
    for q, (title, _, items) in enumerate(b.read_config(b.CONFIG)["geom"]):
        rows = []
        for caption, rel in items:
            r = rel[0].replace("MRTFA", p).replace("_BYEXP_line.png", "_OVERLAY_line_all.png")
            r = re.sub(r"_shells_[^/]*\.png$", "_shells.png", r)
            rows.append((caption.replace("MRTF-A", t), r))
        out.append((title.replace(", by experiment", ", 10 min vs 3 hr"), GEOM_LEADS[q].format(t=t), rows))
    return out


def geodesic(p: str, t: str) -> list:
    """(title, path under the compile) of the geodesic figures, the MRTF-A deck's GEODESIC on this compile, led by
    the geo025 region violins figure-reviewer passed for the report (2026-11-05)."""
    g, pv, pr, gp = ("geom_density/geo025_region_violins/", "geom_density/pocket_violins/",
                     "geom_density/profiles/pocket/", "grid_panels/across_time/")
    return [
        (f"{t} in the deepest invagination's region, inside, outer 0.25 μm",
         g + f"{p}_geomdens_geo025region_inperinuc025_violin.png"),
        (f"{t} in the deepest invagination's region, outside, 0.5 μm perinuclear",
         g + f"{p}_geomdens_geo025region_perinuc05_violin.png"),
        ("DNA in the deepest invagination's region, inside, outer 0.25 μm",
         g + "DNA_geomdens_geo025region_inperinuc025_violin.png"),
        (f"{t} vs geodesic distance from the pocket center", pr + f"{p}_geomdens_pocketdist_shells_OVERLAY_line.png"),
        ("DNA vs geodesic distance from the pocket center", pr + "DNA_geomdens_pocketdist_shells_OVERLAY_line.png"),
        (f"{t} near the pocket center, inside, outer 0.25 μm", pv + f"{p}_geomdens_pocketdisk_inperinuc025_violin.png"),
        (f"{t} near the pocket center, outside, 0.5 μm perinuclear", pv + f"{p}_geomdens_pocketdisk_perinuc05_violin.png"),
        ("DNA near the pocket center, inside, outer 0.25 μm", pv + "DNA_geomdens_pocketdisk_inperinuc025_violin.png"),
        (f"{t} pocket profile fits, inside, outer 0.25 μm", pr + f"{p}_geomdens_pocketdist_fits_inperinuc025_violin.png"),
        (f"{t} pocket profile fits, outside, 0.5 μm perinuclear", pr + f"{p}_geomdens_pocketdist_fits_perinuc05_violin.png"),
        ("DNA pocket profile fits, inside, outer 0.25 μm", pr + "DNA_geomdens_pocketdist_fits_inperinuc025_violin.png"),
        (f"{t} near the deepest invagination, 0.5 μm shell", gp + f"{p}_deepest_region_geo025_periph_ratio_05um.png"),
        ("DNA near the deepest invagination, 0.5 μm shell", gp + "nuc_deepest_region_geo025_periph_ratio_05um.png"),
        (f"{t} near the deepest invagination, 1 μm shell", gp + f"{p}_deepest_region_geo025_periph_ratio_1um.png"),
        ("DNA near the deepest invagination, 1 μm shell", gp + "nuc_deepest_region_geo025_periph_ratio_1um.png"),
        (f"{t} intensity slope at the deepest invagination", gp + f"{p}_deepest_invag_geo025_gradient_slope.png"),
        ("DNA intensity slope at the deepest invagination", gp + "nuc_deepest_invag_geo025_gradient_slope.png"),
        (f"{t} slope at the deepest invagination / elsewhere", gp + f"{p}_deepest_invag_geo025_gradient_ratio.png"),
        ("DNA slope at the deepest invagination / elsewhere", gp + "nuc_deepest_invag_geo025_gradient_ratio.png"),
        ("Radius of the deepest invagination's geodesic region", gp + "deepest_region_geo025_radius_um.png"),
    ]


def figure_slides(prs, compile_dir: Path, items: list, section: str, lead: str) -> int:
    """One slide per (title, path) present; the section lead goes in the notes."""
    n = 0
    for title, rel in items:
        path = compile_dir / rel
        if not path.is_file():
            print(f"WARNING: missing panel {path}")
            continue
        plot_slide(prs, title, section, [str(path)])
        prs.slides[-1].notes_slide.notes_text_frame.text = lead
        n += 1
    return n


def showcase_slides(prs, cfg: dict, kept: dict, p: str, t: str, out_dir: Path) -> int:
    """b.SHOW_TOP kept cells per condition with the most TF near the deepest invagination, one cell a slide, the
    MRTF-A showcase's four panels (TF, DNA, actin, merge) at the plane of the TF_nuc render."""
    import tifffile
    metric = f"{p}_deepest_region_geo025_periph_ratio_025um"
    fields = ["cell_id", metric, "nuc_deepest_region_geo025_periph_ratio_025um", "chull_max_D",
              "actin_bottom_slice_num", "deepest_invag_z_syn"]
    spec = rt.channel_rgb(p)
    names = [t if n == p else n for n, _, _ in spec] + ["Merge"]
    colors = [c for _, _, c in spec] + [b.SHOW_PANELS[-1][2]]
    out_dir.mkdir(parents=True, exist_ok=True)
    cache = out_dir / "planes.json"
    planes = json.loads(cache.read_text(encoding="utf-8")) if cache.is_file() else {}
    labels = dict(cfg["conditions"])
    n = 0
    for cond, prog in cfg["wells"].items():
        mat = sio.loadmat(m._winlong(prog / cfg["results"]), squeeze_me=True, variable_names=fields)
        cols = {f: np.asarray(mat[f], float).ravel() for f in fields}
        rows = sorted(((cols[metric][k], int(c), {f: cols[f][k] for f in fields})
                       for k, c in enumerate(cols["cell_id"])
                       if np.isfinite(c) and int(c) in kept.get(cond, set()) and np.isfinite(cols[metric][k])),
                      key=lambda r: -r[0])[:b.SHOW_TOP]
        files = rt.cell_files(prog.parents[2])
        for k, (value, cid, v) in enumerate(rows):
            stacks = {(p if ch == "TF" else ch): tifffile.imread(m._winlong(f)) for ch, f in files[cid].items()}
            est = v["actin_bottom_slice_num"] - 1 + v["deepest_invag_z_syn"] / m.XZ_ZSTEP_UM
            est = int(round(est)) if np.isfinite(est) else None
            tile = m.percell_files(prog / "physical_scale_images" / f"{p}_nuc" / "cells").get(cid)
            if tile is None:
                print(f"WARNING: no {p}_nuc tile for {cond} cell {cid}; left out")
                continue
            key, stamp = f"{cond}|{cid}", os.path.getmtime(m._winlong(tile))
            if planes.get(key, {}).get("tile_mtime") == stamp:
                z, zm, score = planes[key]["z"], planes[key]["z_match"], planes[key]["score"]
            else:
                z, z_syn, zm, score = b.render_plane(stacks["DNA"], tile, est)
                planes[key] = {"z": z, "z_syn": z_syn, "z_match": zm, "score": score, "tile_mtime": stamp}
            print(f"showcase {cond} cell {cid}: {metric} {value:.2f}, plane {z + 1} (render z_syn "
                  f"{planes[key]['z_syn']}, DNA match plane {zm + 1}, score {score:.2f})")
            if abs(zm - z) > 1:
                print(f"WARNING: DNA match plane {zm + 1} is {abs(zm - z)} planes from the render's plane {z + 1} "
                      f"for {cond} cell {cid}; check it")
            paths = []
            for name, im in zip(names, b.three_color(stacks, z, channel_rgb=spec)):
                paths.append(out_dir / f"{cond}_cell{cid}_{name.replace('κ', 'k')}.png")
                m.with_scalebar(im, b.SHOW_UP / b.SHOW_PSIZE).save(m._winlong(paths[-1]))
            dna = v["nuc_deepest_region_geo025_periph_ratio_025um"]
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            for name, color, path, box in zip(names, colors, paths, b.SHOW_BOXES):
                top = ns.place(slide, str(path), box, b.SHOW_IMG_TOP, b.SHOW_IMG_BOTTOM)
                frame = slide.shapes.add_textbox(Inches(box["left"]), Inches(top - 0.45), Inches(box["width"]),
                                                 Inches(0.42)).text_frame
                para = frame.paragraphs[0]
                para.text, para.alignment = name, PP_ALIGN.CENTER
                para.runs[0].font.size, para.runs[0].font.bold, para.runs[0].font.color.rgb = Pt(20), True, color
            frame = slide.shapes.add_textbox(Inches(0.1), Inches(b.SHOW_IMG_BOTTOM + 0.02), Inches(13.13),
                                             Inches(1.4)).text_frame
            frame.word_wrap = True
            lines = ((f"{labels[cond]}, cell {cid}  ·  {t} near deepest invagination ({metric}): {value:.2f}", 18, True),
                     (f"DNA, same region (nuc_deepest_region_geo025_periph_ratio_025um): {dna:.2f}   {t} / DNA: "
                      f"{value / dna:.2f}   Invag. depth (chull_max_D): {v['chull_max_D']:.2f} μm", 14, False))
            for j, (text, size, bold) in enumerate(lines):
                para = frame.paragraphs[0] if j == 0 else frame.add_paragraph()
                para.text, para.alignment = text, PP_ALIGN.CENTER
                para.runs[0].font.size, para.runs[0].font.bold = Pt(size), bold
            ns._add_title(slide, f"Most {t} near the deepest invagination, {labels[cond]}, {k + 1}/{len(rows)}",
                          f"Top {b.SHOW_TOP} kept cells per condition · plane through the deepest invagination · "
                          "per-cell contrast · ratios in the 0.25 μm shell, below 1 = depleted · bars = 5 μm", fit=True)
            n += 1
    cache.write_text(json.dumps(planes, indent=1), encoding="utf-8")
    return n


def build(tf: str, cfg: dict, compile_dir, out: Path) -> int:
    p, t = TF[tf]
    m.CONDITIONS = cfg["conditions"]
    labels = dict(cfg["conditions"])
    kept = kept_cells(compile_dir) if compile_dir else None
    for c in cfg["wells"] if kept is not None else []:
        if not kept.get(c):
            print(f"WARNING: no kept cells for {c} in cell_tracing.csv")
    exp = {"name": t, "wells": cfg["wells"], "results": cfg["results"]}
    # The MRTF-A deck's MONTAGES with the TF's render families and the TF's name in MRTF-A's legend slot.
    montages = [(tag, fam.replace("MRTFA", p) if fam else None, kind, title.replace("MRTF-A", t),
                 [(t if name == "MRTF-A" else name, color) for name, color in chans], per_page,
                 footer.replace("MRTF-A", t))
                for tag, fam, kind, title, chans, per_page, footer in b.MONTAGES]
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(m.SLIDE_W), Inches(m.SLIDE_H)
    sink = {"prs": prs, "tags": []}
    n = 0
    for tag, family, kind, title, channels, per_page, footer in montages:
        by_cond = {labels[c]: b.gel_tiles(exp, c, family, kind, None if kept is None else kept.get(c, set()),
                                          b.PER_PAGE) for c in cfg["wells"]}
        n_pages = -(-b.PER_PAGE // per_page)
        pages = [{lab: tl[k * per_page:(k + 1) * per_page] for lab, tl in by_cond.items()} for k in range(n_pages)]
        titles = [title + (f" {k + 1}/{n_pages}" if n_pages > 1 else "") for k in range(n_pages)]
        n += m.build_packed_deck(out, tag, titles, pages, {}, m.XZ_PPU, f"{footer} Bars = 5 μm.".strip(),
                                 channels=channels, slide_bar=False, sink=sink,
                                 per_shelf=1 if per_page == b.SPLIT_PER_PAGE else b.PER_SHELF, one_row=True)
    if compile_dir:
        n += showcase_slides(prs, cfg, kept, p, t, DATA_ROOT / "deck_tiles" / compile_dir.name / "showcase")
        geo_section = f"{t} and DNA around the deepest invagination, geodesic"
        geo_lead = b.GEO_LEAD.replace("MRTF-A", t).replace("One figure or panel per experiment, never pooled.",
                                                            "10 min and 3 hr in one figure, never pooled.")
        results = panels(p, t)
        first = results[0][0]
        geom_done = False
        for section, field, caption in results:
            if section != first and not geom_done:
                n += figure_slides(prs, compile_dir, geodesic(p, t), geo_section, geo_lead)
                for title, lead, items in geom_sections(p, t):
                    n += figure_slides(prs, compile_dir, items, title, lead)
                geom_done = True
            pending = b.PENDING.get(field.replace(p, "MRTFA", 1), "")
            n += figure_slides(prs, compile_dir, [(caption, f"grid_panels/across_time/{field}.png")],
                               section + pending, "")
    if out.exists():
        b.backup_presentation(str(out))
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    print(f"{n} slides -> {out}")
    return n


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tf", required=True, choices=sorted(TF))
    ap.add_argument("--compile", type=Path, help="compile folder (default: newest with cell_tracing.csv)")
    ap.add_argument("--smoke", action="store_true", help="montage slides only, every cell, into tests\\")
    ap.add_argument("--out", type=Path, help="deck to write (default: the deck folder)")
    args = ap.parse_args()
    p, t = TF[args.tf]
    cfg = read_config(CONFIG_DIR / f"CTL_3siCTLs_{p}_compilation.m")
    if args.smoke:
        build(args.tf, cfg, None, args.out or DATA_ROOT / "tests" / f"{args.tf}_deck_smoke" / f"{args.tf}_deck_smoke.pptx")
        return 0
    compile_dir = args.compile or latest_compile(f"CTL_3siCTLs_{p}_")
    print(f"compile: {compile_dir}")
    out = args.out or OUT_DIR / f"{p} in 3SI CTLs, 10 min vs 3 hr on 1010 ({compile_dir.name[-8:]}).pptx"
    build(args.tf, cfg, compile_dir, out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
