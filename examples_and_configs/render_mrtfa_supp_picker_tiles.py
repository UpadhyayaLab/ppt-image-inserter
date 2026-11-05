#!/usr/bin/env python
"""
Bare tiles for FF's representative-cell picker, MRTF-A supplement to Fig 6 (FF 2026-11-05: "build me a
representative cell picker"). Feb 5, 2024 only. Per gel, the cells closest to that gel's median of the picker
value (MRTF-A perinuclear near-pocket enrichment, pocketdisk_lt2um perinuc05, from TCell-3D's per-cell CSV):
typical cells, not the top ones.

Each cell: MRTF-A | DNA | actin (grayscale) | merge, at the plane of the pipeline's MRTFA_nuc render (the z in
its filename, synthetic, taken to the raw grid; the DNA template match is a cross-check), one common px/μm for
every tile, a 5 μm bar on the merge only, no text. Written to
<out>/tiles/<cond>/cell_<id>_{MRTFA,DNA,actin,merge}.png, plus <out>/tiles/index.csv.

The renderer, plane match and channel map are build_mrtfa_gels_deck's showcase code.

Usage (base miniforge python):
    python examples_and_configs/render_mrtfa_supp_picker_tiles.py                 # 10 per gel from the CSV
    python examples_and_configs/render_mrtfa_supp_picker_tiles.py --cells Gel1_12kPa:52   # timing smoke
The CSV needs columns condition (config condition name, e.g. Gel1_12kPa, or its label, e.g. 12 kPa),
cell_id and the --value column.
"""

import argparse
import csv
import json
import os
import sys
import time
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_mrtfa_gels_deck as b  # noqa: E402

m = b.m
EXPERIMENT = "Feb 5, 2024"
OUT = b.DATA_ROOT / "tests" / "mrtfa_supp_picker"
PER_GEL = 10
PANELS = ("MRTFA", "DNA", "actin", "merge")   # three_color's order: MRTF-A, DNA, actin, merge
TOP_FRAC = 0.75             # plane above this fraction of the nucleus height (bottom to top slice) is flagged


def pick(csv_path: Path, value_col: str, labels: dict, n: int) -> list:
    """[(cond, cell id, value, gel median, rank)] for the n cells per gel closest to the gel's median."""
    by_cond = {}
    names = {lab: c for c, lab in labels.items()}
    with open(csv_path, encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            v = float(r[value_col]) if r[value_col] not in ("", "NaN", "nan") else np.nan
            if np.isfinite(v):
                cond = names.get(r["condition"], r["condition"])
                by_cond.setdefault(cond, []).append((int(float(r["cell_id"])), v))
    out = []
    for cond in labels:
        rows = by_cond.get(cond, [])
        if not rows:
            print(f"WARNING: no {value_col} values for {cond} in {csv_path}")
            continue
        med = float(np.median([v for _, v in rows]))
        for rank, (cid, v) in enumerate(sorted(rows, key=lambda r: abs(r[1] - med))[:n], 1):
            out.append((cond, cid, v, med, rank))
    return out


def ordered(list_path: Path, labels: dict, todo_only: bool = False) -> list:
    """[(cond, cell id, value, gel median, order)] from a candidate list, in its own order (never re-ranked).
    todo_only keeps the rows whose "rendered" column is False."""
    names = {lab: c for c, lab in labels.items()}
    with open(list_path, encoding="utf-8-sig", newline="") as fh:
        rows = [(names.get(r["condition"], r["condition"]), int(float(r["cell_id"])), np.nan, np.nan,
                 int(float(r["order"]))) for r in csv.DictReader(fh)
                if not todo_only or r.get("rendered", "False").strip().lower() == "false"]
    return sorted(rows, key=lambda r: (list(labels).index(r[0]), r[4]))


def nucleus_planes(exp: dict) -> dict:
    """{(cond, cell id): (bottom, top)} nucleus slice numbers (1-based) from each well's results .mat."""
    import scipy.io as sio
    out = {}
    for cond, prog in exp["wells"].items():
        mat = sio.loadmat(m._winlong(prog / exp["results"]), squeeze_me=True,
                          variable_names=["cell_id", "nuc_bottom_slice_num", "nuc_top_slice_num"])
        for c, lo, hi in zip(*(np.asarray(mat[k], float).ravel()
                               for k in ("cell_id", "nuc_bottom_slice_num", "nuc_top_slice_num"))):
            if np.isfinite(c):
                out[(cond, int(c))] = (lo, hi)
    return out


def render(exp: dict, cond: str, cid: int, tiles: Path, planes: dict, neighbors: bool = False,
           only_moved: bool = False) -> tuple:
    """Write one cell's four bare tiles (and with neighbors, the DNA tile one plane below and above, same crop
    and contrast) at the MRTFA_nuc render's own plane; returns (plane 1-based, z_syn, match plane 1-based, match
    score, previous plane 1-based or None). With only_moved, tiles are written only when the plane differs from
    the one cached by the last run (or a tile is missing)."""
    import tifffile
    prog = exp["wells"][cond]
    stacks = {k: tifffile.imread(m._winlong(prog.parent / f"{c}-cell{cid}.tif"))
              for k, c in b.SHOW_CHANNELS[exp["name"]].items()}
    tile = m.percell_files(prog / "physical_scale_images" / "MRTFA_nuc" / "cells").get(cid)
    if tile is None:
        raise FileNotFoundError(f"no MRTFA_nuc render for {cond} cell {cid}")
    key = f"{exp['name']}|{cond}|{cid}"
    stamp = os.path.getmtime(m._winlong(tile))
    prev = planes.get(key, {}).get("z")
    # "z_match" marks the filename-z cache; entries from the template-match-only runs are redone
    if planes.get(key, {}).get("tile_mtime") == stamp and "z_match" in planes[key]:
        z, z_syn, zm, score = (planes[key][k] for k in ("z", "z_syn", "z_match", "score"))
    else:
        z, z_syn, zm, score = b.render_plane(stacks["DNA"], tile, None)
        planes[key] = {"z": z, "z_syn": z_syn, "z_match": zm, "score": score, "tile_mtime": stamp}
    if abs(zm - z) > 1:
        print(f"WARNING: DNA match plane {zm + 1} is {abs(zm - z)} planes from the render's plane {z + 1} "
              f"for {cond} cell {cid}; check it")
    if score < 0.6:
        print(f"WARNING: weak DNA match ({score:.2f}) for {cond} cell {cid}; check its plane")
    d = tiles / cond
    d.mkdir(parents=True, exist_ok=True)
    extra = [(z - 1, "DNA_zm1"), (z + 1, "DNA_zp1")] if neighbors else []
    extra = [(ze, n) for ze, n in extra if 0 <= ze < stacks["DNA"].shape[0]]
    names = list(PANELS) + [n for _, n in extra]
    missing = any(not os.path.isfile(m._winlong(d / f"cell_{cid}_{n}.png")) for n in names)
    if not only_moved or prev != z or missing:
        for name, im in zip(names, b.three_color(stacks, z, tuple(ze for ze, _ in extra))):
            if name == "merge":
                im = m.with_scalebar(im, b.SHOW_UP / b.SHOW_PSIZE)
            im.save(m._winlong(d / f"cell_{cid}_{name}.png"))
    return z + 1, z_syn, zm + 1, score, None if prev is None else prev + 1


def figure_tiles(exp: dict, cond: str, cid: int, plane: int, out_dir: Path) -> Path:
    """FF's chosen cell at his plane (1-based raw, z = 1 the coverslip): the picker's four bare tiles, 5 μm bar on
    the merge only, into out_dir, plus tile_info.csv (plane, px/μm, crop box, contrast). Returns the CSV path."""
    import tifffile
    prog = exp["wells"][cond]
    stacks = {k: tifffile.imread(m._winlong(prog.parent / f"{c}-cell{cid}.tif"))
              for k, c in b.SHOW_CHANNELS[exp["name"]].items()}
    if not 1 <= plane <= stacks["DNA"].shape[0]:
        raise ValueError(f"plane {plane} outside the {stacks['DNA'].shape[0]}-plane stack of {cond} cell {cid}")
    out_dir.mkdir(parents=True, exist_ok=True)
    info = {}
    paths = []
    # FF 2026-11-05 "no need for actin for the figure": also a MRTF-A + DNA merge, same crop and contrast
    names = list(PANELS) + ["merge_MRTFA_DNA"]
    for name, im in zip(names, b.three_color(stacks, plane - 1, info=info, extra_merges=(("MRTF-A", "DNA"),))):
        if name.startswith("merge"):
            im = m.with_scalebar(im, b.SHOW_UP / b.SHOW_PSIZE)
        paths.append(out_dir / f"cell_{cid}_z{plane}_{name}.png")
        im.save(m._winlong(paths[-1]))
    (r0, r1), (c0, c1) = info["box"]
    row = {"experiment": exp["name"], "condition": cond, "cell_id": cid, "plane_z_raw_1based": plane,
           "z_syn": round(b.raw_to_syn(plane), 2), "px_per_um": round(b.SHOW_UP / b.SHOW_PSIZE, 3),
           "raw_um_per_px": b.SHOW_PSIZE, "upsample": b.SHOW_UP, "scalebar_um": 5,
           "crop_row0": r0, "crop_row1": r1, "crop_col0": c0, "crop_col1": c1,
           "tile_w_px": (c1 - c0) * b.SHOW_UP, "tile_h_px": (r1 - r0) * b.SHOW_UP}
    for ch, (lo, hi) in info["contrast"].items():
        row[f"{ch}_lo"], row[f"{ch}_hi"] = round(float(lo), 1), round(float(hi), 1)
    # one row per merge recipe: the four picker tiles (merge with actin), then the MRTF-A + DNA merge
    rows = [dict(row, merge_channels="MRTF-A+DNA+Actin", files=";".join(p.name for p in paths[:4])),
            dict(row, merge_channels="MRTF-A+DNA", files=paths[4].name)]
    csv_path = out_dir / "tile_info.csv"
    with open(csv_path, "w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    return csv_path


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--csv", type=Path, default=OUT / "percell_Feb_5_2024.csv")
    ap.add_argument("--value", default="MRTFA_pocketdisk_lt2um_perinuc05", help="picker value column of the CSV")
    ap.add_argument("--n", type=int, default=PER_GEL)
    ap.add_argument("--cells", nargs="*", help="cond:id ... instead of the CSV pick (timing smoke)")
    ap.add_argument("--list", type=Path, help="candidate list (condition, cell_id, order) rendered in its order "
                                              "instead of the median pick")
    ap.add_argument("--sub", default="tiles", help="tile folder under --out")
    ap.add_argument("--neighbors", action="store_true", help="also the DNA tile one plane below and above")
    ap.add_argument("--figure", help="cond:id:plane (1-based raw) -> <out>/figure_cell<id>_z<plane>/, the chosen "
                                     "figure cell's tiles and tile_info.csv")
    ap.add_argument("--append", action="store_true", help="add the rows to an existing index.csv")
    ap.add_argument("--todo-only", action="store_true", help="with --list: only rows whose 'rendered' is False")
    ap.add_argument("--top-frac", type=float, default=TOP_FRAC, help="flag planes above this fraction of the "
                                                                     "nucleus height")
    ap.add_argument("--only-moved", action="store_true", help="re-render only cells whose plane changed since "
                                                              "the last run (index and planes are redone for all)")
    ap.add_argument("--out", type=Path, default=OUT)
    args = ap.parse_args()
    cfg = b.read_config(b.CONFIG)
    exp = next(e for e in cfg["experiments"] if e["name"] == EXPERIMENT)
    labels = {c: dict(cfg["conditions"])[c] for c in exp["wells"]}
    if args.figure:
        cond, cid, plane = args.figure.split(":")
        out = figure_tiles(exp, cond, int(cid), int(plane), args.out / f"figure_cell{cid}_z{plane}")
        print(f"figure tiles -> {out.parent}")
        return 0
    if args.cells:
        picks = [(c.split(":")[0], int(c.split(":")[1]), np.nan, np.nan, 0) for c in args.cells]
    elif args.list:
        picks = ordered(args.list, labels, args.todo_only)
    else:
        picks = pick(args.csv, args.value, labels, args.n)
    nuc = nucleus_planes(exp)
    tiles = args.out / args.sub
    tiles.mkdir(parents=True, exist_ok=True)
    cache = tiles / "planes.json"
    planes = json.loads(cache.read_text(encoding="utf-8")) if cache.is_file() else {}
    rows, moved, failed = [], [], 0
    for cond, cid, v, med, rank in picks:
        t0 = time.time()
        try:
            z, z_syn, zm, score, prev = render(exp, cond, cid, tiles, planes, args.neighbors, args.only_moved)
        except Exception as err:   # one bad cell does not stop the rest
            print(f"FAILED {cond} cell {cid}: {err}")
            failed += 1
            continue
        if prev is not None and prev != z:
            moved.append(f"{labels[cond]} cell {cid}: plane {prev} -> {z}")
        lo, hi = nuc.get((cond, cid), (np.nan, np.nan))
        # nuc_*_slice_num are synthetic (isotropic) indices, like the render's z_syn
        zs = z_syn if z_syn is not None else b.raw_to_syn(z)
        frac = (zs - lo) / (hi - lo) if hi > lo else np.nan   # 0 = nucleus bottom, 1 = its top
        flag = "near top of nucleus" if frac > args.top_frac else ""
        print(f"{cond} cell {cid}: plane {z} (z_syn {z_syn}, {frac:.2f} of the nucleus height), DNA match plane "
              f"{zm} ({score:.2f}), {time.time() - t0:.1f} s" + (f"  FLAG: {flag}" if flag else ""))
        row = {"condition": labels[cond], "cell_id": cid, "plane_z": z, "z_syn": z_syn}
        row.update({"order": rank, "list": args.list.stem} if args.list else
                   {"typicality_rank": rank, "value": v, "condition_median": med})
        row.update({"match_plane_z": zm, "match_score": round(score, 3), "nuc_bottom_slice": lo,
                    "nuc_top_slice": hi, "plane_frac_of_nucleus": round(frac, 2), "flag": flag})
        rows.append(row)
    cache.write_text(json.dumps(planes, indent=1), encoding="utf-8")
    if not args.cells:
        index = tiles / "index.csv"
        if args.append and index.is_file():   # keep earlier rows; a re-rendered cell replaces its old row
            with open(index, encoding="utf-8", newline="") as fh:
                done = {(r["condition"], str(r["cell_id"])) for r in rows}
                rows = [r for r in csv.DictReader(fh) if (r["condition"], r["cell_id"]) not in done] + rows
        fields = list(dict.fromkeys(k for r in rows for k in r)) or ["condition"]   # union: older rows lack new columns
        with open(index, "w", encoding="utf-8", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=fields, restval="")
            w.writeheader()
            w.writerows(rows)
    print(f"{len(moved)} cell(s) moved from the last run's plane" + "".join(f"\n  {s}" for s in moved))
    print(f"{len(rows)} cells -> {tiles}" + (f", {failed} failed" if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
