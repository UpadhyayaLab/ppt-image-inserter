#!/usr/bin/env python
"""
Bare picker tiles for the NFkB / NFAT 3SI CTL supplements, the render_mrtfa_supp_picker_tiles.py look on the
transcription-factor layout (FF 2026-11-05, via "MRTF dataset configs": mirror the MRTF-A supplement picker).

Each cell: TF | DNA | actin (grayscale) | 3-color merge | TF + DNA merge, one plane, one crop, one px/μm for every
tile, a 5 μm bar on the merges only, no text. The renderer is build_mrtfa_gels_deck.three_color with the TF in
MRTF-A's red slot (FF 2026-11-05 "yes red"), DNA cyan, actin gray. Written to
<out>/<sub>/<cond>/cell_<id>_{<TF>,DNA,actin,merge,merge_<TF>_DNA}.png, plus <out>/<sub>/index.csv.

Layout: <root>/<cond>/cropped/channels/C<k>-<NN>_<id>.tif, C1 = TF, C2 = Hoechst, C3 = actin; cell ids 0-based.
psize 0.065 μm, zstep 0.3 μm, z = 1 the coverslip.

The plane is given, never chosen here: cond:id:plane on --cells, or the --nadir CSV's --plane-col (1-based raw
slice; TCell's nadir export, mrtfa_nadir_px_<date>.csv convention). A nadir CSV with a condition column is
matched on (condition, cell); without one it needs exactly one --conds.

Usage (base miniforge python):
    python examples_and_configs/render_tf_picker_tiles.py --tf nfkb --cells G14A2_1010_10min_nfkb_:0:12  # smoke
    python examples_and_configs/render_tf_picker_tiles.py --tf nfkb --list cands.csv --nadir nadir.csv
"""

import argparse
import csv
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_mrtfa_gels_deck as b  # noqa: E402

m = b.m
ROOT = Path("L:/FF/transcription_factors/3siCTLs_NFATNFkB")
TF_NAME = {"nfkb": "NFkB", "nfat": "NFAT"}
CHANNEL_FILES = {"TF": "C1", "DNA": "C2", "Actin": "C3"}
PSIZE = 0.065               # μm per raw xy pixel (FF via "MRTF dataset configs", matches b.SHOW_PSIZE)


def channel_rgb(tf: str) -> list:
    """SHOW_RGB with the TF in MRTF-A's row (red), DNA and actin unchanged."""
    return [(tf if name == "MRTF-A" else name, w, c) for name, w, c in b.SHOW_RGB]


def conditions(root: Path, tf: str, only_1010: bool) -> list:
    """Condition folders for this TF (<...>_<tf>_ or <...>_<tf>_NNN), the 1010 arms only unless told otherwise
    (the papers compare 10 min vs 3 hr on 1010; PLL was dropped from the compile)."""
    pat = re.compile(rf"_{tf}_(\d+)?$", re.IGNORECASE)
    out = sorted(p.name for p in root.iterdir() if p.is_dir() and pat.search(p.name))
    return [c for c in out if "_1010_" in c] if only_1010 else out


def cell_files(cond_dir: Path) -> dict:
    """{cell id: {"TF"|"DNA"|"Actin": path}} from cropped/channels."""
    out = {}
    names = {v: k for k, v in CHANNEL_FILES.items()}
    for p in (cond_dir / "cropped" / "channels").iterdir():
        hit = re.match(r"(C\d)-\d+_(\d+)\.tif$", p.name, re.IGNORECASE)
        if hit and hit.group(1).upper() in names:
            out.setdefault(int(hit.group(2)), {})[names[hit.group(1).upper()]] = p
    return out


def read_nadir(path: Path, plane_col: str) -> dict:
    """{(condition or None, cell id): row} with row["plane"] = int(row[plane_col])."""
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            cid = int(float(r.get("cell_id", r.get("cell"))))
            r["plane"] = int(round(float(r[plane_col])))
            out[(r.get("condition") or None, cid)] = r
    return out


def render(tf: str, cond_dir: Path, files: dict, cid: int, plane: int, out_dir: Path, corner: str) -> dict:
    """One cell's five tiles at `plane` (1-based raw); returns its index row."""
    import tifffile
    stacks = {(tf if k == "TF" else k): tifffile.imread(m._winlong(p)) for k, p in files.items()}
    nz = stacks["DNA"].shape[0]
    if not 1 <= plane <= nz:
        raise ValueError(f"plane {plane} outside the {nz}-plane stack")
    info = {}
    names = [tf, "DNA", "actin", "merge", f"merge_{tf}_DNA"]
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, im in zip(names, b.three_color(stacks, plane - 1, info=info, extra_merges=((tf, "DNA"),),
                                             channel_rgb=channel_rgb(tf))):
        if name.startswith("merge"):
            im = m.with_scalebar(im, b.SHOW_UP / PSIZE, corner=corner)
        im.save(m._winlong(out_dir / f"cell_{cid}_{name}.png"))
    (r0, r1), (c0, c1) = info["box"]
    row = {"condition": cond_dir.name, "cell_id": cid, "plane_z_raw_1based": plane, "n_planes": nz,
           "px_per_um": round(b.SHOW_UP / PSIZE, 3), "scalebar_um": 5, "scalebar_corner": corner,
           "crop_row0": r0, "crop_row1": r1, "crop_col0": c0, "crop_col1": c1}
    for ch, (lo, hi) in info["contrast"].items():
        row[f"{ch}_lo"], row[f"{ch}_hi"] = round(float(lo), 1), round(float(hi), 1)
    row["tf_file"] = files["TF"].name
    return row


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tf", required=True, choices=sorted(TF_NAME))
    ap.add_argument("--root", type=Path, default=ROOT)
    ap.add_argument("--conds", nargs="*", help="condition folders (default: this TF's 1010 arms)")
    ap.add_argument("--all-arms", action="store_true", help="include PLL arms in the default conditions")
    ap.add_argument("--cells", nargs="*", help="cond:id:plane ... (plane 1-based raw)")
    ap.add_argument("--list", type=Path, help="candidate CSV (condition, cell_id[, order]); planes from --nadir")
    ap.add_argument("--nadir", type=Path, help="nadir CSV (cell or cell_id[, condition], --plane-col)")
    ap.add_argument("--plane-col", default="z", help="nadir CSV column holding the 1-based raw plane")
    ap.add_argument("--bar-corner", default="strip", choices=("strip", "upper-right"))
    ap.add_argument("--out", type=Path, help="default <root>/tests/<tf>_picker")
    ap.add_argument("--sub", default="tiles")
    ap.add_argument("--workers", type=int, default=8, help="process pool size, capped at 12; 1 = serial")
    args = ap.parse_args()
    tf = TF_NAME[args.tf]
    out = (args.out or args.root / "tests" / f"{args.tf}_picker") / args.sub
    conds = args.conds or conditions(args.root, args.tf, not args.all_arms)

    extra, skipped = {}, []   # (cond, id) -> candidate columns carried into index.csv; cells with no plane
    if args.cells:
        picks = [(c, int(i), int(p)) for c, i, p in (s.split(":") for s in args.cells)]
    elif args.list:
        if not args.nadir:
            ap.error("--list needs --nadir for the planes")
        nadir = read_nadir(args.nadir, args.plane_col)
        with open(args.list, encoding="utf-8-sig", newline="") as fh:
            cand = list(csv.DictReader(fh))
        cand.sort(key=lambda r: (r["condition"], float(r.get("order", 0) or 0)))
        picks = []
        for r in cand:
            cond, cid = r["condition"], int(float(r["cell_id"]))
            hit = nadir.get((cond, cid))
            if hit is None and len(conds) == 1 and cond == conds[0]:
                hit = nadir.get((None, cid))
            extra[(cond, cid)] = {k: v for k, v in r.items() if k not in ("condition", "cell_id")}
            if hit is None:
                # no plane is chosen here (a fallback plane would be an invention): listed, not rendered
                print(f"WARNING: no nadir plane for {cond} cell {cid}; skipped")
                skipped.append({"condition": cond, "cell_id": cid, "status": "skipped: no nadir plane",
                                **extra[(cond, cid)]})
                continue
            extra[(cond, cid)].update(nadir_x_px=hit["x_px"], nadir_y_px=hit["y_px"])
            picks.append((cond, cid, hit["plane"]))
    else:
        ap.error("give --cells or --list")

    files_by_cond = {c: cell_files(args.root / c) for c in dict.fromkeys(c for c, _, _ in picks)}
    jobs, rows, failed = [], [], 0
    for cond, cid, plane in picks:
        files = files_by_cond[cond].get(cid)
        if files is None or set(files) != set(CHANNEL_FILES):
            print(f"FAILED {cond} cell {cid}: channels found: {sorted(files or {})}")
            failed += 1
            continue
        jobs.append((tf, args.root / cond, files, cid, plane, out / cond, args.bar_corner))
    t0 = time.time()
    if args.workers > 1 and len(jobs) > 1:
        from concurrent.futures import ProcessPoolExecutor
        with ProcessPoolExecutor(min(args.workers, 12, len(jobs))) as pool:
            futs = [(j, pool.submit(render, *j)) for j in jobs]
            results = []
            for j, f in futs:
                try:
                    results.append((j, f.result(), None))
                except Exception as err:   # one bad cell does not stop the rest
                    results.append((j, None, err))
    else:   # serial fallback (--workers 1)
        results = []
        for j in jobs:
            try:
                results.append((j, render(*j), None))
            except Exception as err:
                results.append((j, None, err))
    for j, row, err in results:
        if err is not None:
            print(f"FAILED {j[1].name} cell {j[3]}: {err}")
            failed += 1
        else:
            row["status"] = "rendered"
            row.update(extra.get((row["condition"], row["cell_id"]), {}))
            rows.append(row)
            print(f"{j[1].name} cell {j[3]}: plane {j[4]}")
    print(f"{len(jobs)} cells rendered in {time.time() - t0:.1f} s")
    rows += skipped
    if rows:
        out.mkdir(parents=True, exist_ok=True)
        with open(out / "index.csv", "w", encoding="utf-8", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=list(dict.fromkeys(k for r in rows for k in r)), restval="")
            w.writeheader()
            w.writerows(rows)
    print(f"{len(rows)} cells -> {out}" + (f", {failed} failed" if failed else ""))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
