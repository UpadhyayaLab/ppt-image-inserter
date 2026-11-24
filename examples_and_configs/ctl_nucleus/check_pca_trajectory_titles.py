#!/usr/bin/env python
"""Check each slide of the naive PCA trajectory deck: its title must name the experiment of BOTH its panels.

For every slide, the mmdd in each panel's filename (row_pca_trajA_<mmdd>.png, row_pca_traj3dA_<mmdd>_v1.png, read from
the picture metadata the builder stores) is looked up in chromatin-analysis scripts/datasets.py (naive datasets only);
the title must equal "<date>, <marker> experiment" of that dataset, and both panels must carry the same mmdd. Also
checks one slide per dataset, no slide twice. A first slide without panels is the title slide and is skipped.

Usage:
    python examples_and_configs/ctl_nucleus/check_pca_trajectory_titles.py <deck config.yaml> [--datasets <datasets.py>]
Exit code 1 on any mismatch.
"""
import argparse
import importlib.util
import os
import re
import sys

import yaml
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from ppt_image_inserter import extract_image_metadata  # noqa: E402

DATASETS_PY = "C:/Users/Upadhyaya_Lab/Code/chromatin-analysis/scripts/datasets.py"
PANEL_RE = re.compile(r"^row_pca_traj(?:3d)?A_(\d{4})(?:_v\d+)?\.png$")


def load_naive(path):
    """{mmdd: (date, marker)} for every naive dataset in datasets.py."""
    spec = importlib.util.spec_from_file_location("datasets", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    out = {}
    for d in mod.DATASETS.values():
        if d["cell_state"] != "naive":
            continue
        mmdd = d["date"][5:7] + d["date"][8:10]
        if mmdd in out:
            raise SystemExit(f"two naive datasets share mmdd {mmdd}")
        out[mmdd] = (d["date"], d["marker"])
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("config")
    ap.add_argument("--datasets", default=DATASETS_PY)
    args = ap.parse_args()
    cfg = yaml.safe_load(open(args.config, encoding="utf-8"))
    deck = cfg.get("output_path") or cfg["presentation"]
    naive = load_naive(args.datasets)

    pics = {}
    for m in extract_image_metadata(deck):
        pics.setdefault(m["slide_index"], []).append(os.path.basename(m["original_path"] or m["filename"] or ""))

    prs = Presentation(deck)
    fails, seen = [], []
    for i, slide in enumerate(prs.slides):
        title = slide.shapes.title.text_frame.text.strip() if slide.shapes.title is not None else ""
        if not title:   # builder-added title box: the first text box at the top
            boxes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
            title = min(boxes, key=lambda s: s.top).text_frame.text.strip() if boxes else ""
        files = pics.get(i, [])
        if i == 0 and not files:   # title slide
            print(f"slide 1: title slide  {title!r}")
            continue
        mmdds = []
        for f in files:
            m = PANEL_RE.match(f)
            mmdds.append(m.group(1) if m else None)
        problem = None
        if len(files) != 2:
            problem = f"{len(files)} panels, expected 2"
        elif None in mmdds:
            problem = f"panel name not recognized: {files}"
        elif mmdds[0] != mmdds[1]:
            problem = f"panels from different experiments: {files}"
        elif mmdds[0] not in naive:
            problem = f"mmdd {mmdds[0]} is not a naive dataset in datasets.py"
        else:
            date, marker = naive[mmdds[0]]
            expected = f"{date}, {marker} experiment"   # FF 2026-11-24: "2024-01-29, H3K27me3 experiment and such"
            if title != expected:
                problem = f"title {title!r} != expected {expected!r}"
            seen.append(mmdds[0])
        status = "FAIL" if problem else "OK"
        print(f"slide {i + 1}: {status}  title={title!r}  panels={files}" + (f"  -> {problem}" if problem else ""))
        if problem:
            fails.append(i + 1)

    dupes = sorted({x for x in seen if seen.count(x) > 1})
    missing = sorted(set(naive) - set(seen))
    if dupes:
        print(f"FAIL: experiment on more than one slide: {dupes}")
    if missing:
        print(f"NOTE: naive datasets with no slide: {missing}")
    ok = not fails and not dupes
    print(f"Title check: {len(seen) - len(dupes)}/{len(prs.slides) - 1} experiment slides OK, {len(fails)} failed" + ("" if ok else ", FAILED"))
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
