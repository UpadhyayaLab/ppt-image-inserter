"""
check_scalebar_pixel_widths.py

One-shot diagnostic. Opens a montage PNG, finds every horizontal run of
near-white pixels that looks like a scalebar (width 50-300 px, thickness
2-8 px on near-black background), and reports the pixel widths grouped
by tile-row position. If Stage AF is honest, every scalebar should be
the same width (150 px for PPUM=30 px/μm × 5 μm).

Usage:
    python examples_and_configs/check_scalebar_pixel_widths.py <montage.png>
"""

import sys
from collections import Counter

import numpy as np
from PIL import Image

WHITE_THRESHOLD = 220        # pixel >= this is "white"
BAR_MIN_WIDTH = 50           # ignore short white runs (could be text, axis ticks)
BAR_MAX_WIDTH = 400          # ignore very long white runs (could be empty white panels)


def find_scalebars(arr: np.ndarray) -> list:
    """Return list of (row, col_start, col_end) for each candidate scalebar.

    Strategy: scan each row; find contiguous runs of pixels >= WHITE_THRESHOLD;
    keep runs whose length is in [BAR_MIN_WIDTH, BAR_MAX_WIDTH].
    """
    H, W = arr.shape
    bars = []
    is_white = arr >= WHITE_THRESHOLD
    for y in range(H):
        row = is_white[y]
        # find runs
        d = np.diff(row.astype(np.int8))
        starts = np.flatnonzero(d == 1) + 1
        ends   = np.flatnonzero(d == -1) + 1
        if row[0]:
            starts = np.insert(starts, 0, 0)
        if row[-1]:
            ends = np.append(ends, W)
        for s, e in zip(starts, ends):
            w = e - s
            if BAR_MIN_WIDTH <= w <= BAR_MAX_WIDTH:
                bars.append((y, int(s), int(e)))
    return bars


def main(path: str) -> None:
    print(f"Reading {path}")
    img = Image.open(path).convert("L")
    arr = np.array(img)
    H, W = arr.shape
    print(f"Image: {W} x {H} px (grayscale, threshold>={WHITE_THRESHOLD} = white)\n")

    bars = find_scalebars(arr)
    if not bars:
        print("No candidate scalebars found.")
        return

    # Bars stacked vertically (thick line) appear as multiple rows with
    # near-identical (start, end). Cluster by (start, end) to dedupe thick bars
    # into a single (start, end, top_row, bottom_row) entry.
    cluster_key = lambda b: (b[1], b[2])
    clusters = {}
    for (y, s, e) in bars:
        clusters.setdefault((s, e), []).append(y)

    width_counter = Counter()
    print(f"Found {len(clusters)} distinct horizontal bars (by start/end col):")
    print(f"{'row_top':>8} {'row_bot':>8} {'col_lo':>8} {'col_hi':>8} {'width_px':>10}")
    for (s, e), rows in sorted(clusters.items(), key=lambda kv: (min(kv[1]), kv[0][0])):
        if len(rows) < 2:
            continue  # ignore single-row runs (likely noise / text)
        rows_sorted = sorted(rows)
        # Only report contiguous bands (real bars are stacked rows)
        contiguous = []
        cur = [rows_sorted[0]]
        for r in rows_sorted[1:]:
            if r == cur[-1] + 1:
                cur.append(r)
            else:
                if len(cur) >= 2:
                    contiguous.append((cur[0], cur[-1]))
                cur = [r]
        if len(cur) >= 2:
            contiguous.append((cur[0], cur[-1]))
        for (rtop, rbot) in contiguous:
            w = e - s
            print(f"{rtop:>8} {rbot:>8} {s:>8} {e:>8} {w:>10}")
            width_counter[w] += 1

    print("\nWidth histogram:")
    for w, n in width_counter.most_common():
        print(f"  width = {w:>4} px  ->  {n} bars")

    widths = list(width_counter.elements())
    if widths:
        import statistics as st
        print(
            f"\nSummary across {len(widths)} bars: "
            f"min={min(widths)}, max={max(widths)}, "
            f"median={int(st.median(widths))}, "
            f"stdev={st.stdev(widths) if len(widths) > 1 else 0:.2f}"
        )


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python check_scalebar_pixel_widths.py <montage.png>")
        sys.exit(1)
    main(sys.argv[1])
