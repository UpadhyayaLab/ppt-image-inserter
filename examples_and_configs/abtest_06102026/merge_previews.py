#!/usr/bin/env python
"""
Quick PREVIEWS of an H3K27me3 + H3K27ac merge for the 06/10/2026 antibody test (well A1, the only well with both
marks), for FF's decision page on merge colors, contrast and families. Not a pipeline render: no MATLAB, no
background subtraction beyond a per-slice floor, nearest raw z-slice. Every picture says so.

Data: raw single-cell crops <exp>/A1_H3K27me3_H3K27ac_/cropped/channels/C{1..4}-<FOV>_<id>.tif
      (C1 H3K27me3, C2 actin, C3 H3K27ac, C4 Hoechst; dataset config Jurkats_H3K27_AbTest_20260610_A1_dataset.m,
      FF's channel map 2026-11-20). Pixel 0.065 μm, z-step 0.3 μm.
Slice: the pipeline's broadest nuclear slice (z in physical_scale_images/H3K27me3_nuc_bz/cells/Cell_<id>_z<z>.png,
      on the isotropic stack), mapped to the nearest raw slice: round((z - 1) * 0.065 / 0.3).
Cells: 3 kept A1 cells (cells 2 and 5 are excluded by the compile, mesh genus).

Writes the option PNGs, the montage-deck slide PNGs copied as the page's "All slides", and page/page_data.json for
decision_page/build_decision_page.py.

Usage (base miniforge python):
    python examples_and_configs/abtest_06102026/merge_previews.py <out dir>
"""
import json
import re
import shutil
import sys
from pathlib import Path

import numpy as np
import tifffile
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from ppt_image_inserter import deck_strings  # noqa: E402

EXP = Path("L:/FF/Nucleus_H3K27me3/06102026_jurkatsCD3plus_onPLL_testingAntibodiesH3K27me3H3K27ac_")
RAW = EXP / "A1_H3K27me3_H3K27ac_" / "cropped" / "channels"
BZ = EXP / "prog_fixed_cells" / "A1_H3K27me3_H3K27ac_" / "physical_scale_images" / "H3K27me3_nuc_bz" / "cells"
DECK = Path("K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Histone Mods_RNAPII/H3K27_AbTest_20260610/"
            "H3K27 antibody test, montages, 06102026.pptx")
DECK_PNG = EXP / "deck_png_20261121" / "montages"
PSIZE, ZSTEP = 0.065, 0.3
CELLS = [1, 6, 9]
KEPT = [0, 1, 3, 4, 6, 7, 8, 9, 10, 11]      # A1 minus excluded cells 2 and 5 (excluded_cells_applied.csv)
SCALE = 3                                    # display upsampling (nearest), so one tile is ~3x raw size
PCT = (0, 99.9)                              # contrast: 0 = the background floor, top = 99.9th pct in the nucleus
NOTE = "Preview, not pipeline render"

MAG, GRN, RED, CYN, BLU = (1, 0, 1), (0, 1, 0), (1, 0, 0), (0, 1, 1), (0, 0.45, 1)
COLORS = {  # choice -> (H3K27me3 color, H3K27ac color, button text)
    "mg": (MAG, GRN, "H3K27me3 magenta, H3K27ac green"),
    "gm": (GRN, MAG, "H3K27me3 green, H3K27ac magenta"),
    "rc": (RED, CYN, "H3K27me3 red, H3K27ac cyan"),
}
NAMES = {MAG: "magenta", GRN: "green", RED: "red", CYN: "cyan", BLU: "blue"}


def font(size, bold=False):
    return ImageFont.truetype("arialbd.ttf" if bold else "arial.ttf", size)


def raw_path(ch, cid):
    hits = list(RAW.glob(f"C{ch}-*_{cid}.tif"))
    if len(hits) != 1:
        raise SystemExit(f"cell {cid} channel {ch}: {len(hits)} files")
    return hits[0]


def cell_slices(cid):
    """(H3K27me3, H3K27ac, DNA) raw slices at the broadest nuclear slice, floored, plus a nucleus mask."""
    hits = [p for p in BZ.iterdir() if re.fullmatch(rf"Cell_{cid}_z(\d+)\.png", p.name)]
    z_iso = int(re.fullmatch(rf"Cell_{cid}_z(\d+)\.png", hits[0].name).group(1))
    sl = []
    for ch in (1, 3, 4):
        st = tifffile.imread(raw_path(ch, cid)).astype(float)
        z = int(round((z_iso - 1) * PSIZE / ZSTEP))
        z = min(max(z, 0), st.shape[0] - 1)
        sl.append(st[z])
    mask = sl[2] > np.percentile(sl[2], 60)        # Hoechst-bright region = nucleus (preview only)
    outside = sl[2] < np.percentile(sl[2], 30)     # clearly outside the nucleus
    # Floor = each channel's own background (median outside the nucleus), so background reads black.
    out = [np.clip(s - np.median(s[outside]), 0, None) for s in sl]
    return out[0], out[1], out[2], mask


def crop_box(mask, size):
    ys, xs = np.nonzero(mask)
    cy, cx = int(np.median(ys)), int(np.median(xs))
    h, w = mask.shape
    y0 = min(max(cy - size // 2, 0), h - size)
    x0 = min(max(cx - size // 2, 0), w - size)
    return y0, x0


def norm(img, lo, hi):
    return np.clip((img - lo) / max(hi - lo, 1e-9), 0, 1)


def rgb(chans):
    """chans: [(normalized image, color)] -> uint8 RGB, additive."""
    out = np.zeros(chans[0][0].shape + (3,))
    for im, col in chans:
        out += im[..., None] * np.array(col)[None, None, :]
    return (np.clip(out, 0, 1) * 255).astype(np.uint8)


def tile(arr):
    im = Image.fromarray(arr if arr.ndim == 3 else np.stack([arr] * 3, -1))
    return im.resize((im.width * SCALE, im.height * SCALE), Image.NEAREST)


def bar(im):
    """5 μm scale bar, bottom left."""
    d = ImageDraw.Draw(im)
    px = 5 / PSIZE * SCALE
    d.rectangle([14, im.height - 26, 14 + px, im.height - 16], fill="white")
    return im


def wrap(lines, fnt, width):
    """Greedy word wrap of each key line at spaces so no line is wider than `width` px."""
    d = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    out = []
    for line in lines:
        cur = ""
        for word in line.split(" "):
            trial = f"{cur} {word}".strip()
            if cur and d.textlength(trial, font=fnt) > width:
                out.append(cur)
                cur = word
            else:
                cur = trial
        out.append(cur)
    return out


def sheet(rows, key, out):
    """rows: [(cell id, [panel images])]; key: text lines under the note. One row per cell."""
    # Text scales with the sheet width: the page shows each sheet ~350 px wide, so fixed 30-40 px text read as ~6 px.
    pad = 16
    pw, ph = rows[0][1][0].size
    ncol = len(rows[0][1])
    side = ncol == 1                     # one panel per cell: the cells side by side
    W = 20 + (len(rows) if side else ncol) * (pw + pad)
    # Key and cell labels: W // 17 gives a cap height of ~14 px when the sheet is shown 350 px wide (figure-reviewer
    # 2026-11-21: W // 26 and W // 30 read ~9.6 and ~8.3 px). Key lines wrap at spaces to fit the sheet.
    tn, tk, tc = W // 16, W // 17, W // 17          # note, key, cell label sizes
    key = wrap(key, font(tk), W - 30)
    head = 20 + int(tn * 1.3) + int(tk * 1.25) * len(key) + 10
    lab = int(tc * 1.3)
    H = head + (lab + ph + pad if side else len(rows) * (lab + ph + pad))
    sh = Image.new("RGB", (W, H), "white")
    d = ImageDraw.Draw(sh)
    d.text((20, 14), NOTE, font=font(tn, True), fill=(180, 0, 0))
    for k, line in enumerate(key):
        d.text((20, 20 + int(tn * 1.3) + int(tk * 1.25) * k), line, font=font(tk), fill="black")
        if d.textlength(line, font=font(tk)) > W - 30:
            raise SystemExit(f"key line too wide for its sheet: {line!r}")
    for r, (cid, panels) in enumerate(rows):
        x0, y0 = (20 + r * (pw + pad), head) if side else (20, head + r * (lab + ph + pad))
        d.text((x0, y0), f"Cell {cid}", font=font(tc, True), fill="black")
        for c, p in enumerate(panels):
            sh.paste(bar(p.copy()), (x0 + c * (pw + pad), y0 + lab))
    sh.save(out)


def main(out_dir):
    out = Path(out_dir)
    page = out / "page"
    page.mkdir(parents=True, exist_ok=True)
    data = {c: cell_slices(c) for c in sorted(set(CELLS) | set(KEPT))}
    size = 260                                     # 17 μm square crop, the same for every cell
    # Fixed contrast: one range per mark, from the nucleus pixels of all kept A1 cells at their broadest slice.
    fixed = [(0.0, np.percentile(np.concatenate([data[c][k][data[c][3]] for c in KEPT]), PCT[1])) for k in range(3)]

    def panels(cid, cols, contrast, family):
        m3, ac, dna, mask = data[cid]
        y0, x0 = crop_box(mask, size)
        cut = lambda a: a[y0:y0 + size, x0:x0 + size]   # noqa: E731
        lim = (fixed if contrast == "fixed" else
               [(0.0, np.percentile(a[mask], PCT[1])) for a in (m3, ac, dna)])
        n3, nac, nd = (cut(norm(a, *l)) for a, l in zip((m3, ac, dna), lim))
        merge2 = tile(rgb([(n3, cols[0]), (nac, cols[1])]))
        if family == "merge":
            return [merge2]
        grays = [tile((n3 * 255).astype(np.uint8)), tile((nac * 255).astype(np.uint8))]
        if family == "three":
            return grays + [merge2]
        return grays + [tile((nd * 255).astype(np.uint8)),
                        tile(rgb([(n3, cols[0]), (nac, cols[1]), (nd, BLU)]))]

    def key_lines(cols, dna=False):
        return ([f"H3K27me3 {NAMES[cols[0]]}, H3K27ac {NAMES[cols[1]]}"] + (["DNA blue"] if dna else []))

    pngs = {}
    base = COLORS["mg"][:2]
    for ch, (c3, cac, _) in COLORS.items():
        f = f"opt_colors_{ch}.png"
        sheet([(c, panels(c, (c3, cac), "cell", "merge")) for c in CELLS],
              key_lines((c3, cac)) + ["Contrast per cell", "Bar = 5 μm"], page / f)
        pngs[("colors", ch)] = f
    for ch, txt in (("cell", "Contrast per cell"), ("fixed", "Contrast fixed per mark across the well")):
        f = f"opt_contrast_{ch}.png"
        sheet([(c, panels(c, base, ch, "merge")) for c in CELLS], key_lines(base) + [txt, "Bar = 5 μm"], page / f)
        pngs[("contrast", ch)] = f
    fam_key = {"merge": "Merge only", "three": "H3K27me3 | H3K27ac | merge",
               "four": "H3K27me3 | H3K27ac | DNA | merge"}
    for ch, txt in fam_key.items():
        f = f"opt_families_{ch}.png"
        sheet([(c, panels(c, base, "cell", ch)) for c in CELLS],
              [txt] + key_lines(base, ch == "four") + ["Contrast per cell", "Bar = 5 μm"], page / f)
        pngs[("families", ch)] = f

    # "All slides": the montage deck as built, its own exported slides.
    prs = Presentation(str(DECK))
    lines = {}
    for where, line in deck_strings(prs):
        lines.setdefault(int(re.search(r"\d+", where).group()), []).append(line)
    deck = []
    for n in range(1, len(prs.slides) + 1):
        png = f"deck_{n:02d}.png"
        shutil.copy(DECK_PNG / f"slide_{n:02d}.png", page / png)
        deck.append({"n": n, "id": f"montage_slide_{n:02d}", "png": png, "lines": lines.get(n, []),
                     "unlisted": "", "fact": None})
    nxt = len(deck) + 1                                 # the merge slides would be new slides after the deck
    decisions = {code: {"applies": [2, 3], "options": {}} for code in ("colors", "contrast", "families")}
    for (code, ch), f in pngs.items():
        decisions[code]["options"][ch] = [{"n": nxt, "png": f, "lines": [], "fact": None}]
    (page / "page_data.json").write_text(json.dumps({"deck": deck, "decisions": decisions}, ensure_ascii=False,
                                                    indent=1), encoding="utf-8")
    print(f"fixed contrast ranges (H3K27me3, H3K27ac, DNA): {[tuple(round(v, 1) for v in r) for r in fixed]}")
    print(f"wrote {len(pngs)} option sheets and {len(deck)} deck slides to {page}")


if __name__ == "__main__":
    main(sys.argv[1])
