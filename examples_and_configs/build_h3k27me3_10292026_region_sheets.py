#!/usr/bin/env python
"""
One sheet per example nucleus for the H3K27me3 deck (FF 2026-11-01): the 3 region sizes around
the 2 points, one region per tile, Region Nadir above Centrosome Projection. Tiles from the
microscopy-nodes-blender-vis session, same camera and crop within a cell, 56 px/μm.

Usage (base miniforge python):
    python examples_and_configs/build_h3k27me3_10292026_region_sheets.py
Writes <experiment>/tests/region_sheets_20261101/cell_<id>_regions.png.
"""

import sys
import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, str(Path(__file__).resolve().parent))
from insert_noco_10262026_montage_slides import with_scalebar  # noqa: E402

TILES = Path("L:/FF/blender_renderings/_raw/h3k1029_seed_disks_20261101/noco1029_dmso")
OUT = Path("L:/FF/Nucleus_centrosome/nocodazole_fixed/"
           "10292026_jurkats_noco_h3k27me3640LP45_p549LP45_egfpcentrin2488LP45_h405LP40_/tests/region_sheets_20261101")
CELLS = (10, 34)
PPU = 56               # px/μm of every tile
BAR_THICK = 22
# FF 2026-11-01: each name in its region's color. The Nadir region's (48, 8, 96) is unreadable on the
# black ground, so its name takes a lighter violet of the same hue.
SEEDS = (("nadir", "Region\nNadir", (160, 110, 245)), ("cent", "Centrosome\nProjection", (255, 0, 255)))
NAMES = {"Region Nadir": SEEDS[0][2], "Centrosome Projection": SEEDS[1][2]}
SIZES = (("r010", "10% of mean distance"), ("r025", "25% of mean distance"), ("r050", "50% of mean distance"))
FONT_FRAC = 0.082      # label height / tile height, readable at FF's fit-to-width view
GAP = 24
# Italic text (FF 2026-11-01). The 50% region around the Centrosome Projection is faces_by_cent, the
# Cent.-Prox. Reg. (microscopy-nodes-blender-vis gate: 0 faces differ on both cells).
NOTE = ("Region = nuclear surface within a set distance of the white dot, measured along the surface. "
        "Distance = 10, 25 or 50% of the average distance from the dot to all faces on the nucleus, "
        "weighted by face area.",
        "Region Nadir: deepest point of the Cent.-Prox. Reg., farthest from the convex hull surface.",
        "Centrosome Projection: surface point closest to the centrosome; its 50% region is the Cent.-Prox. Reg.")
NOTE_CHARS = 118       # characters per line of the italic text
NOTE_FONT_FRAC = 0.018  # font size / sheet width


def note_line(draw, center: tuple, line: str, font) -> None:
    """One centered line of the italic text, a leading seed name in its color."""
    name = next((n for n in NAMES if line.startswith(n)), "")
    x = center[0] - draw.textlength(line, font=font) / 2
    draw.text((x, center[1]), name, fill=NAMES.get(name), font=font, anchor="lm")
    draw.text((x + draw.textlength(name, font=font), center[1]), line[len(name):], fill="white", font=font,
              anchor="lm")


def sheet(cell: int) -> Path:
    """The cell's 2 x 3 sheet, black ground, white labels."""
    tiles = [[with_scalebar(Image.open(TILES / f"disk_{seed}_{size}" / f"noco1029_dmso_cell_{cell}_disk_{seed}_{size}.png")
                            .convert("RGB"), PPU, thick=BAR_THICK) for size, _ in SIZES] for seed, _, _ in SEEDS]
    w, h = tiles[0][0].size
    font = ImageFont.truetype("arialbd.ttf", round(h * FONT_FRAC))
    left, top = round(h * 0.62), round(h * 0.13)
    width = left + 3 * w + 2 * GAP
    italic = ImageFont.truetype("ariali.ttf", round(width * NOTE_FONT_FRAC))
    lines = [ln for para in NOTE for ln in textwrap.wrap(para, NOTE_CHARS)]
    line_h = round(italic.size * 1.3)
    out = Image.new("RGB", (width, top + 2 * h + GAP + line_h * (len(lines) + 1)), (0, 0, 0))
    draw = ImageDraw.Draw(out)
    for k, ln in enumerate(lines):
        note_line(draw, (width // 2, top + 2 * h + GAP + line_h * (k + 1)), ln, italic)
    for j, (_, label) in enumerate(SIZES):
        draw.text((left + j * (w + GAP) + w // 2, top // 2), label, fill="white", font=font, anchor="mm")
    for i, (_, label, color) in enumerate(SEEDS):
        y = top + i * (h + GAP)
        draw.multiline_text((left // 2, y + h // 2), label, fill=color, font=font, anchor="mm", align="center")
        for j in range(3):
            out.paste(tiles[i][j], (left + j * (w + GAP), y))
    path = OUT / f"cell_{cell}_regions.png"
    out.save(path)
    print(path, out.size)
    return path


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    for c in CELLS:
        sheet(c)
