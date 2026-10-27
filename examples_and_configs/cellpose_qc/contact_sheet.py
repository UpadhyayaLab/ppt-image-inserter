"""Tile a condition's cached QC composites into one labeled contact sheet, for a
quick look at every FOV's mask at once (the decks split them 9 per slide).

Usage (PPT_editing env):
    python contact_sheet.py <cache_dir> <out.jpg> [--suffix _with_nuc] [--cols 3]
"""
import argparse
import os
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ap = argparse.ArgumentParser()
ap.add_argument("cache_dir")
ap.add_argument("out")
ap.add_argument("--suffix", default="_with_nuc")
ap.add_argument("--cols", type=int, default=3)
ap.add_argument("--tile-w", type=int, default=1000)
args = ap.parse_args()

fs = sorted(p for p in Path(args.cache_dir).glob(f"*{args.suffix}.jpg")
            if re.match(r"^\d+" + re.escape(args.suffix) + r"$", p.stem))
W = args.tile_w
H = W // 2
rows = (len(fs) + args.cols - 1) // args.cols
sheet = Image.new("RGB", (args.cols * W, rows * H), "gray")
draw = ImageDraw.Draw(sheet)
font = ImageFont.truetype("arial.ttf", max(16, W // 28))
for i, f in enumerate(fs):
    x, y = (i % args.cols) * W, (i // args.cols) * H
    sheet.paste(Image.open(f).resize((W - 6, H - 6)), (x + 3, y + 3))
    draw.text((x + int(W * 0.44), y + 10), "FOV " + f.stem[: -len(args.suffix)],
              fill="yellow", font=font)
os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
sheet.save(args.out, quality=88)
print(f"{len(fs)} FOVs -> {args.out}")
