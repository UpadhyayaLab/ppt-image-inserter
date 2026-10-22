#!/usr/bin/env python3
r"""sketch_grant_wireframe.py  To-scale box wireframes (16:9, 13.333 x 7.5 in) for the
two grant-report slides (FF 2026-09-29): pMLC and actin (3/5/7/10 min set), CAT vs FMC
at 10 min, representative cells + pooled quantification. Boxes only, large labels
(paper-figures skill: sketch first, FF picks, then the deck builder builds).

Panel aspect ratios are the real PNGs': reps and violins square, profiles 1.19:1.
Output: K:\FF\PPT\PPT_autogeneration\CART\grant\wireframes\wireframe_<slide>.png
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(r'K:\FF\PPT\PPT_autogeneration\CART\grant\wireframes')
W, H, DPI = 13.333, 7.5, 110
PROF_AR = 1.19
X0 = 0.38                  # left margin; both slides share the right edge X0 + 12.58
NL = chr(10)

# (label, x, y, w, h) in inches.
SLIDES = {
    # Row 1: reps (4 cols of 1.745 in tiles + 0.8 in row labels) | profile (1.19:1);
    # Row 2: four square violins spanning the same left/right edges.
    'pMLC': [
        ('Title', 0.3, 0.1, 12.7, 0.55),
        ('A  Reps 2x4' + NL + '(Actin | pMLC | Merge | Regions)', X0, 0.8, 7.96, 3.55),
        ('B  Profile' + NL + 'Actin + pMLC', X0 + 8.11, 0.8, 3.55 * PROF_AR, 3.55),
        ('C  pMLC MFI' + NL + '/ CAT Median', X0, 4.45, 2.95, 2.95),
        ('D  pMLC' + NL + 'Central', X0 + 3.127, 4.45, 2.95, 2.95),
        ('E  pMLC' + NL + 'Peripheral', X0 + 6.254, 4.45, 2.95, 2.95),
        ('F  pMLC' + NL + 'Distal', X0 + 9.381, 4.45, 2.95, 2.95),
    ],
    # One tall row: reps 2x2 (2.9 in tiles) | 2x2 grid of 2.9 in plot slots.
    'Actin': [
        ('Title', 0.3, 0.1, 12.7, 0.55),
        ('A  Reps 2x2' + NL + '(Actin | Regions)', X0, 0.9, 6.66, 5.86),
        ('B  Profile' + NL + 'Actin', X0 + 6.86, 0.9, 2.9, 2.9),
        ('C  Synapse' + NL + 'Area', X0 + 9.86, 0.9, 2.9, 2.9),
        ('D  Actin' + NL + 'Zones', X0 + 6.86, 3.86, 2.9, 2.9),
        ('E  IRM Area?' + NL + '(or R_median)', X0 + 9.86, 3.86, 2.9, 2.9),
    ],
}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    font = ImageFont.truetype('calibrib.ttf', 26)
    for name, boxes in SLIDES.items():
        im = Image.new('RGB', (int(W * DPI), int(H * DPI)), 'white')
        d = ImageDraw.Draw(im)
        d.rectangle([0, 0, im.width - 1, im.height - 1], outline='black', width=3)
        for lab, x, y, w, h in boxes:
            box = [x * DPI, y * DPI, (x + w) * DPI, (y + h) * DPI]
            d.rectangle(box, outline=(60, 90, 160), width=3, fill=(235, 241, 252))
            txt = lab if lab == 'Title' else f'{lab}{NL}{w:.1f} x {h:.1f} in'
            d.multiline_text(((box[0] + box[2]) / 2, (box[1] + box[3]) / 2), txt, fill='black',
                             font=font, anchor='mm', align='center')
        p = OUT / f'wireframe_{name}.png'
        im.save(p)
        print(p)


if __name__ == '__main__':
    main()
