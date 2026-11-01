#!/usr/bin/env python
"""
Export a deck's slides to PNG through PowerPoint, all of them or only the ones named, so a small
edit re-exports just its slides. PNG names: slide_NN.png, 3 digits past 99. Never quits PowerPoint.

Usage (base miniforge python):
    python examples_and_configs/export_deck_slides.py <deck> <out dir> [width px] [--slides 2,31,39]
"""

import os
import sys

import win32com.client as w


def main() -> int:
    args = [a for a in sys.argv[1:] if a != "--slides"]
    wanted = None
    if "--slides" in sys.argv:
        wanted = {int(n) for n in sys.argv[sys.argv.index("--slides") + 1].split(",")}
        args.remove(sys.argv[sys.argv.index("--slides") + 1])
    deck, out = os.path.abspath(args[0]), os.path.abspath(args[1])
    width = int(args[2]) if len(args) > 2 else 2000
    os.makedirs(out, exist_ok=True)
    app = w.Dispatch("PowerPoint.Application")
    for k in range(1, app.Presentations.Count + 1):
        try:
            if os.path.normcase(app.Presentations(k).FullName) == os.path.normcase(deck):
                print(f"[skip] {deck} is open in PowerPoint")
                return 1
        except Exception:      # a presentation that refuses automation is not this deck
            continue
    pres = app.Presentations.Open(deck, True, False, False)  # ReadOnly, Untitled, WithWindow
    try:
        total = pres.Slides.Count
        pad = 3 if total > 99 else 2
        numbers = sorted(wanted) if wanted else range(1, total + 1)
        for i in numbers:
            pres.Slides(i).Export(os.path.join(out, f"slide_{i:0{pad}d}.png"), "PNG", width, width * 9 // 16)
    finally:
        pres.Close()
    print(f"exported {len(numbers)} of {total} slides to {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
