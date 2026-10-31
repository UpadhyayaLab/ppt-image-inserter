"""Export every slide of one or more .pptx decks to PNG (1920x1080) through PowerPoint, for review pages.

Same rules as build_pmlc_profiles_deck.export_pngs (CART_pMLC_ppt): PowerPoint is single-instance, so Dispatch
attaches to FF's open window -- NEVER app.Quit(); only Close() the presentation opened here, and skip a deck FF
already has open. Output: <out_root>/<deck stem>/NN.png (touches only those names, removing any past the last slide).

Usage (base miniforge3 has pywin32; the PPT_editing env does not):
    C:\\Users\\Upadhyaya_Lab\\miniforge3\\python.exe export_deck_pngs.py --out "<dataset root>\\deck_png_<date>" deck1.pptx [...]
"""
import argparse
import os
from pathlib import Path

import win32com.client as w


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--out", required=True, help="root folder; one subfolder per deck")
    ap.add_argument("--width", type=int, default=1920)
    ap.add_argument("--height", type=int, default=1080)
    ap.add_argument("decks", nargs="+")
    args = ap.parse_args()

    app = w.Dispatch("PowerPoint.Application")
    open_now = set()
    for k in range(1, app.Presentations.Count + 1):
        try:                                  # protected-view / locked decks deny FullName
            open_now.add(os.path.normcase(app.Presentations(k).FullName))
        except Exception:                     # noqa: BLE001 -- unreadable: cannot be ours to skip
            continue
    for deck in args.decks:
        deck = os.path.abspath(deck)
        if os.path.normcase(deck) in open_now:
            print(f"[skip] {deck} is open in PowerPoint; close it there to export")
            continue
        out = Path(args.out) / Path(deck).stem
        out.mkdir(parents=True, exist_ok=True)
        pres = app.Presentations.Open(deck, ReadOnly=True, WithWindow=False)
        try:
            n = pres.Slides.Count
            for i in range(1, n + 1):
                pres.Slides(i).Export(str(out / f"{i:02d}.png"), "PNG", args.width, args.height)
        finally:
            pres.Close()                      # our presentation only; the app stays running
        for old in out.glob("*.png"):         # an older export of a longer deck
            if old.stem.isdigit() and int(old.stem) > n:
                old.unlink()
        print(f"exported {n} slides -> {out}", flush=True)


if __name__ == "__main__":
    main()
