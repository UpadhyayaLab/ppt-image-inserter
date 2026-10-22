#!/usr/bin/env python3
r"""build_grant_combined_deck.py  Three-slide grant-report deck (FF 2026-09-29): the actin and
pMLC slides of build_grant_deck.py followed by the CD19 internalization slide of
build_grant_cd19_slide.py, WITHOUT panel letters. Every slide is built by its own module's
function, so this deck cannot drift from the two single decks; neither of those decks (nor
their builders) is touched -- the letters are dropped here only, by swapping the shared
letter() helper for a no-op while this deck is built.

Output: K:\FF\PPT\PPT_autogeneration\CART\grant\CART_grant_actin_pMLC_CD19_<yyyymmdd>.pptx
Run:    python build_grant_combined_deck.py [--export <dir>]
"""
from __future__ import annotations

import argparse
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_grant_deck as bgd          # noqa: E402  (helpers + actin/pMLC slides)
import build_grant_cd19_slide as cd19   # noqa: E402  (CD19 slide)

bpd = bgd.bpd
OUT = bgd.OUT_DIR / f'CART_grant_actin_pMLC_CD19_{datetime.date.today():%Y%m%d}.pptx'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--export', help='also export one PNG per slide to this dir')
    a = ap.parse_args()
    if OUT.exists() and bgd.backup_presentation is not None:
        bgd.backup_presentation(str(OUT), backup_base=str(bgd.OUT_DIR / 'backups'))
    bgd.letter = lambda *args, **kwargs: None      # FF 2026-09-29: no panel letters
    deck = bpd.Deck()
    bgd.actin_slide(deck)
    bgd.pmlc_slide(deck)
    cd19.cd19_slide(deck)
    bgd.calibri(deck.prs)
    deck.prs.save(str(OUT))
    print(f'saved {OUT} ({len(deck.prs.slides)} slides)')
    if a.export:
        bpd.OUT = OUT                       # export_pngs opens bpd.OUT
        bpd.export_pngs(a.export, deck.titles)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
