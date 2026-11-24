"""Example slide for FF's pick: four numbers per FOV beside each movie on one TGT movie slide (review deck only).

The layout and its checks live in fov_numbers.py (shared with the deck builder, config key `fov_numbers`). FF approved
the example on 2026-11-23 (review\\fov_numbers_example_v7, built when the values were from the old threshold and
marked "Provisional values"; pass --provisional to mark them again).

    python examples_and_configs\\tgt\\fov_numbers_example.py --config <deck config> --csv <fov_slide_numbers.csv>
        --slide 5 --out K:\\FF\\PPT\\PPT_autogeneration\\TGT\\review\\fov_numbers_example_v8.pptx
"""
import argparse
import shutil
import sys
import tempfile
from pathlib import Path

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import build_cat_fmc_slides as b  # noqa: E402
import fov_numbers as fn  # noqa: E402
from ppt_image_inserter.movies import build_movie_deck_via_com  # noqa: E402

COLUMNS = ("away_rise", "ratio_last", "low_last_pooled", "high_last_pooled")   # fov_slide_numbers v2 column names


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--config", type=Path, required=True)
    ap.add_argument("--csv", type=Path, required=True)
    ap.add_argument("--columns", nargs=4, default=COLUMNS, help="csv columns of the four labels, in label order")
    ap.add_argument("--slide", type=int, default=5, help="deck slide number of the movie slide to use")
    ap.add_argument("--provisional", action="store_true", help='mark the values "Provisional values"')
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()
    if args.out.exists():
        sys.exit(f"STOP: {args.out} exists and is never overwritten")
    cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
    entry = b.config_plan(cfg)[args.slide - 1]
    if entry[0] != "movie" or {c for c, _ in entry[2]} != {"CAT", "FMC"}:
        sys.exit(f"STOP: slide {args.slide} is not a movie slide with CAT and FMC movies")
    rows, removed = fn.load_rows(args.csv), b.removed_marks(cfg)
    scratch = Path(tempfile.mkdtemp(prefix="fov_numbers_"))
    try:
        maker = b.MovieSlideMaker(scratch, removed)
        s, shown = fn.numbers_slide(maker(entry[1], entry[2]), maker, rows, removed, args.columns, args.provisional)
        maker.write_posters([s])   # MovieSlideMaker writes posters only on request (incremental builds)
        build_movie_deck_via_com([s], str(args.out), b.SLIDE_W, b.SLIDE_H, loop=True)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    from pptx import Presentation
    problems = fn.check_numbers_slide(Presentation(str(args.out)).slides[0], shown)
    print(f"{args.out}: deck slide {args.slide}, movies {maker.movie_w:.1f} x {maker.movie_h:.1f} pt (deck size)")
    print(f"sizes: values {fn.VAL_PT:g} pt bold, labels {fn.LAB_PT:g} pt, (last frame) {fn.QUAL_PT:g} pt, definitions "
          f"{fn.DEF_PT:g} pt")
    print(f"CHECK no text under {fn.MIN_PT:g} pt, no overlaps, every value equals the csv: "
          + ("PASS" if not problems else "FAIL" + "".join(f"\n  {p}" for p in problems)))
    for _, paths in entry[2]:
        for p in paths:
            r = fn.row_for(p.stem, rows, removed)
            print(f"  {p.stem}: " + ("removed, no numbers" if r["removed"] == "yes" else
                                     " | ".join(v or "(blank)" for v in fn.values(r, args.columns))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
