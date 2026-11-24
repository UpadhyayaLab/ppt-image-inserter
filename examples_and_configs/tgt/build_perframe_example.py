"""Example slides for FF's pick: one pair of movies per slide, shown twice (TGT per-movie scaling on top, as in the
deck; TGT per-frame scaling below), with two candidate row labels (review deck only, never the main deck).

FF 2026-11-23, relayed: "every slide should have just one pair of movies. one is just like the current one. the
second should employ per-frame scaling below - and we need to indicate that"; "so number of data slides will be
doubled"; "IRM can be unchanged".

Top row: the deck's movie slide layout (build_cat_fmc_slides.MovieSlideMaker + fov_numbers.numbers_slide), the first
pair (row) of the chosen deck slide, with its four numbers, red marks and definition lines. Bottom row: the same two
FOVs from render_cat_fmc_movies.py's "frame" version (same run, frames and outlines; TGT low and high = 1st and 99.9th
percentile of that frame), no numbers. The labels sit in the bottom row's empty middle column, at the number-label
size or larger. One slide: a two-line key, bold.

    python examples_and_configs/tgt/build_perframe_example.py --config <deck config> --slide 5
        --out K:/FF/PPT/PPT_autogeneration/TGT/review/perframe_example/perframe_example_v1.pptx
"""
import argparse
import shutil
import sys
import tempfile
from dataclasses import replace
from pathlib import Path

import imageio.v3 as iio
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
import build_cat_fmc_slides as b  # noqa: E402
import fov_numbers as fn  # noqa: E402
from ppt_image_inserter.movies import SlideSpec, build_movie_deck_via_com  # noqa: E402

# FF 2026-11-23, relayed: "on eevery slide now: the scaling must be clearly stated"; wording from the director's
# "TGT: per-movie scaling" / "TGT: per-frame scaling" with "Top"/"Bottom" to tie each line to its row, not approved.
# The top row's middle column is full (numbers block 183.5 of 196 pt), so the key sits in the bottom row's.
LABELS = {"key": ["Top: TGT per-movie scaling", "Bottom: TGT per-frame scaling"]}
KEY_PT, KEY_PITCH = 20.0, 24.0   # above the number labels' 16 pt, bold, black: readable across a room


def frame_movie(mp4):
    """The per-frame version of a deck movie: <root>/movie/<day>/x.mp4 -> <root>/frame/<day>/x.mp4."""
    return mp4.parents[2] / "frame" / mp4.parent.name / mp4.name


def pairs_per_day(plan):
    """Slides per day in the one-pair layout: one per row of every current movie slide."""
    out = {}
    for e in plan:
        if e[0] == "movie":
            out[e[3]] = out.get(e[3], 0) + max(len(ps) for _, ps in e[2])
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--config", type=Path, required=True)
    ap.add_argument("--slide", type=int, required=True, help="deck slide number of the movie slide (its first row)")
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()
    if args.out.exists():
        sys.exit(f"STOP: {args.out} exists and is never overwritten")
    cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
    plan = b.config_plan(cfg)
    entry = plan[args.slide - 1]
    if entry[0] != "movie":
        sys.exit(f"STOP: slide {args.slide} is not a movie slide")
    top = [(c, ps[0]) for c, ps in entry[2] if c and ps]
    bottom = [(c, frame_movie(p)) for c, p in top]
    for (c, p), (_, q) in zip(top, bottom):   # the per-frame movie: same FOV (file name), same frame count
        if not q.is_file():
            sys.exit(f"STOP: {q} not found")
        n_p, n_q = (sum(1 for _ in iio.imiter(x)) for x in (p, q))
        ok = p.name == q.name and n_p == n_q
        print(f"CHECK {c}: {p.name} ({n_p} frames) above {q.parent.parent.name}/{q.name} ({n_q} frames): "
              + ("PASS" if ok else "FAIL"))
        if not ok:
            sys.exit("STOP: the per-frame movie does not match the movie above it")
    n_day = pairs_per_day(plan)[entry[3]]
    title = f"{entry[3].replace('_', ' ')} (slide 1 of {n_day})"
    rows, removed = fn.load_rows(cfg["fov_numbers"]["csv"]), b.removed_marks(cfg)
    scratch = Path(tempfile.mkdtemp(prefix="perframe_example_"))
    try:
        maker = b.MovieSlideMaker(scratch, removed)
        full = maker(title, [(c, [p, q]) for (c, p), (_, q) in zip(top, bottom)])   # posters and both rows' places
        y0 = min(m.top_pt for m in full.movies)
        boxes = [tb for tb in full.textboxes   # a removed FOV's red mark under the top row only
                 if not (tb.color_rgb == b.MARK_BGR and tb.top_pt > y0 + maker.movie_h + 1)]
        top_spec = SlideSpec(boxes, [m for m in full.movies if m.top_pt == y0], full.pictures)
        s, shown = fn.numbers_slide(top_spec, maker, rows, removed, cfg["fov_numbers"]["columns"])
        w, h, dy = maker.movie_w, maker.movie_h, fn.ROWS_TOP - b.ROWS_TOP
        lefts = [fn.EDGE, b.SLIDE_W - fn.EDGE - w]
        row2 = [replace(m, left_pt=lefts[0 if m.left_pt < b.SLIDE_W / 2 else 1], top_pt=m.top_pt + dy)
                for m in full.movies if m.top_pt != y0]
        gap_l, gap_r = fn.EDGE + w + fn.CLEAR, b.SLIDE_W - fn.EDGE - w - fn.CLEAR
        row2_top = row2[0].top_pt
        slides = []
        for key, lines in LABELS.items():
            for ln in lines:
                if b.wrapped_lines(ln, KEY_PT, gap_r - gap_l, font_path=r"C:\Windows\Fonts\calibrib.ttf",
                                   insets_pt=0.0) > 1:
                    sys.exit(f"STOP: label {ln!r} does not fit one line at {KEY_PT:g} pt bold")
            lab_h = KEY_PITCH * len(lines)
            lab = fn.box("\r".join(lines), gap_l, row2_top + (h - lab_h) / 2, gap_r - gap_l, lab_h, KEY_PT,
                         "center", bold=True, pitch=KEY_PITCH)
            slides.append(SlideSpec(tuple(s.textboxes) + (lab,), tuple(s.movies) + tuple(row2), s.pictures))
        args.out.parent.mkdir(parents=True, exist_ok=True)
        maker.write_posters(slides)   # MovieSlideMaker writes posters only on request (incremental builds)
        build_movie_deck_via_com(slides, str(args.out), b.SLIDE_W, b.SLIDE_H, loop=True)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    from pptx import Presentation
    prs = Presentation(str(args.out))
    for key, sl in zip(LABELS, prs.slides):
        problems = fn.check_numbers_slide(sl, shown)
        print(f"CHECK example {key}: no text under {fn.MIN_PT:g} pt, no overlaps, every value equals the csv, "
              f"definition lines once: " + ("PASS" if not problems else "FAIL" + "".join(f"\n  {p}" for p in problems)))
    print(f"{args.out}: {len(prs.slides)} slides; title {title!r}; movies {w:.0f} x {h:.0f} pt; key {KEY_PT:g} pt bold")
    return 0


if __name__ == "__main__":
    sys.exit(main())
