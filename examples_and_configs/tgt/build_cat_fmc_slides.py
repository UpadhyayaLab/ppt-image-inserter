"""CAT and FMC movies on PowerPoint slides: 2 fields per cell type per slide, CAT left, FMC right.

FF 2026-11-22: "we do not need these pages. can use ppt"; "2 fovs per cell type per page sounds good but: IRM should
be to left of [TGT]"; "pls produce 1 slide to start"; "show me full day movies"; "one deck total".
Movies: render_cat_fmc_movies.py output (IRM with cell outlines left, TGT right, IRM frames only), version "movie"
(each movie on its own TGT brightness range) by default. Slide 13.33 x 7.5 in: title (32 pt); column headers CAT (blue)
and FMC (red); two rows of movies under each header. Movies autoplay and loop. Posters (first frames) are JPEG
(FF 2026-11-22: "yes" to JPEG posters; PNG posters were 40 % of the d5 deck).
Built with PPT_image_inserter's build_movie_deck_via_com (PowerPoint COM), which closes its deck and quits PowerPoint
when no other presentation is open.

Three modes:
  --fields W1:1 W1:2   one slide with those two fields of each cell type; title "20260925 d5: W1 FOV1 and FOV2".
  --full-day           every movie of one day; titles "20260925 d5 (slide X of Y)".
  --config <yaml>      one deck: title page, details slides, then every movie of each day in the config, in order;
                       movie slide titles count slides across the whole deck (FF: "Deck slides, whole deck").
Field order (--full-day, --config): each cell type W1 FOV1 onward, then W2 and so on. Each slide takes the next 2 CAT
(left) and next 2 FMC (right) fields; when one cell type runs out, the other's remaining fields fill slides column by
column (left top, left bottom, right top, right bottom) under its own header (FF: "last slide can just be 3 fmc that
is fine").
The deck is a new file; the script stops if it exists.
Moved from TGT_analysis\\scripts\\first_look 2026-11-22 (FF: "the script needs to be in ppt-inserter / period").
Run from the PPT_image_inserter repo root (env PPT_editing):
    python examples_and_configs\\tgt\\build_cat_fmc_slides.py --config examples_and_configs\\configs\\live\\tgt\\cat_fmc_all_days_deck.yaml
"""
import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path

import imageio.v3 as iio
import yaml
from PIL import Image

PPT_TOOL = Path(__file__).resolve().parents[2]   # repo root (examples_and_configs\tgt\ -> PPT_image_inserter)
sys.path.insert(0, str(PPT_TOOL))
from ppt_image_inserter.movies import (  # noqa: E402
    MovieSpec, PictureSpec, SlideSpec, TextboxSpec, build_movie_deck_via_com, make_run_scratch_dir)

DECK_DIR = Path(r"K:\FF\PPT\PPT_autogeneration\TGT")   # FF 2026-11-22: "it needs to be in ppt autogen folder"
SLIDE_W, SLIDE_H = 960.0, 540.0              # points (13.333 x 7.5 in)
MOVIE_RATIO = 1032 / 652                      # render_cat_fmc_movies.py frame: 2 x 512 + 8 gap, 512 + 2 x 70 strips
TITLE_TOP, TITLE_H, TITLE_PT = 6.0, 46.0, 32.0
HEAD_TOP, HEAD_H, HEAD_PT = 52.0, 32.0, 24.0
ROWS_TOP, ROWS_BOTTOM, ROW_GAP, COL_GAP = 88.0, 532.0, 8.0, 30.0
COLOR_BGR = {"CAT": 0xD97333, "FMC": 0x4040D9}   # CAT [0.20 0.45 0.85], FMC [0.85 0.25 0.25] as BGR integers
GREY_BGR = 0x707070                               # the cytokine deck's grey (subtitle, caption)
CELL_TYPES = ("CAT", "FMC")
POSTER_JPEG_QUALITY = 90
APPROVED_TEXT = PPT_TOOL / "examples_and_configs" / "configs" / "live" / "tgt" / "cat_fmc_approved_text.yaml"
FONT = "Calibri"   # the reference deck's theme font (build_biad2255_deck.py: "the deck's theme font (Calibri)")
# Title page and details slides copy the reference deck (FF: "base these on our recent cytokine ppt"; "titles need to
# be centered"). Boxes as measured on it with python-pptx (inches x 72 = points); the code that drew it:
# insert_channel_misalignment_slides.title_slide / text_slide / add_footer.
REFERENCE_DECK = Path(r"K:\FF\PPT\PPT_autogeneration\BIAD2255_smFISH\BIAD2255_smFISH_20261119.pptx")
# Title page: title 0.60/0.80 in, 12.13 x 3.30 in, 54 pt bold, centered, anchored bottom; subtitle at 4.40 in,
# 1.40 in tall, 32 pt grey, centered, anchored top.
TP_LEFT, TP_W, TP_TITLE_TOP, TP_TITLE_H, TP_TITLE_PT = 43.2, 873.4, 57.6, 237.6, 54.0
TP_SUB_TOP, TP_SUB_H, TP_SUB_PT = 316.8, 100.8, 32.0
# Details: title band 0.12/0.05 in, 13.09 x 1.30 in, bold, centered, anchored middle; 32 pt (repo rule: titles fixed
# 32 pt; the reference used 40). Bullets in one box at 0.52/1.59 in, 12.29 x 5.49 in, 26 pt (the reference's size),
# 18 pt after each. Caption 0.12/7.06 in, 13.09 x 0.40 in, 18 pt grey, right-aligned.
DT_LEFT, DT_W, DT_TITLE_TOP, DT_TITLE_H = 8.64, 942.5, 3.6, 93.6
BUL_LEFT, BUL_W, BUL_TOP, BUL_H, BUL_PT, BUL_GAP = 37.4, 884.9, 114.5, 395.3, 26.0, 18.0
CAP_TOP, CAP_H, CAP_PT = 508.3, 28.8, 18.0


def fields_title(day, fields):
    """'20260925 d5: W1 FOV1 and FOV2' (FF 2026-11-22: "in slide titles specific fovs should be indicated")."""
    (w1, p1), (w2, p2) = (f.split(":") for f in fields)
    second = f"FOV{int(p2)}" if w1 == w2 else f"{w2} FOV{int(p2)}"
    return f"{day.replace('_', ' ')}: {w1} FOV{int(p1)} and {second}"


def day_movies(movie_dir, day, ctype):
    """This cell type's movies, in order W1 FOV1 onward, then W2."""
    pat = re.compile(rf"{re.escape(day)}_{ctype}_W(\d+)_p(\d+)\.mp4$")
    hits = [(int(m[1]), int(m[2]), p) for p in movie_dir.glob(f"{day}_{ctype}_W*_p*.mp4") if (m := pat.match(p.name))]
    return [p for _, _, p in sorted(hits)]


def full_day_plan(cat, fmc, one_type="both_columns", start="left"):
    """List of slides; each slide is a list of columns (cell type, [movies top to bottom]); ("", []) = empty column.
    Fields left of one cell type when the other runs out (FF's open choices A and C):
      one_type "both_columns": 4 per slide, column by column; "right_only": 2 per slide in its usual column.
      start "left" | "right": with both_columns, the column filled first."""
    slides = []
    while cat and fmc:
        slides.append([("CAT", cat[:2]), ("FMC", fmc[:2])])
        cat, fmc = cat[2:], fmc[2:]
    for ctype, rest in (("CAT", cat), ("FMC", fmc)):
        if one_type == "right_only":
            for i in range(0, len(rest), 2):
                col = (ctype, rest[i:i + 2])
                slides.append([col, ("", [])] if ctype == "CAT" else [("", []), col])
            continue
        for i in range(0, len(rest), 4):
            chunk = rest[i:i + 4]
            cols = [(ctype, chunk[j:j + 2]) for j in range(0, len(chunk), 2)]
            if start == "right":
                cols = [cols[1] if len(cols) > 1 else ("", []), cols[0]]
            slides.append(cols)
    return slides


def day_plan(movies_root, day, version, layout=None, wells=None):
    """wells (optional): keep only these wells, e.g. ["FMC_W6"] (the flagged-wells deck, FF 2026-11-23)."""
    layout = layout or {}
    movie_dir = Path(movies_root) / version / day
    by_type = {c: day_movies(movie_dir, day, c) for c in CELL_TYPES}
    if wells:
        by_type = {c: [p for p in v if any(p.name.startswith(f"{day}_{w}_p") for w in wells)]
                   for c, v in by_type.items()}
    print(f"{movie_dir}: " + ", ".join(f"{c} {len(v)} movies" for c, v in by_type.items()))
    return full_day_plan(by_type["CAT"], by_type["FMC"], layout.get("one_type", "both_columns"),
                         layout.get("start", "left"))


FIG_MARGIN = 9.0   # pt around a figure slide's image (left, right, bottom)


def figure_slide(title, png, title_pt=TITLE_PT):
    """A metric figure slide: the details slides' centered title band, then one figure as large as fits below it
    (aspect kept, centered). No other text (FF 2026-11-22, relayed: "One figure per slide, as large as fits")."""
    png = Path(png)
    if not png.is_file():
        sys.exit(f"STOP: figure {png} not found")
    with Image.open(png) as im:
        ratio = im.width / im.height
    top = DT_TITLE_TOP + DT_TITLE_H
    room_w, room_h = SLIDE_W - 2 * FIG_MARGIN, SLIDE_H - top - FIG_MARGIN
    w = min(room_w, room_h * ratio)
    h = w / ratio
    title_box = TextboxSpec(title, DT_LEFT, DT_TITLE_TOP, DT_W, DT_TITLE_H, title_pt, align="center",
                            anchor="middle", font_name=FONT)
    return SlideSpec([title_box], [], [PictureSpec(str(png), (SLIDE_W - w) / 2, top + (room_h - h) / 2, w, h)])


def divider_slide(text):
    """A day divider (FF's open choice B): the day, 54 pt bold, centered on a plain slide."""
    return SlideSpec([TextboxSpec(text, TP_LEFT, 0, TP_W, SLIDE_H, TP_TITLE_PT, align="center", anchor="middle",
                                  font_name=FONT)], [])


def title_page(title, subtitle):
    return SlideSpec([
        TextboxSpec(title, TP_LEFT, TP_TITLE_TOP, TP_W, TP_TITLE_H, TP_TITLE_PT, align="center", anchor="bottom",
                    font_name=FONT),
        TextboxSpec(subtitle, TP_LEFT, TP_SUB_TOP, TP_W, TP_SUB_H, TP_SUB_PT, bold=False, align="center",
                    anchor="top", color_rgb=GREY_BGR, font_name=FONT)], [])


def wrapped_lines(text, pt, width_pt, font_path=r"C:\Windows\Fonts\calibri.ttf", insets_pt=14.4, slack=1.05):
    """Lines PowerPoint needs for text in a Calibri box of width_pt (greedy word wrap, measured with PIL;
    widths padded by slack so the estimate errs toward more lines)."""
    from PIL import ImageFont
    font = ImageFont.truetype(font_path, 1000)
    room = width_pt - insets_pt
    lines, cur = 1, ""
    for word in text.split():
        trial = f"{cur} {word}".strip()
        if font.getlength(trial) / 1000 * pt * slack > room and cur:
            lines, cur = lines + 1, word
        else:
            cur = trial
    return lines


def paragraph(b):
    """One line of a details slide's bullet box: a string is a bullet ("\u2022 " added); {heading: text} is a heading line
    with no bullet mark (FF 2026-11-22, slide 3 "Brightness, fixed for each movie")."""
    return b["heading"] if isinstance(b, dict) else "\u2022 " + b


def bullets_height(bullets, pt):
    """Estimated height (pt) of the bullet box's text at this size: wrapped lines x 1.2 line spacing + gaps."""
    paras = [paragraph(b) for b in bullets]
    return sum(wrapped_lines(p, pt, BUL_W) for p in paras) * pt * 1.2 + BUL_GAP * (len(paras) - 1)


def deck_bullet_pt(details):
    """One bullet size for every details slide, as in the reference deck: its 26 pt if every slide's bullets fit the
    box, else the largest size down to 22 pt that fits (24 pt is the reference's smallest; 22 pt added 2026-11-22
    when FF's slide 3 text needed it, reported to FF)."""
    for pt in (BUL_PT, 25.0, 24.0, 23.0, 22.0):
        if all(bullets_height(d["bullets"], pt) <= BUL_H for d in details):
            return pt
    sys.exit("STOP: details bullets do not fit their box even at 24 pt")


def details_slide(title, bullets, caption="", pt=BUL_PT, title_pt=TITLE_PT, reference=""):
    """The reference deck's text slide: centered title band, all bullets in one box (18 pt after each, so a wrapped
    bullet pushes the next one down), grey caption at the bottom right. reference: a short citation in italics at the
    bottom left (FF 2026-11-23), in the caption's band, size and grey."""
    boxes = [TextboxSpec(title, DT_LEFT, DT_TITLE_TOP, DT_W, DT_TITLE_H, title_pt, align="center", anchor="middle",
                         font_name=FONT)]
    boxes.append(TextboxSpec("\r".join(paragraph(b) for b in bullets), BUL_LEFT, BUL_TOP, BUL_W, BUL_H, pt,
                             bold=False, align="left", anchor="top", space_after_pt=BUL_GAP, font_name=FONT))
    if caption:
        boxes.append(TextboxSpec(caption, DT_LEFT, CAP_TOP, DT_W, CAP_H, CAP_PT, bold=False, align="right",
                                 color_rgb=GREY_BGR, font_name=FONT))
    if reference:
        boxes.append(TextboxSpec(reference, DT_LEFT, CAP_TOP, DT_W, CAP_H, CAP_PT, bold=False, align="left",
                                 color_rgb=GREY_BGR, font_name=FONT, italic=True))
    return SlideSpec(boxes, [])


def check_against_reference(prs, text_slides, ref_path=REFERENCE_DECK):
    """Title page (slide 0) and the text slides at 0-based indices text_slides vs the reference deck (its slide 1 and
    slide 5), measured with python-pptx: alignment and vertical anchor of titles, subtitle and caption; bold; font.
    Title size is not compared (repo rule 32 pt vs the reference's 40 pt). The italic reference line has no
    counterpart in the reference deck and is skipped. Returns one line per difference; empty when they match."""
    from pptx import Presentation
    ref = Presentation(str(ref_path))

    def boxes(slide):
        return [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]

    def fmt(sh):
        p = sh.text_frame.paragraphs[0]
        return {"align": p.alignment, "anchor": sh.text_frame.vertical_anchor, "bold": bool(p.runs[0].font.bold)}

    problems = []

    def compare(name, ours, theirs):
        a, b = fmt(ours), fmt(theirs)
        problems.extend(f"{name}: {k} {a[k]} (reference {b[k]})" for k in a if a[k] != b[k])
        fonts = {r.font.name for p in ours.text_frame.paragraphs for r in p.runs}
        if fonts != {FONT}:
            problems.append(f"{name}: font {sorted(map(str, fonts))} (reference theme font {FONT})")

    r1, o1 = boxes(ref.slides[0]), boxes(prs.slides[0])
    compare("title page title", o1[0], r1[0])
    compare("title page subtitle", o1[1], r1[1])
    r5 = boxes(ref.slides[4])
    for i in text_slides:
        o = [sh for sh in boxes(prs.slides[i]) if not sh.text_frame.paragraphs[0].runs[0].font.italic]
        compare(f"slide {i + 1} title", o[0], r5[0])
        if len(o) > 2:
            compare(f"slide {i + 1} caption", o[-1], r5[-1])
    return problems


class MovieSlideMaker:
    """Movie slides from (title, columns); posters (JPEG first frames) go in a local scratch folder."""

    def __init__(self, scratch):
        self.scratch = scratch
        self.movie_h = (ROWS_BOTTOM - ROWS_TOP - ROW_GAP) / 2
        self.movie_w = self.movie_h * MOVIE_RATIO
        self.left0 = (SLIDE_W - 2 * self.movie_w - COL_GAP) / 2
        self.fps_seen = set()

    def __call__(self, title, columns):
        boxes = [TextboxSpec(title, 0, TITLE_TOP, SLIDE_W, TITLE_H, TITLE_PT, font_name=FONT)]
        movies = []
        for col, (ctype, paths) in enumerate(columns):
            if not ctype:   # empty column
                continue
            left = self.left0 + col * (self.movie_w + COL_GAP)
            boxes.append(TextboxSpec(ctype, left, HEAD_TOP, self.movie_w, HEAD_H, HEAD_PT,
                                     color_rgb=COLOR_BGR[ctype], font_name=FONT))
            for row, mp4 in enumerate(paths):
                if not mp4.is_file():
                    sys.exit(f"STOP: {mp4} not found")
                frame = iio.imread(mp4, index=0)
                if abs(frame.shape[1] / frame.shape[0] - MOVIE_RATIO) > 0.01:
                    sys.exit(f"STOP: {mp4.name} is {frame.shape[1]} x {frame.shape[0]}, not the expected shape")
                self.fps_seen.add(iio.immeta(mp4).get("fps"))
                poster = self.scratch / f"{mp4.stem}_poster.jpg"
                Image.fromarray(frame).convert("RGB").save(poster, "JPEG", quality=POSTER_JPEG_QUALITY)
                movies.append(MovieSpec(str(mp4), str(poster), left, ROWS_TOP + row * (self.movie_h + ROW_GAP),
                                        self.movie_w, self.movie_h))
        return SlideSpec(boxes, movies)


def describe(columns, day):
    return " | ".join(f"{c} " + ", ".join(p.stem.split(f"{day}_")[-1] for p in paths) for c, paths in columns)


def config_plan(cfg, version="movie"):
    """Every slide of the all-days deck, in order: ("slide", SlideSpec) for the title page, details slides and day
    dividers, ("movie", title, columns, day) for movie slides. Movie titles count slides across the whole deck.
    cfg["layout"] (optional) holds FF's open layout choices: details_title_pt, one_type, start, day_dividers."""
    layout = cfg.get("layout") or {}
    tp = cfg["title_page"]
    entries = [("slide", title_page(tp["title"], tp["subtitle"]))]
    bullet_pt = deck_bullet_pt(cfg["details"])
    print(f"details bullets {bullet_pt:g} pt (reference deck 26 pt)")
    for d in cfg["details"]:
        entries.append(("slide", details_slide(d["title"], d["bullets"], d.get("caption", ""), bullet_pt,
                                               layout.get("details_title_pt", TITLE_PT), d.get("reference", ""))))
    for d in cfg["days"]:
        if layout.get("day_dividers"):
            entries.append(("slide", divider_slide(d["day"].replace("_", " "))))
        entries += [("movie", None, cols, d["day"])
                    for cols in day_plan(d["movies_root"], d["day"], version, layout, d.get("wells"))]
    for f in cfg.get("figures") or []:   # metric slides at the end (FF 2026-11-23: "quant results ... moved to end")
        entries.append(("slide", figure_slide(f["title"], f["png"])))
    n_total = len(entries)
    return [(e[0], f"{e[3].replace('_', ' ')} (slide {k} of {n_total})", e[2], e[3]) if e[0] == "movie" else e
            for k, e in enumerate(entries, 1)]


def build_entries(entries, out, approved):
    """Build the deck from entries, run the movie/autoplay check and the text checks; returns the saved deck
    (python-pptx Presentation)."""
    scratch = Path(make_run_scratch_dir(Path(tempfile.gettempdir()) / "tgt_cat_fmc_posters"))
    try:   # posters in the local temp folder (removed after the build), never next to the movies on Y:
        make = MovieSlideMaker(scratch)
        slides = [e[1] if e[0] == "slide" else make(e[1], e[2]) for e in entries]
        out.parent.mkdir(parents=True, exist_ok=True)
        n_autoplay = build_movie_deck_via_com(slides, str(out), SLIDE_W, SLIDE_H, loop=True)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    movies = [e for e in entries if e[0] == "movie"]
    n_movies = sum(len(s.movies) for s in slides)
    expected = sum(len(paths) for e in movies for _, paths in e[2])
    print(f"{out}: {len(slides)} slides ({len(slides) - len(movies)} text, {len(movies)} movie), {n_movies} movies of "
          f"{make.movie_w:.0f} x {make.movie_h:.0f} pt, fps {sorted(make.fps_seen)}, autoplay set on {n_autoplay}")
    for k, e in enumerate(entries, 1):
        if e[0] == "movie":
            print(f"  slide {k}: {e[1]} | {describe(e[2], e[3])}")
    if n_autoplay != len(movies) or n_movies != expected:
        sys.exit(f"CHECK FAIL: expected {expected} movies and autoplay on {len(movies)} slide(s)")
    print(f"CHECK {expected} movies, autoplay on {n_autoplay} movie slide(s): PASS")
    from pptx import Presentation as _P
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    n_pics_expected = sum(len(s.pictures) for s in slides)
    n_pics = sum(1 for sl in _P(str(out)).slides for sh in sl.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
    if n_pics != n_pics_expected:
        sys.exit(f"CHECK FAIL: {n_pics} still pictures in the deck, expected {n_pics_expected}")
    if n_pics_expected:
        print(f"CHECK {n_pics} figure picture(s): PASS")
    # Text checks report, they never stop a build (repo CLAUDE.md). lint_presentation skips slide 1 (title slide).
    from pptx import Presentation
    from ppt_image_inserter.approved_text import deck_strings, load_approved, unapproved
    from ppt_image_inserter.text_lint import lint_presentation
    prs = Presentation(str(out))
    problems = lint_presentation(prs) + unapproved(deck_strings(prs), load_approved(approved))
    print(f"TEXT CHECK: {len(problems)} problem(s)" + "".join(f"\n  {p}" for p in problems))
    return prs


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = ap.add_mutually_exclusive_group(required=True)
    mode.add_argument("--config", type=Path, help="all-days deck config (YAML)")
    mode.add_argument("--fields", nargs=2, help="two well:position pairs, e.g. W1:1 W1:2")
    mode.add_argument("--full-day", action="store_true", help="every movie of the day, 2 per cell type per slide")
    ap.add_argument("--movies-root", type=Path, help="--fields / --full-day: render output folder")
    ap.add_argument("--day", help="--fields / --full-day: e.g. 20260925_d5")
    ap.add_argument("--version", default="movie", choices=("movie", "day", "frame"))
    ap.add_argument("--out", type=Path, help="deck path (--config: overrides the config's output)")
    ap.add_argument("--slides", help="--config: build only these deck slide numbers, e.g. 1,2,3,4,5,16 (previews)")
    args = ap.parse_args()

    approved = APPROVED_TEXT
    if args.config:
        cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
        out = args.out or Path(cfg["output"])
        approved = args.config.parent / cfg["approved_text"]
    else:
        if not (args.movies_root and args.day and args.out):
            sys.exit("STOP: --fields and --full-day need --movies-root, --day and --out")
        out = args.out
    if DECK_DIR not in out.resolve().parents:
        sys.exit(f"STOP: decks go under {DECK_DIR} (PPT_image_inserter output rule), not {out}")
    if out.exists():
        sys.exit(f"STOP: {out} exists and is never overwritten")

    # deck entries in order: ("slide", SlideSpec) or ("movie", title, columns, day)
    if args.config:
        entries = config_plan(cfg, args.version)
        if args.slides:   # preview deck: only these deck slides, titles numbered as in the full deck
            keep = {int(n) for n in args.slides.split(",")}
            entries = [e for k, e in enumerate(entries, 1) if k in keep]
    elif args.full_day:
        plan = day_plan(args.movies_root, args.day, args.version)
        entries = [("movie", f"{args.day.replace('_', ' ')} (slide {k} of {len(plan)})", cols, args.day)
                   for k, cols in enumerate(plan, 1)]
    else:
        movie_dir = args.movies_root / args.version / args.day
        cols = [(c, [movie_dir / f"{args.day}_{c}_{w}_p{int(p):02d}.mp4" for w, p in (f.split(":") for f in args.fields)])
                for c in CELL_TYPES]
        entries = [("movie", fields_title(args.day, args.fields), cols, args.day)]

    prs = build_entries(entries, out, approved)
    n_lead = 1 + len(cfg["details"]) if args.config else 0   # title page, details
    n_fig = len(cfg.get("figures") or []) if args.config else 0   # metric slides, last in the deck
    text_slides = list(range(1, n_lead)) + list(range(len(entries) - n_fig, len(entries)))
    if args.config and all(entries[i][0] == "slide" for i in [0] + text_slides):
        # figure slides: only their title is compared (same band as details)
        fmt_problems = check_against_reference(prs, text_slides)
        print(f"FORMAT CHECK vs {REFERENCE_DECK.name}: {len(fmt_problems)} difference(s)"
              + "".join(f"\n  {p}" for p in fmt_problems))
    print(f"deck size {out.stat().st_size / 2**20:.1f} MB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
