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
Incremental builds (2026-11-23, FF: "how does it take 15 min to make this tiny change?"; "improve the infra to speed up
changes broadly"): --from <existing deck> builds only the slides whose spec changed (text, values, notes, marks, slides
added, removed or moved) and splices them into a copy of that deck (ppt_image_inserter.incremental.update_deck); every
other slide is kept byte for byte and checked by content hash. Needs the deck's sidecar <deck>.slides.json, written by
every full build of this script (or --seed <deck>, which checks an older deck against its specs first). Every build
logs each stage's time (STAGE lines and a STAGE TIMES table).
    python examples_and_configs	gtuild_cat_fmc_slides.py --config <yaml> --pairs --from <live deck> --out <live>_building.pptx
    python -m ppt_image_inserter.incremental swap <live>_building.pptx <live deck>     (after the checks pass)
Moved from TGT_analysis\\scripts\\first_look 2026-11-22 (FF: "the script needs to be in ppt-inserter / period").
Run from the PPT_image_inserter repo root (env PPT_editing):
    python examples_and_configs\\tgt\\build_cat_fmc_slides.py --config examples_and_configs\\configs\\live\\tgt\\cat_fmc_sept_2026_deck.yaml
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
from ppt_image_inserter import incremental  # noqa: E402
from ppt_image_inserter.stage_timer import StageTimer  # noqa: E402

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
WORKERS = 12   # threads reading movies (posters, movie checks); the global cap is 12; 1 = serial
POSTER_RULE = f"poster: first frame of the movie, JPEG q{POSTER_JPEG_QUALITY}"   # in every slide hash (incremental)
# Expected seconds per build stage, for the FAR OVER flag (StageTimer). Only measured values go here, each with its
# run; a stage not listed logs "expected: unmeasured".
# Measured 2026-11-23 on the 119-slide movie-pairs deck (398 movies, 481 MB, deck on K:): full build
# tests\incrementalull_build2.log (345.6 s); incremental tests t0 / t1 / t2 .log (40 / 43 / 88 s); seed 19 s.
# (base, per_item) = base + per_item x the first number in the stage name (StageTimer).
EXPECTED_STAGE_S = {
    "plan": 1.0, "slide specs": 1.0,
    "posters (": (0.0, 0.035),                       # per movie: 13.2 s / 398 (full), 2.9 s / 80 (t2)
    "PowerPoint build and save (": (2.0, 0.70),      # per movie: 277.4 s / 398 (full); 1 text slide 2.1 s (t1)
    "autoplay": 23.0,                                # one rewrite of the 481 MB file: 22.2 s
    "check source pictures": 2.0, "write sidecar": 2.5, "read sidecar": 1.2, "hash slide specs": 1.0,
    "splice into the existing deck and save": 22.0,  # 19.7-22.2 s, mostly writing 481 MB to K:
    "check splice": 5.0, "picture checks": 3.0, "text, format and value checks": 5.0,
    "deck checks (": (0.6, 0.18),                    # per slide whose movies are read: 18.4 s / 101, 4.8 s / 20
    "check deck against specs": 18.0,                # seed: 16-18 s (md5 of every movie)
}
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
# Removed-FOV mark under a movie (proposal shown to FF 2026-11-23, not yet picked): one line, red.
MARK_H, MARK_PT, MARK_BGR = 22.0, 16.0, 0x0000C0   # RGB (192, 0, 0) as BGR


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


LABEL_W_FRAC = 0.10   # each CAT / FMC label box, as a fraction of the picture width


def figure_slide(title, png, title_pt=TITLE_PT, labels_at=None, note=""):
    """A metric figure slide: the details slides' centered title band, then one figure as large as fits below it
    (aspect kept, centered) (FF 2026-11-22, relayed: "One figure per slide, as large as fits"). labels_at (x, y):
    "CAT" and "FMC" text boxes side by side, as the movie slides' headers (24 pt bold, same colors), their left edge
    and top at these fractions of the picture (FF 2026-11-23: legend "in colored text in ppt not part of figure").
    note: one line in the reference deck's caption box (18 pt grey, right-aligned, bottom); the picture then fits
    above the box, so the two never overlap (FF 2026-11-23, relayed). A list of lines (the low/high slides' second
    line, FF 2026-11-23 "ok on 5 and 6") grows the box upward one caption height per extra line; a line that would
    wrap stops the build."""
    lines = [note] if isinstance(note, str) else list(note)
    lines = [ln for ln in lines if ln]
    for ln in lines:
        if wrapped_lines(ln, CAP_PT, DT_W) > 1:
            sys.exit(f"STOP: note line wraps at {CAP_PT:g} pt: {ln!r}")
    note_top = CAP_TOP - CAP_H * (len(lines) - 1)
    png = Path(png)
    if not png.is_file():
        sys.exit(f"STOP: figure {png} not found")
    with Image.open(png) as im:
        ratio = im.width / im.height
    top = DT_TITLE_TOP + DT_TITLE_H
    bottom = note_top if lines else SLIDE_H - FIG_MARGIN
    room_w, room_h = SLIDE_W - 2 * FIG_MARGIN, bottom - top
    w = min(room_w, room_h * ratio)
    h = w / ratio
    title_box = TextboxSpec(title, DT_LEFT, DT_TITLE_TOP, DT_W, DT_TITLE_H, title_pt, align="center",
                            anchor="middle", font_name=FONT)
    pic = PictureSpec(str(png), (SLIDE_W - w) / 2, top + (room_h - h) / 2, w, h)
    boxes = [title_box]
    if labels_at:
        lw = LABEL_W_FRAC * w
        for k, ctype in enumerate(CELL_TYPES):
            boxes.append(TextboxSpec(ctype, pic.left_pt + labels_at[0] * w + k * lw, pic.top_pt + labels_at[1] * h,
                                     lw, HEAD_H, HEAD_PT, color_rgb=COLOR_BGR[ctype], font_name=FONT))
    if lines:
        boxes.append(TextboxSpec("\r".join(lines), DT_LEFT, note_top, DT_W, CAP_H * len(lines), CAP_PT, bold=False,
                                 align="right", color_rgb=GREY_BGR, font_name=FONT))
    return SlideSpec(boxes, [], [pic])


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


# Sub-bullets (FF 2026-11-23: "subbullets. less flat"; "use subbullets not - this should be case everywhere"): a bullet
# mark like the main bullets, never a dash, 4 pt smaller, indented under the main bullet's text, 6 pt after a line
# that is followed by a sub-bullet. On a slide with sub-bullets the main bullets are bold and the sub-bullets are not
# (FF 2026-11-23: "bold stuff in main bllets -but not in subbelets"). The reference deck has no sub-bullets.
SUB_MARK, SUB_DROP_PT, SUB_GAP = "\u2022 ", 4.0, 6.0
FONT_FILE, FONT_FILE_BOLD = r"C:\Windows\Fonts\calibri.ttf", r"C:\Windows\Fonts\calibrib.ttf"


def paragraphs(b):
    """The lines of one details-slide entry as (text, level). A string is a bullet ("\u2022 " added); {heading: text}
    is a heading line with no mark (FF 2026-11-22, slide 3 "Brightness, fixed for each movie"); {text, sub: [...]} is
    a bullet with sub-bullets one level down (SUB_MARK added)."""
    if isinstance(b, str):
        return [("\u2022 " + b, 0)]
    if "heading" in b:
        return [(b["heading"], 0)]
    return [("\u2022 " + b["text"], 0)] + [(SUB_MARK + t, 1) for t in b.get("sub") or []]


def flat_paragraphs(bullets):
    return [p for b in bullets for p in paragraphs(b)]


def has_sub(bullets):
    """True when the slide has sub-bullets: then its main bullets are bold."""
    return any(level == 1 for _, level in flat_paragraphs(bullets))


def bullets_height(bullets, pt):
    """Estimated height (pt) of the bullet box's text at this size: wrapped lines x 1.2 line spacing (main bullets in
    bold when the slide has sub-bullets; sub-bullets at pt - SUB_DROP_PT, in the box width less the bullet indent),
    plus the space after each line (SUB_GAP before a sub-bullet, else BUL_GAP)."""
    paras = flat_paragraphs(bullets)
    bold = has_sub(bullets)
    h = 0.0
    for k, (text, level) in enumerate(paras):
        size = pt - SUB_DROP_PT * level
        font = FONT_FILE_BOLD if bold and level == 0 else FONT_FILE
        h += wrapped_lines(text, size, BUL_W - bullet_indent_pt(pt) * level, font_path=font) * size * 1.2
        if k + 1 < len(paras):
            h += SUB_GAP if paras[k + 1][1] == 1 else BUL_GAP
    return h


def bullet_indent_pt(pt, mark="• ", font_path=FONT_FILE):
    """Width (pt) of a typed bullet mark at this size (bold font file for a bold main bullet): the hanging indent of a
    bullet paragraph, so its wrapped lines start under the text (figure review 2026-11-23: wrapped line started under
    the mark)."""
    from PIL import ImageFont
    return ImageFont.truetype(font_path, 100).getlength(mark) * pt / 100


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
    paras, bold_main = flat_paragraphs(bullets), has_sub(bullets)
    boxes.append(TextboxSpec("\r".join(t for t, _ in paras), BUL_LEFT, BUL_TOP, BUL_W, BUL_H, pt,
                             bold=False, align="left", anchor="top", space_after_pt=BUL_GAP, font_name=FONT,
                             hanging_indent_pt=bullet_indent_pt(pt, font_path=FONT_FILE_BOLD if bold_main else FONT_FILE),
                             para_levels=[level for _, level in paras],
                             sub_indent_pt=bullet_indent_pt(pt - SUB_DROP_PT, SUB_MARK),
                             sub_font_size_pt=pt - SUB_DROP_PT, sub_space_after_pt=SUB_GAP,
                             main_bold=True if bold_main else None))
    if caption:
        boxes.append(TextboxSpec(caption, DT_LEFT, CAP_TOP, DT_W, CAP_H, CAP_PT, bold=False, align="right",
                                 color_rgb=GREY_BGR, font_name=FONT))
    if reference:
        boxes.append(TextboxSpec(reference, DT_LEFT, CAP_TOP, DT_W, CAP_H, CAP_PT, bold=False, align="left",
                                 color_rgb=GREY_BGR, font_name=FONT, italic=True))
    return SlideSpec(boxes, [])


def check_bullet_format(prs):
    """FF's slide-text rules (2026-11-23), checked on every slide; one line per problem, empty when clean:
    no dash ("-", "–", "—") used as a bullet mark ("use subbullets not -"); in a text box with sub-bullets
    (paragraphs indented further than the box's least-indented bullet), main bullets bold and sub-bullets not bold
    ("bold stuff in main bllets -but not in subbelets"); no "field" in any form or case ("say FOVs never ay Fields of
    View"; the director, relayed 2026-11-23: every "field" that means a FOV changes)."""
    import re
    problems = []
    for n, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            if re.search(r"field", sh.text_frame.text, re.I):
                problems.append(f"slide {n}: 'field' in {sh.text_frame.text[:60]!r}")
            paras = [p for p in sh.text_frame.paragraphs if p.text.strip()]
            for p in paras:
                if re.match(r"\s*[-–—]\s", p.text):
                    problems.append(f"slide {n}: dash used as a bullet mark: {p.text[:50]!r}")
            bul = [(int(p._p.pPr.get("marL", 0)) if p._p.pPr is not None else 0, p) for p in paras
                   if p.text.startswith("•")]
            if len({m for m, _ in bul}) < 2:
                continue   # no sub-bullets in this box
            top = min(m for m, _ in bul)
            for m, p in bul:
                bold = [r.font.bold for r in p.runs if r.text.strip()]
                if m == top and not all(b is True for b in bold):
                    problems.append(f"slide {n}: main bullet not bold: {p.text[:50]!r}")
                if m > top and any(b is True for b in bold):
                    problems.append(f"slide {n}: sub-bullet bold: {p.text[:50]!r}")
    return problems


def check_figure_labels(prs, fig_slides):
    """Each figure slide (0-based indices) has one "CAT" and one "FMC" box, inside its picture and not below the
    picture's bottom edge (FF 2026-11-23: "where there is space in figs not below figures"). One line per problem."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    problems = []
    for i in fig_slides:
        shapes = list(prs.slides[i].shapes)
        pics = [s for s in shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
        if len(pics) != 1:
            problems.append(f"slide {i + 1}: {len(pics)} pictures, expected 1")
            continue
        p = pics[0]
        for ctype in CELL_TYPES:
            hits = [s for s in shapes if s.has_text_frame and s.text_frame.text.strip() == ctype]
            if len(hits) != 1:
                problems.append(f"slide {i + 1}: {len(hits)} '{ctype}' boxes, expected 1")
                continue
            b = hits[0]
            if not (b.left >= p.left and b.left + b.width <= p.left + p.width and b.top >= p.top
                    and b.top + b.height <= p.top + p.height):
                problems.append(f"slide {i + 1}: '{ctype}' box is not inside the picture")
            if shapes.index(b) < shapes.index(p):
                problems.append(f"slide {i + 1}: '{ctype}' box is behind the picture (hidden)")
        for s in shapes:   # the note (any other box below the title band) must not overlap the picture
            if s.has_text_frame and s.text_frame.text.strip() not in CELL_TYPES and s.top > p.top:
                if s.top < p.top + p.height:
                    problems.append(f"slide {i + 1}: note '{s.text_frame.text[:30]}' overlaps the picture")
    return problems


def check_against_reference(prs, text_slides, ref_path=REFERENCE_DECK, title_only=()):
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
        if len(o) > 2 and i not in title_only:   # figure slides: their CAT / FMC labels are no caption
            compare(f"slide {i + 1} caption", o[-1], r5[-1])
    return problems


class MovieSlideMaker:
    """Movie slides from (title, columns); posters (JPEG first frames) go in a local scratch folder, written by
    write_posters only for the slides that are built (an incremental build reads no frame of an unchanged slide)."""

    def __init__(self, scratch, removed=None):
        """removed: {movie stem: mark text}. When given, every movie gets a line of room under it (all slides keep
        one movie size) and each removed movie its mark there, in red (FF 2026-11-23: "on the deck slides we need
        to highlight removed FOVs"; the mark itself is FF's pick)."""
        self.scratch = scratch
        self.removed = removed or {}
        self.mark_h = MARK_H if self.removed else 0.0
        self.movie_h = (ROWS_BOTTOM - ROWS_TOP - ROW_GAP - 2 * self.mark_h) / 2
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
                # version folder in the name: a movie and its per-frame copy share a stem (movie-pairs deck)
                poster = self.scratch / f"{mp4.parent.parent.name}_{mp4.stem}_poster.jpg"
                top = ROWS_TOP + row * (self.movie_h + self.mark_h + ROW_GAP)
                movies.append(MovieSpec(str(mp4), str(poster), left, top, self.movie_w, self.movie_h))
                if mp4.stem in self.removed:
                    boxes.append(TextboxSpec(self.removed[mp4.stem], left - COL_GAP / 2, top + self.movie_h,
                                             self.movie_w + COL_GAP, self.mark_h, MARK_PT, bold=False,
                                             color_rgb=MARK_BGR, font_name=FONT,
                                             subscripts=subscript_ranges(self.removed[mp4.stem])))
        return SlideSpec(boxes, movies)

    def write_posters(self, slides, workers=None):
        """Write the poster (JPEG first frame) of every movie on these slides, in `workers` threads (default WORKERS;
        1 = serial); stops on a movie of the wrong shape. Records each movie's fps (fps_seen)."""
        from concurrent.futures import ThreadPoolExecutor
        workers = WORKERS if workers is None else workers

        def poster(mv):
            mp4 = Path(mv.movie_path)
            frame = iio.imread(mp4, index=0)
            if abs(frame.shape[1] / frame.shape[0] - MOVIE_RATIO) > 0.01:
                return f"STOP: {mp4.name} is {frame.shape[1]} x {frame.shape[0]}, not the expected shape", None
            Image.fromarray(frame).convert("RGB").save(mv.poster_path, "JPEG", quality=POSTER_JPEG_QUALITY)
            return None, iio.immeta(mp4).get("fps")

        mvs = [mv for s in slides for mv in s.movies]
        if workers > 1 and len(mvs) > 1:
            with ThreadPoolExecutor(min(workers, len(mvs))) as ex:
                results = list(ex.map(poster, mvs))
        else:
            results = [poster(mv) for mv in mvs]
        for err, fps in results:
            if err:
                sys.exit(err)
            self.fps_seen.add(fps)


def describe(columns, day):
    return " | ".join(f"{c} " + ", ".join(p.stem.split(f"{day}_")[-1] for p in paths) for c, paths in columns)


# Movie-pairs deck (tgt_cat_fmc_sept_2026_movie_pairs.pptx): FF 2026-11-23, relayed: "every slide should have just one
# pair of movies. one is just like the current one. the second should employ per-frame scaling below - and we need to
# indicate that"; "on eevery slide now: the scaling must be clearly stated"; go: "Yes, build it". Key wording: the
# director's, shown to FF, who did not object; not approved verbatim. Example: review\perframe_example (PASS).
PAIR_KEY = ("Top: TGT per-movie scaling", "Bottom: TGT per-frame scaling")
PAIR_KEY_PT, PAIR_KEY_PITCH = 20.0, 24.0   # above the number labels' 16 pt, bold, black


def frame_movie(mp4):
    """The per-frame version of a deck movie (render_cat_fmc_movies.py, same run): <root>/movie/<day>/x.mp4 ->
    <root>/frame/<day>/x.mp4. TGT low and high = 1st and 99.9th percentile of that frame; IRM unchanged."""
    return mp4.parents[2] / "frame" / mp4.parent.name / mp4.name


def pair_entries(e):
    """One movie entry (two rows) -> one entry per row: each column holds [movie, its per-frame movie]."""
    n_rows = max(len(ps) for _, ps in e[2])
    out = []
    for r in range(n_rows):
        cols = [(c, [ps[r], frame_movie(ps[r])]) if c and len(ps) > r else ("", []) for c, ps in e[2]]
        if any(ps for _, ps in cols):
            out.append((e[0], None, cols, e[3]))
    return out


def one_type(e):
    """True for a movie-pairs entry whose movies are all of one cell type."""
    return e[0] == "movie" and len({c for c, ps in e[2] if ps}) == 1


def repair_one_type(entries):
    """Slides with one cell type only: consecutive FOVs of the same well paired, lower FOV left; a well's odd last FOV
    alone, on the left (FF 2026-11-23, relayed: "Pair consecutive"; the director: "A pair never mixes two wells",
    single movie "on the left on every such slide"). Per day, the one-type slides form one block, replaced in place."""
    out, k = [], 0
    while k < len(entries):
        e = entries[k]
        if not one_type(e):
            out.append(e)
            k += 1
            continue
        j = k
        while j < len(entries) and one_type(entries[j]) and entries[j][3] == e[3]:
            j += 1
        block = [(c, ps) for x in entries[k:j] for c, ps in x[2] if ps]
        key = lambda cp: (cp[0], int(re.search(r"_W(\d+)_p(\d+)$", cp[1][0].stem)[1]),   # noqa: E731
                          int(re.search(r"_W(\d+)_p(\d+)$", cp[1][0].stem)[2]))
        block.sort(key=key)
        wells = {}
        for cp in block:
            wells.setdefault(key(cp)[:2], []).append(cp)
        for w in sorted(wells):
            fovs = wells[w]
            for i in range(0, len(fovs), 2):
                pair = fovs[i:i + 2]
                out.append(("movie", None, pair + [("", [])] * (2 - len(pair)), e[3]))
        k = j
    return out


def check_one_type(entries):
    """Every one-type movie slide: two consecutive FOVs of one well (lower left), or one FOV on the left that is its
    well's last one-type FOV. Returns problem lines."""
    problems = []
    for k, e in enumerate(entries, 1):
        if not one_type(e):
            continue
        cols = [(c, ps) for c, ps in e[2]]
        ids = [re.search(r"_(CAT|FMC)_W(\d+)_p(\d+)$", ps[0].stem).groups() for c, ps in cols if ps]
        if len(ids) == 2:
            (c1, w1, p1), (c2, w2, p2) = ids
            if not (c1 == c2 and w1 == w2 and int(p2) == int(p1) + 1):
                problems.append(f"slide {k}: {ids} not two consecutive FOVs of one well")
        elif not cols[0][1]:
            problems.append(f"slide {k}: single movie not on the left")
    return problems


def pair_slide(spec, make, rows, removed, columns):
    """A movie-pairs slide from MovieSlideMaker output (row 1: the movies, row 2: their per-frame versions): row 1 in
    the deck's numbers layout (fov_numbers.numbers_slide) with its numbers and red marks, row 2 the same columns with
    no numbers or marks, and the scaling key in row 2's empty middle column. Returns (SlideSpec, values shown)."""
    import fov_numbers as fn
    from dataclasses import replace
    y0 = min(m.top_pt for m in spec.movies)
    boxes = [tb for tb in spec.textboxes if not (tb.color_rgb == MARK_BGR and tb.top_pt > y0 + make.movie_h + 1)]
    s, shown = fn.numbers_slide(SlideSpec(boxes, [m for m in spec.movies if m.top_pt == y0], spec.pictures), make,
                                rows, removed, columns)
    w, h, dy = make.movie_w, make.movie_h, fn.ROWS_TOP - ROWS_TOP
    lefts = [fn.EDGE, SLIDE_W - fn.EDGE - w]
    row2 = [replace(m, left_pt=lefts[0 if m.left_pt < SLIDE_W / 2 else 1], top_pt=m.top_pt + dy)
            for m in spec.movies if m.top_pt != y0]
    gap_l, gap_r = fn.EDGE + w + fn.CLEAR, SLIDE_W - fn.EDGE - w - fn.CLEAR
    for ln in PAIR_KEY:
        if wrapped_lines(ln, PAIR_KEY_PT, gap_r - gap_l, font_path=r"C:\Windows\Fonts\calibrib.ttf", insets_pt=0.0) > 1:
            sys.exit(f"STOP: key line {ln!r} does not fit one line at {PAIR_KEY_PT:g} pt bold")
    key_h = PAIR_KEY_PITCH * len(PAIR_KEY)
    row2_top = fn.ROWS_TOP + h + make.mark_h + ROW_GAP
    key = fn.box("\r".join(PAIR_KEY), gap_l, row2_top + (h - key_h) / 2, gap_r - gap_l, key_h, PAIR_KEY_PT, "center",
                 bold=True, pitch=PAIR_KEY_PITCH)
    return SlideSpec(tuple(s.textboxes) + (key,), tuple(s.movies) + tuple(row2), s.pictures), shown


def _pair_file_checks(k, want_top, want_bot, vids):
    """check_pairs, the part that reads the movie files (run in a thread per slide): frame counts top / bottom, the
    embedded movies byte-identical to the files, each still image its own movie's first frame."""
    import hashlib
    import io
    import numpy as np
    md5 = lambda b: hashlib.md5(b).hexdigest()   # noqa: E731
    problems = []
    for a, b_ in zip(want_top, want_bot):
        if a.name != b_.name or b_.parent.parent.name != "frame" or a.parent.parent.name == "frame":
            problems.append(f"slide {k + 1}: {b_} is not the per-frame movie of {a}")
        elif sum(1 for _ in iio.imiter(a)) != sum(1 for _ in iio.imiter(b_)):
            problems.append(f"slide {k + 1}: {a.name}: frame counts differ top / bottom")
    y_top = vids[0][0] if vids else 0
    got_top = [v[2] for v in vids if v[0] == y_top]
    got_bot = [v[2] for v in vids if v[0] != y_top]
    if got_top != [md5(p.read_bytes()) for p in want_top] or got_bot != [md5(p.read_bytes()) for p in want_bot]:
        problems.append(f"slide {k + 1}: embedded movies differ from the files (top {len(got_top)}, "
                        f"bottom {len(got_bot)})")
    # each still image shown before play is its own movie's first frame (2026-11-23: a movie and its per-frame
    # copy shared a poster file name, so the top row showed the per-frame poster)
    for v, p in zip([v for v in vids if v[0] == y_top] + [v for v in vids if v[0] != y_top], want_top + want_bot):
        shown_img = np.asarray(Image.open(io.BytesIO(v[3])).convert("L"), float) if v[3] else None
        first = np.asarray(Image.fromarray(iio.imread(p, index=0)).convert("L"), float)
        if shown_img is None or shown_img.shape != first.shape or np.abs(shown_img - first).mean() > 3.0:
            problems.append(f"slide {k + 1}: the still image of {p.parent.parent.name}/{p.name} is not its "
                            "first frame")
    return problems


def check_pairs(prs, entries, only=None, workers=WORKERS):
    """Movie-pairs deck: every FOV once on top and once below, the one below its own per-frame movie (same name, from
    a frame folder, same frame count), each slide's embedded movies byte-identical to those files (top row = the
    movies, bottom row = the per-frame movies), the key on every movie slide. Returns (counts, problem lines).
    only: 0-based slides whose movies are checked (an incremental build: the slides it built; the slides it kept are
    byte-identical to the deck checked before); None = every movie slide. The movie files are read in `workers`
    threads, one slide each (1 = serial)."""
    import hashlib
    from concurrent.futures import ThreadPoolExecutor
    problems, tops, bottoms, jobs = [], [], [], []
    md5 = lambda b: hashlib.md5(b).hexdigest()   # noqa: E731
    movie_slides = [(k, e) for k, e in enumerate(entries) if e[0] == "movie"]
    n_key = 0
    for k, e in movie_slides:
        want_top = [ps[0] for _, ps in e[2] if ps]
        want_bot = [ps[1] for _, ps in e[2] if ps]
        tops += want_top
        bottoms += want_bot
        slide = prs.slides[k]
        if sum(1 for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text == "\n".join(PAIR_KEY)) == 1:
            n_key += 1
        else:
            problems.append(f"slide {k + 1}: scaling key not there exactly once")
        if only is not None and k not in only:
            continue
        vids = []   # (top, left, md5 of the embedded movie, still image bytes), read here, in the main thread
        for sh in slide.shapes:
            link = sh._element.xpath(".//a:videoFile/@r:link")
            if link:
                blip = sh._element.xpath(".//a:blip/@r:embed")
                vids.append((sh.top, sh.left, md5(slide.part.related_part(link[0]).blob),
                             slide.part.related_part(blip[0]).blob if blip else b""))
        vids.sort(key=lambda v: (v[0], v[1]))
        jobs.append((k, want_top, want_bot, vids))
    if workers > 1 and len(jobs) > 1:
        with ThreadPoolExecutor(min(workers, len(jobs))) as ex:
            results = list(ex.map(lambda j: _pair_file_checks(*j), jobs))
    else:
        results = [_pair_file_checks(*j) for j in jobs]
    problems += [p for r in results for p in r]
    stems_top, stems_bot = [p.stem for p in tops], [p.stem for p in bottoms]
    if len(set(stems_top)) != len(stems_top) or sorted(stems_top) != sorted(stems_bot):
        problems.append("a FOV is on top more than once, or the top and bottom FOVs differ")
    return {"movie slides": len(movie_slides), "FOVs on top": len(set(stems_top)), "FOVs below": len(set(stems_bot)),
            "key": n_key, "movies checked on slides": len(jobs)}, problems


def movie_notes_for(entry, notes):
    """The config's movie_notes that go on this movie slide: its day matches and a column of that cell type has a
    movie (FF 2026-11-23, relayed: "'20260924 d4' the cat fov here look very weird lots of punctate tgt. this needs to
    be noted"; place: "On d4 movie slides")."""
    return [n for n in (notes or []) if entry[3] == n["day"] and any(c == n["ctype"] and ps for c, ps in entry[2])]


def movie_note_box(text, make, numbers_layout):
    """The note in the red-mark slot under the bottom movie of the left column (empty on a day with no removed FOV),
    grey, one line at the mark size; the same position on every slide."""
    import fov_numbers as fn
    top = (fn.ROWS_TOP if numbers_layout else ROWS_TOP) + 2 * make.movie_h + make.mark_h + ROW_GAP
    left0 = fn.EDGE if numbers_layout else make.left0
    width = fn.MARK_W if numbers_layout else make.movie_w + COL_GAP
    left = min(max(left0 + make.movie_w / 2 - width / 2, 0.0), SLIDE_W - width)
    if wrapped_lines(text, MARK_PT, width, insets_pt=0.0) > 1:
        sys.exit(f"STOP: movie note {text!r} does not fit one line at {MARK_PT:g} pt")
    return TextboxSpec(text, left, top, width, make.mark_h or MARK_H, MARK_PT, bold=False, color_rgb=GREY_BGR,
                       font_name=FONT, anchor="top", side_margin_pt=0.0, vert_margin_pt=0.0)


def config_plan(cfg, version="movie", pairs=False):
    """Every slide of the all-days deck, in order: ("slide", SlideSpec) for the title page, details slides and day
    dividers, ("movie", title, columns, day) for movie slides. Movie titles count that day's movie slides.
    cfg["layout"] (optional) holds FF's open layout choices: details_title_pt, one_type, start, day_dividers."""
    layout = cfg.get("layout") or {}
    tp = cfg["title_page"]
    entries = [("slide", title_page(tp["title"], tp["subtitle"]))]
    figs = cfg.get("figures") or []
    for f in figs:
        if not (f.get("title") or "").strip():
            sys.exit(f"STOP: metric slide {f.get('png', '')} has no title yet (FF picks it)")
        if "png" in f and not f["png"]:
            sys.exit(f"STOP: metric slide '{f['title']}' has no figure path yet")
        if "png" not in f and not f.get("bullets"):
            sys.exit(f"STOP: text slide '{f['title']}' has no approved body yet")
    for d in cfg["details"]:
        if not d.get("bullets"):
            sys.exit(f"STOP: text slide '{d['title']}' has no approved body yet")
    bullet_pt = deck_bullet_pt(list(cfg["details"]) + [f for f in figs if "png" not in f])
    print(f"details bullets {bullet_pt:g} pt (reference deck 26 pt)")
    details = cfg["details"]
    if pairs:   # movie-pairs deck: a details entry with pairs_split becomes those slides, at the deck's bullet size
        details = [x for d in details for x in (d.get("pairs_split") or [d])]
        for d in details:
            if bullets_height(d["bullets"], bullet_pt) > BUL_H:
                sys.exit(f"STOP: text slide '{d['title']}' does not fit its box at {bullet_pt:g} pt")
    for d in details:
        entries.append(("slide", details_slide(d["title"], d["bullets"], d.get("caption", ""), bullet_pt,
                                               layout.get("details_title_pt", TITLE_PT), d.get("reference", ""))))
    for d in cfg["days"]:
        if layout.get("day_dividers"):
            entries.append(("slide", divider_slide(d["day"].replace("_", " "))))
        entries += [("movie", None, cols, d["day"])
                    for cols in day_plan(d["movies_root"], d["day"], version, layout, d.get("wells"))]
    if pairs:   # one slide per row of the current movie slides: (column, [movie, its per-frame movie])
        entries = [x for e in entries for x in ([e] if e[0] != "movie" else pair_entries(e))]
        entries = repair_one_type(entries)
    for f in figs:   # metric slides at the end (FF 2026-11-23: "quant results ... moved to end")
        if "png" in f:
            # labels_at per figure (e.g. the per-well figures, laid out unlike the per-day ones), else the deck's
            entries.append(("slide", figure_slide(f["title"], f["png"],
                                                  labels_at=f.get("labels_at", cfg.get("figure_labels_at")),
                                                  note=f.get("note", ""))))
        else:   # a text slide among the metric slides (e.g. the definitions slide)
            entries.append(("slide", details_slide(f["title"], f["bullets"], f.get("caption", ""), bullet_pt,
                                                   layout.get("details_title_pt", TITLE_PT), f.get("reference", ""))))
    # Movie titles count that day's movie slides (FF 2026-11-23, relayed: "it should be (slide X of Y) where Y is the
    # number of slides for that particular day"), as the d5 deck read on 11-22: "20260925 d5 (slide 1 of 11)".
    per_day = {}
    for e in entries:
        if e[0] == "movie":
            per_day[e[3]] = per_day.get(e[3], 0) + 1
    seen, out = {}, []
    for e in entries:
        if e[0] == "movie":
            seen[e[3]] = seen.get(e[3], 0) + 1
            e = (e[0], f"{e[3].replace('_', ' ')} (slide {seen[e[3]]} of {per_day[e[3]]})", e[2], e[3])
        out.append(e)
    return out


def check_day_titles(prs):
    """On the saved deck: every slide with a movie has one title "<day> (slide X of Y)"; per day, X runs 1..Y in deck
    order with no gaps and Y equals the number of movie slides of that day. Returns ({day: Y}, problem lines)."""
    title_re = re.compile(r"^(?P<day>.+) \(slide (?P<x>\d+) of (?P<y>\d+)\)$")
    found, problems = [], []
    for k, slide in enumerate(prs.slides, 1):
        if not any(sh._element.xpath(".//a:videoFile") for sh in slide.shapes):
            continue
        titles = [m for sh in slide.shapes if sh.has_text_frame for m in [title_re.match(sh.text_frame.text)] if m]
        if len(titles) != 1:
            problems.append(f"slide {k}: {len(titles)} '(slide X of Y)' titles")
            continue
        found.append((k, titles[0]["day"], int(titles[0]["x"]), int(titles[0]["y"])))
    counts = {}
    for _, day, _, _ in found:
        counts[day] = counts.get(day, 0) + 1
    seen = {}
    for k, day, x, y in found:
        seen[day] = seen.get(day, 0) + 1
        if x != seen[day] or y != counts[day]:
            problems.append(f"slide {k}: '{day} (slide {x} of {y})', expected slide {seen[day]} of {counts[day]}")
    return counts, problems


def fov_name(well, position):
    """The FOV's name as its movie caption writes it: ("FMC_W1", 1) -> "FMC W1 FOV1"."""
    return f"{well.replace('_', ' ')} FOV{int(position)}"


SUBSCRIPT_RE = re.compile(r"(?<=I)(under|away)\b")   # "Iunder", "Iaway" in slide text: the subscript part


def subscript_ranges(text):
    """1-based (start, length) of each "under" / "away" written as a subscript of I (TextboxSpec.subscripts)."""
    return [(m.start() + 1, len(m.group())) for m in SUBSCRIPT_RE.finditer(text)] or None


def removed_marks(cfg):
    """{movie stem: mark text} from the config's removed_fovs groups (day, well, positions, mark). A mark may hold
    {fov}: the FOV's name as the movie's caption writes it ("FMC W1 FOV1", fov_name), and {value}: that FOV's
    `value_column` from the config's fov_numbers csv, as the csv writes it (read by script)."""
    rows = None
    out = {}
    for g in cfg.get("removed_fovs") or []:
        for p in g["positions"]:
            mark = g["mark"].replace("{fov}", fov_name(g["well"], p))
            if "{value}" in mark:
                if rows is None:
                    import fov_numbers as fn
                    rows = fn.load_rows(cfg["fov_numbers"]["csv"])
                r = rows[(g["day"], g["well"].replace("_", " "), p)]
                if r["removed"] != "yes" or not r[g["value_column"]].strip():
                    sys.exit(f"STOP: {g['day']} {g['well']} p{p:02d}: csv removed={r['removed']}, "
                             f"{g['value_column']}={r[g['value_column']]!r}")
                mark = mark.format(value=r[g["value_column"]].strip())
            out[f"{g['day']}_{g['well']}_p{p:02d}"] = mark
    return out


def check_removed_marks(out, entries, removed):
    """Exactly the removed movies carry a mark, each mark directly under its own movie, and that movie's embedded
    file byte-identical to the removed stack's mp4. Stops on failure."""
    from pptx import Presentation
    paths = {p.stem: p for e in entries if e[0] == "movie" for _, ps in e[2] for p in ps
             if p.parent.parent.name != "frame"}   # a movie-pairs deck marks only the movie, not its per-frame copy
    want = {s for s in removed if s in paths}
    found = []
    prs = Presentation(str(out))
    for k, slide in enumerate(prs.slides):
        vids = []
        for sh in slide.shapes:
            link = sh._element.xpath(".//a:videoFile/@r:link")
            if link:
                vids.append((sh, slide.part.related_part(link[0]).blob))
        for sh in slide.shapes:
            if not (sh.has_text_frame and sh.text_frame.text in set(removed.values())):
                continue
            above = [(v, blob) for v, blob in vids if abs(v.top + v.height - sh.top) < 12700
                     and v.left < sh.left + sh.width / 2 < v.left + v.width]
            if len(above) != 1:
                sys.exit(f"CHECK FAIL: slide {k + 1}: mark '{sh.text_frame.text}' is not under exactly one movie")
            stem = next((s for s in want if paths[s].read_bytes() == above[0][1]), None)
            if stem is None or removed[stem] != sh.text_frame.text:
                sys.exit(f"CHECK FAIL: slide {k + 1}: mark '{sh.text_frame.text}' is under a movie not listed with it")
            # each mark names the movie above it (FF 2026-11-23, relayed: "it is not clear which fov is removed")
            m = re.match(r"^.+_(?P<ct>CAT|FMC)_(?P<w>W\d+)_p(?P<p>\d+)$", stem)
            name = fov_name(f"{m['ct']}_{m['w']}", m["p"])
            if not sh.text_frame.text.startswith(name + " "):
                sys.exit(f"CHECK FAIL: slide {k + 1}: mark '{sh.text_frame.text}' does not start with the name of "
                         f"the movie above it ({name})")
            found.append(stem)
    if sorted(found) != sorted(want):
        sys.exit(f"CHECK FAIL: marked {sorted(found)}, expected {sorted(want)}")
    print(f"CHECK {len(found)} removed movie(s) marked, each under its own stack and naming it: PASS")


def deck_specs(entries, make, removed=None, numbers=None, movie_notes=None, pairs=False):
    """Every slide's SlideSpec, in deck order (no poster is written here), and {slide index: values shown} for the
    per-slide numbers check."""
    shown = {}
    slides = [e[1] if e[0] == "slide" else make(e[1], e[2]) for e in entries]
    if numbers:
        import fov_numbers as fn
        rows = fn.load_rows(numbers["csv"])
        for k, e in enumerate(entries):
            if e[0] == "movie":
                if pairs:
                    slides[k], shown[k] = pair_slide(slides[k], make, rows, removed or {}, numbers["columns"])
                else:
                    slides[k], shown[k] = fn.numbers_slide(slides[k], make, rows, removed or {}, numbers["columns"])
    for k, e in enumerate(entries):   # config movie_notes: one grey line on the movie slides of a day and cell type
        if e[0] == "movie":
            for n in movie_notes_for(e, movie_notes):
                slides[k] = SlideSpec(tuple(slides[k].textboxes) + (movie_note_box(n["text"], make, bool(numbers)),),
                                      slides[k].movies, slides[k].pictures)
    return slides, shown


def check_source_images(out, slides):
    """Every figure picture in the deck byte-identical to its source PNG (PowerPoint downsamples some on save;
    build_movie_deck_via_com(restore_pictures=True) puts the source bytes back). Stops on a mismatch."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(str(out))
    for k, spec in enumerate(slides):
        pics = [s for s in prs.slides[k].shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
        for shape, pic in zip(pics, spec.pictures):
            if shape.image.blob != Path(pic.image_path).read_bytes():
                sys.exit(f"CHECK FAIL: slide {k + 1} picture is not byte-identical to {pic.image_path}")


def build_entries(entries, out, approved, removed=None, numbers=None, movie_notes=None, pairs=False, base=None,
                  timer=None):
    """Build the deck from entries, run the movie/autoplay check and the text checks; returns (the saved deck as a
    python-pptx Presentation, the 0-based slides built this run). removed: {movie stem: mark text} (removed_marks).
    numbers: the config's `fov_numbers` ({csv, columns}): every movie slide then carries each FOV's four values beside
    its movie (fov_numbers.py, the layout FF approved 2026-11-23), checked against the csv after the build.
    base: an existing deck of these entries: build only the changed slides (incremental.update_deck); None: full build,
    which also writes the sidecar a later incremental build reads."""
    timer = timer or StageTimer()
    scratch = Path(make_run_scratch_dir(Path(tempfile.gettempdir()) / "tgt_cat_fmc_posters"))
    try:   # posters in the local temp folder (removed after the build), never next to the movies on Y:
        make = MovieSlideMaker(scratch, removed)
        with timer.stage("slide specs (layout, values, notes)"):
            slides, shown = deck_specs(entries, make, removed, numbers, movie_notes, pairs)
        out.parent.mkdir(parents=True, exist_ok=True)
        if base:
            info = incremental.update_deck(slides, str(base), str(out), SLIDE_W, SLIDE_H, loop=True, extra=POSTER_RULE,
                                           prepare=make.write_posters, restore_pictures=True, timer=timer)
            built = [k - 1 for k in info["built"]]
            print(f"INCREMENTAL from {base}: built {len(info['built'])} slide(s) {info['built']}; kept "
                  f"{len(info['kept'])} in place; moved {len(info['moved'])} (old -> new) {info['moved']}; dropped "
                  f"{len(info['dropped'])} {info['dropped']}")
            print("CHECK every kept slide identical to its old slide and every built slide to its fresh build (content "
                  "hash: slide XML, movies, pictures, posters): " + ("PASS" if not info["problems"] else
                  f"FAIL ({len(info['problems'])})" + "".join(f"\n  {p}" for p in info["problems"])))
            n_autoplay = info["autoplay"]
            want_autoplay = sum(1 for k in built if entries[k][0] == "movie")
        else:
            with timer.stage(f"posters ({sum(len(s.movies) for s in slides)} movies)"):
                make.write_posters(slides)
            n_autoplay = build_movie_deck_via_com(slides, str(out), SLIDE_W, SLIDE_H, loop=True,
                                                  restore_pictures=True, timer=timer)
            with timer.stage("write sidecar (slide hashes)"):
                incremental.write_sidecar(str(out), [incremental.spec_hash(s, SLIDE_W, SLIDE_H, True, POSTER_RULE, True)
                                                     for s in slides])
            built = list(range(len(slides)))
            want_autoplay = sum(1 for e in entries if e[0] == "movie")
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    movies = [e for e in entries if e[0] == "movie"]
    n_movies = sum(len(s.movies) for s in slides)
    expected = sum(len(paths) for e in movies for _, paths in e[2])
    print(f"{out}: {len(slides)} slides ({len(slides) - len(movies)} text, {len(movies)} movie), {n_movies} movies of "
          f"{make.movie_w:.0f} x {make.movie_h:.0f} pt, fps {sorted(make.fps_seen)} (built slides), autoplay set on "
          f"{n_autoplay} built movie slide(s)")
    for k, e in enumerate(entries, 1):
        if e[0] == "movie":
            print(f"  slide {k}: {e[1]} | {describe(e[2], e[3])}")
    if n_autoplay != want_autoplay or n_movies != expected:
        sys.exit(f"CHECK FAIL: expected {expected} movies and autoplay on {want_autoplay} built slide(s)")
    print(f"CHECK {expected} movies, autoplay on {n_autoplay} built movie slide(s): PASS")
    from pptx import Presentation as _P
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    n_pics_expected = sum(len(s.pictures) for s in slides)
    with timer.stage("picture checks"):
        n_pics = sum(1 for sl in _P(str(out)).slides for sh in sl.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
        if n_pics != n_pics_expected:
            sys.exit(f"CHECK FAIL: {n_pics} still pictures in the deck, expected {n_pics_expected}")
        if n_pics_expected:
            check_source_images(out, slides)
            print(f"CHECK {n_pics} figure picture(s), each byte-identical to its source PNG: PASS")
    if removed:
        check_removed_marks(out, entries, removed)
    # Text checks report, they never stop a build (repo CLAUDE.md). lint_presentation skips slide 1 (title slide).
    text_stage = timer.stage("text, format and value checks")
    text_stage.__enter__()
    from pptx import Presentation
    from ppt_image_inserter.approved_text import deck_strings, load_approved, unapproved
    from ppt_image_inserter.text_lint import lint_presentation
    prs = Presentation(str(out))
    strings = deck_strings(prs)
    if numbers:   # the per-FOV values are csv data, checked against the csv below, not wording
        import fov_numbers as fn
        strings = [(w, t) for w, t in strings if not fn.VALUE_RE.fullmatch(t)]
    problems = lint_presentation(prs) + unapproved(strings, load_approved(approved))
    print(f"TEXT CHECK: {len(problems)} problem(s)" + "".join(f"\n  {p}" for p in problems))
    fmt = check_bullet_format(prs)
    print("CHECK bullet marks, bold main bullets, no 'field' in any form: "
          + ("PASS" if not fmt else f"FAIL ({len(fmt)})" + "".join(f"\n  {p}" for p in fmt)))
    counts, title_problems = check_day_titles(prs)
    print("CHECK movie titles per day, (slide X of Y) with X = 1..Y and Y = the day's movie slides "
          f"({', '.join(f'{d}: {y}' for d, y in counts.items())}): "
          + ("PASS" if not title_problems else f"FAIL ({len(title_problems)})"
             + "".join(f"\n  {p}" for p in title_problems)))
    sigma = [f"slide {n}: {t!r}" for n, sl in enumerate(prs.slides, 1) for sh in sl.shapes if sh.has_text_frame
             for t in [sh.text_frame.text] if re.search("σ|sigma|labels 3 and 4|Data from Kiet", t, re.I)]
    print(f"CHECK no 'σ', 'sigma', 'labels 3 and 4' or 'Data from Kiet' in the deck text: " + ("PASS" if not sigma else
                                                          f"FAIL ({len(sigma)})" + "".join(f"\n  {p}" for p in sigma)))
    if numbers:
        bad = {k: fn.check_numbers_slide(prs.slides[k], v) for k, v in shown.items()}
        bad = {k: v for k, v in bad.items() if v}
        print(f"CHECK on {len(shown)} movie slide(s): every value equals the csv, no text under {fn.MIN_PT:g} pt, "
              "no overlaps: " + ("PASS" if not bad else f"FAIL on {len(bad)} slide(s)" + "".join(
                  f"\n  slide {k + 1}: {p}" for k, ps in bad.items() for p in ps)))
    text_stage.__exit__(None, None, None)
    return prs, built


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
    ap.add_argument("--pairs", action="store_true", help="--config: movie-pairs deck (one pair per slide, its "
                    "per-frame movies below; needs --out)")
    ap.add_argument("--from", dest="base", type=Path, help="--config: incremental build of --out from this existing "
                    "deck: only slides whose spec changed are built (needs the deck's sidecar <deck>.slides.json)")
    ap.add_argument("--seed", type=Path, help="--config: check this existing deck against the config's slides and, if "
                    "it matches, write its sidecar so later builds can be incremental; builds nothing")
    args = ap.parse_args()
    timer = StageTimer(EXPECTED_STAGE_S)

    approved = APPROVED_TEXT
    if args.config:
        cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
        if cfg.get("stop_build"):   # a config not ready to build says why
            sys.exit(f"STOP: {cfg['stop_build']}")
        out = args.out or Path(cfg["output"])
        approved = args.config.parent / cfg["approved_text"]
    else:
        if not (args.movies_root and args.day and args.out):
            sys.exit("STOP: --fields and --full-day need --movies-root, --day and --out")
        out = args.out
    if (args.base or args.seed) and not args.config:
        sys.exit("STOP: --from and --seed need --config")
    if args.base and args.slides:
        sys.exit("STOP: --from builds the whole deck incrementally; --slides is for preview decks")
    if not args.seed:
        if DECK_DIR not in out.resolve().parents:
            sys.exit(f"STOP: decks go under {DECK_DIR} (PPT_image_inserter output rule), not {out}")
        if out.exists():
            sys.exit(f"STOP: {out} exists and is never overwritten")

    # deck entries in order: ("slide", SlideSpec) or ("movie", title, columns, day)
    if args.config:
        if args.pairs and not args.out and not args.seed:
            sys.exit("STOP: --pairs needs --out (the config's output is the main deck)")
        with timer.stage("plan (config, movie lists)"):
            entries = config_plan(cfg, args.version, pairs=args.pairs)
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

    if args.seed:   # an older deck: check it against its specs, then record its slide hashes
        scratch = Path(make_run_scratch_dir(Path(tempfile.gettempdir()) / "tgt_cat_fmc_posters"))
        try:
            with timer.stage("slide specs (layout, values, notes)"):
                slides, _ = deck_specs(entries, MovieSlideMaker(scratch, removed_marks(cfg)), removed_marks(cfg),
                                       cfg.get("fov_numbers"), cfg.get("movie_notes"), args.pairs)
            with timer.stage("check deck against specs (text, positions, movie and picture bytes), write sidecar"):
                problems = incremental.seed_sidecar(str(args.seed), slides, SLIDE_W, SLIDE_H, True, POSTER_RULE, True)
        finally:
            shutil.rmtree(scratch, ignore_errors=True)
        print(f"CHECK {args.seed.name} matches the config's {len(slides)} slides: " + (
            f"PASS, sidecar written: {incremental.sidecar_path(args.seed)}" if not problems else
            f"FAIL ({len(problems)}), no sidecar written" + "".join(f"\n  {p}" for p in problems[:30])))
        timer.report()
        return 0 if not problems else 1

    prs, built = build_entries(entries, out, approved, removed_marks(cfg) if args.config else None,
                               cfg.get("fov_numbers") if args.config else None,
                               cfg.get("movie_notes") if args.config else None, pairs=args.pairs, base=args.base,
                               timer=timer)
    n_read = sum(1 for k in built if entries[k][0] == "movie") if args.pairs else 0
    checks = timer.stage(f"deck checks ({n_read} slides' movies read; pairs, notes, format, labels)")
    checks.__enter__()
    if args.pairs:
        ot = check_one_type(entries)
        print(f"CHECK one-cell-type slides ({sum(1 for e in entries if one_type(e))}): consecutive FOVs of one well, "
              "lower left; a single movie on the left: " + ("PASS" if not ot else f"FAIL ({len(ot)})"
                                                            + "".join(f"\n  {p}" for p in ot)))
        counts, pair_problems = check_pairs(prs, entries, only=set(built))
        print(f"CHECK movie pairs ({', '.join(f'{k} {v}' for k, v in counts.items())}): every FOV once on top and once "
              "below, below = its per-frame movie (same name, frame folder, same frame count), embedded movies "
              "byte-identical to the files, key on every movie slide: "
              + ("PASS" if not pair_problems else f"FAIL ({len(pair_problems)})"
                 + "".join(f"\n  {p}" for p in pair_problems[:20])))
    if args.config and cfg.get("movie_notes"):   # each note on exactly the slides movie_notes_for picks
        for n in cfg["movie_notes"]:
            want = [k for k, e in enumerate(entries, 1) if e[0] == "movie" and movie_notes_for(e, [n])]
            got = [k for k, sl in enumerate(prs.slides, 1)
                   if any(sh.has_text_frame and sh.text_frame.text == n["text"] for sh in sl.shapes)]
            print(f"CHECK note {n['text']!r} on {len(got)} slide(s), the {n['day']} movie slides with a {n['ctype']} "
                  f"FOV ({len(want)}): " + ("PASS" if got == want else f"FAIL (on {got}, expected {want})"))
    if args.slides:   # a preview deck: the whole-deck format and label checks do not apply
        checks.__exit__(None, None, None)
        print(f"deck size {out.stat().st_size / 2**20:.1f} MB (preview of slides {args.slides})")
        timer.report()
        return 0
    # title page and details: every slide before the first movie slide (a pairs_split entry adds slides)
    n_lead = next((k for k, e in enumerate(entries) if e[0] == "movie"), 0) if args.config else 0
    n_fig = len(cfg.get("figures") or []) if args.config else 0   # metric slides, last in the deck
    fig_slides = list(range(len(entries) - n_fig, len(entries)))
    text_slides = list(range(1, n_lead)) + fig_slides
    if args.config and all(entries[i][0] == "slide" for i in [0] + text_slides):
        # figure slides: only their title is compared (same band as details)
        fmt_problems = check_against_reference(prs, text_slides, title_only=set(fig_slides))
        print(f"FORMAT CHECK vs {REFERENCE_DECK.name}: {len(fmt_problems)} difference(s)"
              + "".join(f"\n  {p}" for p in fmt_problems))
    pic_slides = [i for i, f in zip(fig_slides, cfg.get("figures") or []) if "png" in f] if args.config else []
    if args.config and pic_slides and cfg.get("figure_labels_at"):
        label_problems = check_figure_labels(prs, pic_slides)
        print(f"CHECK CAT/FMC labels and notes on {len(pic_slides)} figure slide(s): "
              + ("PASS" if not label_problems else "FAIL" + "".join(f"\n  {p}" for p in label_problems)))
    checks.__exit__(None, None, None)
    print(f"deck size {out.stat().st_size / 2**20:.1f} MB")
    timer.report()
    return 0


if __name__ == "__main__":
    sys.exit(main())
