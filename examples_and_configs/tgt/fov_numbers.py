"""Per-FOV numbers on TGT movie slides: four values per movie beside it, the layout FF approved on 2026-11-23
("example slide looks good. standing go for remaing full ppt", relayed), built as review\\fov_numbers_example_v7.

FF 2026-11-23, relayed by "TGT Sept EMCCD copy and first look": "on each data slide we need to record key quantitative
results of each fov to left and right of it"; "we need single numbers"; "it must be per-fov period". The four labels
are FF's, with "under" and "away" as subscripts of I and 3x MAD in labels 3 and 4 ("say 3X MAD and write out full
form"). Values are read by script from the per-FOV csv, never typed: labels 1 and 2 with 2 decimals as in the csv,
labels 3 and 4 as whole percent. An empty csv cell stays blank; a removed FOV keeps its red mark and gets no values.

Layout (the director's brief, 2026-11-23): movies at the deck's size, CAT column at the left edge, FMC at the right,
CAT / FMC headers on the title line. Each row of movies gets one block filling the gap between its two movies: the CAT
movie's values at the block's left edge, the four labels in the middle, the FMC movie's values at the right edge,
each value centered vertically on its label. Definitions at the bottom, dark, full width.

Used by build_cat_fmc_slides.py (config key `fov_numbers`) and fov_numbers_example.py.
"""
import csv
import re
import sys
from dataclasses import replace
from itertools import combinations
from pathlib import Path

from PIL import ImageFont

import build_cat_fmc_slides as b
from ppt_image_inserter.movies import SlideSpec, TextboxSpec

# FF's four labels as (text, subscript) runs, and whether "(last frame)" goes under the expression.
LABELS = [
    ([("I", 0), ("away", 1), (": last / first", 0)], False),
    ([("I", 0), ("under", 1), (" / I", 0), ("away", 1)], True),
    ([("area with I", 0), ("under", 1), (" < I", 0), ("away", 1), (" − 3× MAD", 0)], True),
    ([("area with I", 0), ("under", 1), (" > I", 0), ("away", 1), (" + 3× MAD", 0)], True),
]
QUALIFIER = "(last frame)"
PERCENT = (False, False, True, True)   # labels 3 and 4 are area fractions, shown as whole percent
# The two definition lines, FF's approval (2026-11-23, relayed): "really redo both lines to make clearer and more
# concise"; "also use italics at ottom"; "ok to the two lines". In "Iaway" (line 2) "away" is a subscript; nothing else.
DEF_LINES = ("I: TGT signal. under: pixels inside a cell; away: pixels at least 0.5 μm from every object; both in "
             "every IRM frame.",
             "Ratios: means of I. Areas: % of under pixels, with I smoothed (Gaussian, 1 pixel) and Iaway the median. "
             "MAD: median absolute deviation of I away.")
DEFINITIONS = "\r".join(DEF_LINES)
DEF_SUBS = [(DEFINITIONS.index("Iaway") + 2, len("away"))]   # 1-based start of "away" in "Iaway"
PROVISIONAL = "Provisional values"   # example only, while the values were from the old threshold

# Sizes (pt) and exact line pitches. Brief: values 24 bold, labels 18, qualifier 14, definitions >= 14. Labels at 18
# are 231 pt wide (Calibri, measured) and the block has 212 pt for them, so labels are 16, the largest that fits.
VAL_PT, VAL_PITCH = 24.0, 29.0
LAB_PT, LAB_PITCH = 16.0, 20.0
QUAL_PT, QUAL_PITCH = 14.0, 17.0
DEF_PT, DEF_PITCH = 14.0, 17.0
PROV_PT = 14.0
ITEM_GAP = 10.0                   # between the four label items
EDGE = 6.0                        # slide edge to the nearest movie
CLEAR = 10.8                      # block to each movie: 0.15 in
VAL_W, PAD = 43.0, 4.0            # value column (widest csv value at 24 pt bold: 43 pt); value to label column
MARK_W = 444.0                    # red mark box, no inner side margins: longest mark at 16 pt is 433 pt (Calibri)
TITLE_W, HEAD_W = 540.0, 80.0     # title box between the CAT / FMC header boxes, all on the title line
ROWS_TOP = b.TITLE_TOP + b.TITLE_H + 2
DEF_H = 2 * DEF_PITCH
MIN_PT = 14.0                     # no text on a movie slide under this size (the director's check)
STEM_RE = re.compile(r"^(?P<day>.+)_(?P<ct>CAT|FMC)_(?P<w>W\d+)_p(?P<p>\d+)$")
VALUE_RE = re.compile(r"\d+(\.\d+)?%?")   # a shown value (2 decimals or whole percent)


def load_rows(path):
    """{(day, "CAT W1", position): row} from the per-FOV csv; one row per FOV."""
    rows = {}
    for r in csv.DictReader(open(path, encoding="utf-8")):
        key = (r["day"], r["well"], int(r["position"]))
        if key in rows:
            sys.exit(f"STOP: two csv rows for {key}")
        rows[key] = r
    return rows


def row_for(stem, rows, removed):
    """The csv row of a movie, checked against the stack name and the deck's removed marks."""
    m = STEM_RE.match(stem)
    if not m:
        sys.exit(f"STOP: movie name {stem} not understood")
    r = rows.get((m["day"], f"{m['ct']} {m['w']}", int(m["p"])))
    if r is None:
        sys.exit(f"STOP: no csv row for {stem}")
    if f"_{m['ct']}_{m['w']}_" not in r["stack"] or not r["stack"].endswith(f"_p{m['p']}"):
        sys.exit(f"STOP: csv stack {r['stack']} does not match movie {stem}")
    if (r["removed"] == "yes") != (stem in removed):
        sys.exit(f"STOP: {stem}: csv removed={r['removed']}, deck mark {'present' if stem in removed else 'absent'}")
    return r


def values(r, columns):
    """The four values as shown: 2 decimals as in the csv, or whole percent; an empty cell stays ""."""
    out = []
    for c, pct in zip(columns, PERCENT):
        v = r[c].strip()
        out.append("" if v == "" else (f"{round(float(v) * 100):d}%" if pct else v))
    return out


def box(text, left, top, width, height, pt, align, bold=False, color=None, subs=None, pitch=None, anchor="middle",
        italic=False):
    """A text box exactly its lines' height (no top/bottom or side inner margins), so boxes can sit close."""
    return TextboxSpec(text, left, top, width, height, pt, bold=bold, align=align, anchor=anchor, font_name=b.FONT,
                       color_rgb=color, side_margin_pt=0.0, vert_margin_pt=0.0, subscripts=subs,
                       line_spacing_pt=pitch or height, italic=italic)


def label_h(qual):
    return LAB_PITCH + (QUAL_PITCH if qual else 0)


def item_offsets():
    """Top of each label item's expression line relative to the block top, and the block height. Positions as the
    director approved (fov_numbers_example_v5); each value is centered vertically on its whole label (both lines),
    which keeps it inside the space its item has."""
    tops, y, half_extra = [], (VAL_PITCH - LAB_PITCH) / 2, (VAL_PITCH - LAB_PITCH) / 2
    for _, qual in LABELS:
        tops.append(y)
        bottom = max(y + label_h(qual), y + LAB_PITCH / 2 + VAL_PITCH / 2)
        y = bottom + ITEM_GAP + half_extra
    return tops, y - ITEM_GAP - half_extra


def label_boxes(left, width, block_top, tops):
    boxes = []
    for (runs, qual), t in zip(LABELS, tops):
        text, subs = "", []
        for s, is_sub in runs:
            if is_sub:
                subs.append((len(text) + 1, len(s)))
            text += s
        boxes.append(box(text, left, block_top + t, width, LAB_PITCH, LAB_PT, "center", subs=subs))
        if qual:
            boxes.append(box(QUALIFIER, left, block_top + t + LAB_PITCH, width, QUAL_PITCH, QUAL_PT, "center"))
    return boxes


def value_boxes(vals, left, align, color, block_top, tops):
    return [box(v, left, block_top + t + label_h(qual) / 2 - VAL_PITCH / 2, VAL_W, VAL_PITCH, VAL_PT, align,
                bold=True, color=color)
            for v, t, (_, qual) in zip(vals, tops, LABELS) if v != ""]


def title_fits(title):
    """True when a movie title fits its box on the title line at the deck's title size (bold Calibri, measured)."""
    font = ImageFont.truetype(b.FONT_FILE_BOLD, 100)
    return font.getlength(title) / 100 * b.TITLE_PT <= TITLE_W - 14.4


def numbers_slide(spec, maker, rows, removed, columns, provisional=False):
    """The deck's movie slide (MovieSlideMaker output at the deck's movie size) in the numbers layout. Returns the
    SlideSpec and the values shown, one list of four per non-removed movie."""
    w, h, dy = maker.movie_w, maker.movie_h, ROWS_TOP - b.ROWS_TOP
    lefts = [EDGE, b.SLIDE_W - EDGE - w]

    def col(left):
        return 0 if left < b.SLIDE_W / 2 else 1

    boxes = []
    for tb in spec.textboxes:
        if tb.width_pt >= b.SLIDE_W:   # title, narrowed so the headers fit beside it on the same line
            if not title_fits(tb.text):
                sys.exit(f"STOP: title {tb.text!r} is wider than its {TITLE_W:.0f} pt box")
            boxes.append(replace(tb, left_pt=(b.SLIDE_W - TITLE_W) / 2, width_pt=TITLE_W, anchor="middle"))
        elif tb.color_rgb == b.MARK_BGR:   # right under its own movie (top-anchored, no inner margins)
            center = lefts[col(tb.left_pt + tb.width_pt / 2)] + w / 2
            boxes.append(replace(tb, left_pt=min(max(center - MARK_W / 2, 0.0), b.SLIDE_W - MARK_W),
                                 width_pt=MARK_W, top_pt=tb.top_pt + dy, anchor="top", side_margin_pt=0.0,
                                 vert_margin_pt=0.0))
        else:   # CAT / FMC header, on the title line above its column
            center = lefts[col(tb.left_pt)] + w / 2
            boxes.append(replace(tb, left_pt=center - HEAD_W / 2, width_pt=HEAD_W, top_pt=b.TITLE_TOP,
                                 height_pt=b.TITLE_H, anchor="middle"))
    movies = [replace(mv, left_pt=lefts[col(mv.left_pt)], top_pt=mv.top_pt + dy) for mv in spec.movies]

    gap_l, gap_r = EDGE + w + CLEAR, b.SLIDE_W - EDGE - w - CLEAR
    lab_left, lab_w = gap_l + VAL_W + PAD, gap_r - gap_l - 2 * (VAL_W + PAD)
    tops, block_h = item_offsets()
    if block_h > h:
        sys.exit(f"STOP: the block ({block_h:.0f} pt) is taller than a movie ({h:.0f} pt)")
    by_row = {}
    for mv in movies:
        by_row.setdefault(round(mv.top_pt), []).append(mv)
    shown = []
    for top, mvs in sorted(by_row.items()):
        block_top = top + (h - block_h) / 2
        any_values = False
        for mv in mvs:
            stem = Path(mv.movie_path).stem
            r = row_for(stem, rows, removed)
            if r["removed"] == "yes":
                continue   # its value column stays empty
            vals, ctype = values(r, columns), STEM_RE.match(stem)["ct"]
            shown.append(vals)
            any_values = True
            if col(mv.left_pt) == 0:
                boxes += value_boxes(vals, gap_l, "left", b.COLOR_BGR[ctype], block_top, tops)
            else:
                boxes += value_boxes(vals, gap_r - VAL_W, "right", b.COLOR_BGR[ctype], block_top, tops)
        if any_values:
            boxes += label_boxes(lab_left, lab_w, block_top, tops)
    if provisional:   # centered between the two rows of movies, clear of the red marks
        row_tops = sorted(by_row)
        mid = (row_tops[0] + h + row_tops[-1]) / 2 if len(row_tops) > 1 else row_tops[0] + h + maker.mark_h / 2
        boxes.append(box(PROVISIONAL, b.SLIDE_W / 2 - 70, mid - 9, 140, 18, PROV_PT, "center", bold=True))
    boxes.append(box(DEFINITIONS, 8, b.SLIDE_H - DEF_H - 3, b.SLIDE_W - 16, DEF_H, DEF_PT, "center", pitch=DEF_PITCH,
                     anchor="bottom", subs=DEF_SUBS, italic=True))
    return SlideSpec(boxes, movies, spec.pictures), shown


def check_numbers_slide(slide, shown):
    """Script checks on one saved movie slide: no text under 14 pt, no shape overlaps another, the values shown equal
    the csv-derived ones. Returns problem lines (empty when all pass)."""
    from pptx.util import Pt
    problems = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for run in p.runs:
                    if run.text.strip() and (run.font.size is None or run.font.size < Pt(MIN_PT)):
                        problems.append(f"text under {MIN_PT:g} pt or size unknown: {run.text!r} ({run.font.size})")
    rect = [(sh.left / 12700, sh.top / 12700, (sh.left + sh.width) / 12700, (sh.top + sh.height) / 12700,
             sh.text_frame.text[:30] if sh.has_text_frame else f"movie {sh.name}") for sh in slide.shapes]
    for a, c in combinations(rect, 2):
        if min(a[2], c[2]) - max(a[0], c[0]) > 0.01 and min(a[3], c[3]) - max(a[1], c[1]) > 0.01:
            problems.append(f"overlap: {a[4]!r} and {c[4]!r}")
    defs = [sh for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text == "\n".join(DEF_LINES)]
    if len(defs) != 1:
        problems.append(f"{len(defs)} boxes with the two approved definition lines (expected 1)")
    elif not all(r.font.italic for p in defs[0].text_frame.paragraphs for r in p.runs if r.text.strip()):
        problems.append("definition lines not italic throughout")
    got = [sh.text_frame.text for sh in slide.shapes if sh.has_text_frame and VALUE_RE.fullmatch(sh.text_frame.text)]
    want = [v for vals in shown for v in vals if v != ""]
    if sorted(got) != sorted(want):
        problems.append(f"values {sorted(got)} != csv {sorted(want)}")
    return problems
