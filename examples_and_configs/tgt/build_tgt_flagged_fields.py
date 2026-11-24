"""TGT flagged fields deck: only the flagged fields, the reason for each flag on its slides.

FF 2026-11-23, relayed verbatim by "TGT Sept EMCCD copy and first look": "that's anawful ppt it should have only the
flagged ones and include info on why they are flagged! I have no diea why or which nes are supposedly problematic".
Slides: the title page of the main TGT deck, then per section (one flagged well) its movies in the main deck's movie
format (IRM with outlines left, TGT right; 2 per column, column by column; column header in the cell type's color),
with a one-line caption under each movie when the section has a medians table, and the section's reason lines at the
bottom of its first slide. Every slide keeps the same movie size (room for captions and up to 3 reason lines).
Checks (stop on failure): movie count; each slide's embedded movies byte-identical to their sources; each caption's
numbers equal the medians table, re-read independently; text lint and approved-text check are reported.
The deck is a new file; the script stops if it exists (back up and swap by hand).
Run from the PPT_image_inserter repo root (env PPT_editing):
    python examples_and_configs\\tgt\\build_tgt_flagged_fields.py --config examples_and_configs\\configs\\live\\tgt\\tgt_flagged_fields_deck.yaml --out <new file>
"""
import argparse
import csv
import hashlib
import re
import shutil
import sys
import tempfile
from collections import defaultdict
from pathlib import Path

import imageio.v3 as iio
import yaml
from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_cat_fmc_slides import (  # noqa: E402
    COL_GAP, COLOR_BGR, DECK_DIR, DT_LEFT, DT_W, FONT, HEAD_H, HEAD_PT, HEAD_TOP, MOVIE_RATIO, POSTER_JPEG_QUALITY,
    ROW_GAP, ROWS_BOTTOM, ROWS_TOP, SLIDE_H, SLIDE_W, TITLE_H, TITLE_PT, TITLE_TOP, deck_bullet_pt, details_slide,
    check_bullet_format, full_day_plan, title_page, wrapped_lines)
from ppt_image_inserter.movies import (  # noqa: E402
    MovieSpec, SlideSpec, TextboxSpec, build_movie_deck_via_com, make_run_scratch_dir)

CAPTION_PTS, CAPTION_H = (16.0, 15.0, 14.0), 22.0   # one line under each movie: largest size that fits one line
LINE_PT, LINE_H, MAX_LINES = 18.0, 24.0, 3   # reason lines at the bottom of a section's first slide
GRID_MAX_W = 880.0   # widest movie grid (pt): 40 pt margin each side on the 960 pt slide
CAPTION_RE = re.compile(r"^p(\d\d): whole-FOV TGT median (\d+) to (\d+) counts$")   # "field" -> "FOV" 2026-11-23 (config caption)


def geometry(cap, n_lines, n_movies=4):
    """Movie height, width and grid left edge for a section: room for its captions (cap = caption box height, 0 for
    none) and its reason lines, the same on every slide of the section. A section of one movie shows it alone,
    centered, as large as fits; a section of two shows them side by side in one row; otherwise the 2 x 2 grid."""
    rows, cols = {1: (1, 1), 2: (1, 2)}.get(n_movies, (2, 2))
    h = (ROWS_BOTTOM - ROWS_TOP - n_lines * LINE_H - rows * cap - (rows - 1) * ROW_GAP) / rows
    w = min(h * MOVIE_RATIO, (GRID_MAX_W - (cols - 1) * COL_GAP) / cols)
    return w / MOVIE_RATIO, w, (SLIDE_W - cols * w - (cols - 1) * COL_GAP) / 2


def medians(csv_path, day, well):
    """{position: (first-frame median, last-frame median)} of this well's movies, from the per-frame table."""
    frames = defaultdict(list)
    with open(csv_path, newline="") as fh:
        for r in csv.DictReader(fh):
            if r["movie"].startswith(f"{day}_{well}_"):
                frames[int(r["position"].lstrip("p"))].append((int(r["frame"]), int(r["median"])))
    return {p: (min(v)[1], max(v)[1]) for p, v in frames.items()}


def ratios(csv_path, field_fmt, positions):
    """{position: (value at the first elapsed_min, value at the last)} from a field_lines table (day, construct,
    field, elapsed_min, value); field_fmt names a field, e.g. "FMC 20260923 FMC W4 p{pos:02d}"."""
    rows = defaultdict(list)
    with open(csv_path, newline="") as fh:
        for r in csv.DictReader(fh):
            rows[r["field"]].append((float(r["elapsed_min"]), float(r["value"])))
    out = {}
    for p in positions:
        v = rows.get(field_fmt.format(pos=p))
        if not v:
            sys.exit(f"STOP: {field_fmt.format(pos=p)} not in {csv_path}")
        out[p] = (min(v)[1], max(v)[1])
    return out


def check_lowest(sec):
    """The section's fields are the lowest last values of their day and construct in its ratio table, leaving out
    the listed removed fields, and every other field ends above the stated value. Stops on failure."""
    lc = sec["lowest_check"]
    last = defaultdict(list)
    day, construct = lc["day"], lc["construct"]
    with open(sec["ratio_csv"], newline="") as fh:
        for r in csv.DictReader(fh):
            if r["day"] == day and r["construct"] == construct and r["field"] not in lc["exclude"]:
                last[r["field"]].append((float(r["elapsed_min"]), float(r["value"])))
    ends = sorted((max(v)[1], f) for f, v in last.items())
    mine = {sec["ratio_field"].format(pos=p) for p in sec["positions"]}
    n = len(mine)
    if {f for _, f in ends[:n]} != mine or ends[n][0] <= lc["above"]:
        sys.exit(f"CHECK FAIL: lowest {n} are {ends[:n]}, next {ends[n]}; claimed {sorted(mine)} and others "
                 f"above {lc['above']}")
    print(f"CHECK '{sec['title']}': its fields are the lowest {n} ({', '.join(f'{v:.3f}' for v, _ in ends[:n])}); "
          f"next lowest {ends[n][0]:.3f} > {lc['above']}: PASS")


def movie_path(movie_dir, day, well, pos):
    return Path(movie_dir) / day / f"{day}_{well}_p{pos:02d}.mp4"


def section_slides(sec, movie_dir, scratch):
    """Slides of one section: [(SlideSpec, [source mp4 paths])]."""
    paths = [movie_path(movie_dir, sec["day"], sec["well"], p) for p in sec["positions"]]
    ctype = sec["well"].split("_")[0]
    meds = medians(sec["medians_csv"], sec["day"], sec["well"]) if sec.get("medians_csv") else {}
    lines = list(sec.get("lines") or [])
    if sec.get("unflagged_positions"):
        last = [meds[p][1] for p in sec["unflagged_positions"]]
        lines = [s.format(lo=min(last), hi=max(last)) for s in lines]
    if len(lines) > MAX_LINES:
        sys.exit(f"STOP: section '{sec['title']}' has {len(lines)} lines, room for {MAX_LINES}")
    if meds:
        captions = {p: sec["caption"].format(pos=p, first=meds[p][0], last=meds[p][1]) for p in sec["positions"]}
    elif sec.get("ratio_csv"):
        rat = ratios(sec["ratio_csv"], sec["ratio_field"], sec["positions"])
        captions = {p: sec["caption"].format(pos=p, first=rat[p][0], last=rat[p][1]) for p in sec["positions"]}
    else:   # fixed caption text per position (numbers not from a table: stated in the config's source comment)
        captions = {int(p): c for p, c in (sec.get("captions") or {}).items()}
    # caption size: the largest of CAPTION_PTS that fits one line; else the smallest size on as many lines as needed
    n_mov = len(sec["positions"])
    _, w1, _ = geometry(CAPTION_H if captions else 0.0, len(lines), n_mov)
    cap_pt = next((pt for pt in CAPTION_PTS if all(wrapped_lines(c, pt, w1 + COL_GAP) == 1
                                                   for c in captions.values())), None)
    n_cap = 1
    if captions and cap_pt is None:
        cap_pt = CAPTION_PTS[-1]
        n_cap = max(wrapped_lines(c, cap_pt, w1 + COL_GAP) for c in captions.values())
    cap_h = CAPTION_H * n_cap if captions else 0.0
    movie_h, movie_w, left0 = geometry(cap_h, len(lines), n_mov)
    cap_w = movie_w + COL_GAP   # a caption may run into half the gap on each side
    if captions and max(wrapped_lines(c, cap_pt, cap_w) for c in captions.values()) > n_cap:
        sys.exit(f"STOP: captions of '{sec['title']}' do not fit {n_cap} line(s) at {cap_pt:g} pt")
    if captions:
        print(f"'{sec['title']}': movies {movie_w:.0f} x {movie_h:.0f} pt, captions {cap_pt:g} pt")
    else:
        print(f"'{sec['title']}': movies {movie_w:.0f} x {movie_h:.0f} pt")
    if n_mov == 2:   # one row, side by side (geometry)
        plan = [[(ctype, [paths[0]]), (ctype, [paths[1]])]]
    else:
        plan = full_day_plan(paths if ctype == "CAT" else [], paths if ctype == "FMC" else [])
    out = []
    for k, cols in enumerate(plan):
        boxes = [TextboxSpec(sec["title"], 0, TITLE_TOP, SLIDE_W, TITLE_H, TITLE_PT, font_name=FONT)]
        movies, used = [], []
        for col, (ct, col_paths) in enumerate(cols):
            if not ct:
                continue
            left = left0 + col * (movie_w + COL_GAP)
            boxes.append(TextboxSpec(ct, left, HEAD_TOP, movie_w, HEAD_H, HEAD_PT, color_rgb=COLOR_BGR[ct],
                                     font_name=FONT))
            for row, mp4 in enumerate(col_paths):
                if not mp4.is_file():
                    sys.exit(f"STOP: {mp4} not found")
                frame = iio.imread(mp4, index=0)
                poster = scratch / f"{mp4.stem}_poster.jpg"
                Image.fromarray(frame).convert("RGB").save(poster, "JPEG", quality=POSTER_JPEG_QUALITY)
                top = ROWS_TOP + row * (movie_h + cap_h + ROW_GAP)
                movies.append(MovieSpec(str(mp4), str(poster), left, top, movie_w, movie_h))
                used.append(mp4)
                if captions:
                    pos = int(mp4.stem.rsplit("_p", 1)[1])
                    boxes.append(TextboxSpec(captions[pos], left - COL_GAP / 2, top + movie_h, cap_w, cap_h,
                                             cap_pt, bold=False, font_name=FONT))
        if k == 0 and lines:
            boxes.append(TextboxSpec("\r".join(lines), DT_LEFT, ROWS_BOTTOM - len(lines) * LINE_H, DT_W,
                                     len(lines) * LINE_H, LINE_PT, bold=False, align="center", anchor="top",
                                     font_name=FONT))
        out.append((SlideSpec(boxes, movies), used))
    return out


def check_summary(cfg):
    """Each summary line names its section's well and positions ("positions a to b" = that range, "all N positions" =
    N positions); for a section with a medians table, every flagged position ends below its first frame and every
    unflagged one does not. Stops on failure."""
    for sec in cfg["sections"]:
        line, pos = sec["summary"], sec["positions"]
        day, well = sec["day"].replace("_", " "), sec["well"].replace("_", " ")
        if not line.startswith(f"{day} {well}, "):
            sys.exit(f"CHECK FAIL: summary line '{line}' does not name {day} {well}")
        m_range = re.search(r"positions (\d+) to (\d+)", line)
        m_one = re.search(r"position (\d+):", line)
        m_all = re.search(r"all (\d+) positions", line)
        m_list = re.search(r"positions ((?:\d+(?:, | and ))+\d+):", line)
        ok = (m_range and pos == list(range(int(m_range[1]), int(m_range[2]) + 1))) or (m_all and len(pos) == int(
            m_all[1])) or (m_one and pos == [int(m_one[1])]) or (
            m_list and pos == [int(n) for n in re.findall(r"\d+", m_list[1])])
        if not ok:
            sys.exit(f"CHECK FAIL: summary line '{line}' does not match positions {pos}")
        if sec.get("medians_csv"):
            table = medians(sec["medians_csv"], sec["day"], sec["well"])
            lost = sorted(p for p, (first, last) in table.items() if last < first)
            if lost != sorted(pos):
                sys.exit(f"CHECK FAIL: positions ending below their start are {lost}, flagged {pos}")
    print(f"CHECK {len(cfg['sections'])} summary line(s) match their sections; flagged positions are exactly those "
          "ending below their first-frame median: PASS")


def md5(data):
    return hashlib.md5(data).hexdigest()


def check_deck(out, built, cfg):
    """Stops on failure. built: [(SlideSpec, [source paths])] in deck order (title page has none)."""
    from pptx import Presentation
    prs = Presentation(str(out))
    n_expected = sum(len(s["positions"]) for s in cfg["sections"])
    n_total = 0
    for k, (_, srcs) in enumerate(built):
        parts = {r.target_part.partname: r.target_part for r in prs.slides[k].part.rels.values()
                 if not r.is_external and r.target_part.content_type.startswith("video")}
        got = sorted(md5(p.blob) for p in parts.values())
        want = sorted(md5(Path(s).read_bytes()) for s in srcs)
        if got != want:
            sys.exit(f"CHECK FAIL: slide {k + 1} movies are not byte-identical to {[Path(s).name for s in srcs]}")
        n_total += len(got)
    if n_total != n_expected:
        sys.exit(f"CHECK FAIL: {n_total} movies in the deck, expected {n_expected}")
    print(f"CHECK {n_total} movies, each byte-identical to its source mp4: PASS")
    for sec in cfg["sections"]:   # every movie of a section is one of its flagged positions
        allowed = {md5(movie_path(cfg["movie_dir"], sec["day"], sec["well"], p).read_bytes()): p
                   for p in sec["positions"]}
        n_sec = sum(1 for k, (_, srcs) in enumerate(built) for s in srcs if Path(s).name.startswith(
            f"{sec['day']}_{sec['well']}_p") and md5(Path(s).read_bytes()) in allowed)
        if n_sec != len(sec["positions"]):
            sys.exit(f"CHECK FAIL: section '{sec['title']}' has {n_sec} of its {len(sec['positions'])} positions")
        print(f"CHECK '{sec['title']}': only positions {sec['positions']}: PASS")
    n_cap = 0
    for sec in cfg["sections"]:   # caption numbers vs the table, re-read here
        if not sec.get("medians_csv"):
            continue
        table = medians(sec["medians_csv"], sec["day"], sec["well"])
        n_before = n_cap
        for slide in prs.slides:
            if not any(sh.has_text_frame and sh.text_frame.text == sec["title"] for sh in slide.shapes):
                continue
            for sh in slide.shapes:
                m = sh.has_text_frame and CAPTION_RE.match(sh.text_frame.text)
                if m:
                    pos, first, last = int(m[1]), int(m[2]), int(m[3])
                    if table[pos] != (first, last):
                        sys.exit(f"CHECK FAIL: caption p{pos:02d} says {first} to {last}, table {table[pos]}")
                    n_cap += 1
        if n_cap - n_before != len(sec["positions"]):
            sys.exit(f"CHECK FAIL: {n_cap - n_before} captions in '{sec['title']}', "
                     f"expected {len(sec['positions'])}")
    print(f"CHECK {n_cap} caption(s), numbers equal the medians table: PASS")
    for sec in cfg["sections"]:   # ratio captions: exact text recomputed from the table, re-read here
        if not sec.get("ratio_csv"):
            continue
        rat = ratios(sec["ratio_csv"], sec["ratio_field"], sec["positions"])
        want = sorted(sec["caption"].format(pos=p, first=rat[p][0], last=rat[p][1]) for p in sec["positions"])
        got = sorted(sh.text_frame.text for slide in prs.slides
                     if any(s.has_text_frame and s.text_frame.text == sec["title"] for s in slide.shapes)
                     for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text in want)
        if got != want:
            sys.exit(f"CHECK FAIL: '{sec['title']}' captions {got}, expected {want}")
        print(f"CHECK {len(got)} caption(s) of '{sec['title']}' equal the ratio table "
              f"({'; '.join(f'p{p:02d} {a:.4f} to {b:.4f}' for p, (a, b) in rat.items())}): PASS")
        if sec.get("lowest_check"):
            check_lowest(sec)
    return prs


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--config", type=Path, required=True)
    ap.add_argument("--out", type=Path, help="deck path (default: the config's output)")
    args = ap.parse_args()
    cfg = yaml.safe_load(args.config.read_text(encoding="utf-8"))
    out = args.out or Path(cfg["output"])
    if DECK_DIR not in out.resolve().parents:
        sys.exit(f"STOP: decks go under {DECK_DIR}, not {out}")
    if out.exists():
        sys.exit(f"STOP: {out} exists and is never overwritten")
    check_summary(cfg)
    scratch = Path(make_run_scratch_dir(Path(tempfile.gettempdir()) / "tgt_flagged_posters"))
    try:
        tp = cfg["title_page"]
        built = [(title_page(tp["title"], tp["subtitle"]), [])]
        summary = [s["summary"] for s in cfg["sections"]]
        built.append((details_slide(cfg["summary_title"], summary, pt=deck_bullet_pt([{"bullets": summary}])), []))
        for sec in cfg["sections"]:
            built += section_slides(sec, cfg["movie_dir"], scratch)
        n_autoplay = build_movie_deck_via_com([s for s, _ in built], str(out), SLIDE_W, SLIDE_H, loop=True)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    print(f"{out}: {len(built)} slides, autoplay set on {n_autoplay}")
    prs = check_deck(out, built, cfg)
    from ppt_image_inserter.approved_text import deck_strings, load_approved, unapproved
    from ppt_image_inserter.text_lint import lint_presentation
    problems = lint_presentation(prs) + unapproved(deck_strings(prs), load_approved(args.config.parent /
                                                                                     cfg["approved_text"]))
    print(f"TEXT CHECK: {len(problems)} problem(s)" + "".join(f"\n  {p}" for p in problems))
    fmt = check_bullet_format(prs)
    print("CHECK bullet marks, bold main bullets, no 'field' in any form: "
          + ("PASS" if not fmt else f"FAIL ({len(fmt)})" + "".join(f"\n  {p}" for p in fmt)))
    print("TEXT AS BUILT:")
    for k, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                for line in sh.text_frame.text.split("\n"):
                    print(f"  slide {k}: {line}")
    print(f"deck size {out.stat().st_size / 2**20:.1f} MB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
