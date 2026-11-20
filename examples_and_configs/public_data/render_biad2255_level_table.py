"""Render the BIAD2255 "Relative level near nadir" table as a slide image, in the Fig 6 summary table format
(FF 2026-11-19: "reuse table format of Histone mark data agent presentation").

Format, imported unchanged: the Fig 6 table's CSS and Edge screenshot (render_fig6_table_slides.py) and its value cell
(build_fig6_summary_table.cell: mean ± SEM, blue below 1 / red above 1, stars from the BH-adjusted signed-rank p).
Values: the rendering session's CSV (TCell-3D-Morphodynamics dev/analyses/public_perturbation_tcell/biad2255_deck/
render_timecourse_level_table.m): per time point and channel, n_cells, mean, sem, signrank_p_adj (BH over the six time
points of one channel, as Fig 6 adjusts over the arms of one panel).

Usage:  python render_biad2255_level_table.py <level_mean_sem_signrank.csv> <out.png> [approved_text.yaml]
            [--channels IFNG,TNF,DAPI] [--ref "the layer average"]
Defaults: the outside ring (IFNG, TNF; "1 = the ring average"). Inner layers (FF 2026-11-19: inner 0.25 and 0.5 μm
"as for figure 6 of paper", DAPI shown beside the mRNAs as Fig 6 shows DNA beside the mark) pass --channels and --ref.
Prints any text drawn into the image that the approved-text file lacks (repo CLAUDE.md: builders check text they
draw into images).
"""
import csv
import html
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "fig6_metrics"))
sys.path.insert(0, str(HERE.parents[1]))
import render_fig6_table_slides as F6  # noqa: E402  (also puts the Fig 6 table builder on the path)
from build_fig6_summary_table import cell  # noqa: E402

CHANNELS = ["IFNG", "TNF"]   # outside ring, no DAPI: FF 2026-11-19 "leave DAPI out of this (for anything perinuclear)"
REF = "the ring average"
TIMES = ["0 h", "1 h", "2 h", "3 h", "4 h", "6 h"]
KEY_LINES = [
    "Mean ± SEM over cells; 1 = {ref}",
    "below 1", "above 1",
    "* <0.05, ** <0.01, *** <0.001 vs 1 (signed-rank, BH-adjusted); gray = n.s.",   # Fig 6 approved key line
]


def load(path):
    rows = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            rows[(r["time"], r["channel"])] = r
    missing = [(t, c) for t in TIMES for c in CHANNELS if (t, c) not in rows]
    if missing:
        raise SystemExit(f"ERROR: CSV lacks {missing}")
    n_by_time = {t: {rows[(t, c)]["n_cells"] for c in CHANNELS} for t in TIMES}
    if any(len(v) != 1 for v in n_by_time.values()):
        raise SystemExit(f"ERROR: n differs between channels at one time point: {n_by_time}")
    return rows


def table_html(rows):
    # no group header: the slide title says "Relative level near nadir" (it repeated in the image)
    h = ['<table><thead><tr><th class="l">Time</th><th>n</th><th class="first">' + CHANNELS[0] + "</th>"
         + "".join(f"<th>{c}</th>" for c in CHANNELS[1:]) + "</tr></thead><tbody>"]
    for t in TIMES:
        r = {c: rows[(t, c)] for c in CHANNELS}
        tds = [cell(float(r[c]["mean"]), float(r[c]["sem"]), r[c]["signrank_p_adj"], True) for c in CHANNELS]
        tds[0] = F6.first(tds[0])
        h.append(f'<tr><td class="l">{t}</td><td class="n">{r[CHANNELS[0]]["n_cells"]}</td>{"".join(tds)}</tr>')
    return "".join(h) + "</tbody></table>"


def page(rows):
    key = ('<div class="key"><span>' + html.escape(KEY_LINES[0].format(ref=REF)) + "&nbsp;&nbsp;&nbsp;"
           '<span class="sw" style="background:rgba(37,99,235,.55)"></span>below 1&nbsp;&nbsp;'
           '<span class="sw" style="background:rgba(220,38,38,.55)"></span>above 1</span>'
           "<span>" + html.escape(KEY_LINES[3]) + "</span></div>")
    # larger than the Fig 6 tables: 6 rows here, so the table can fill the slide at a readable size
    css = F6.CSS + ("body,th,td{font-size:24px} .se{font-size:24px} sup{font-size:18px} th,td{padding:2px 14px}"
                    ".key{font-size:22px;width:auto;flex-direction:column;align-items:center} .key span{white-space:nowrap}"
                    # 2 columns: the key is wider than the table; center both (reviewer: table sat off-center)
                    "table{margin-left:auto;margin-right:auto}")
    return (f'<meta charset="utf-8"><title>Relative level near nadir</title><style>{css}</style><div class="box">'
            + table_html(rows) + key + "</div>").replace("µ", F6.MU)


def main(csv_path, out_png, approved=None):
    rows = load(csv_path)
    out = Path(out_png)
    out.parent.mkdir(parents=True, exist_ok=True)
    hp = out.with_suffix(".html")
    hp.write_text(page(rows), encoding="utf-8")
    subprocess.run([F6.EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                    "--window-size=1400,700", f"--screenshot={out}", hp.resolve().as_uri()],
                   check=True, capture_output=True, timeout=120)
    if not out.exists():
        raise SystemExit(f"ERROR: no screenshot written: {out}")
    from PIL import Image
    im = Image.open(out).convert("RGB")
    bbox = Image.eval(im, lambda v: 255 - v).getbbox()      # trim the white page around the table
    # extra white at the bottom keeps the key clear of the slide caption once placed (reviewer: they touched)
    im.crop((max(bbox[0] - 16, 0), max(bbox[1] - 16, 0), bbox[2] + 16, bbox[3] + 130)).save(out)
    print(f"{out} ({len(TIMES)} rows x {len(CHANNELS)} channels)")
    if approved:
        from ppt_image_inserter.approved_text import is_approved, load_approved
        ok = load_approved(approved)
        drawn = ["Time", "n"] + CHANNELS + TIMES + [KEY_LINES[0].format(ref=REF)] + KEY_LINES[1:]
        bad = [t for t in drawn if not is_approved(t, ok)]
        for t in bad:
            print(f"[TEXT] table image: not in the approved-text file: {t!r}")


if __name__ == "__main__":
    args = sys.argv[1:]
    for flag in ("--channels", "--ref"):
        if flag in args:
            i = args.index(flag)
            val = args[i + 1]
            del args[i:i + 2]
            if flag == "--channels":
                CHANNELS = val.split(",")
            else:
                REF = val
    if len(args) not in (2, 3):
        sys.exit(__doc__)
    main(*args)
