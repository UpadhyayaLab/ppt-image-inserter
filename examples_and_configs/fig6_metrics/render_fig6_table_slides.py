"""Render the Figure 6 summary table (mean +- SEM) as 5 slide images, one per mark group.

FF 2026-11-19: slides = H3K27me3 Jurkat, H3K27me3 CTL (3SI + naive), H3K9me3 (all), H3K27ac (all), RNAP II (all);
"we should keep table simple. just show the rel level near nadir. that's all" -> one measure, mark and DNA
("show DNA along with h3k27me3"); "you should include centrosome version of it when that was imaged. in other cases -
just leave blank" / "we need to show what is c in figure 6" -> a "Near centrosome" group = paper Fig 6C's measure
(seed study S2, "Region Nadir near Centrosome"), from <figs_dir>/C_cent_mean_sem.csv (fig6_C_cent_mean_sem.m, Histone
mark data session); only experiments in that CSV have it, so only slides holding one get the extra columns.

Values come from the Histone mark data session's table builder (collect(), cell()) in TCell-3D-Morphodynamics,
imported unchanged: the C level (deepest invagination region / whole-surface average), mean +- SEM over cells.

Every line of text drawn into these images comes from deck_text.yaml (TEXT; pass another dict to render a variant).

Usage:  python render_fig6_table_slides.py <figs_dir> <out_dir> [<deck config.yaml>]
Writes <out_dir>/table_<n>_<slug>.html and .png (headless Edge screenshot), contents.png (slide ranges computed from
the deck config's titles) and prints rows per slide.
"""
import csv
import html
import subprocess
import sys
from pathlib import Path

import yaml

BUILDER = Path("C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/dev/reporting/fig6_metrics_page")
sys.path.insert(0, str(BUILDER))
from build_fig6_summary_table import TITLE, MARK, cell, collect  # noqa: E402
from build_fig6_metrics_page import KINDS  # noqa: E402

EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
TEXT = yaml.safe_load(open(Path(__file__).with_name("deck_text.yaml"), encoding="utf-8"))
DECK_CFG = Path(__file__).resolve().parents[1] / "configs/fixed/fig6_metrics/config_fig6_metrics_by_metric_20261118.yaml"
MU = "\u03bc"   # Greek mu (repo rule), the builder writes the micro sign
SHOW = {"RNAPII": "RNAP II"}
ORDER = [k for _, keys in KINDS for k in keys]
SECTION = {k: kind for kind, keys in KINDS for k in keys}

SLIDES = [   # (title, slug, keys)
    ("H3K27me3, Jurkat", "k27me3_jurkat", [k for k in ORDER if MARK[k] == "H3K27me3" and SECTION[k] == "Jurkat"]),
    ("H3K27me3, CTL (3SI and naive)", "k27me3_ctl", [k for k in ORDER if MARK[k] == "H3K27me3" and SECTION[k] != "Jurkat"]),
    ("H3K9me3", "k9me3", [k for k in ORDER if MARK[k] == "H3K9me3"]),
    ("H3K27ac", "k27ac", [k for k in ORDER if MARK[k] == "H3K27ac"]),
    ("RNAP II", "rnap2", [k for k in ORDER if MARK[k] == "RNAPII"]),
]

CSS = """
body{background:#fff;color:#1d1d1f;font:15px/1.2 "Segoe UI",system-ui,sans-serif;margin:0;padding:8px}
.box{display:inline-block}
table{border-collapse:collapse;font-variant-numeric:tabular-nums}
th,td{padding:0 6px;border-bottom:1px solid #d9d9de;text-align:right;white-space:nowrap}
td.exp{white-space:normal;width:150px;min-width:150px;font-size:14px;font-weight:600;text-align:left;vertical-align:top;background:#f4f4f6;padding:2px 6px}
tr.top td{border-top:2px solid #a9a9b0}
th{font-weight:600;background:#f4f4f6} th.l,td.l{text-align:left}
thead tr:first-child th{border-bottom:none} th.grp{text-align:center;border-left:2px solid #d9d9de}
td.first,th.first{border-left:2px solid #d9d9de}
td.n{color:#6e6e73} td.na{color:#6e6e73;text-align:center} td.ns{color:#6e6e73}
sup{font-size:14px;line-height:0;margin-left:1px;color:#1d1d1f} .se{font-size:15px;color:#6e6e73}
.key{display:flex;flex-wrap:wrap;gap:4px 14px;color:#4a4a4f;font-size:15px;margin:4px 0 0;width:0;min-width:100%}
.sw{display:inline-block;width:22px;height:12px;vertical-align:-1px;margin-right:5px;border-radius:2px}
"""

def first(td):
    return td.replace('<td class="', '<td class="first ', 1)


def halves(keys, data):
    """Split the experiments into two side-by-side tables at the boundary that best balances the rows."""
    sizes = [len(data.get(k, {})) for k in keys]
    if sum(sizes) <= 14:
        return [keys]
    best = min(range(1, len(keys)), key=lambda i: abs(sum(sizes[:i]) - sum(sizes[i:])))
    return [keys[:best], keys[best:]]


def norm(label):
    return label.replace("µ", MU)


def load_cent(figs):
    """paper Fig 6C values: {key: {condition: {"mark"|"DNA": (mean, sem, n)}}}; empty if the CSV is absent."""
    p = Path(figs) / "C_cent_mean_sem.csv"
    out = {}
    if p.exists():
        with open(p, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                out.setdefault(r["key"], {}).setdefault(norm(r["condition"]), {})[r["channel"]] = (
                    float(r["mean"]), float(r["sem"]), int(r["n_cells"]))
    return out


def table_html(keys, data, mk, cent, text):
    t = {k: html.escape(v) for k, v in text["table"].items()}
    with_cent = any(k in cent for k in keys)
    if with_cent:
        h = ['<table><thead><tr><th class="l" rowspan="2">Experiment</th><th class="l" rowspan="2">Arm</th>'
             f'<th class="grp" colspan="3">{t["group_whole"]}</th><th class="grp" colspan="3">{t["group_cent"]}</th></tr>'
             f'<tr><th class="first">n</th><th>{mk}</th><th>DNA</th><th class="first">n</th><th>{mk}</th><th>DNA</th></tr>'
             '</thead><tbody>']
    else:
        h = ['<table><thead><tr><th class="l" rowspan="2">Experiment</th><th class="l" rowspan="2">Arm</th>'
             f'<th rowspan="2">n</th><th class="grp" colspan="2">{t["group_level"]}</th></tr>'
             f'<tr><th class="first">{mk}</th><th>DNA</th></tr></thead><tbody>']
    n_rows = n_cent = 0
    for key in keys:
        arms = data.get(key, {})
        for j, (label, a) in enumerate(arms.items()):
            exp = f'<td class="exp" rowspan="{len(arms)}">{html.escape(TITLE[key])}</td>' if j == 0 else ""
            n_cls = ' class="n first"' if with_cent else ' class="n"'
            row = (f'<tr{" class=top" if j == 0 else ""}>{exp}<td class="l">{html.escape(label)}</td>'
                   f'<td{n_cls}>{a.get("n", "–")}</td>'
                   + (cell(a.get("C_mark"), a.get("C_mark_se"), a.get("C_mark_p"), True) if with_cent else
                      first(cell(a.get("C_mark"), a.get("C_mark_se"), a.get("C_mark_p"), True)))
                   + cell(a.get("C_dna"), a.get("C_dna_se"), a.get("C_dna_p"), True))
            if with_cent:
                c = cent.get(key, {}).get(norm(label))
                if c:
                    row += (f'<td class="n first">{c["mark"][2]}</td>' + cell(c["mark"][0], c["mark"][1])
                            + cell(c["DNA"][0], c["DNA"][1]))
                    n_cent += 1
                else:
                    row += '<td class="n first"></td><td></td><td></td>'
            h.append(row + "</tr>")
            n_rows += 1
    return "".join(h) + "</tbody></table>", n_rows, n_cent


def key_lines(keys, data, cent, text=None):
    """The key lines a table slide shows, in order (deck_text.yaml table.key_*)."""
    t = (text or TEXT)["table"]
    na = any(a.get("C_mark_p") == "NA" or a.get("C_dna_p") == "NA" for k in keys for a in data.get(k, {}).values())
    return ([t["key_main"], t["key_below"], t["key_above"], t["key_stats"]] + ([t["key_na"]] if na else [])
            + ([t["key_cent"]] if any(k in cent for k in keys) else []))


def header_lines(keys, cent, text=None):
    t = (text or TEXT)["table"]
    return [t["group_whole"], t["group_cent"]] if any(k in cent for k in keys) else [t["group_level"]]


def slide_html(title, keys, data, cent, text=None):
    text = text or TEXT
    t = {k: html.escape(v) for k, v in text["table"].items()}
    na = len(key_lines(keys, data, cent, text)) > 4 + any(k in cent for k in keys)
    mk = SHOW.get(MARK[keys[0]], MARK[keys[0]])
    groups = [keys] if any(k in cent for k in keys) else halves(keys, data)
    parts = [table_html(ks, data, mk, cent, text) for ks in groups]
    key = (f'<div class="key"><span>{t["key_main"]}</span>'
           f'<span><span class="sw" style="background:rgba(37,99,235,.55)"></span>{t["key_below"]}&nbsp;&nbsp;'
           f'<span class="sw" style="background:rgba(220,38,38,.55)"></span>{t["key_above"]}</span>'
           f'<span>{t["key_stats"]}{"; " + t["key_na"] if na else ""}</span></div>')
    cent_key = f'<div class="key"><span>{t["key_cent"]}</span></div>' if any(k in cent for k in keys) else ""
    h = (f'<meta charset="utf-8"><title>{html.escape(title)}</title><style>{CSS}</style><div class="box">'
         '<div style="display:flex;gap:28px;align-items:flex-start">' + "".join(p[0] for p in parts) + "</div>"
         + key + cent_key + "</div>")
    return h.replace("µ", MU), sum(p[1] for p in parts), sum(p[2] for p in parts)


def pages(data, cent):
    """(title, slug, keys) per slide; a slide with the centrosome columns becomes two slides, split at the
    experiment boundary that balances rows (FF 2026-11-19: "Split into 2 slides")."""
    out = []
    for title, slug, keys in SLIDES:
        if any(k in cent for k in keys):
            hs = halves(keys, data)
            out += [(f"{title} ({i} of {len(hs)})", f"{slug}_{i}of{len(hs)}", ks) for i, ks in enumerate(hs, 1)]
        else:
            out.append((title, slug, keys))
    return out


def title_group(title):
    """The part of a title before ": " (its prefix), or the whole title."""
    return title.split(": ", 1)[0]


def contents_rows(titles, text=None):
    """[(range, row text)] for the contents slide: each row covers the slides (numbered from 1, title slide first)
    whose title prefix (before ": ") is the row's `titles` (a title without ": " only has to start with it); fails unless every row covers one unbroken block and the rows cover
    slides 3 to the end exactly once (FF 2026-11-19: title page "needs more info", "break into 2 slides").
    titles: the deck's slide titles from slide 2 on."""
    rows, covered = [], []
    for r in (text or TEXT)["contents"]["rows"]:
        p = r["titles"]
        hit = [n for n, t in enumerate(titles, 2) if title_group(t) == p or (": " not in t and t.startswith(p))]
        if not hit or hit != list(range(hit[0], hit[-1] + 1)):
            raise SystemExit(f"contents row {r['text']!r}: titles starting {p!r} are slides {hit}, not one block")
        covered += hit
        rows.append((f"{hit[0]}–{hit[-1]}" if len(hit) > 1 else str(hit[0]), r["text"]))
    if sorted(covered) != list(range(3, len(titles) + 2)):
        raise SystemExit(f"contents rows cover slides {sorted(covered)}, not 3 to {len(titles) + 1} exactly once")
    return rows


def contents_lines(titles, text=None):
    """The contents slide's lines as the slide shows them: rows "a–b  text", then the notes."""
    return ([f"{a}  {t}" for a, t in contents_rows(titles, text)] + list((text or TEXT)["contents"]["notes"]))


def render_contents(out, titles, text=None):
    rows = "".join(f'<tr><td class="k">{html.escape(s)}</td><td>{html.escape(t)}</td></tr>'
                   for s, t in contents_rows(titles, text))
    rows += "".join(f'<tr><td></td><td class="note">{html.escape(n)}</td></tr>' for n in (text or TEXT)["contents"]["notes"])
    page = (f'<meta charset="utf-8"><title>Contents</title><style>body{{background:#fff;color:#1d1d1f;'
            f'font:30px/1.5 "Segoe UI",system-ui,sans-serif;margin:0;padding:8px}}table{{border-collapse:collapse;'
            f'width:auto}}td{{padding:5px 10px;vertical-align:top}}td.k{{font-weight:600;white-space:nowrap}}'
            f'td.note{{color:#6e6e73;padding-top:14px}}</style><table>{rows}</table>')
    hp, pp = Path(out) / "contents.html", Path(out) / "contents.png"
    hp.write_text(page, encoding="utf-8")
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                    "--window-size=1100,600", f"--screenshot={pp}", hp.resolve().as_uri()],
                   check=True, capture_output=True, timeout=120)
    from PIL import Image   # the window is fixed (1100 x 600 at 2x): text touching the bottom edge is cut off
    im = Image.open(pp).convert("L")
    if min(im.crop((0, im.height - 6, im.width, im.height)).getdata()) < 200:
        raise SystemExit(f"{pp}: contents text reaches the bottom edge (cut off)")
    print(pp.name, "| contents slide")


def main(figs, out_dir, cfg_path=DECK_CFG):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    cfg = yaml.safe_load(open(cfg_path, encoding="utf-8"))
    render_contents(out, [e["title"] for e in cfg["images"]])
    for old in list(out.glob("table_*.html")) + list(out.glob("table_*.png")):   # this script's own outputs
        old.unlink()
    data, _ = collect(figs)
    cent = load_cent(figs)
    n_cent_rows = sum(len(v) for v in cent.values())
    placed = total = 0
    for i, (title, slug, keys) in enumerate(pages(data, cent), 1):
        page, n, nc = slide_html(title, keys, data, cent)
        placed += nc
        hp, pp = out / f"table_{i}_{slug}.html", out / f"table_{i}_{slug}.png"
        hp.write_text(page, encoding="utf-8")
        subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                        "--window-size=2600,1150", f"--screenshot={pp}", hp.resolve().as_uri()],
                       check=True, capture_output=True, timeout=120)
        if not pp.exists():
            raise RuntimeError(f"no screenshot written: {pp}")
        no_c = sum(1 for k in keys for a in data.get(k, {}).values() if "C_mark" not in a)
        print(f"{pp.name} | {title}: {len(keys)} experiments, {n} arm rows, {no_c} without a nadir level, "
              f"{nc} with a centrosome value")
        total += n
    print("total arm rows", total, "| centrosome arms placed", placed, "of", n_cent_rows)
    if placed != n_cent_rows:
        raise SystemExit("ERROR: some centrosome rows matched no table arm")


if __name__ == "__main__":
    main(*sys.argv[1:4])
