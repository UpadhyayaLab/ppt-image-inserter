"""Render the Figure 6 summary table (mean +- SEM) as 5 slide images, one per mark group.

FF 2026-11-19: slides = H3K27me3 Jurkat, H3K27me3 CTL (3SI + naive), H3K9me3 (all), H3K27ac (all), RNAP II (all);
"we should keep table simple. just show the rel level near nadir. that's all" -> one measure, mark and DNA
("show DNA along with h3k27me3"); "you should include centrosome version of it when that was imaged. in other cases -
just leave blank" / "we need to show what is c in figure 6" -> a "Near centrosome" group = paper Fig 6C's measure
(seed study S2, "Region Nadir near Centrosome"), from <figs_dir>/C_cent_mean_sem.csv (fig6_C_cent_mean_sem.m, Histone
mark data session); only experiments in that CSV have it, so only slides holding one get the extra columns.

Values come from the Histone mark data session's table builder (collect(), cell()) in TCell-3D-Morphodynamics,
imported unchanged: the C level (deepest invagination region / whole-surface average), mean +- SEM over cells.

Usage:  python render_fig6_table_slides.py <figs_dir> <out_dir>
Writes <out_dir>/table_<n>_<slug>.html and .png (headless Edge screenshot) and prints rows per slide.
"""
import csv
import html
import subprocess
import sys
from pathlib import Path

BUILDER = Path("C:/Users/Upadhyaya_Lab/Code/TCell-3D-Morphodynamics/dev/reporting/fig6_metrics_page")
sys.path.insert(0, str(BUILDER))
from build_fig6_summary_table import TITLE, MARK, cell, collect  # noqa: E402
from build_fig6_metrics_page import KINDS  # noqa: E402

EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
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

# Wording from the verified definitions (TCell-3D-Morphodynamics dev/reporting/fig6_metrics_page/fig6_definitions.yaml:
# relative_level_near_nadir_C, near_centrosome_value, table_statistics), approved by FF 2026-11-19.
KEY = ('<div class="key"><span>Mean in the 0.25 μm layer near the nadir ÷ the same layer, whole nucleus '
       "(1 = no enrichment); mean ± SEM over cells</span>"
       '<span><span class="sw" style="background:rgba(37,99,235,.55)"></span>below 1&nbsp;&nbsp;'
       '<span class="sw" style="background:rgba(220,38,38,.55)"></span>above 1</span>'
       "<span>* &lt;0.05, ** &lt;0.01, *** &lt;0.001 vs 1 (signed-rank, BH-adjusted); gray = n.s.{na}</span></div>")
CENT_KEY = ('<div class="key"><span>Near centrosome: disk around the deepest point of the centrosome-proximal patch; '
            "no test; blank = not analyzed</span></div>")


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


def table_html(keys, data, mk, cent):
    with_cent = any(k in cent for k in keys)
    if with_cent:
        h = ['<table><thead><tr><th class="l" rowspan="2">Experiment</th><th class="l" rowspan="2">Arm</th>'
             '<th class="grp" colspan="3">Whole nucleus</th><th class="grp" colspan="3">Near centrosome</th></tr>'
             f'<tr><th class="first">n</th><th>{mk}</th><th>DNA</th><th class="first">n</th><th>{mk}</th><th>DNA</th></tr>'
             '</thead><tbody>']
    else:
        h = ['<table><thead><tr><th class="l" rowspan="2">Experiment</th><th class="l" rowspan="2">Arm</th>'
             '<th rowspan="2">n</th><th class="grp" colspan="2">Relative level near nadir</th></tr>'
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


def slide_html(title, keys, data, cent):
    mk = SHOW.get(MARK[keys[0]], MARK[keys[0]])
    groups = [keys] if any(k in cent for k in keys) else halves(keys, data)
    parts = [table_html(ks, data, mk, cent) for ks in groups]
    h = (f'<meta charset="utf-8"><title>{html.escape(title)}</title><style>{CSS}</style><div class="box">'
         '<div style="display:flex;gap:28px;align-items:flex-start">' + "".join(p[0] for p in parts) + "</div>"
         + KEY.replace("{na}", "; n/a = too few cells to test" if any(
             a.get("C_mark_p") == "NA" or a.get("C_dna_p") == "NA" for k in keys for a in data.get(k, {}).values()) else "")
         + (CENT_KEY if any(k in cent for k in keys) else "")
         + "</div>")
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


# Contents slide (FF 2026-11-19: title page "needs more info", "break into 2 slides"; "keep labels concise, period.
# no need to mention C, D, E, F"; "do not mention voxels"). Slide numbers count the title slide (1) and this slide
# (2). Nadir facts from the Histone mark data session 2026-11-19 (compute_nuclear_concavity.m:150 whole
# nucleus; slides 3-4 = the centrosome-proximal seed). (slides, text, title regex): the deck check confirms exactly those slides' titles match the regex.
CONTENTS = [   # one short phrase per row (FF 2026-11-19: "clean up slide 2!!!! this is slop"); no character limit (FF 2026-11-19)
    ("3–4", "Nadir: surface point farthest from the hull boundary", "The regions measured"),
    ("5–10", "Summary tables", "Relative level near nadir, mean ± SEM"),
    ("11–31", "Relative level near nadir", "Relative level near nadir: "),
    ("32–52", "Relative level vs distance from nadir, along the surface", "Relative level vs distance from nadir"),
    ("53–73", "Relative level vs depth below convex hull, ÷ its 0–0.25 μm bin", "Relative level vs depth below convex hull"),
]
# Definition lines, verbatim as approved by FF 2026-11-19 (fig6_definitions.yaml: shell_and_face_density,
# relative_level_near_nadir_C, region_near_nadir_C). Check: <= 100 characters each.
CONTENTS_NOTES = [
    "Level: mean in the 0.25 μm layer inside the nuclear surface ÷ that layer over the whole nucleus",
    "Near nadir: deeper than half the nadir's depth, within 25% of the mean surface distance from it",
]


def render_contents(out):
    rows = "".join(f'<tr><td class="k">{html.escape(s)}</td><td>{html.escape(t)}</td></tr>' for s, t, _ in CONTENTS)
    rows += "".join(f'<tr><td></td><td class="note">{html.escape(n)}</td></tr>' for n in CONTENTS_NOTES)
    page = (f'<meta charset="utf-8"><title>Contents</title><style>body{{background:#fff;color:#1d1d1f;'
            f'font:30px/1.5 "Segoe UI",system-ui,sans-serif;margin:0;padding:8px}}table{{border-collapse:collapse;'
            f'width:auto}}td{{padding:5px 10px;vertical-align:top}}td.k{{font-weight:600;white-space:nowrap}}'
            f'td.note{{color:#6e6e73;padding-top:14px}}</style><table>{rows}</table>')
    hp, pp = out / "contents.html", out / "contents.png"
    hp.write_text(page, encoding="utf-8")
    subprocess.run([EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=2",
                    "--window-size=1100,600", f"--screenshot={pp}", hp.resolve().as_uri()],
                   check=True, capture_output=True, timeout=120)
    print(pp.name, "| contents slide")


def main(figs, out_dir):
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    render_contents(out)
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
    main(*sys.argv[1:3])
