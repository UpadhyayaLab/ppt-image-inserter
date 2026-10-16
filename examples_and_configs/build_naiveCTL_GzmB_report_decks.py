"""
build_naiveCTL_GzmB_report_decks.py

Send-ready Granzyme B decks for one naive-CTL GzmB experiment, in the style of
FF's June 17, 2025 package (K:/FF/PPT/Centrosome_Lytic_Granules_PPT/):

  GrzmB_naive_CTL_<date>.pptx                  <- like GrzmB_naive_CTL_20260611.pptx
  Actin_GrzmB_naive_CTL_<date>_montages.pptx   <- like Actin_GrzmB_naive_CTL_20260611_montages.pptx

Both start from FF's June files, so the title / "Experimental info" / "Analysis
pipeline" slides keep their fonts and formatting (text edited at run level).

Analysis deck: same slide order, titles and annotation boxes as June. Each June
"effects of activation" slide becomes two slides (no activation comparison in
these experiments): glass vs glass 1-16 timecourses (2-up) and glass vs glass
1-16 at matched time (3-up, or 4-up with four timepoints). Images are inserted
by batch_insert_images.py from a generated YAML.

Montage deck: actin XZ MIP, GzmB XZ MIP (clim 10-100), actin at synapse, GzmB
at synapse (clim 10-100), one slide per condition, via the helpers of
insert_GzmB_qc_xz_mip_synapse_slides.py (first real montage chunk, black slide).

Usage:
    conda run -n PPT_editing python examples_and_configs/build_naiveCTL_GzmB_report_decks.py --dataset 10292025
"""

import argparse
import copy
import io
import os
import shutil
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import insert_GzmB_qc_xz_mip_synapse_slides as gq  # noqa: E402

JUNE_DIR = "K:/FF/PPT/Centrosome_Lytic_Granules_PPT/"
JUNE_ANALYSIS = JUNE_DIR + "GrzmB_naive_CTL_20260611.pptx"
JUNE_MONTAGES = JUNE_DIR + "Actin_GrzmB_naive_CTL_20260611_montages.pptx"
OUT_DIR = "K:/FF/PPT/PPT_autogeneration/Naive_CTL/"
CONFIG_DIR = os.path.join(HERE, "configs", "fixed", "CTLs")

DATASETS = {
    "10292025": dict(
        exp_date="October 29, 2025",
        mark="H3K27ac",
        results_root=("L:/FF/Naive_CTLs/GzB_histone_marks/tifsnaiveCTLs_H3K27acGrzmB/"
                      "10292025_naiveCTLs_/compiled_results/"
                      "naiveCTLs_GzmB_10292025_12conditions_20260930"),
        gel_times="28, 50, 74 ",
        glass_times="28, 50, 74 hr",
        pll=False,
        substrate=["substrate_28h", "substrate_50h", "substrate_74h"],
        substrate_label="28 / 50 / 74 h",
        coating=["coating_28h", "coating_50h", "coating_74h"],
        coating_label="28 / 50 / 74 h",
        exclusion_note=None,
    ),
    "10152025": dict(
        exp_date="October 15, 2025",
        mark="H3K9me3",
        results_root=("L:/FF/Naive_CTLs/GzB_histone_marks/tifsnaiveCTLs_H3K9me3GrzmB/"
                      "10152025_naiveCTLs_/compiled_results/"
                      "naiveCTLs_GzmB_10152025_15conditions_20260930"),
        gel_times="24, 48, 72 ",
        glass_times="6, 24, 48, 72 hr",
        pll=True,
        substrate=["substrate_24h", "substrate_48h", "substrate_72h"],
        substrate_label="24 / 48 / 72 h",
        coating=["coating_6h", "coating_24h", "coating_48h", "coating_72h"],
        coating_label="6 / 24 / 48 / 72 h",
        exclusion_note=("7 objects in PLL 2 hr with outlying Granzyme B signal were "
                        "excluded (6 debris objects, 1 imaging artifact)."),
    ),
}

# Montage kinds in June's order: (title prefix, channel subdir, montage subpath).
MONTAGE_KINDS = [
    ("Actin XZ MIP", "actin", "xz_mip/montages"),
    ("GzmB XZ MIP (clim 10-100)", "GzmB", "xz_mip/montages_clim10_100"),
    ("Actin at Synapse", "actin", "synapse/mask/montages"),
    ("GzmB at Synapse (clim 10-100)", "GzmB", "synapse/slice/montages_clim10_100"),
]

# Annotation boxes, verbatim from June (position/size of June slide 12).
NOTE_GATE = ("Cells with peak Granzyme B signal of < 20 after bg subtraction and "
             "smoothing removed from consideration")
NOTE_AUTOCORR = ("Cells with peak Granzyme B signal of < 20 after bg subtraction and "
                 "smoothing or synapse MFI < 5 removed from consideration")
# The 3-slice MIP g(r) is gated on the 3-slice MIP synapse MFI (> 10), not 5.
NOTE_AUTOCORR_3MIP = ("Cells with peak Granzyme B signal of < 20 after bg subtraction and "
                      "smoothing or 3-slice MIP synapse MFI < 10 removed from consideration")
NOTE_YLIM = ("Hard to interpret with these ylimits: see later binary plots of % of "
             "cells with Granzyme B")
NOTE_BOX = (0.06, 6.64, 7.01, 0.77)
NOTE_PT = 20

# Layout slides appended to the analysis template (picture geometry from June).
SQUARE_TOP = 1.18
LAYOUTS = {
    "wide": [(0.32, 1.90, 12.70, 4.41)],
    "3up": [(0.22, 2.30, 4.10, 4.10), (4.62, 2.30, 4.10, 4.10), (9.02, 2.30, 4.10, 4.10)],
    "square": [(3.97, SQUARE_TOP, 5.40, 5.40)],
    "2up": [(0.81, SQUARE_TOP, 5.40, 5.40), (6.77, SQUARE_TOP, 5.40, 5.40)],
    "4up": [(0.117 + k * 3.3, 2.75, 3.20, 3.20) for k in range(4)],
}
N_INTRO = 3  # title, experimental info, analysis pipeline


# ---------------------------------------------------------------------------
# pptx helpers
# ---------------------------------------------------------------------------
def delete_slides_except(prs, keep):
    """Delete every slide whose 0-based index is not in keep."""
    sld_ids = prs.slides._sldIdLst
    for i, sld_id in reversed(list(enumerate(list(sld_ids)))):
        if i not in keep:
            prs.part.drop_rel(sld_id.rId)
            sld_ids.remove(sld_id)


def replace_in_runs(slide, old, new, count_expected=1):
    """Replace text inside single runs (keeps run formatting)."""
    n = 0
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for run in para.runs:
                if old in run.text:
                    run.text = run.text.replace(old, new)
                    n += 1
    if n != count_expected:
        raise SystemExit(f"Expected {count_expected} run(s) containing {old!r}, found {n}.")


def find_paragraph(slide, startswith):
    for sh in slide.shapes:
        if sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                if para.text.startswith(startswith):
                    return para
    raise SystemExit(f"No paragraph starting with {startswith!r}.")


def add_paragraph_after(para, runs, after=None):
    """Clone para (first two runs: bold label + plain text) with new text, placed
    after `after` (default: after para itself)."""
    new_p = copy.deepcopy(para._p)
    (after or para)._p.addnext(new_p)
    r_elems = [c for c in new_p if c.tag.endswith("}r")]
    for r in r_elems[len(runs):]:
        new_p.remove(r)
    for r, text in zip(r_elems, runs):
        r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}t").text = text


def edit_intro_slides(prs, ds, montage_ranges):
    """Edit June's title / info (/ pipeline) slides for this experiment."""
    s_title, s_info = prs.slides[0], prs.slides[1]
    replace_in_runs(s_title, "June 17, 2025", ds["exp_date"])

    replace_in_runs(s_info, ", actin (phalloidin), β-tubulin antibody",
                    ", actin (phalloidin). Also imaged: {} antibody (not analyzed here)".format(ds["mark"]))
    glass_p = find_paragraph(s_info, "Glass: ")
    glass_tail = "2, 6, 24, 48, 72 hr. Also non-activated (PLL) control at 2 hr."
    # June's PLL sentence stays on the Glass line, as in June.
    replace_in_runs(s_info, glass_tail, ds["glass_times"] +
                    (". Also non-activated (PLL) control at 2 hr." if ds["pll"] else ""))
    for gel in ("1.5 kPa: ", "12 kPa: "):
        para = find_paragraph(s_info, gel)
        hits = [r for r in para.runs if r.text == "24, 48, 72 "]
        if len(hits) != 1:
            raise SystemExit(f"Gel timepoint run not found in {gel!r} line.")
        hits[0].text = ds["gel_times"]
    add_paragraph_after(glass_p, ["Glass 1-16: ", ds["glass_times"]])
    if ds["exclusion_note"]:
        last = find_paragraph(s_info, "Glass 1-16: ")
        add_paragraph_after(last, ["Note: ", ds["exclusion_note"]])

    if len(prs.slides) > 2 and montage_ranges:
        s_pipe = prs.slides[2]
        replace_in_runs(s_pipe, "See slides 3-14 of", "See slides {} of".format(montage_ranges["Actin XZ MIP"]))
        replace_in_runs(s_pipe, "slides 27-38 of", "slides {} of".format(montage_ranges["Actin at Synapse"]))
        replace_in_runs(s_pipe, "slides 39-50 of", "slides {} of".format(montage_ranges["GzmB at Synapse (clim 10-100)"]))
        # Clone the bold-label + plain-text formatting of the "Granzyme B total
        # signal:" paragraph, but place the new line last.
        add_paragraph_after(find_paragraph(s_pipe, "Granzyme B total signal:"), [
            "Changes since the June analysis: ",
            "cells whose actin reaches the top of the stack are excluded from whole-cell and "
            "axial Granzyme B metrics; neighboring cells merged into a cell's mask at the crop "
            "edge are trimmed using Cellpose. Granzyme B parameters and thresholds unchanged."],
            after=find_paragraph(s_pipe, "Synapse plane segmented"))


def add_note(slide, text):
    left, top, w, h = NOTE_BOX
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for run in tf.paragraphs[0].runs:
        run.font.size = Pt(NOTE_PT)


# ---------------------------------------------------------------------------
# Analysis deck
# ---------------------------------------------------------------------------
def slide_spec(ds):
    """Ordered [(title, [rel image paths], layout, note)] in June's order."""
    root = ds["results_root"]
    coat_layout = "4up" if len(ds["coating"]) == 4 else "3up"

    def grid(group, raw):
        return "grid_panels/{}/{}.png".format(group, raw)

    def allcond(raw, flat_dir):
        rel = grid("all_conditions", raw)
        if flat_dir and not os.path.exists(os.path.join(root, rel)):
            rel = "{}/{}.png".format(flat_dir, raw)
        return rel

    def panels(raw, pretty, note, allcond_layout="wide", flat_dir=None, allcond_imgs=None):
        """The five panel slides of one metric, as {kind: slide}."""
        coating, coat_lab, coat_lay = ds["coating"], ds["coating_label"], coat_layout
        if note and len(coating) == 4:
            # Gated metrics: glass 6 hr has ~1 GzmB+ cell, so its violin drops out
            # and the 6 hr panel shows glass 1-16 alone -> show 24/48/72 hr only.
            coating, coat_lay = coating[1:], "3up"
            coat_lab = coat_lab.split(" / ", 1)[1]
        return {
            "all": ("{} — all conditions".format(pretty),
                    allcond_imgs or [allcond(raw, flat_dir)], allcond_layout, note),
            "substrate": ("{} — substrate comparison ({})".format(pretty, ds["substrate_label"]),
                          [grid(g, raw) for g in ds["substrate"]], "3up", note),
            "timecourse": ("{} — timecourses: 1.5 kPa, 12 kPa, glass".format(pretty),
                           [grid(g, raw) for g in ("gel_1p5kPa_timecourse", "gel_12kPa_timecourse",
                                                   "glass_timecourse")], "3up", note),
            "glass116_tc": ("{} — timecourses: glass vs glass 1-16".format(pretty),
                            [grid(g, raw) for g in ("glass_timecourse", "glass_1-16_timecourse")],
                            "2up", note),
            "coating": ("{} — glass vs glass 1-16 ({})".format(pretty, coat_lab),
                        [grid(g, raw) for g in coating], coat_lay, note),
        }

    ORDER = ["all", "substrate", "timecourse", "glass116_tc", "coating"]
    spec = []

    # Total signal (ungated) + % high total (June slides 4-11).
    p = panels("GzmB_total_sig", "Granzyme B: total signal", None)
    p["timecourse"] = p["timecourse"][:3] + (NOTE_YLIM,)
    p["glass116_tc"] = p["glass116_tc"][:3] + (NOTE_YLIM,)
    spec += [p[k] for k in ORDER]
    p = panels("GzmB_total_sig_fraction", "Granzyme B: % cells with high total signal (> 1M)",
               None, flat_dir="polarization")
    spec += [p[k] for k in ORDER]

    # Z50 / Z75 interleaved per panel kind, then the norm. cell height pair (June 12-23).
    for suffix, tag in (("", ""), ("_norm_cell_height", " (norm. cell height)")):
        p50 = panels("GzmB_z50_cell_bottom_distance" + suffix, "Granzyme B Z₅₀" + tag, NOTE_GATE)
        p75 = panels("GzmB_z75_cell_bottom_distance" + suffix, "Granzyme B Z₇₅" + tag, NOTE_GATE)
        for k in ORDER:
            spec += [p50[k], p75[k]]

    # Axial profiles: all-conditions 2-up (average + all cells), then the average
    # set and the all-cells set (June 24-28).
    avg, allc = "GzmB_axial_profile_auc1", "GzmB_axial_profile_auc1_all_cells_with_average"
    pa = panels(avg, "Granzyme B: axial intensity profile", NOTE_GATE)
    pc = panels(allc, "Granzyme B: axial intensity profile (all cells)", NOTE_GATE)
    spec.append(("Granzyme B: axial intensity profile — all conditions",
                 [allcond(avg, "profiles"), allcond(allc, "profiles")], "2up", NOTE_GATE))
    spec += [pa[k] for k in ORDER[1:]] + [pc[k] for k in ORDER[1:]]

    # Z50 fraction polarized, 3 um (June 29-31).
    p = panels("GzmB_z50_cell_bottom_distance_3um_polarization",
               "Granzyme B Z₅₀: fraction polarized (3 μm thresh)", NOTE_GATE, flat_dir="polarization")
    spec += [p[k] for k in ORDER]

    # Synapse total signal, plain + 3-slice MIP (June 32-39). Ungated (the GzmB
    # peak gate excludes these fields), so no gate note.
    for raw, pretty in (("GzmB_synapse_total_sig", "Granzyme B: total signal at synapse"),
                        ("GzmB_synapse_total_sig_3mip", "Granzyme B: total signal at synapse (3-slice MIP)")):
        p = panels(raw, pretty, None)
        spec += [p[k] for k in ORDER]

    # Pair autocorrelation: all cells, then 3-slice MIP average (June 40-45).
    p = panels("GzmB_synapse_autocorr_c_all_cells_with_average",
               "Granzyme B: pair autocorrelation at synapse (all cells)", NOTE_AUTOCORR,
               allcond_layout="square", flat_dir="profiles")
    spec += [p[k] for k in ORDER]
    p = panels("GzmB_synapse_autocorr_c_3mip",
               "Granzyme B: pair autocorrelation at synapse, 3-slice MIP", NOTE_AUTOCORR_3MIP,
               allcond_layout="square", flat_dir="profiles")
    spec += [p[k] for k in ORDER]

    missing = [r for _, rels, _, _ in spec for r in rels if not os.path.exists(os.path.join(root, r))]
    if missing:
        raise SystemExit("Missing images:\n  " + "\n  ".join(missing))
    return spec


def build_analysis_template(ds, path, montage_ranges):
    shutil.copyfile(JUNE_ANALYSIS, path)
    prs = Presentation(path)
    june_layout = prs.slides[3].slide_layout
    blob = next(sh for sh in prs.slides[3].shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE).image.blob
    delete_slides_except(prs, set(range(N_INTRO)))
    edit_intro_slides(prs, ds, montage_ranges)
    layout_index = {}
    for name, boxes in LAYOUTS.items():
        slide = prs.slides.add_slide(june_layout)
        for ph in list(slide.placeholders):
            ph._element.getparent().remove(ph._element)
        for (l, t, w, h) in boxes:
            slide.shapes.add_picture(io.BytesIO(blob), Inches(l), Inches(t), Inches(w), Inches(h))
        layout_index[name] = len(prs.slides) - 1
    prs.save(path)
    return layout_index


def yq(text):
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def build_analysis_deck(name, ds, montage_ranges):
    out_path = OUT_DIR + "GrzmB_naive_CTL_{}.pptx".format(name)
    template = OUT_DIR + "GrzmB_naive_CTL_{}_template.pptx".format(name)
    layout_index = build_analysis_template(ds, template, montage_ranges)
    spec = slide_spec(ds)

    n_slides = N_INTRO + len(LAYOUTS)
    lines = [
        "# AUTO-GENERATED by examples_and_configs/build_naiveCTL_GzmB_report_decks.py --dataset " + name,
        "presentation: " + yq(template),
        "output_path: " + yq(out_path),
        "template_slide: {}".format(layout_index["wide"]),
        "preserve_slides: [{}]".format(", ".join(str(i) for i in range(n_slides))),
        "base_dir: " + yq(ds["results_root"]),
        "auto_position: true",
        "add_label: false",
        "",
        "images:",
    ]
    for title, rels, layout, _ in spec:
        lines.append("  - images: [{}]".format(", ".join(yq(r) for r in rels)))
        lines.append("    template_slide: {}".format(layout_index[layout]))
        lines.append("    title: {}".format(yq(title)))
    yaml_path = os.path.join(CONFIG_DIR, "config_GrzmB_naive_CTL_{}_report.yaml".format(name))
    with open(yaml_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")

    subprocess.run([sys.executable, os.path.join(HERE, "batch_insert_images.py"), yaml_path], check=True)

    # Post-process: drop the layout slides, add June's annotation boxes.
    prs = Presentation(out_path)
    delete_slides_except(prs, set(range(N_INTRO)) | set(range(n_slides, len(prs.slides))))
    body = list(prs.slides)[N_INTRO:]
    if len(body) != len(spec):
        raise SystemExit(f"{len(body)} content slides, expected {len(spec)}.")
    for slide, (title, _, _, note) in zip(body, spec):
        text = " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)
        if title not in text:
            raise SystemExit(f"Slide order mismatch: expected {title!r}, found {text!r}.")
        if note:
            add_note(slide, note)
    prs.save(out_path)
    os.remove(template)
    print(f"Analysis deck: {out_path}  ({len(prs.slides)} slides)")
    return out_path


# ---------------------------------------------------------------------------
# Montage deck
# ---------------------------------------------------------------------------
def build_montage_deck(name, ds, qc):
    out_path = OUT_DIR + "Actin_GrzmB_naive_CTL_{}_montages.pptx".format(name)
    shutil.copyfile(JUNE_MONTAGES, out_path)
    prs = Presentation(out_path)
    delete_slides_except(prs, {0, 1})
    edit_intro_slides(prs, ds, None)

    root = Path(qc["dataset_root"])
    ranges, problems = {}, []
    for kind, channel, sub in MONTAGE_KINDS:
        first = len(prs.slides) + 1
        for label, folder in qc["conditions"]:
            chunk, fell_back = gq.find_first_real_chunk(root / folder / "prog_CTL_analysis" / channel / sub)
            if chunk is None or fell_back:
                problems.append(f"{kind} / {label}: {'missing' if chunk is None else 'smoke fallback'}")
            slide, _ = gq.build_slide(prs, f"{kind}: {label}  ({qc['date_tag']})", chunk)
            if first == 3 and len(prs.slides) == 3:
                gq.add_textbox(slide, "All scale bars 5 μm", 0.41, 0.10, 2.48, 0.41,
                               font_pt=22, color=RGBColor(0xFF, 0xFF, 0xFF), bold=True)
        ranges[kind] = f"{first}-{len(prs.slides)}"
    prs.save(out_path)
    print(f"Montage deck: {out_path}  ({len(prs.slides)} slides)  ranges: {ranges}")
    if problems:
        print("Montage problems:\n  " + "\n  ".join(problems))
    return ranges


def main():
    ap = argparse.ArgumentParser(description="Send-ready naive-CTL GzmB decks (June 2026 style)")
    ap.add_argument("--dataset", required=True, choices=sorted(DATASETS))
    name = ap.parse_args().dataset
    ds, qc = DATASETS[name], gq.DATASETS[name]
    ranges = build_montage_deck(name, ds, qc)
    build_analysis_deck(name, ds, ranges)


if __name__ == "__main__":
    main()
