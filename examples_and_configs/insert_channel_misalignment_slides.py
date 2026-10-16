"""
insert_channel_misalignment_slides.py

Group-meeting deck on the ERa / RNAP2 channel misalignment (MCF7 project):
the bead offsets for BOTH bead sets (2024-05-22 TetraSpeck = Aashli's data,
2026-09 bead_test), the bead-free correction end to end, and the discussion
points (other microscopes, correction in post-processing, what is affected).

Material is copied from the existing reports, never redrawn:
  intranuclear_colocalization/docs/Bead_Offset_Report_TetraSpeck_20240522.pdf
  intranuclear_colocalization/docs/Bead_Offset_Report.pdf
  intranuclear_colocalization/docs/Registration_Method_Validation.pdf
Figures that exist as full-resolution PNGs in the study folders are used as is
(white margins trimmed); figures that exist only inside a PDF are rendered at
300 dpi with Ghostscript and cropped. Every number in a slide title is computed
here from the studies' own CSV / Summary files (or read from the PDF text), and
printed, so the titles cannot drift from the reports.

Usage:
    conda run -n PPT_editing python examples_and_configs/insert_channel_misalignment_slides.py
    ... --list     # dry run: print every slide title and source, build nothing
"""

import json
import os
import re
import subprocess
import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import insert_chromark_h3k27me3_summary_slides as M  # noqa: E402

# ---------------------------------------------------------------------------
# Sources
# ---------------------------------------------------------------------------
DOCS = Path("C:/Users/Upadhyaya_Lab/Code/intranuclear_colocalization/docs")
PDF = {
    "tetra": DOCS / "Bead_Offset_Report_TetraSpeck_20240522.pdf",
    "bead2026": DOCS / "Bead_Offset_Report.pdf",
    "valid": DOCS / "Registration_Method_Validation.pdf",
}
BEAD = {
    "2024": Path("L:/FF/MCF7_fixed_cells/2026_Channel_Misalignment_Tests/"
                 "20240522_Bead-Density_Titration/100nm_TetraSpeck_1to10/"
                 "Surface_Step0p3um_Range5um_FOV001/Bead_Offset_Study_20260917_123515"),
    "2026": Path("L:/FF/MCF7_fixed_cells/bead_test/Bead_Offset_Study_20260902_155249"),
}
BEAD_LABEL = {"2024": "2024-05-22 TetraSpeck beads, gel surface (Aashli)",
              "2026": "2026-09 bead_test beads on glass (Vishavdeep)"}
GEL_ROOT = Path("L:/FF/Beads_alignment/20201014_Bead_alignment_test_gel")   # Vishavdeep's 2026-10-14 gel beads
GEL_LABEL = "2026-10-14 gel beads (Vishavdeep)"
CELLS = Path("L:/FF/MCF7_fixed_cells/20260613_MCF7_ERa_RNAP2_1kPa_E2_vs_EtOH/Coloc_Analysis")
ROUTE1 = CELLS / "Route1_Coherence_Study_benitez_20260903_104608"
VERIFY = CELLS / "Registration_Verify_ours_20260903_125750"
GS = "C:/Program Files/gs/gs10.00.0/bin/gswin64c.exe"

OUT_DIR = Path("K:/FF/PPT/PPT_autogeneration/MCF7_channel_misalignment")
OUT = OUT_DIR / "Channel_Misalignment_GroupMeeting.pptx"
IMG_DIR = OUT_DIR / "source_images"
PANEL_DIR = IMG_DIR / "beadpanels"   # from intranuclear_colocalization/scripts/render_bead_slide_panels.py
PANELS = json.loads((PANEL_DIR / "panels.json").read_text())
CH_NAMES = ["640 nm", "561 nm", "488 nm", "405 nm"]
PAGE_DIR = IMG_DIR / "pages"
CROP_DIR = IMG_DIR / "crops"

PX_NM = 65.0
FIELDS = ["EtOH/01", "EtOH/02", "E2/01", "E2/02"]
SKIP_NUCLEI = {("E2/01", 14)}   # FF 2026-10-02: unusual nucleus, left out of the examples

# PDF figure crops: (key, pdf, page, (x0, y0, x1, y1) as fractions of the page)
CROPS = {
    "tetra_shift_vs_dist": ("tetra", 2, (0.055, 0.190, 0.945, 0.765)),
    "bead2026_shift_vs_dist": ("bead2026", 2, (0.050, 0.285, 0.945, 0.740)),
    "tetra_vs_2026": ("tetra", 16, (0.040, 0.425, 0.960, 0.705)),
    "valid_planted": ("valid", 5, (0.040, 0.300, 0.955, 0.635)),
    "valid_consistency": ("valid", 24, (0.030, 0.315, 0.910, 0.610)),
    "valid_fit": ("valid", 29, (0.030, 0.180, 0.600, 0.455)),
    "valid_overlap": ("valid", 25, (0.035, 0.255, 0.940, 0.775)),
    "valid_ccf_etoh": ("valid", 27, (0.170, 0.440, 0.840, 0.785)),
    "valid_ccf_e2": ("valid", 28, (0.170, 0.345, 0.840, 0.715)),
}

# ---------------------------------------------------------------------------
# Layout: large text, large images
# ---------------------------------------------------------------------------
W, H = M.SLIDE_W, M.SLIDE_H
BLACK = RGBColor(0, 0, 0)
GREY = RGBColor(0x70, 0x70, 0x70)
BLUE = RGBColor(0x1F, 0x5F, 0xA8)
MARGIN = 0.12
TITLE_TOP, TITLE_H = 0.05, 1.30
IMG_TOP = TITLE_TOP + TITLE_H + 0.04
FOOT_H = 0.30
IMG_H = H - IMG_TOP - FOOT_H - 0.02
IMG_W = W - 2 * MARGIN
FOOT_PT = 13
BULLET_PT = 34
SIDE_ASPECT = 1.5      # images narrower than this (w/h) go full height, title beside them
SIDE_TEXT_W = 4.3


def title_pt(text):
    n = len(text)
    return 44 if n <= 45 else 40 if n <= 80 else 36


def add_title(slide, text, pt=None):
    M.add_textbox(slide, text, MARGIN, TITLE_TOP, IMG_W, TITLE_H,
                  font_pt=pt or title_pt(text), color=BLACK, bold=True)
    box = slide.shapes[-1]
    box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE


def add_footer(slide, text, left=MARGIN, width=IMG_W):
    M.add_textbox(slide, text, left, H - FOOT_H, width, FOOT_H,
                  font_pt=FOOT_PT, color=GREY, align=PP_ALIGN.RIGHT)


def image_slide(prs, title, img, source):
    s = M._new_slide(prs)
    with Image.open(img) as im:
        aspect = im.width / im.height
    if aspect >= SIDE_ASPECT:
        add_title(s, title)
        M.add_image_in_box(s, str(img), MARGIN, IMG_TOP, IMG_W, IMG_H)
        add_footer(s, source)
        return s
    # tall figure: full slide height on the left, title in large type on the right
    img_h = H - 2 * MARGIN - FOOT_H
    img_w = min(W - SIDE_TEXT_W - 3 * MARGIN, img_h * aspect)
    M.add_image_in_box(s, str(img), MARGIN, MARGIN, img_w, img_h)
    tx = MARGIN + img_w + 0.25
    tw = W - tx - MARGIN
    M.add_textbox(s, title, tx, MARGIN, tw, img_h, font_pt=40, color=BLACK, bold=True,
                  align=PP_ALIGN.LEFT)
    s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    add_footer(s, source)
    return s


def pair_slide(prs, title, imgs, labels, source):
    s = M._new_slide(prs)
    add_title(s, title)
    gap = 0.15
    with Image.open(imgs[0]) as im:
        wide = im.width / im.height > 1.6
    if wide:   # wide figures: stacked, label at the left
        lab_w = 2.2
        row_h = (IMG_H - gap) / 2
        for i, (img, lab) in enumerate(zip(imgs, labels)):
            top = IMG_TOP + i * (row_h + gap)
            M.add_textbox(s, lab, MARGIN, top, lab_w, row_h, font_pt=28, color=BLUE, bold=True)
            s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            M.add_image_in_box(s, str(img), MARGIN + lab_w, top, IMG_W - lab_w, row_h)
        add_footer(s, source)
        return s
    ncol = len(imgs)
    col_w = (IMG_W - (ncol - 1) * gap) / ncol
    for i, (img, lab) in enumerate(zip(imgs, labels)):
        left = MARGIN + i * (col_w + gap)
        M.add_textbox(s, lab, left, IMG_TOP, col_w, 0.5, font_pt=24, color=BLUE, bold=True)
        M.add_image_in_box(s, str(img), left, IMG_TOP + 0.5, col_w, IMG_H - 0.5)
    add_footer(s, source)
    return s


def text_slide(prs, title, bullets, source=""):
    s = M._new_slide(prs)
    add_title(s, title, pt=40)
    box = s.shapes.add_textbox(Inches(MARGIN + 0.4), Inches(IMG_TOP + 0.2),
                               Inches(IMG_W - 0.8), Inches(IMG_H - 0.3))
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        sub = b.startswith("  ")
        p.text = ("\u2013 " if sub else "\u2022 ") + b.strip()
        p.level = 1 if sub else 0
        p.space_after = Pt(18)
        r = p.runs[0]
        r.font.size = Pt(BULLET_PT - 4 if sub else BULLET_PT)
        r.font.color.rgb = GREY if sub else BLACK
    if source:
        add_footer(s, source)
    return s


def table_slide(prs, title, header, rows, source):
    s = M._new_slide(prs)
    add_title(s, title)
    nr, nc = len(rows) + 1, len(header)
    row_h = min(0.62, (IMG_H - 0.2) / nr)
    tw = min(IMG_W, 3.9 * nc)
    tpt = 26 if nr <= 7 else 22
    shape = s.shapes.add_table(nr, nc, Inches((W - tw) / 2), Inches(IMG_TOP + 0.15),
                               Inches(tw), Inches(row_h * nr))
    tbl = shape.table
    for j, h in enumerate(header):
        tbl.cell(0, j).text = h
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            tbl.cell(i, j).text = v
    for i in range(nr):
        for j in range(nc):
            for p in tbl.cell(i, j).text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(tpt)
                    r.font.bold = (i == 0 or j == 0)
    add_footer(s, source)
    return s


def locator_slide(prs, title, key, f, source):
    """Left: whole field (top) and its arrow map (bottom). Right: the 6 circled beads. Labels native."""
    pm = PANELS["{}_F{}".format(key, f)]
    s = M._new_slide(prs)
    add_title(s, title)
    top = IMG_TOP
    col_h = H - top - FOOT_H - 0.05
    side = (col_h - 0.1) / 2
    arrows = PANEL_DIR / "{}_F{}_arrows_bare.png".format(key, f)
    if arrows.exists():
        M.add_image_in_box(s, str(PANEL_DIR / "{}_F{}_field.png".format(key, f)), MARGIN, top, side, side)
        M.add_image_in_box(s, str(arrows), MARGIN, top + side + 0.1, side, side)
        M.add_textbox(s, "shift ×40", MARGIN + side + 0.1, top + 2 * side + 0.1 - 0.5, 2.4, 0.5, font_pt=20,
                      color=RGBColor(0xD9, 0x1A, 0x1A), bold=True, align=PP_ALIGN.LEFT)
    else:                               # no arrow map (gel sets): the field fills the left column
        side = col_h
        M.add_image_in_box(s, str(PANEL_DIR / "{}_F{}_field.png".format(key, f)), MARGIN, top, side, side)
    M.add_textbox(s, "★ = zero point" if pm["in_field"] else "zero point off field", MARGIN + side + 0.1,
                  top + (col_h - 0.1) / 2 - 0.9, 1.45, 0.9, font_pt=20, color=BLACK, align=PP_ALIGN.LEFT)
    x0 = MARGIN + side + 1.7
    tile = min((W - x0 - MARGIN - 0.5) / 3, (col_h - 1.25) / 2)
    gapx = (W - x0 - MARGIN - 3 * tile) / 2
    for j, bd in enumerate(pm["beads"][:6]):
        r, c = divmod(j, 3)
        left = x0 + c * (tile + gapx)
        t = top + r * (tile + 0.65)
        M.add_textbox(s, "{:.1f} px".format(bd["shift"]), left, t, tile, 0.6, font_pt=34, color=BLACK, bold=True)
        M.add_image_in_box(s, str(PANEL_DIR / "{}_F{}_b{}_ov.png".format(key, f, bd["i"])), left, t + 0.6, tile, tile)
        M.add_textbox(s, str(bd["i"]), left + 0.05, t + 0.62, 0.5, 0.45, font_pt=22, color=RGBColor(255, 110, 110),
                      bold=True, align=PP_ALIGN.LEFT)
    add_footer(s, source)
    return s


def beadcompare_slide(prs, title, rows, source):
    """One row per bead set (label, key, f): its 6 locator beads, 640/488 overlay, |shift| above each."""
    s = M._new_slide(prs)
    add_title(s, title)
    lab_w, head_h, gap = 1.6, 0.55, 0.12
    tile = min((IMG_W - lab_w - 5 * gap) / 6, (IMG_H - len(rows) * head_h - gap) / len(rows))
    x0 = MARGIN + lab_w
    for r, (lab, key, f) in enumerate(rows):
        pm = PANELS["{}_F{}".format(key, f)]
        t = IMG_TOP + r * (head_h + tile + gap)
        M.add_textbox(s, lab, MARGIN, t + head_h, lab_w, tile, font_pt=30, color=BLUE, bold=True,
                      align=PP_ALIGN.LEFT)
        s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for j, bd in enumerate(pm["beads"][:6]):
            left = x0 + j * (tile + gap)
            M.add_textbox(s, "{:.1f} px".format(bd["shift"]), left, t, tile, head_h, font_pt=24, color=BLACK,
                          bold=True)
            M.add_image_in_box(s, str(PANEL_DIR / "{}_F{}_b{}_ov.png".format(key, f, bd["i"])), left, t + head_h,
                               tile, tile)
    add_footer(s, source)
    return s


def beadgrid_slide(prs, title, key, f, rows_i, source):
    """Rows = beads, columns = each channel alone + the 640/488 overlay."""
    pm = PANELS["{}_F{}".format(key, f)]
    n_ch = pm["n_ch"]
    s = M._new_slide(prs)
    add_title(s, title)
    lab_w, head_h = 2.0, 0.55
    ncol = n_ch + 1
    tile = min((IMG_W - lab_w - 0.1 * ncol) / ncol, (IMG_H - head_h - 0.1 * len(rows_i)) / len(rows_i))
    x0 = MARGIN + lab_w + (IMG_W - lab_w - ncol * (tile + 0.1)) / 2
    heads = CH_NAMES[:n_ch] + ["640 + 488"]
    if n_ch >= 4:                       # FF 2026-10-02: flag it, do not troubleshoot it here
        heads[3] = "405 nm*"
        source = source + " · *405 nm looks odd: weak, noisy bead signal (not used for the shifts)"
    for c, hname in enumerate(heads):
        M.add_textbox(s, hname, x0 + c * (tile + 0.1), IMG_TOP, tile, head_h, font_pt=26, color=BLUE, bold=True)
    for r, i in enumerate(rows_i):
        bd = pm["beads"][i - 1]
        t = IMG_TOP + head_h + r * (tile + 0.1)
        M.add_textbox(s, "{:.1f} px".format(bd["shift"]), MARGIN, t, lab_w, tile, font_pt=34, color=BLACK, bold=True)
        s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        for c in range(ncol):
            img = "{}_F{}_b{}_{}.png".format(key, f, i, "ov" if c == n_ch else "c{}".format(c + 1))
            M.add_image_in_box(s, str(PANEL_DIR / img), x0 + c * (tile + 0.1), t, tile, tile)
    add_footer(s, source)
    return s


def equation_slide(prs, title, eq_lines, defs):
    s = M._new_slide(prs)
    add_title(s, title, pt=40)
    eb = s.shapes.add_textbox(Inches(MARGIN), Inches(IMG_TOP + 0.3), Inches(IMG_W), Inches(1.2))
    ep = eb.text_frame.paragraphs[0]
    ep.alignment = PP_ALIGN.CENTER
    for seg in re.split(r"(_\w)", " ".join(eq_lines)):     # "_x" -> subscript x
        if not seg:
            continue
        r = ep.add_run()
        r.text = seg[1:] if seg.startswith("_") else seg
        r.font.size, r.font.bold, r.font.color.rgb = Pt(44), True, BLUE
        if seg.startswith("_"):
            r.font._element.set("baseline", "-25000")
    box = s.shapes.add_textbox(Inches(MARGIN + 0.6), Inches(IMG_TOP + 1.9), Inches(IMG_W - 1.2), Inches(IMG_H - 2.0))
    tf = box.text_frame
    tf.word_wrap = True
    for i, d in enumerate(defs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = d
        p.space_after = Pt(12)
        p.runs[0].font.size = Pt(30)
    return s


def rawcorr_slide(prs, title, stem, source, which="full"):
    """As recorded | corrected, side by side and as large as the slide allows: the whole nucleus
    (which="full") or the yellow zoom box (which="zoom")."""
    s = M._new_slide(prs)
    add_title(s, title)
    gap = 0.3
    side = min((IMG_W - gap) / 2, IMG_H - 0.55)
    x0 = MARGIN + (IMG_W - (2 * side + gap)) / 2
    for i, (name, lab) in enumerate((("before", "as recorded"), ("after", "corrected"))):
        left = x0 + i * (side + gap)
        M.add_textbox(s, lab, left, IMG_TOP, side, 0.5, font_pt=28, color=BLACK, bold=True)
        M.add_image_in_box(s, str(PANEL_DIR / "{}_{}_{}.png".format(stem, name, which)), left, IMG_TOP + 0.5,
                           side, side)
    add_footer(s, source)
    return s


def beforeafter_slide(prs, title, imgs, source):
    s = M._new_slide(prs)
    add_title(s, title)
    gap = 0.3
    side = min((IMG_W - gap) / 2, IMG_H - 0.55)
    x0 = MARGIN + (IMG_W - (2 * side + gap)) / 2
    for i, (img, lab) in enumerate(zip(imgs, ("as recorded", "corrected"))):
        left = x0 + i * (side + gap)
        M.add_textbox(s, lab, left, IMG_TOP, side, 0.5, font_pt=28, color=BLACK, bold=True)
        M.add_image_in_box(s, str(img), left, IMG_TOP + 0.5, side, side)
    add_footer(s, source)
    return s


def divider_slide(prs, title, subtitle=""):
    s = M._new_slide(prs, bg=M.DIVIDER_BG)
    M.add_textbox(s, title, 0.6, 1.2, W - 1.2, 2.9, font_pt=54, color=BLACK, bold=True)
    s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    if subtitle:
        M.add_textbox(s, subtitle, 0.6, 4.3, W - 1.2, 1.6, font_pt=30, color=GREY)
        s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.TOP
    return s


def title_slide(prs, title, subtitle):
    s = M._new_slide(prs)
    M.add_textbox(s, title, 0.6, 0.8, W - 1.2, 3.3, font_pt=54, color=BLACK, bold=True)
    s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
    M.add_textbox(s, subtitle, 0.6, 4.4, W - 1.2, 1.4, font_pt=32, color=GREY)
    s.shapes[-1].text_frame.vertical_anchor = MSO_ANCHOR.TOP
    return s


# ---------------------------------------------------------------------------
# Images: trim white margins, render + crop PDF pages
# ---------------------------------------------------------------------------
def trim(im, pad=12):
    rgb = im.convert("RGB")
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    diff = ImageChops.difference(rgb, bg).convert("L").point(lambda v: 255 if v > 12 else 0)
    box = diff.getbbox()
    if not box:
        return rgb
    x0, y0, x1, y1 = box
    return rgb.crop((max(0, x0 - pad), max(0, y0 - pad),
                     min(rgb.width, x1 + pad), min(rgb.height, y1 + pad)))


def prep(src, box=None):
    """Trimmed copy of an original PNG, optionally cropped to box = fractions (x0, y0, x1, y1).
    Originals are never modified."""
    src = Path(src)
    if not src.exists():
        raise FileNotFoundError(src)
    tag = src.parent.name if src.parent.name not in ("figures", "montages") else src.parent.parent.name
    suffix = "" if box is None else "__crop" + "_".join("{:03d}".format(int(round(100 * v))) for v in box)
    out = CROP_DIR / "{}__{}{}.png".format(tag[-15:], src.stem, suffix)
    im = Image.open(src)
    if box is not None:
        w, h = im.size
        im = im.crop((int(box[0] * w), int(box[1] * h), int(box[2] * w), int(box[3] * h)))
    trim(im).save(out)
    return out


NULL_BOX = (0.0, 0.0, 0.60, 1.0)       # QC_NullOverview: the camera-frame map only
METHOD_BOX = (0.0, 0.0, 1.0, 0.52)     # QC_MethodScatter: the x-shift row only


def page_png(pdf_key, page):
    p = PAGE_DIR / "{}_p{:02d}.png".format(pdf_key, page)
    if not p.exists():
        PAGE_DIR.mkdir(parents=True, exist_ok=True)
        subprocess.run([GS, "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=png16m", "-r300",
                        "-dTextAlphaBits=4", "-dGraphicsAlphaBits=4",
                        "-dFirstPage={}".format(page), "-dLastPage={}".format(page),
                        "-sOutputFile={}".format(p), str(PDF[pdf_key])], check=True)
    return p


def crop(key):
    pdf_key, page, (fx0, fy0, fx1, fy1) = CROPS[key]
    out = CROP_DIR / "{}.png".format(key)
    im = Image.open(page_png(pdf_key, page))   # re-cropped every build, so an edited box takes effect
    w, h = im.size
    trim(im.crop((int(fx0 * w), int(fy0 * h), int(fx1 * w), int(fy1 * h)))).save(out)
    return out


def pdf_text(pdf_key, page, tries=3):
    for _ in range(tries):          # Ghostscript's txtwrite has come back empty now and then; retry
        t = _pdf_text_once(pdf_key, page)
        if t.strip():
            return t
    raise RuntimeError("no text from {} p.{}".format(pdf_key, page))


def _pdf_text_once(pdf_key, page):
    r = subprocess.run([GS, "-q", "-dNOPAUSE", "-dBATCH", "-sDEVICE=txtwrite",
                        "-dFirstPage={}".format(page), "-dLastPage={}".format(page),
                        "-sOutputFile=-", str(PDF[pdf_key])], capture_output=True, text=True,
                       encoding="utf-8", errors="replace", check=True)
    return r.stdout


# ---------------------------------------------------------------------------
# Numbers, computed from the studies' own files
# ---------------------------------------------------------------------------
def read_csv(path):
    return np.genfromtxt(path, delimiter=",", names=True, dtype=None, encoding="utf-8")


def bead_stats(run):
    d = read_csv(run / "PerBead_Shifts.csv")
    d = d[d["pair"] == "c1-c3"]
    fovs = sorted(set(int(f) for f in d["fov"]))
    out = {"fovs": fovs, "n": len(d)}
    # joint fit: shared magnification m, own translation per FOV (centroid shifts)
    nF = len(fovs)
    A = np.zeros((2 * len(d), 1 + 2 * nF))
    b = np.zeros(2 * len(d))
    for i, row in enumerate(d):
        k = fovs.index(int(row["fov"]))
        A[2 * i, 0], A[2 * i, 1 + 2 * k], b[2 * i] = row["x"], 1, row["dx_cen"]
        A[2 * i + 1, 0], A[2 * i + 1, 2 + 2 * k], b[2 * i + 1] = row["y"], 1, row["dy_cen"]
    sol, *_ = np.linalg.lstsq(A, b, rcond=None)
    res = (b - A @ sol).reshape(-1, 2)
    out["m"] = sol[0]
    out["resid_axis"] = float(np.sqrt(np.mean(res ** 2)))
    nulls = {f: -sol[1 + 2 * k: 3 + 2 * k] / sol[0] for k, f in enumerate(fovs)}
    P = np.array(list(nulls.values()))
    out["null_spread"] = float(max(np.hypot(*(p - q)) for p in P for q in P))
    per = {}
    meth_diff = 0.0
    ncc = read_csv(run / "GlobalNCC.csv")
    ncc = ncc[ncc["pair"] == "c1-c3"]
    ncc_diff = 0.0
    for f in fovs:
        s = d[d["fov"] == f]
        mag = np.hypot(s["dx_cen"], s["dy_cen"])
        Af = np.column_stack([np.r_[s["x"], s["y"]],
                              np.r_[np.ones(len(s)), np.zeros(len(s))],
                              np.r_[np.zeros(len(s)), np.ones(len(s))]])
        mf = np.linalg.lstsq(Af, np.r_[s["dx_cen"], s["dy_cen"]], rcond=None)[0][0]
        meds = {k: np.array([np.median(s["dx_" + k]), np.median(s["dy_" + k])])
                for k in ("cen", "gauss", "xc")}
        for a in meds:
            for c in meds:
                meth_diff = max(meth_diff, float(np.max(np.abs(meds[a] - meds[c]))))
        g = ncc[ncc["fov"] == f]
        gv = np.array([float(g["gdx"][0]), float(g["gdy"][0])])
        ncc_diff = max(ncc_diff, float(np.max(np.abs(gv - meds["cen"]))))
        per[f] = {"n": len(s), "p02": float(np.percentile(mag, 2)), "p98": float(np.percentile(mag, 98)),
                  "max": float(mag.max()), "m": mf, "med_cen": meds["cen"].tolist(), "ncc": gv.tolist()}
    out["per"] = per
    out["meth_diff"] = meth_diff
    out["ncc_diff"] = ncc_diff
    mag_all = np.hypot(d["dx_cen"], d["dy_cen"])
    out["p98_all"] = float(np.percentile(mag_all, 98))
    out["med_all"] = float(np.median(mag_all))
    out["null_in"] = sum(1 for p in P if 0 <= p[0] <= 2048 and 0 <= p[1] <= 2048)
    loc = run / "LocatorBeads.csv"
    if loc.exists():
        L = read_csv(loc)
        out["loc"] = {}
        for f in fovs:
            rows = sorted((r for r in L if int(r["fov"]) == f), key=lambda r: int(r["rank"]))
            mags = []
            for r in rows:
                s = d[(d["fov"] == f) & (d["beadID"] == r["beadID"])]
                mags.append(float(np.hypot(s["dx_cen"][0], s["dy_cen"][0])))
            out["loc"][f] = (mags[0], mags[-1], float(rows[-1]["dNull"]))
    return out


def cell_stats():
    r = read_csv(ROUTE1 / "PerNucleus_Route1.csv")
    kept = (r["qFG"] >= 0.3) & (r["clippedFG"] == 0)
    fields = {}
    for f in FIELDS:
        s = r[(r["fov"] == f) & kept]
        u = np.column_stack([s["dxFG"], s["dyFG"]])
        R = float(np.hypot(*np.mean(u / np.hypot(u[:, 0], u[:, 1])[:, None], axis=0)))
        wm = (s["qFG"][:, None] * u).sum(0) / s["qFG"].sum()
        fields[f] = {"n": len(s), "n_all": int(np.sum(r["fov"] == f)), "R": R,
                     "R_rand": float(np.sqrt(np.pi) / (2 * np.sqrt(len(s)))),
                     "mean_len": float(np.hypot(*wm))}
    ex = r[(r["fov"] == "EtOH/01") & (r["nucID"] == 13)][0]
    summ = (ROUTE1 / "Summary.txt").read_text()
    mm = re.search(r"magnification m = \+([\d.]+)%\s+\(95% CI \[\+([\d.]+) \+([\d.]+)\]%\)", summ)
    v = read_csv(VERIFY / "PerNucleus.csv")
    vk = v[v["kept"] == 1]
    corr = np.hypot(vk["corr_dx"], vk["corr_dy"])
    per_field = {f: (float(np.median(vk[vk["fov"] == f]["resid_before"])),
                     float(np.median(vk[vk["fov"] == f]["resid_after"])),
                     int(np.sum(vk["fov"] == f))) for f in FIELDS}
    return {
        "fields": fields, "n_kept": int(kept.sum()), "n_all": len(r),
        "ex": (float(ex["dxFG"]), float(ex["dyFG"]), float(ex["qFG"])),
        "m": float(mm.group(1)), "m_lo": float(mm.group(2)), "m_hi": float(mm.group(3)),
        "before": float(np.median(vk["resid_before"])), "after": float(np.median(vk["resid_after"])),
        "n_fell": int(np.sum(vk["resid_after"] < vk["resid_before"])), "n_v": len(vk),
        "corr_med": float(np.median(corr)), "corr_min": float(corr.min()), "corr_max": float(corr.max()),
        "per_field": per_field, "verify": v,
    }


def nm(px):
    return "{:.0f} nm".format(px * PX_NM)


# bead detection, shown on slide 4 and again in the gel section (FF 2026-10-15: "explain bead
# detection again even if it is done at beginning"); detectBeads in bead_offset_study.m
DETECT_2D = [
    "One slice, 640 nm, slice median subtracted",
    "Smooth (σ 1 px); local maxima above median + 8 robust SD",
    "Drop: edge (24 px), saturated (≥ 60,000), any pair < 16 px (1 μm) apart",
    "Gives whole-pixel positions; each shift method then measures the bead in both colors",
]
DETECT_2D_SRC = "detectBeads, bead_offset_study.m · robust SD = 1.4826 × median absolute deviation"


def gel_slides():
    """Section on the 2026-10-14 gel beads: what they look like next to the 2024 beads (also a gel,
    imaged at its surface where beads are sparse; log.txt of 20240522), then can the
    zero point be predicted from other fields? Images: Field_Gallery of
    intranuclear_colocalization/scripts/bead_field_gallery.m. Numbers and plots: the newest
    (by time) Null_Prediction_* run of scripts/bead_null_prediction.m. Stage A (focus, estimator,
    stage direction) is shown alone until stage B (the prediction) exists."""
    out = [("divider", "4 · Can we predict the zero point?",
            "Gel beads, Oct 2026 (Vishavdeep) · 3 regions, slice 4 of 7, 640 vs 488 nm")]
    gal = GEL_ROOT / "Field_Gallery"
    glass = sorted(gal.glob("Whole_2024 gel surface_*.png"))   # 2024: a gel, imaged at its surface
    gels = {r: sorted(gal.glob("Whole_gel {}_*.png".format(r))) for r in ("bottom", "middle", "top")}
    if glass and gels["middle"]:
        g, m = glass[0], gels["middle"][0]
        gsrc = "same display rule for both: background = slice median, 0 to 99.9th percentile · 65 nm/px"
        out.append(("pair", "2024: single beads · Oct 2026: a carpet of beads, cracked",
                    [g, m], ["2024 gel, surface (beads 1:10)", "Oct 2026 gel"],
                    "whole field, 640 nm · bar 20 μm · " + gsrc))
        cg, cm = (p.with_name(p.name.replace("Whole_", "Crop_")) for p in (g, m))
        out.append(("pair", "Up close: no isolated beads in the Oct 2026 gel",
                    [cg, cm], ["2024", "Oct 2026"],
                    "central 19.5 μm, 640 nm green / 488 nm magenta · bar 2 μm · " + gsrc))
        reg = [gels[r][0] for r in ("bottom", "middle", "top") if gels[r]]
        if len(reg) == 3:
            out.append(("pair", "Gel: one field from each region", reg,
                        ["bottom", "middle", "top"], "whole field, 640 nm, slice 4 · bar 20 μm · display as for 2024"))
    # single beads before anything quantitative (FF 2026-10-15); panels from
    # intranuclear_colocalization/scripts/render_bead_slide_panels.py --sets gel_middle ...
    out.append(("text", "Single beads: how they are found", DETECT_2D, DETECT_2D_SRC))
    if "gel_middle_F01" in PANELS:
        bsrc = "tiles 1.2 μm · 640 green / 488 magenta · line = shift 640 → 488"
        out.append(("beadcompare", "Single beads: 2024 vs the most isolated gel beads",
                    [("2024", "2024", "1"), ("2026-10-14 gel", "gel_middle", "01")],
                    bsrc + " · gel: from its most isolated 10 % · shift: 2024 centroid, gel cross-correlation"))
        out.append(("locator", "Gel, middle field 01: six isolated beads", "gel_middle", "01",
                    GEL_LABEL + " · slice 4 · ★ = zero point: where the two colors line up · " + bsrc + " (cross-correlation)"))
    # the two kinds of measurement on one crop (FF 2026-10-15: "this needs to be clear in ppt");
    # Methods_*.png of bead_field_gallery.m
    meth = sorted(gal.glob("Methods_gel middle_*.png"))
    if meth:
        out.append(("image", "Every shift is measured on slice 4 of 7: beads or tiles", prep(meth[0]),
                    GEL_LABEL + " · every field · other slices: focus only · beads: centroid, Gaussian or cross-corr. · tiles: cross-corr. · bar 10 μm"))
    runs = [r for r in GEL_ROOT.glob("Null_Prediction_*")
            if (r / "StageA_Summary.json").exists() or (r / "NullPrediction_Summary.json").exists()]
    newest = lambda fn: max((r for r in runs if (r / fn).exists()), key=lambda r: (r / fn).stat().st_mtime,
                            default=None)
    a_run, b_run = newest("StageA_Summary.json"), newest("NullPrediction_Summary.json")
    if a_run is None and b_run is None:
        print("[gel] no Null_Prediction run yet: analysis slides skipped")
    # stage B (FF go 2026-10-15): newest Null_Prediction_B_* folder, four runs; tile_all is the
    # main one. Loaded here: its zero-point map opens the subsection on why the zero point moves
    bp = max(GEL_ROOT.glob("Null_Prediction_B_*"), key=lambda r: r.name, default=None)
    BR = {}
    if bp is not None:
        for d in sorted(bp.iterdir()):
            if (d / "NullPrediction_Summary.json").exists():
                BR[d.name] = (d, json.loads((d / "NullPrediction_Summary.json").read_text()))
    if a_run is not None:
        A = json.loads((a_run / "StageA_Summary.json").read_text())
        pm = A["pooledM_pct"]
        F = read_csv(a_run / "FieldTable.csv")
        # a field's focus: the median focus of its beads found in 3D when stage A was run with
        # beadFocusDir (FF 2026-10-15), else bead_offset_study's field focus
        fcol = "focus_slice" if "focus_slice" in F.dtype.names else "focus640_slice"
        foc = [float(r[fcol]) for r in F]
        mv = [(float(r["move_dx_um"]), float(r["t_x"])) for r in F if r["move_dx_um"] not in ("", "NaN")]
        med = lambda v: sorted(v)[len(v) // 2] if len(v) % 2 else sum(sorted(v)[len(v) // 2 - 1:len(v) // 2 + 1]) / 2
        tneg = med([t for d, t in mv if d < 0])
        tpos = med([t for d, t in mv if d > 0])
        # rows (bead_null_prediction loadRuns: runs of one X step sign): per row its direction,
        # field count and median whole-field shift; a row "follows" when its x shift has the
        # sign opposite to its step (-X rows positive), as the pooled medians do
        rows = {}
        if "row" in F.dtype.names:
            for r in F:
                rows.setdefault((str(r["region"]), int(r["row"])), []).append((int(r["rowDir"]), float(r["t_x"])))
        rows3 = {k: v for k, v in rows.items() if len(v) >= 3}
        n_follow = sum(1 for v in rows3.values() if med([t for _, t in v]) * v[0][0] < 0)
        n_small = sum(len(v) for k, v in rows.items() if k not in rows3)
        # by the row's step direction: b in x and each field's own m (takeaway: m stays, b moves);
        # within-row scatter: median distance of a field's b from its row's median
        rdir = lambda r: int(float(r["rowDir"])) if "rowDir" in F.dtype.names else 0
        b_neg = med([float(r["t_x"]) for r in F if rdir(r) < 0])
        b_pos = med([float(r["t_x"]) for r in F if rdir(r) > 0])
        mcol = "ownM_pct" if "ownM_pct" in F.dtype.names else None
        m_neg = med([float(r[mcol]) for r in F if rdir(r) < 0]) if mcol else float("nan")
        m_pos = med([float(r[mcol]) for r in F if rdir(r) > 0]) if mcol else float("nan")
        b_within = med([abs(t - med([u for _, u in v])) for v in rows3.values() for _, t in v]) if rows3 else float("nan")
        print("[gel] stage A {} ({}): {} fields, pooled m {}, t_x -X/+X {:+.2f}/{:+.2f} px".format(
            a_run.name, A["method"], A["nFields"],
            ", ".join("{} {:.3f}%".format(k, v) for k, v in pm.items()), tneg, tpos))
        res = [float(r["ownResid_px"]) for r in F]
        # how focus is measured, before any focus plot (FF 2026-10-15: "this needs to be
        # comprehensive"); no cutoff anywhere, the choice is FF's
        # two slides, no overlap (FF 2026-10-15: "47-48 ... somewhat redundant"): what the two
        # focus numbers are, then how the beads for bead focus are found
        # how the beads are found, then what bead focus is (FF 2026-10-15: detection first)
        out.append(("text", "Finding every bead in 3D", [
            "All 7 slices: minus median, smoothed in 2D (σ 1 px)",
            "Bead: brighter than its 26 neighbors in x, y, z",
            "and above median + 8 robust SD",
            "Drop edge and saturated beads · no 1 μm rule",
            "Peak at slice 1 or 7: may lie outside the stack",
        ], "detectBeads3D, bead_focus_map.m · 640 nm · robust SD = 1.4826 × MAD · edge 24 px · saturated ≥ 60,000"))
        out.append(("text", "How focus is measured", [
            "Slices 1–7, 0.3 μm apart (ND2 metadata) · shifts: slice 4 only",
            "Bead focus: the slice where a bead is brightest",
            "Brightness on a slice: highest of the 3×3 pixels at the bead, minus the slice median",
            "Refined between slices by a parabola through three slices",
            "A field's focus: the median focus of its beads",
        ], GEL_LABEL + " · 640 nm · beads found in 3D"))
        # per-bead focus (FF 2026-10-15: "the surface is not flat"); intranuclear_colocalization/
        # scripts/bead_focus_map.m, newest Bead_Focus_* run
        fr = max(GEL_ROOT.glob("Bead_Focus_*"), key=lambda r: r.name, default=None)
        FS = json.loads((fr / "Summary.json").read_text()) if fr is not None and (fr / "Summary.json").exists() else None
        if FS is not None and "methodField" in FS:
            out.append(("image", "One field: a pile of beads at slice 1, few at 7",
                        prep(fr / "Focus_Method.png"),
                        "{} · field {} · beads found in 3D · orange: peak at the first or last slice, "
                        "not refined".format(GEL_LABEL, FS["methodField"])))
        # slice 1 vs slice 7 (FF 2026-10-15: "I can believe it is not just an artifact but show me")
        if FS is not None and "nFieldsFirstGtLast" in FS:
            esrc = "{} · field {} · 26 μm crop, same display range for both slices · bar 5 μm".format(
                GEL_LABEL, FS["methodField"])
            # one slide: same field and crop as the unmarked Edge_Slices.png (FF 2026-10-16: combine)
            out.append(("image", "2026-10-14 gel, {}: slice-1 beads line the cracks".format(
                FS["methodField"]),
                        prep(fr / "Edge_Slices_Marked.png"),
                        esrc + " · circles: beads whose peak is that slice"))
            out.append(("image", "Beads peaking at slice 1, 4 or 7: real peaks",
                        prep(fr / "Edge_Profiles.png"),
                        "{} · all fields · each bead's brightness over its own peak, averaged · below 0: dimmer than the slice median".format(GEL_LABEL)))
            out.append(("image", "More beads peak at slice 1 than at 7 in {} of {} fields".format(
                FS["nFieldsFirstGtLast"], FS["nFieldsEdge"]), prep(fr / "Edge_Counts.png"),
                "{} · median over fields: {:.1f} % at slice 1, {:.1f} % at slice 7".format(
                    GEL_LABEL, FS["pctAtFirst_median"], FS["pctAtLast_median"])))
        fq = sorted(foc)
        p10, p90 = fq[round(0.1 * (len(fq) - 1))], fq[round(0.9 * (len(fq) - 1))]
        out.append(("image", "Median bead focus: 80 % of fields at slices {:.1f}–{:.1f}".format(p10, p90),
                    prep(a_run / "QC_Focus.png"), GEL_LABEL + " · one value per field: the median focus of its beads"))
        if FS is not None:
            print("[gel] focus map {}: within {:.2f}, between {:.2f} slices, |dz| nn {:.2f} vs random {:.2f}".format(
                fr.name, FS["within_p10p90_median"], FS["between_p10p90"], FS["nnAbsDiff_median"],
                FS["randAbsDiff_median"]))
            fsrc = "{} · {:,} beads found in 3D".format(GEL_LABEL, int(FS["nBeads"]))
            out.append(("image", "Focus spread: ≥ {:.1f} slices within a field, {:.1f} between".format(
                FS["within_p10p90_median"], FS["between_p10p90"]), prep(fr / "Focus_Spread.png"),
                fsrc + " · within: median field's middle 80 %, cut off at slices 1 and 7 · between: middle 80 % of medians · slice 0.3 μm"))
            out.append(("image", "Not flat: nearby beads agree better ({:.2f} vs {:.2f} slices)".format(
                FS["nnAbsDiff_median"], FS["randAbsDiff_median"]), prep(fr / "Focus_Map_1.png"),
                fsrc + " · image: max projection · median |dz|: nearest vs random bead · slice = 0.3 μm"))
        def ranks(v):                                  # average ranks, for Spearman
            o = sorted(range(len(v)), key=lambda i: v[i]); r = [0.0] * len(v); i = 0
            while i < len(o):
                j = i
                while j + 1 < len(o) and v[o[j + 1]] == v[o[i]]:
                    j += 1
                for k in range(i, j + 1):
                    r[o[k]] = (i + j) / 2
                i = j + 1
            return r
        ok = [(abs(z - float(r["zUsed"])), e) for r, z, e in zip(F, foc, res) if z == z and e == e]
        ra, rb = ranks([a for a, _ in ok]), ranks([b for _, b in ok])
        ma, mb = sum(ra) / len(ra), sum(rb) / len(rb)
        rho = sum((x - ma) * (y - mb) for x, y in zip(ra, rb)) / (
            sum((x - ma) ** 2 for x in ra) * sum((y - mb) ** 2 for y in rb)) ** 0.5
        out.append(("image", "Field focused at slice {:.0f}: only blur at slice 4".format(min(z for z in foc if z == z)),
                    prep(a_run / "QC_Focus_Examples.png"),
                    GEL_LABEL + " · 640 nm, slice 4"))
        # what the zero point is, before the fit uses it (FF 2026-10-16: definitions first, no
        # numbers; b for the whole-field shift, t in the code, like an intercept)
        out.append(("text", "The zero point is where the two colors line up", [
            "Color shift at image position r:   d(r) = m · (r − a) + b",
            "  Magnification difference m, about a fixed axis a",
            "  Whole-field shift b: one per field, the same at every pixel",
            "Zero point: where d = 0   →   r₀ = a − b / m",
            "If b = 0, the zero point is the axis",
            "A small b moves it far from the axis, because m is small",
        ], GEL_LABEL + " · shifts measured on slice 4 of every field"))
        # the gel fit itself, drawn like the 2024 / 2026 shift-vs-distance slides (FF 2026-10-16);
        # title: one magnification for every field (FF 2026-10-16: "needs to be highly emphasized")
        if (a_run / "QC_Shift_vs_Distance.png").exists() and "tile" in pm:
            out.append(("image", "One magnification for all {} fields: {:.2f} %".format(len(F), pm["tile"]),
                        prep(a_run / "QC_Shift_vs_Distance.png"),
                        GEL_LABEL + " · dot = one tile, distance from its field's zero point (fitted from that field's tiles) "
                        "· dashed = one slope for all fields"))
        if "tile" in pm:
            out.append(("image", "Crowded beads: the centroid does not work", prep(a_run / "QC_Method_m.png"),
                GEL_LABEL + " · magnification % = px per 1000 px ÷ 10 · tiles {:.2f}% ≈ 2026 glass beads "
                "· Gaussian also low ({:.2f}%) · dot: one field · top: all fields together".format(pm["tile"], pm["gauss"])))
        # misfit needs the fit, so it follows the first fitted slide (FF 2026-10-16: "move it later")
        out.append(("image", "Misfit: flat near slice 4, high in a few far-off fields",
            prep(a_run / "QC_Focus_vs_Fit.png"),
            GEL_LABEL + " · misfit: RMS of tiles around the field's own fit · Spearman ρ vs |focus − 4| = {:.2f}{}".format(
                rho, " · fields at slices {:g}–{:g} only: same results".format(*[round(v, 1) for v in BR["tile_mid80"][1]["focusKeep"]])
                if "tile_mid80" in BR else "")))
        # why does the zero point move? (FF 2026-10-16 redesign: observation, what it means,
        # where on the stage, rows compared, what is not known)
        if "tile_all" in BR:
            Bm = BR["tile_all"][1]
            mfrac = Bm["m_pct"] / 100
        if "tile_all" in BR:
            out.append(("image", "Zero points differ by field: median deviation {:,.0f} px".format(Bm["nullDistMed_px"]),
                        prep(BR["tile_all"][0] / "NullPred_Map.png"),
                        "{} · all {} fields · tiles · deviation: distance to ★, the mean of all fields".format(
                            GEL_LABEL, Bm["nFieldsKept"])))
        # what a row is, before the stage maps and the row counts (FF 2026-10-16: "7 rows is out of nowhere")
        if rows3:
            nreg = len({k[0] for k in rows})
            # FF 2026-10-16: say that this is only how the imaging was done
            out.append(("text", "How the fields were imaged: the stage scans in rows", [
                "Acquisition only (ND2 log): no shifts on this slide",
                "Each region: rows of fields stepped along X, alternately +X, −X",
                "Row: consecutive fields reached by X steps of one sign",
                "{} rows, {} per region; {} have ≥ 3 fields".format(len(rows), len(rows) // nreg, len(rows3))
                if len(rows) % nreg == 0 else "{} rows; {} have ≥ 3 fields".format(len(rows), len(rows3)),
                "Question: does b depend on the row's step direction?",
            ], GEL_LABEL + " · stage path and order: ND2 log · {} fields in short rows at the turns".format(n_small)))
        # the stage path, zero point x | y
        if (a_run / "QC_Stage_Path.png").exists():
            out.append(("image", "Zero point x alternates from row to row; y less so", prep(a_run / "QC_Stage_Path.png"),
                        GEL_LABEL + " · ND2 log · arrow: step from the previous field "
                        "· axis: the mean zero point · regions bottom → top by stage Y"))
        if rows3 and (a_run / "QC_Row_Shift.png").exists():
            # FF 2026-10-16: "make clear it is a hardware lead"
            # title: the pattern, not medians that read as per-field values (FF 2026-10-16: "horrible title");
            # footer: row medians by direction and the field scatter
            rmed = {d: sorted(med([t for _, t in v]) for v in rows3.values() if v[0][0] == d) for d in (-1, 1)}
            sgn = lambda v: "{:+.1f}".format(v).replace("-", "−")
            out.append(("image", "Hardware lead: in every row, the median x shift flips sign with the step direction",
                        prep(a_run / "QC_Row_Shift.png"),
                        GEL_LABEL + " · rows of ≥ 3 fields: medians {} to {} px (−X), {} to {} px (+X) · within a row, a field is "
                        "typically {:.1f} px from its row median · number: fields per row".format(
                            sgn(rmed[-1][0]), sgn(rmed[-1][-1]), sgn(rmed[1][-1]), sgn(rmed[1][0]), b_within)))
            if (a_run / "QC_Row_M.png").exists():
                out.append(("image", "Magnification is the same in every row and direction",
                            prep(a_run / "QC_Row_M.png"),
                            GEL_LABEL + " · dot: one field's own fit · bar: row median · dashed: one fit to all fields"))
            # what the direction result says and what it cannot, one slide (FF 2026-10-16: merge)
            out.append(("text", "A hardware lead, not yet a mechanism", ([
                "Magnification: the same in −X and +X rows ({:.3f} vs {:.3f} %)".format(m_neg, m_pos)] if mcol else []) + [
                "Whole-field x shift: each row's median flips sign with the step",
                "So something shifts the whole image without rescaling it",
                "But n = {} rows, confounded with each row's Y position and time".format(len(rows3)),
                "Fits \"drift between the two color exposures\". Test: the same fields, approached from both directions",
            ], GEL_LABEL + " · magnification: median of each field's own fit · rows of ≥ 3 fields; {} fields in shorter rows at the turns".format(n_small)))
    # stage B: tile_all in full, all four runs in one comparison slide
    if "tile_all" in BR:
        b_run, B = BR["tile_all"]
        al = B["sets"]["all"]
        print("[gel] stage B {}: {} runs; tile_all m {:.3f}%, zero-point error {:.0f} px, residual {:.2f}/{:.2f}/{:.2f} px".format(
            bp.name, len(BR), B["m_pct"], al["nullErr_med"], al["none_med"], al["model_med"], al["oracle_med"]))
        out.append(("text", "Predicting the zero point: how", [
            "Split the fields at random in two halves, 100 times",
            "Fit one half: magnification and axis, shared by all its fields",
            "Predict the other half: zero point = that axis",
            "Error: each field's own zero point − the prediction",
            "Residual per tile: no correction · other half's model · own fit",
            "All {} fields; a check on the {} at slices {:g}–{:g} follows".format(
                B["nFieldsKept"], BR["tile_mid80"][1]["nFieldsKept"], *[round(v, 1) for v in BR["tile_mid80"][1]["focusKeep"]])
            if "tile_mid80" in BR else "All {} fields".format(B["nFieldsKept"]),
        ], "{} · bead_null_prediction.m stage B · seed 1".format(GEL_LABEL)))
        src = "{} · all {} fields · tiles · {} random 50/50 splits".format(GEL_LABEL, B["nFieldsKept"], B["nSplits"])
        out.append(("image", "Predicted zero point: off by {:,.0f} px, like the scatter".format(al["nullErr_med"]),
                    prep(b_run / "NullPred_Error.png"),
                    src + " · prediction = the other half's mean zero point · last bar: all errors at or above that value"))
        # lead-in for the residual bars (FF 2026-10-16: "needs a real lead in")
        out.append(("text", "Does the other half's model still correct the shifts?", [
            "No zero point, but magnification and axis may carry over",
            "Correct each held-out tile with the other half's magnification and axis",
            "Compare with no correction and with the own fit",
            "Own fit = best case: each held-out field with its own shift",
            "bottom, middle, top: splits within that region only",
        ], src + " · residual: what is left of the 640 → 488 shift"))
        out.append(("image", "Other half's m and axis: {:.1f} → {:.1f} px; the rest is b".format(
            al["none_med"], al["model_med"]), prep(b_run / "NullPred_Residual.png"),
            src + " · residual per tile · own fit (each field's own b): {:.1f} px".format(al["oracle_med"])))
        lab = {"tile_all": "tiles, all", "tile_mid80": "tiles, middle 80 %",
               "xc_all": "beads (xc), all", "xc_mid80": "beads (xc), middle 80 %"}
        rows = []
        for k in ("tile_all", "tile_mid80", "xc_all", "xc_mid80"):
            if k in BR:
                a2 = BR[k][1]["sets"]["all"]
                rows.append("{}: off {:,.0f} px · {:.1f} → {:.1f} px (own {:.1f})".format(
                    lab[k], a2["nullErr_med"], a2["none_med"], a2["model_med"], a2["oracle_med"]))
        mods = [BR[k][1]["sets"]["all"]["model_med"] for k in BR]
        ttl = ("Fewer fields, single beads: same residual (±{:.2f} px)".format(max(mods) - min(mods))
               if max(mods) - min(mods) < 0.1 else "Four variants: tiles or single beads, all or 80 % of fields")
        out.append(("text", ttl, rows,
                    "{} · medians over splits · off: zero point · px: none → other half (own fit) · middle 80 %: slices {:g}–{:g}".format(
                        GEL_LABEL, *BR.get("tile_mid80", (None, {"focusKeep": [0, 0]}))[1]["focusKeep"])))
        out.append(("text", "What the gel beads say", [
            "One magnification, {:.2f} %, in every field and step direction".format(B["m_pct"]),
            "Zero point not predictable from other fields: b sets it",
            "Other fields' m and axis: {:.1f} → {:.1f} px per tile; the rest is b".format(al["none_med"], al["model_med"]),
            "So b must be measured per field, as the cell correction does",
            "Hardware lead: b follows the stage X step direction",
        ], GEL_LABEL + " · tiles, all fields"))
    return out


# ---------------------------------------------------------------------------
# Deck
# ---------------------------------------------------------------------------
def build_plan():
    """[(kind, args...)] for every slide; images resolved and numbers computed here."""
    B = {k: bead_stats(v) for k, v in BEAD.items()}
    C = cell_stats()
    t25 = pdf_text("valid", 25)
    ov = re.findall(r"ER\s+area overlapping RNAP2\s+(EtOH|E2)\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+([+\-][\d.]+) %\s+\w+\s+([\d.]+)", t25)
    ccf = {}
    for pg, cond in ((27, "EtOH"), (28, "E2")):
        mt = re.search(r"CCF at zero shift\s+\w+\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+([+\-][\d.]+) %\s+\w+\s+([\d.]+)",
                       pdf_text("valid", pg))
        ccf[cond] = mt.groups()
    t5 = pdf_text("valid", 5)
    pl = re.search(r"read back at (\d+) % \(planted on ER.?\) and (\d+)", t5)

    for k in ("2024", "2026"):
        b = B[k]
        print("[{}] n={} joint m={:+.3f}% resid/axis={:.2f} px  null spread={:.0f} px  methods<= {:.2f} px  "
              "NCC-vs-centroid<= {:.2f} px  p98={:.2f} px".format(
                  k, b["n"], 100 * b["m"], b["resid_axis"], b["null_spread"], b["meth_diff"],
                  b["ncc_diff"], b["p98_all"]))
        for f, v in b["per"].items():
            print("     FOV {}: n={} m={:+.3f}% |shift| 2-98% = {:.2f}-{:.2f} px".format(
                f, v["n"], 100 * v["m"], v["p02"], v["p98"]))
        if "loc" in b:
            print("     locator beads #1/#6:", {f: "{:.2f}/{:.2f}".format(a, c) for f, (a, c, _) in b["loc"].items()})
    print("[cells] m={}% [{}-{}], kept {}/{}, offset {:.2f}->{:.2f} px, fell {}/{}, corr median {:.2f} ({:.2f}-{:.2f})".format(
        C["m"], C["m_lo"], C["m_hi"], C["n_kept"], C["n_all"], C["before"], C["after"],
        C["n_fell"], C["n_v"], C["corr_med"], C["corr_min"], C["corr_max"]))
    print("[cells] fields:", {f: "n={n} R={R:.2f} rand={R_rand:.2f} mean={mean_len:.2f}".format(**v)
                              for f, v in C["fields"].items()})
    print("[overlap]", ov, "[ccf]", ccf, "[planted]", pl.groups() if pl else None)

    b24, b26 = B["2024"], B["2026"]
    S = []
    add = S.append
    g24, g26 = 1000 * b24["m"], 1000 * b26["m"]     # shift growth, px per 1000 px from the zero point
    gc = 10 * C["m"]                                 # the same for the cells (C["m"] is in %)

    add(("title", "Channel misalignment on the Ti-E / W1 / BSI",
         "Bead measurements, how we correct it, and what to do next"))
    add(("text", "Outline", [
        "1  The offset on beads: 2024 (Aashli) and 2026 (Vishavdeep)",
        "2  2024 vs 2026 compared",
        "3  Why: lateral chromatic aberration",
        "4  Can we predict the zero point? (gel beads, Oct 2026)",
        "5  Correcting cell images without beads",
        "6  Discussion: other microscopes, correction, what is affected",
    ], ""))

    # --- 1. beads: 2024 and 2026 interleaved, slide type by slide type ----------
    add(("divider", "1 · The offset on beads",
         "2024 (Aashli) and 2026 (Vishavdeep) · TetraSpeck beads, 640 vs 488 nm"))
    # how beads are found, every bead set in the deck (FF 2026-10-15); detectBeads in
    # intranuclear_colocalization/scripts/bead_offset_study.m, settings in each run's RunMeta.json
    add(("text", "How beads are found (every data set)", DETECT_2D, DETECT_2D_SRC))
    YR = {"2024": "2024", "2026": "2026"}
    per = {}
    for k in ("2024", "2026"):
        b, lab = B[k], BEAD_LABEL[k]
        fields = []
        for f in b["fovs"]:
            pm = PANELS["{}_F{}".format(k, f)]
            sh = [bd["shift"] for bd in pm["beads"]]
            fields.append(("locator", "{} · field {}: {:.1f} → {:.1f} px".format(YR[k], f, sh[0], sh[-1]), k, f,
                           "{} · field {} · green = 640 nm, magenta = 488 nm, white line = shift".format(lab, f)))
        fz = min((f for f in b["fovs"] if PANELS["{}_F{}".format(k, f)]["in_field"]),
                 key=lambda f: np.hypot(*(np.array(PANELS["{}_F{}".format(k, f)]["null"]) - 1024)))
        pz = PANELS["{}_F{}".format(k, fz)]
        rows_i = [1, 4, 6]
        pp = PANELS["{}_plots".format(k)]
        rest = [
            ("beadgrid", "{} · one bead in every color: {:.1f}, {:.1f}, {:.1f} px".format(
                YR[k], *[pz["beads"][i - 1]["shift"] for i in rows_i]), k, fz, rows_i,
             "{} · field {} · beads 1, 4, 6 of the field slide".format(lab, fz)),
            ("image", "{} · all {:,} beads: +{:.1f} px per 1000 px".format(YR[k], b["n"], 1000 * pp["m"]),
             PANEL_DIR / "{}_shift_vs_dist.png".format(k),
             lab + " · each dot = one bead · dashed = one slope for all fields"),
            ("image", "{} · 3 bead-locating methods agree (≤{:.2f} px)".format(YR[k], b["meth_diff"]),
             PANEL_DIR / "{}_methods.png".format(k),
             lab + " · each dot = one bead: its shift measured two ways · dashed = identical"),
            ("table", "{} · whole-image correlation agrees (≤{:.1f} px)".format(YR[k], b["ncc_diff"]),
             ["field", "beads (x, y)", "whole image (x, y)"],
             [["{}".format(f), "{:+.1f}, {:+.1f} px".format(*b["per"][f]["med_cen"]),
               "{:+.1f}, {:+.1f} px".format(*b["per"][f]["ncc"])] for f in b["fovs"]],
             lab + " · 640 → 488 nm shift averaged over each field, two independent ways"),
        ]
        per[k] = (fields, rest)
    f24, f26 = per["2024"][0], per["2026"][0]
    for i in range(max(len(f24), len(f26))):          # field slides alternate 2024 / 2026
        if i < len(f24):
            add(f24[i])
        if i < len(f26):
            add(f26[i])
    for a24, a26 in zip(per["2024"][1], per["2026"][1]):
        add(a24)
        add(a26)

    # --- 2. 2024 vs 2026 ----------------------------------------------------------
    add(("divider", "2 · 2024 vs 2026 compared", "Same microscope, two bead collections"))
    add(("table", "2026 shifts are larger, but the pattern is the same",
         ["", "2024 (Aashli)", "2026 (Vishavdeep)"], [
             ["fields, beads", "{}, {:,}".format(len(b24["fovs"]), b24["n"]), "{}, {:,}".format(len(b26["fovs"]), b26["n"])],
             ["median shift", "{:.1f} px ({})".format(b24["med_all"], nm(b24["med_all"])),
              "{:.1f} px ({})".format(b26["med_all"], nm(b26["med_all"]))],
             ["largest shifts (98th pct)", "{:.1f} px ({})".format(b24["p98_all"], nm(b24["p98_all"])),
              "{:.1f} px ({})".format(b26["p98_all"], nm(b26["p98_all"]))],
             ["growth per 1000 px", "{:.1f} px".format(g24), "{:.1f} px".format(g26)],
             ["zero point inside the field", "{} of {} fields".format(b24["null_in"], len(b24["fovs"])),
              "{} of {} fields".format(b26["null_in"], len(b26["fovs"]))],
             ["scatter about the fit", "{:.2f} px".format(b24["resid_axis"]), "{:.2f} px".format(b26["resid_axis"])],
             ["median distance from zero point", "{:,.0f} px".format(PANELS["2024_dist"]["median"]),
              "{:,.0f} px".format(PANELS["2026_dist"]["median"])],
         ], "640 vs 488 nm channel · 2024 analyzed at each field's best focus, 2026 at a fixed mid-slice"))
    rows = []
    for i in range(max(len(b24["fovs"]), len(b26["fovs"]))):
        r24 = "{:.2f} px".format(1000 * b24["per"][b24["fovs"][i]]["m"]) if i < len(b24["fovs"]) else "–"
        r26 = "{:.2f} px".format(1000 * b26["per"][b26["fovs"][i]]["m"]) if i < len(b26["fovs"]) else "–"
        rows.append(["field {}".format(i + 1), r24, r26])
    rows.append(["all fields", "{:.2f} px".format(g24), "{:.2f} px".format(g26)])
    add(("table", "Shift growth per 1000 px, every field",
         ["", "2024 (Aashli)", "2026 (Vishavdeep)"], rows,
         "fit per field, centroid shifts, 640 → 488 nm; 2026 fields are FOV 1, 2, 5, 6"))
    fa24, fa26 = PANELS["2024_plots"]["arrow_field"], PANELS["2026_plots"]["arrow_field"]
    add(("image", "Same straight line: {:.1f} vs {:.1f} px per 1000 px".format(g24, g26),
         PANEL_DIR / "shift_vs_dist_2024_vs_2026.png",
         "dots = median shift of the beads in each 100 px distance bin · lines = fitted slope"))
    add(("pair", "Both years: shifts point away from ★ and grow",
         [PANEL_DIR / "2024_F{}_arrows.png".format(fa24), PANEL_DIR / "2026_F{}_arrows.png".format(fa26)],
         ["2024 (Aashli) · field {}".format(fa24), "2026 (Vishavdeep) · field {}".format(fa26)],
         "arrows drawn 40× longer than the shift"))
    d24, d26 = PANELS["2024_dist"]["median"], PANELS["2026_dist"]["median"]
    add(("image", "Median shift: {:.1f} px (2024) vs {:.1f} px (2026)".format(
        PANELS["2024_hist"]["median"], PANELS["2026_hist"]["median"]),
         PANEL_DIR / "shift_hist_2024_vs_2026.png", "every bead, all fields · dashed = median"))
    add(("image", "Why: 2026 zero points sit toward the top-left corner",
         PANEL_DIR / "zero_points_2024_vs_2026.png",
         "★ = each field's zero point (no shift) · numbers = field"))
    add(("text", "Same magnification, different zero point", [
        "Shift grows at the same rate: {:.1f} vs {:.1f} px per 1000 px".format(g24, g26),
        "But 2026 beads sit farther from the zero point: {:,.0f} vs {:,.0f} px (median)".format(d26, d24),
        "So 2026 shifts are ~2× larger: {:.1f} vs {:.1f} px".format(
            PANELS["2026_hist"]["median"], PANELS["2024_hist"]["median"]),
        "The zero point moves with the per-session shift: this is the part that changes",
    ], ""))

    # --- 3. why -----------------------------------------------------------------
    add(("divider", "3 · Why: lateral chromatic aberration", ""))
    add(("text", "Each color is imaged at a slightly different size", [
        "Glass bends colors differently, so each color has a slightly different magnification",
        "So one color's image is a slightly enlarged copy of the other: no shift at the center, growing outward",
        "At 1000 px from the axis: {:.1f}–{:.1f} px = {}–{}".format(g24, g26, nm(g24), nm(g26)),
        "Fixed by the optics: same in every field and in both years",
    ], ""))
    add(("equation", "The model we correct with",
         ["shift(x, y)  =  m · ( x − x_0 ,  y − y_0 )  +  ( t_x , t_y )"],
         ["(x, y): position on the camera (px)",
          "m: magnification difference, {:.1f}–{:.1f} px per 1000 px, fixed by the optics".format(g24, g26),
          "(x0, y0): center of the magnification",
          "(tx, ty): shift of that field, 0–3 px, different every time",
          "Zero point = where the two terms cancel; fits every bead to ~{:.1f} px".format(
              min(b24["resid_axis"], b26["resid_axis"]))]))

    add(("text", "Why would the shift change from field to field?", [
        "The magnification is fixed by the lenses: it cannot change per field",
        "So something tilts or moves one color by 1–2 px, differently in each field",
        "Four candidates (next slides); our data do not yet separate them",
    ], ""))
    add(("text", "1 · Defocus + a slightly tilted beam", [
        "If one color reaches the camera at a small angle, defocus slides it sideways",
        "Each field sits at slightly different focus (slide tilt, gel height, drift)",
        "The two colors also focus ~50 nm apart",
        "Fits: 2024 (best focus) barely moves; 2026 (fixed slice) moves more",
    ], ""))
    add(("text", "2 · The sample acts as a weak prism", [
        "A tilted gel or coverslip surface bends light",
        "Red and blue bend by different amounts",
        "A soft 1 kPa gel is unlikely to be flat: tilt varies field to field",
        "A sample effect, not a microscope fault",
    ], ""))
    add(("text", "3 · Drift between the two color exposures", [
        "If colors are taken one after the other",
        "Stage or sample moves slightly in between (thermal, mechanical)",
        "Amount differs per field and over time",
    ], ""))
    add(("text", "4 · Optics that move between fields", [
        "Filter wheel, turret or dichroic switching",
        "Never re-seats at exactly the same angle",
        "Each re-seat tilts one color's beam slightly differently",
    ], ""))
    add(("text", "How to tell them apart (beads)", [
        "Shift vs each field's focus (have it for 2024) → 1",
        "Beads on bare glass vs on gel → 2",
        "Same field imaged repeatedly, nothing moved → 3",
        "Same field, filter / turret cycled in between → 4",
    ], ""))

    # --- 4. gel beads: predicting the zero point ---------------------------------
    S.extend(gel_slides())

    # --- 5. cells ----------------------------------------------------------------
    add(("divider", "5 · Correcting the cell images",
         "No beads in the dish: the offset is measured from the nuclei"))
    add(("text", "Three steps", [
        "Measure: one offset arrow per nucleus (cross-correlation, widest slice)",
        "Fit: the bead model to all arrows → magnification + shift per field",
        "Correct: move one channel by the model at every pixel",
    ], ""))
    dx, dy, q = C["ex"]
    add(("pair", "One nucleus, widest slice: ERα and RNAP2",
         [PANEL_DIR / "meas_era.png", PANEL_DIR / "meas_pol.png", PANEL_DIR / "meas_merge.png"],
         ["ERα", "RNAP2", "merge (ERα green)"],
         "EtOH/01 nucleus 13 · scale bar 2 μm"))
    add(("pair", "Slide ERα over RNAP2: best match at ({:+.1f}, {:+.1f}) px".format(dx, dy),
         [PANEL_DIR / "meas_ccmap.png", PANEL_DIR / "meas_peak.png"],
         ["match for every shift of ERα", "best row: peak refined below 1 px"],
         "EtOH/01 nucleus 13 · white ○ = best whole pixel, red × = refined · this shift = the nucleus's arrow".format()))
    for f in FIELDS:
        v = PANELS["cells_{}".format(f)]
        add(("image", "{}: scattered, but a clear trend (R̄ {:.2f} vs {:.2f} random)".format(
            f, v["R"], v["R_rand"]), PANEL_DIR / "cells_{}_arrows.png".format(f.replace("/", "_")),
            "left: each nucleus's arrow at its camera position, drawn 120× · right: arrow directions, red = average"))
    add(("image", "Every field has an overall trend, far above random",
         PANEL_DIR / "cells_consistency.png", "R̄ = length of the average of the unit arrows (circular statistics) · black bar = expected for random directions, same n"))
    add(("image", "Fit over all {} nuclei: +{:.2f} px per 1000 px ({:.2f}–{:.2f})".format(
        C["n_kept"], gc, 10 * C["m_lo"], 10 * C["m_hi"]), PANEL_DIR / "cells_fit.png",
        "each dot = one nucleus, its field's shift removed · line = fit · beads: {:.2f} (2024), {:.2f} (2026) px per 1000 px".format(g24, g26)))
    if pl:
        add(("image", "Planted test: read back at ~100% ({}% / {}%)".format(pl.group(1), pl.group(2)),
             PANEL_DIR / "cells_planted.png",
             "known shift / magnification added to one channel of the 127 real nuclei, then re-measured · left: median and quartiles · right: 95% CI"))
    add(("text", "After the correction", [
        "Offset re-measured: {:.2f} → {:.2f} px (median, {} nuclei)".format(C["before"], C["after"], C["n_v"]),
        "Fell in {} of {} nuclei".format(C["n_fell"], C["n_v"]),
        "Correction per nucleus: {:.1f} px median ({:.1f}–{:.1f} px)".format(C["corr_med"], C["corr_min"], C["corr_max"]),
        "  ~1 px left = single-nucleus noise (ERα and RNAP2 are different proteins)",
    ], ""))
    for f in FIELDS:
        bf, af, n = C["per_field"][f]
        fs = f.replace("/", "_")
        add(("beforeafter", "{}: offsets shrink, {:.1f} → {:.1f} px (median)".format(f, bf, af),
             [PANEL_DIR / "fieldba_{}_before.png".format(fs), PANEL_DIR / "fieldba_{}_after.png".format(fs)],
             "{} nuclei · each arrow = one nucleus's ERα–RNAP2 offset, drawn 120× · after = re-measured on the corrected images".format(n)))
    # example nuclei: raw overlays with the correction drawn to scale
    for j in PANELS["rawcorr"]:
        c = float(np.hypot(*j["corr"]))
        add(("rawcorr", "{} nucleus {}: offset {:.1f} → {:.1f} px".format(j["tag"], j["nid"], j["before"], j["after"]),
             j["stem"], "ERα green, RNAP2 magenta, white = both · offset = ERα–RNAP2 shift re-measured in this nucleus "
             "· yellow box = zoom, next slide", "full"))
        add(("rawcorr", "{} nucleus {}, zoom: moved {:.1f} px ({})".format(j["tag"], j["nid"], c, nm(c)),
             j["stem"], "48 × 48 px box (3.1 μm) · correction from the field fit", "zoom"))
    # example nuclei: segmented foci before / after
    v = C["verify"]
    figs = sorted((VERIFY / "figures").glob("Nucleus_*.png"))
    by_field = {f: [] for f in FIELDS}
    for p in figs:
        m = re.match(r"Nucleus_(EtOH|E2)_(\d+)_(\d+)", p.stem)
        f, nid = "{}/{}".format(m.group(1), m.group(2)), int(m.group(3))
        if (f, nid) in SKIP_NUCLEI:
            continue
        row = v[(v["fov"] == f) & (v["nucID"] == nid)][0]
        c = float(np.hypot(row["corr_dx"], row["corr_dy"]))
        fb, fa = float(row["fER_before"]), float(row["fER_after"])
        ov_txt = "; overlap {:.2f} → {:.2f}".format(fb, fa) if np.isfinite(fb) and np.isfinite(fa) else ""
        ttl = ("{} nucleus {}: overlap {:.2f} → {:.2f}".format(f, nid, fb, fa) if ov_txt
               else "{} nucleus {}: segmented foci".format(f, nid))
        item = ("beforeafter", ttl,
                [PANEL_DIR / "seg_{}_seg_before.png".format(p.stem), PANEL_DIR / "seg_{}_seg_after.png".format(p.stem)],
                "segmented: ERα condensates green, RNAP2 foci magenta, overlap white · overlap = fraction of ERα area "
                "on RNAP2 · corrected by {:.1f} px ({})".format(c, nm(c)))
        by_field[f].append((c, bool(ov_txt), item))
    main_figs, backup_figs = [], []
    for f in FIELDS:   # per field: the two largest corrections that have an overlap value
        ranked = sorted(by_field[f], key=lambda t: (not t[1], -t[0]))
        main_figs += [t[2] for t in ranked[:2]]
        backup_figs += [t[2] for t in ranked[2:]]
    S.extend(main_figs)
    o = {c[0]: c for c in ov}
    E = PANELS["effect"]
    add(("image", "Overlap rises ~{:.0f}%, not significant".format(
        50 * sum(E[c]["after"] / E[c]["before"] - 1 for c in ("EtOH", "E2"))),
        PANEL_DIR / "effect_overlap.png",
        "widest slice, same nuclei measured twice · gray lines = single nuclei · bold = mean · paired Wilcoxon"))
    add(("image", "CCF rises ~{:.0f}% at zero shift, not significant".format(
        50 * sum(E["ccf_" + c]["after"] / E["ccf_" + c]["before"] - 1 for c in ("EtOH", "E2"))),
        PANEL_DIR / "effect_ccf.png",
        "widest slice · mean ± SEM over nuclei · dashed = 90°-rotated control · paired Wilcoxon on CCF(0)"))
    add(("text", "Two pipelines, two directions", [
        "Colocalization pipeline: moves RNAP2 onto ERα",
        "VV_RNAP: moves ERα onto RNAP2",
        "  so RNAP2 foci and all RNAP2-only numbers stay exactly the same",
        "Either can be switched: warpChannel = 'POL' / 'ER', or enable = false",
    ], ""))

    # --- 4. discussion ----------------------------------------------------------
    add(("divider", "6 · Discussion", ""))
    add(("text", "Hardware update (Matthew; Nikon: Andy, Sarah)", [
        "Seen on many lenses, silicone and water",
        "Not the W1: same alignment with the W1 on and off",
        "Left: the BSI camera, or something inside the Ti-E",
        "Loaner camera arrives 10/12; Nikon test proposed 10/14 or 10/15",
        "Taking the microscope apart: last resort",
    ], ""))
    add(("text", "What our bead data say", [
        "Radial, no rotation, same in 2024 and 2026 ({:.1f} vs {:.1f} px per 1000 px)".format(g24, g26),
        "That is the signature of lateral chromatic aberration",
        "It is stable, so software corrects it",
        "Nothing here by itself argues for taking the microscope apart",
    ], ""))
    add(("text", "Next step: the same beads on other microscopes", [
        "Same TetraSpeck slide on: Losert lab, Mukhina lab, imaging incubator",
        "Same colors, z-stack through focus, ~5 fields",
        "Our analysis runs unchanged → magnification per color pair, comparable",
        "  bead_offset_study.m + build_bead_report.py",
    ], ""))
    add(("text", "Correcting in post-processing", [
        "Two-color analyses at pixel scale: correct magnification + per-field shift (done, both pipelines)",
        "Needs ≥10 nuclei per field to pin the shift; beads in the same session would remove that need",
        "Single-channel analyses: no correction needed",
    ], ""))
    add(("text", "What is affected most", [
        "Anything comparing two colors pixel by pixel:",
        "  overlap fraction, CCF, foci-centered averages, line-profile groups (center zone ±2 px ≈ the offset)",
        "Worse at the field edge (up to ~4–6 px) than at the center",
        "Not affected: single-channel intensity, foci counts, nuclear shape",
    ], ""))
    add(("text", "Open questions", [
        "Cells give {:.2f} px per 1000 px, beads {:.2f}–{:.2f}: why smaller?".format(gc, g24, g26),
        "What sets the per-field shift?",
        "Beads in the same session as the cells would answer both",
    ], ""))

    # --- backup ------------------------------------------------------------------
    add(("divider", "Backup", ""))
    if (BEAD["2024"] / "QC_AxialFocus.png").exists():
        add(("image", "2024 · focus per color: 488 nm within ~0.05 μm of 640 nm",
             prep(BEAD["2024"] / "QC_AxialFocus.png"), BEAD_LABEL["2024"]))
    for k in ("2024", "2026"):
        add(("image", "{} · per-bead x and y shift, by field".format(k), prep(BEAD[k] / "QC_ShiftHist.png"), BEAD_LABEL[k]))
    S.extend(backup_figs)
    return S


def main():
    list_only = "--list" in sys.argv[1:]
    CROP_DIR.mkdir(parents=True, exist_ok=True)
    plan = build_plan()
    for i, s in enumerate(plan, 1):
        print("{:3d} {:<8} {}".format(i, s[0], s[1]))
    if list_only:
        return
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    for s in plan:
        kind = s[0]
        if kind == "title":
            title_slide(prs, s[1], s[2])
        elif kind == "divider":
            divider_slide(prs, s[1], s[2])
        elif kind == "text":
            text_slide(prs, s[1], s[2], s[3])
        elif kind == "image":
            image_slide(prs, s[1], s[2], s[3])
        elif kind == "pair":
            pair_slide(prs, s[1], s[2], s[3], s[4])
        elif kind == "table":
            table_slide(prs, s[1], s[2], s[3], s[4])
        elif kind == "equation":
            equation_slide(prs, s[1], s[2], s[3])
        elif kind == "beforeafter":
            beforeafter_slide(prs, s[1], s[2], s[3])
        elif kind == "rawcorr":
            rawcorr_slide(prs, s[1], s[2], s[3], s[4])
        elif kind == "locator":
            locator_slide(prs, s[1], s[2], s[3], s[4])
        elif kind == "beadcompare":
            beadcompare_slide(prs, s[1], s[2], s[3])
        elif kind == "beadgrid":
            beadgrid_slide(prs, s[1], s[2], s[3], s[4], s[5])
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        M.backup_presentation(str(OUT), backup_base=str(OUT_DIR / "backups"))
    prs.save(str(OUT))
    print("Done. {} slides -> {}".format(len(plan), OUT))


if __name__ == "__main__":
    main()
