"""
build_biad2255_deck.py

BIAD2255 smFISH deck (Lattanzio et al., EMBO J 2025, public data), driven by a YAML config:
    examples_and_configs/configs/public_data/biad2255_smfish_deck.yaml

Style = the guide deck K:/FF/PPT/PPT_autogeneration/MCF7_channel_misalignment/
Channel_Misalignment_GroupMeeting.pptx: its own slide functions (title, divider, text, image) are
imported from insert_channel_misalignment_slides.py, not redrawn. Images are placed as they are.

Slide kinds in the config: title, divider, text, image (optional crop), pair (wide images stacked,
labels at the left), montages, table (header + rows), placeholder. A slide with draft: true
gets a red tag (config draft_tag) at the bottom left.

Usage (base miniforge: python-pptx + pywin32 for the PNG export):
    python examples_and_configs/public_data/build_biad2255_deck.py <config.yaml>
    ... --list      dry run: print every slide, check images and text, build nothing
    ... --export    also export slide PNGs (PowerPoint COM) to the config's export_dir
"""

import os
import re
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import insert_channel_misalignment_slides as CM

M = CM.M
RED = RGBColor(0xC0, 0x00, 0x00)
TAG_W = 2.6
MICRO_SIGN = "\u00b5"          # never in slide text: Greek mu U+03BC only
CAPTION_FONT = "C:/Windows/Fonts/calibri.ttf"   # the deck's theme font (Calibri), to measure captions


MONTAGE_FOOT_PT = 18          # = insert_noco_10262026_montage_slides.STRIP_FOOT_FONT_PT


def caption_width_in(text, pt=None):
    """Width of a caption in inches at pt (default: the guide deck's footer size), measured with the font file."""
    from PIL import ImageFont
    font = ImageFont.truetype(CAPTION_FONT, 1000)
    return font.getlength(text) / 1000 * (pt or CM.FOOT_PT) / 72


def add_draft_tag(slide, text):
    M.add_textbox(slide, text, CM.MARGIN, CM.H - CM.FOOT_H, TAG_W, CM.FOOT_H,
                  font_pt=CM.FOOT_PT, color=RED, bold=True, align=PP_ALIGN.LEFT)


def placeholder_slide(prs, title, body):
    s = M._new_slide(prs)
    CM.add_title(s, title)
    box = M.add_textbox(s, body, CM.MARGIN + 1.0, CM.IMG_TOP + 0.5, CM.IMG_W - 2.0, CM.IMG_H - 1.0,
                        font_pt=28, color=CM.GREY, align=PP_ALIGN.CENTER)
    box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    box.line.color.rgb = CM.GREY
    box.line.width = Pt(1.5)
    return s


TABLE_PT = 20


def table_slide(prs, title, header, rows, caption):
    """A plain table (guide-deck title and footer): header row bold, one row per entry, every cell
    centered. FF 2026-11-17: a page to validate shows few numbers, in an aligned table."""
    s = M._new_slide(prs)
    CM.add_title(s, title)
    row_h = 0.55
    shape = s.shapes.add_table(len(rows) + 1, len(header), Inches(CM.MARGIN + 0.3), Inches(CM.IMG_TOP + 0.2),
                               Inches(CM.IMG_W - 0.6), Inches(row_h * (len(rows) + 1)))
    tbl = shape.table
    for r, values in enumerate([header] + rows):
        for c, v in enumerate(values):
            cell = tbl.cell(r, c)
            cell.text = str(v)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.size, run.font.bold = Pt(TABLE_PT), r == 0
    if caption:
        M.add_textbox(s, caption, CM.MARGIN, CM.H - CM.FOOT_H, CM.IMG_W, CM.FOOT_H, font_pt=CM.FOOT_PT,
                      color=CM.GREY, align=PP_ALIGN.RIGHT)
    return s


def montage_rows_slide(prs, title, rows, footer, side_by_side=False):
    """Pre-composed montages, one labeled row per condition, in the Oct 2026 Noco / H3K27me3 montage
    look (insert_noco_10262026_montage_slides: black slide, white title, white condition labels,
    gray footer). Every montage is placed at ONE px per inch, so cells compare at physical scale
    when the montages share one px/μm (they carry their own 5 μm bar). rows: [(label, path)]."""
    import insert_noco_10262026_montage_slides as N
    from PIL import Image
    s = N.new_slide(prs, title, None)
    sizes = []
    for _, p in rows:
        with Image.open(p) as im:
            sizes.append(im.size)
    if side_by_side:
        # FF 2026-11-19: "0 hr should be on left, 2 hr on right": one column per condition, its label
        # centered above it, every montage at ONE px per inch
        lab_h, gap = 0.45, 0.25
        col_w = (N.SLIDE_W - 2 * N.MARGIN - gap * (len(rows) - 1)) / len(rows)
        img_h = N.SLIDE_H - N.GRID_TOP - lab_h - N.HULL_FOOT_H
        ppi = max(max(w / col_w, h / img_h) for w, h in sizes)
        top = N.GRID_TOP + (img_h - max(h for _, h in sizes) / ppi) / 2   # block centered between title and footer
        for i, ((label, p), (w, h)) in enumerate(zip(rows, sizes)):
            left = N.MARGIN + i * (col_w + gap)
            box = N.add_textbox(s, label, left, top, col_w, lab_h, font_pt=N.HULL_LABEL_FONT_PT,
                                color=N.WHITE, bold=True)
            box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            box.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            s.shapes.add_picture(str(p), Inches(left + (col_w - w / ppi) / 2), Inches(top + lab_h),
                                 width=Inches(w / ppi), height=Inches(h / ppi))
        print(f"  montage slide '{title}': {ppi:.1f} px/in for every column, {sizes[0][0] / ppi:.2f} in wide")
        rows = []                 # labels and pictures placed; only the footer below
    lab_w, gap = 1.1, 0.12        # label set on two lines ("2 h" / "control"); FF: "things need to be big"
    area_w = N.SLIDE_W - 2 * N.MARGIN - lab_w
    row_h = (N.SLIDE_H - N.GRID_TOP - N.HULL_FOOT_H - gap * (max(len(rows), 1) - 1)) / max(len(rows), 1)
    if rows:
        ppi = max(max(w / area_w, h / row_h) for w, h in sizes)
    for i, ((label, p), (w, h)) in enumerate(zip(rows, sizes)):
        top = N.GRID_TOP + i * (row_h + gap)
        # "2 h control" -> "2 h" / "control": one break per space after the first word group, set as a
        # line break so the narrow column never clips a word
        head, _, rest = label.partition(" h ")
        text = f"{head} h\v{rest}" if rest else label
        box = N.add_textbox(s, text, N.MARGIN, top, lab_w, row_h, font_pt=N.HULL_LABEL_FONT_PT,
                            color=N.WHITE, bold=True)
        box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        box.text_frame.paragraphs[0].alignment = PP_ALIGN.LEFT
        for r in box.text_frame.paragraphs[0].runs:     # add_textbox styles the first run only
            r.font.size, r.font.bold, r.font.color.rgb = Pt(N.HULL_LABEL_FONT_PT), True, N.WHITE
        width_in =max(caption_width_in(w, N.HULL_LABEL_FONT_PT) * 1.08 for w in text.split("\v"))  # bold ≈ +8 %
        if width_in > lab_w - 0.1:
            raise SystemExit(f"[montage] label '{label}' line is {width_in:.2f} in, column {lab_w - 0.1:.2f} in")
        s.shapes.add_picture(str(p), Inches(N.MARGIN + lab_w + (area_w - w / ppi) / 2),
                             Inches(top + (row_h - h / ppi) / 2), width=Inches(w / ppi), height=Inches(h / ppi))
    if rows:
        print(f"  montage slide '{title}': {ppi:.1f} px/in for every row")
    if footer:
        box = N.add_textbox(s, footer, CM.MARGIN + TAG_W, N.SLIDE_H - N.HULL_FOOT_H, N.SLIDE_W - CM.MARGIN - TAG_W
                            - N.MARGIN, N.HULL_FOOT_H, font_pt=N.STRIP_FOOT_FONT_PT, color=N.GRAY)
        box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        box.text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    return s


def slide_text(sp):
    parts = [sp.get("title", ""), sp.get("subtitle", ""), sp.get("caption", ""), sp.get("body", "")]
    if sp.get("kind") == "table":
        return parts + [str(v) for r in [sp["header"]] + sp["rows"] for v in r]
    return (parts + list(sp.get("bullets", [])) + list(sp.get("labels", []))
            + [r["label"] for r in sp.get("rows", [])])


def image_path(cfg, sp, rel, crop_dir):
    """Source image, or for an entry with crop: [x0, y0, x1, y1] (fractions of the image) a cropped,
    white-trimmed copy written to crop_dir (beside the deck; the source is never modified)."""
    src = Path(cfg["base_dir"]) / rel
    if "crop" not in sp:
        return src
    from PIL import Image
    x0, y0, x1, y1 = sp["crop"]
    out = crop_dir / "{}__crop_{}.png".format(src.stem, "_".join(f"{round(100 * v):03d}" for v in sp["crop"]))
    with Image.open(src) as im:
        w, h = im.size
        CM.trim(im.crop((int(x0 * w), int(y0 * h), int(x1 * w), int(y1 * h)))).save(out)
    return out


def check(cfg):
    """Deterministic checks before building; returns a list of problems."""
    base = Path(cfg["base_dir"])
    probs = []
    for i, sp in enumerate(cfg["slides"], 1):
        if sp["kind"] not in ("title", "divider", "text", "image", "pair", "montages", "placeholder", "table"):
            probs.append("slide {}: unknown kind {}".format(i, sp["kind"]))
        if sp["kind"] == "table":
            bad = [r for r in sp["rows"] if len(r) != len(sp["header"])]
            if bad:
                probs.append(f"slide {i}: {len(bad)} table row(s) without {len(sp['header'])} cells")
            if len(sp["rows"]) + 1 > int((CM.IMG_H - 0.3) / 0.55):
                probs.append(f"slide {i}: {len(sp['rows'])} table rows do not fit")
        rels = ([sp["image"]] if sp["kind"] == "image" else [] if sp["kind"] == "table"
                else sp.get("images", []) + [r["image"] for r in sp.get("rows", [])])
        for rel in rels:
            if not (base / rel).is_file():
                probs.append(f"slide {i}: missing image {base / rel}")
        if sp["kind"] == "pair" and len(sp["images"]) != len(sp.get("labels", [])):
            probs.append("slide {}: {} images but {} labels".format(i, len(sp["images"]), len(sp.get("labels", []))))
        for t in slide_text(sp):
            if MICRO_SIGN in t:
                probs.append(f"slide {i}: micro sign U+00B5 in '{t}'")
            if re.search(r"\bdecon", t, re.IGNORECASE):
                probs.append(f"slide {i}: deconvolution mentioned in '{t}'")
        # right-aligned caption must end before the draft tag's box starts (+0.1 in for text margins)
        room = CM.IMG_W - TAG_W - 0.1
        cw = caption_width_in(sp.get("caption", ""), MONTAGE_FOOT_PT if sp["kind"] == "montages" else None)
        if sp.get("draft") and cw > room:
            probs.append(f"slide {i}: caption {cw:.2f} in wide, room beside the DRAFT tag {room:.2f} in")
    return probs


def build(cfg):
    crop_dir = Path(cfg["output"]).parent / "source_images"
    crop_dir.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(CM.W), Inches(CM.H)
    titles = []
    for sp in cfg["slides"]:
        k = sp["kind"]
        if k == "title":
            s = CM.title_slide(prs, sp["title"], sp.get("subtitle", ""))
        elif k == "divider":
            s = CM.divider_slide(prs, sp["title"], sp.get("subtitle", ""))
        elif k == "text":
            s = CM.text_slide(prs, sp["title"], sp["bullets"], sp.get("caption", ""))
        elif k == "image":
            s = CM.image_slide(prs, sp["title"], str(image_path(cfg, sp, sp["image"], crop_dir)), sp.get("caption", ""))
        elif k == "montages":   # one row per condition, one physical scale (Oct 2026 montage look)
            s = montage_rows_slide(prs, sp["title"], [(r["label"], Path(cfg["base_dir"]) / r["image"])
                                                      for r in sp["rows"]], sp.get("caption", ""),
                                   side_by_side=sp.get("side_by_side", False))
        elif k == "table":
            s = table_slide(prs, sp["title"], sp["header"], sp["rows"], sp.get("caption", ""))
        elif k == "pair":   # wide images stacked, label at the left (guide deck's pair_slide)
            s = CM.pair_slide(prs, sp["title"], [str(image_path(cfg, sp, r, crop_dir)) for r in sp["images"]],
                              sp["labels"], sp.get("caption", ""))
        else:
            s = placeholder_slide(prs, sp["title"], sp["body"])
        if sp.get("draft"):
            add_draft_tag(s, cfg["draft_tag"])
        titles.append(sp["title"])
    return prs, titles


def export_pngs(deck, export_dir, titles):
    """Slide PNGs via PowerPoint, same rules as build_pmlc_profiles_deck.export_pngs: attach to the
    running PowerPoint, never Quit it, skip if FF has this deck open, close only what we opened."""
    import win32com.client as w
    app = w.Dispatch("PowerPoint.Application")
    target = os.path.normcase(os.path.abspath(str(deck)))
    for k in range(1, app.Presentations.Count + 1):
        try:
            full = app.Presentations(k).FullName
        except Exception:  # noqa: BLE001  protected-view decks deny FullName
            continue
        if os.path.normcase(full) == target:
            print(f"[skip export] {deck.name} is open in PowerPoint")
            return
    export_dir = Path(export_dir)
    export_dir.mkdir(parents=True, exist_ok=True)
    manifest = export_dir / ".slide_export_manifest.txt"
    if manifest.is_file():
        for n in manifest.read_text(encoding="utf-8").splitlines():
            if n.strip() and (export_dir / n).is_file():
                (export_dir / n).unlink()
    written = []
    pres = app.Presentations.Open(str(deck), WithWindow=False)
    try:
        for i, t in enumerate(titles, start=1):
            name = re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9+\-, ]+", " ", t)).strip()
            png = f"{i:02d} {name}.png"
            pres.Slides(i).Export(os.path.join(str(export_dir), png), "PNG", 1920, 1080)
            written.append(png)
    finally:
        pres.Close()
    manifest.write_text("\n".join(written), encoding="utf-8")
    print(f"Exported {len(written)} slide PNGs -> {export_dir}")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        sys.exit(__doc__)
    cfg = yaml.safe_load(Path(args[0]).read_text(encoding="utf-8"))
    for i, sp in enumerate(cfg["slides"], 1):
        print("{:3d} {:<11} {}".format(i, sp["kind"], sp["title"]))
    probs = check(cfg)
    for p in probs:
        print("[PROBLEM]", p)
    if probs:
        sys.exit(1)
    if "--list" in sys.argv:
        return
    out = Path(cfg["output"])
    if (out.parent / ("~$" + out.name)).exists():
        sys.exit(f"[locked] {out.name} is open in PowerPoint; close it first")
    prs, titles = build(cfg)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.exists():
        M.backup_presentation(str(out), backup_base=str(out.parent / "backups"))
    prs.save(str(out))
    print(f"Done. {len(titles)} slides -> {out}")
    if "--export" in sys.argv:
        export_pngs(out, cfg["export_dir"], titles)


if __name__ == "__main__":
    main()
