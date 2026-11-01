#!/usr/bin/env python
"""
PRELIMINARY montage decks for the 10/29/2026 H3K27me3 Jurkat sets (Noco, Ciliobrevin).

One slide per pipeline montage chunk (prog_fixed_cells/<well>/physical_scale_images/
<family>/montages/montage_cells_<a>_<b>.png), conditions in compile order. Chunk names
change with the cell count, so the slide list is globbed at build time and handed to
batch_insert_images.py as a temporary config.

Every well reads prog_fixed_cells/ (the temporary copies from the actin-mask rerun
lacked cells, bug hunt 2026-10-31). A well with no montages gets a "missing" slide
instead of a silent gap.

Usage:
    python build_h3k27me3_10292026_montage_decks.py [noco|cilio|all]
"""

import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
PROG = "prog_fixed_cells"
TEMPLATE = "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Noco/Effects of Noco, Fixed Cells_MT_template.pptx"
PRELIM = "PRELIMINARY: no figure review yet"

# (family, legend). "{side}" is the experiment's side-view colors: the two sets render
# actin and centrin in swapped colors (checked on the montages, 2026-10-31).
FAMILIES = [
    ("actin_cent_xz_nolines", "Side view: {side}"),
    ("H3K27me3_nuc_bz", "Top view, broadest nuclear slice: H3K27me3 (red), DNA (cyan)"),
]

EXPS = {
    "noco": {
        "root": "L:/FF/Nucleus_centrosome/nocodazole_fixed/10292026_jurkats_noco_h3k27me3640LP45_p549LP45_egfpcentrin2488LP45_h405LP40_",
        "out_dir": "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/Noco/H3K27me3_20261029",
        "out": "Noco_Jurkats_H3K27me3_montages_10292026.pptx",
        "side": "actin (yellow), centrin (magenta)",
        # (condition label, well folder, source root under the experiment; None = show as
        # missing).
        "wells": [("PLL DMSO", "G1A4_h3k27me3_PLL_dmso_", PROG),
                  ("CD3 DMSO", "G1A1_h3k27me3_aCD3_dmso_", PROG),
                  ("CD3 Noco 5 μM", "G1B1_h3k27me3_aCD3_5uM_", PROG),
                  ("CD3 Noco 10 μM", "G1B2_h3k27me3_aCD3_10uM_", PROG)],
        # A template is needed for the preserved title slide; its text is replaced, as in the results deck.
        "preserve": [0, 1],
        "title": "Effects of Noco, H3K27me3, October 29, 2026",
    },
    "cilio": {
        "root": "L:/FF/Nucleus_centrosome/ciliobrevin_fixed/10292026_jurkats_ciliobrevin_h3k27me3640LP45_p549LP45_egfpcentrin2488LP45_h405LP40_",
        "out_dir": "K:/FF/PPT/PPT_autogeneration/Fixed Jurkats, Miscellaneous/CilioD/H3K27me3_20261029",
        "out": "CilioD_Jurkats_H3K27me3_montages_10292026.pptx",
        "side": "actin (magenta), centrin (yellow)",
        "wells": [("PLL DMSO", "G0A4_h3k27me3_PLL_dmso_", PROG),
                  ("CD3 DMSO", "G0A1_h3k27me3_aCD3_dmso_001", PROG),
                  ("CD3 Cilio 50 μM", "G0B2_h3k27me3_aCD3_cilio_", PROG)],
        # No Nocodazole title slide in the CilioD deck (same as its results deck).
        "preserve": [1],
    },
}


def long_path(p):
    """Extended-length form: the well folders run past Windows' 260-character limit."""
    s = str(p).replace("/", "\\")
    return "\\\\?\\" + s if sys.platform == "win32" and not s.startswith("\\\\?\\") else s


def chunks(montage_dir):
    """Montage chunk paths sorted by first cell id, or [] when the folder is missing."""
    try:
        names = os.listdir(long_path(montage_dir))
    except OSError:
        return []
    found = []
    for n in names:
        m = re.fullmatch(r"montage_cells_(\d+)_(\d+)\.png", n)
        if m:
            found.append((int(m.group(1)), int(m.group(2)), Path(montage_dir) / n))
    return sorted(found)


def missing_png(out_dir):
    """A plain 'montages missing' image for wells with nothing on disk."""
    from PIL import Image, ImageDraw, ImageFont
    path = Path(out_dir) / "montages_missing_placeholder.png"
    if not path.exists():
        img = Image.new("RGB", (1600, 800), "white")
        d = ImageDraw.Draw(img)
        d.rectangle([20, 20, 1580, 780], outline=(150, 150, 150), width=6)
        d.text((800, 340), "Montages missing", font=ImageFont.truetype("arialbd.ttf", 64),
               fill="black", anchor="mm")
        d.text((800, 450), "No montage sheets on disk for this well",
               font=ImageFont.truetype("arial.ttf", 40), fill=(90, 90, 90), anchor="mm")
        img.save(path)
    return str(path)


def build(name):
    exp = EXPS[name]
    images = []
    for fam, fam_label in FAMILIES:
        fam_label = fam_label.format(side=exp["side"])
        for cond, well, src in exp["wells"]:
            mdir = Path(exp["root"]) / src / well / "physical_scale_images" / fam / "montages" if src else None
            found = chunks(mdir) if src else []
            if not found:
                print(f"[MISSING] {name} {cond} {fam}: {mdir}")
                images.append({"path": missing_png(exp["out_dir"]), "title": f"{cond}: montages missing",
                               "subtitle": f"{fam_label} · {PRELIM}"})
                continue
            for a, b, p in found:
                images.append({"path": long_path(p), "title": f"{cond}, cells {a}–{b}",
                               "subtitle": f"{fam_label} · bar = 5 μm · {PRELIM}"})
    cfg = {
        "presentation": TEMPLATE,
        "output_path": f"{exp['out_dir']}/{exp['out']}",
        "template_slide": 1,
        "backup_dir": f"{exp['out_dir']}/backups",
        "preserve_slides": exp["preserve"],
        "title_slide_text": exp.get("title"),
        "auto_position": False,
        "position": {"left": 0.3, "top": 1.17, "width": 12.73, "height": 6.15},
        "trim_whitespace": False,
        "add_label": False,
        "remove_template_slide": True,
        "images": images,
    }
    with tempfile.NamedTemporaryFile("w", suffix=".yaml", delete=False, encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)
        cfg_path = f.name
    try:
        rc = subprocess.run([sys.executable, str(HERE / "batch_insert_images.py"), cfg_path]).returncode
    finally:
        os.remove(cfg_path)
    return rc


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    names = list(EXPS) if which == "all" else [which]
    sys.exit(max(build(n) for n in names))
