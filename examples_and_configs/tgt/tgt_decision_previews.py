"""Slides for the TGT movie deck's decision page (decision_page/README.md), made by the deck's own builder
(build_cat_fmc_slides.py: config_plan, build_entries) and export_deck_slides.py.

Two decks, each built and exported once:
- "now": the whole deck as the config builds it today, for the All slides view (text check on). When the spec's deck
  names `live: <deck>` (and `pairs: true` for the movie-pairs deck), that deck is exported instead of a new build, after
  a check that it is the config's build (its sidecar's slide hashes, incremental.py); it is only read;
- "options": for every open decision with layout options, the slides that option changes (the config's `layout`
  with the option's values), up to `max_slides` per option, titles numbered as in that option's deck.
Each slide's lines come from the built deck's own text (python-pptx); a bullet's "• " is dropped so a line matches the
config string it came from (decision_page/text_map.py replaces whole config strings).

Writes <OUT>/page/*.png and <OUT>/page/page_data.json, read by decision_page/build_decision_page.py.
Usage (env PPT_editing):  python tgt_decision_previews.py <decision spec.yaml>
"""
import copy
import shutil
import subprocess
import sys
import time
import json
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import build_cat_fmc_slides as B  # noqa: E402

OUT = B.DECK_DIR / "decision_page"
EXPORT = HERE.parent / "export_deck_slides.py"
BASE_PYTHON = Path(r"C:\Users\Upadhyaya_Lab\miniforge3\python.exe")   # export_deck_slides.py: base miniforge (pywin32)
MOVIE_NOTE = "Labels inside the movies are not listed."


def step(name, t0):
    print(f"[{name}] {time.perf_counter() - t0:.1f} s", flush=True)
    return time.perf_counter()


def signature(e):
    """What a slide shows, without its slide number: text boxes (text, size) or the movies' day and fields."""
    if e[0] == "slide":
        return ("slide",) + tuple((t.text, t.font_size_pt) for t in e[1].textboxes)
    return ("movie", e[3], B.describe(e[2], e[3]))


def slide_lines(prs):
    """Per-slide text lines (1-based slide number -> lines), a bullet's "• " dropped."""
    lines = {}
    for n, slide in enumerate(prs.slides, 1):
        ls = []
        for sh in slide.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    t = p.text.strip()
                    if t:
                        ls.append(t[2:] if t.startswith("\u2022 ") else t)
        lines[n] = ls
    return lines


def export(deck, name, n_expected):
    """Export every slide of deck through PowerPoint; returns the PNG paths in slide order."""
    src = OUT / "src" / name
    shutil.rmtree(src, ignore_errors=True)
    subprocess.run([str(BASE_PYTHON), str(EXPORT), str(deck), str(src), "1600"], check=True)
    pngs = sorted(src.glob("slide_*.png"), key=lambda p: int(p.stem.split("_")[1]))   # slide_100 after slide_99
    if len(pngs) != n_expected:
        raise SystemExit(f"{name}: exported {len(pngs)} slides, expected {n_expected}")
    return pngs


def build_and_export(entries, name, approved, png_dir, pairs=False):
    """Build one deck from entries, export every slide; returns (png paths, per-slide text lines)."""
    deck = OUT / f"{name}.pptx"
    deck.unlink(missing_ok=True)   # this script's own preview deck from its last run
    prs, _ = B.build_entries(entries, deck, approved, pairs=pairs)
    return export(deck, name, len(entries)), slide_lines(prs)


def live_deck(entries, cfg, live, pairs):
    """The live deck instead of a fresh build: it must be the config's build today (its sidecar's slide hashes equal
    the hashes of the slides the config makes, and the deck is unchanged since the sidecar was written, md5).
    Nothing is built and the deck is only read. Returns (png paths, per-slide text lines)."""
    from pptx import Presentation
    scratch = Path(B.make_run_scratch_dir(Path(B.tempfile.gettempdir()) / "tgt_cat_fmc_posters"))
    try:   # deck_specs names the posters only; it writes none
        slides, _ = B.deck_specs(entries, B.MovieSlideMaker(scratch, B.removed_marks(cfg)), B.removed_marks(cfg),
                                 cfg.get("fov_numbers"), cfg.get("movie_notes"), pairs)
    finally:
        shutil.rmtree(scratch, ignore_errors=True)
    want = [B.incremental.spec_hash(s, B.SLIDE_W, B.SLIDE_H, True, B.POSTER_RULE, True) for s in slides]
    have = B.incremental.read_sidecar(str(live))   # stops if the deck changed since its sidecar
    if have != want:
        bad = [k + 1 for k in range(max(len(have), len(want)))
               if k >= len(have) or k >= len(want) or have[k] != want[k]]
        raise SystemExit(f"STOP: {live} is not the config's build ({len(have)} slides in the deck, {len(want)} from "
                         f"the config; slides that differ: {bad[:20]}): build the deck first")
    print(f"CHECK {live.name} is the config's build ({len(want)} slides, sidecar hashes and md5): PASS", flush=True)
    return export(live, "live", len(entries)), slide_lines(Presentation(str(live)))


def main(spec_path):
    t0 = time.perf_counter()
    spec_path = Path(spec_path)
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    cfg_path = spec_path.parent / spec["deck"]["text"]
    cfg = yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    approved = spec_path.parent / spec["deck"]["approved"]
    ap = spec_path.parent / spec["applied"]
    applied = {str(a["code"]) for a in (yaml.safe_load(ap.read_text(encoding="utf-8")) or []) if "code" in a} \
        if ap.exists() else set()
    page_dir = OUT / "page"
    shutil.rmtree(page_dir, ignore_errors=True)
    page_dir.mkdir(parents=True)
    t = step("load", t0)

    # the deck as it is: the live deck when the spec names one (checked to be the config's build), else a fresh build
    pairs = bool(spec["deck"].get("pairs"))
    base = B.config_plan(cfg, pairs=pairs)
    live = spec["deck"].get("live")
    if live and any(o.get("layout") for it in spec["items"] for o in it.get("options", [])):
        raise SystemExit("STOP: layout options need fresh builds; drop deck.live from the spec for them")
    pngs, lines = (live_deck(base, cfg, Path(live), pairs) if live
                   else build_and_export(base, "now", approved, page_dir, pairs))
    deck = []
    for n, e in enumerate(base, 1):
        png = f"deck_{n:02d}.png"
        shutil.copyfile(pngs[n - 1], page_dir / png)
        deck.append({"n": n, "id": f"slide_{n:02d}", "png": png, "lines": lines[n],
                     "unlisted": MOVIE_NOTE if e[0] == "movie" else "", "fact": None})
    t = step(f"now: {len(base)} slides {'read from the live deck' if live else 'built'} and exported", t)

    # one options deck: each open layout option's changed slides
    base_sigs = [signature(e) for e in base]
    entries, where, decisions = [], [], {}
    for it in spec["items"]:
        code = str(it["code"])
        if code in applied or "options" not in it:
            continue
        dd = decisions[code] = {"applies": list(it.get("applies", [])), "options": {}}
        for o in it["options"]:
            if not o.get("layout"):
                continue
            cfg2 = copy.deepcopy(cfg)
            cfg2["layout"] = dict(cfg.get("layout") or {}, **o["layout"])
            plan = B.config_plan(cfg2)
            sigs = set(base_sigs)
            changed = [(n, e) for n, e in enumerate(plan, 1) if signature(e) not in sigs][:it.get("max_slides", 3)]
            if not changed:
                raise SystemExit(f"decision {code} option {o['choice']}: changes no slide")
            new_sigs = {signature(e) for e in plan}
            gone = [n for n, s in enumerate(base_sigs, 1) if s not in new_sigs][:it.get("max_slides", 3)]
            if not it.get("applies"):
                dd["applies"] = sorted(set(dd["applies"]) | set(gone))
            dd["options"][str(o["choice"])] = []
            for n, e in changed:
                entries.append(e)
                where.append((code, str(o["choice"]), n))
    if entries:
        pngs, lines = build_and_export(entries, "options", approved, page_dir, pairs)
        for k, (code, choice, n) in enumerate(where, 1):
            png = f"opt_{code}_{choice}_{n:02d}.png"
            shutil.copyfile(pngs[k - 1], page_dir / png)
            decisions[code]["options"][choice].append({"n": n, "png": png, "lines": lines[k], "fact": None})
    t = step(f"options: {len(entries)} slides built and exported", t)
    out = {"deck": deck, "decisions": decisions}
    (page_dir / "page_data.json").write_text(json.dumps(out, indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"{len(deck)} deck slides, {len(entries)} option slides for {len(decisions)} open decisions -> "
          f"{page_dir / 'page_data.json'}")
    step("total", t0)


if __name__ == "__main__":
    main(sys.argv[1])
