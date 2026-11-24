"""Incremental deck builds: rebuild only the slides whose spec changed, keep every other slide byte for byte.

A full build of a movie deck (build_movie_deck_via_com) embeds every movie again: 11-18 min for the 119-slide TGT
movie-pairs deck (2026-11-23), even when one text slide changed. An incremental build instead:

1. hashes every slide's spec (``spec_hash``): its text boxes, positions and formatting, each movie's file (path,
   size, modification time), each picture's bytes, the slide size, and the COM drawing code itself;
2. reads the hashes of the deck's current slides from the sidecar ``<deck>.slides.json`` that the last build wrote
   (it holds the deck's md5, so a deck changed by hand or restored from a backup is refused, not patched);
3. builds only the slides with a new hash, through the same COM code as a full build, into a small temporary deck;
4. splices them into a copy of the existing deck (python-pptx) in the new order; slides no longer wanted are dropped;
5. checks the result (``check_splice``): every kept slide's content hash (slide XML, with each relationship replaced by
   the md5 of the part it points to) equals the old deck's slide, and every new slide's equals the temporary deck's.

So text edits, new values beside movies, notes, marks, added, removed or reordered slides each cost the slides that
changed, not the whole deck. A later full build gives the same slides (same spec, same COM code).

What it cannot do: patch a deck with no sidecar or one changed since (seed it with ``seed_sidecar`` after checking the
deck against its specs, or build it in full once); change the slide size (every hash changes, so everything is
rebuilt, which is a full build); notes pages are not copied (the builders write none).
"""
import copy
import hashlib
import inspect
import json
import os
import re
import tempfile
from collections import defaultdict, deque
from typing import Callable, Dict, List, Optional, Sequence, Tuple

from . import movies as _movies
from .movies import SlideSpec, _slide_to_manifest, build_movie_deck_via_com

SIDECAR_SUFFIX = ".slides.json"
SIDECAR_FORMAT = 1
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_RT_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
_RT_NOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"


class IncrementalUnavailable(RuntimeError):
    """The deck cannot be patched (no sidecar, or the deck changed since the sidecar was written)."""


def _md5_file(path: str, chunk: int = 1 << 22) -> str:
    h = hashlib.md5()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(chunk), b""):
            h.update(block)
    return h.hexdigest()


def renderer_version() -> str:
    """Hash of the code that draws a slide (the COM script and the on-disk autoplay/loop patches): a change to it
    changes every slide's hash, so nothing drawn by older code is kept."""
    src = _movies._POWERSHELL_TEMPLATE + inspect.getsource(_movies._patch_slide_xml) + inspect.getsource(
        _movies._patch_slide_loop_xml)
    return hashlib.sha1(src.encode("utf-8")).hexdigest()[:16]


def spec_hash(spec: SlideSpec, slide_w: float, slide_h: float, loop: bool, extra: str = "",
              restore_pictures: bool = False) -> str:
    """Hash of everything that decides how a slide is drawn. Movies count by path, size and modification time (a
    re-rendered movie gets a new hash); the poster is the caller's rule applied to the movie, named in ``extra``
    (e.g. "poster: first frame, JPEG q90"), never the temporary poster file. Pictures count by their bytes."""
    m = _slide_to_manifest(spec)
    for mv in m["movies"]:
        st = os.stat(mv["movie_path"])
        mv["movie_path"] = [mv["movie_path"], st.st_size, st.st_mtime_ns]
        mv["poster_path"] = None
    for pic in m["pictures"]:
        pic["image_path"] = [pic["image_path"], _md5_file(pic["image_path"])]
    payload = {"slide": m, "w": slide_w, "h": slide_h, "loop": bool(loop), "extra": extra,
               "restore_pictures": bool(restore_pictures) and bool(spec.pictures),
               "renderer": renderer_version()}
    return hashlib.sha1(json.dumps(payload, sort_keys=True).encode("utf-8")).hexdigest()


def sidecar_path(deck_path: str) -> str:
    return str(deck_path) + SIDECAR_SUFFIX


def write_sidecar(deck_path: str, hashes: Sequence[str]) -> str:
    """Record the slide hashes of a deck just built, with the deck's md5 and size."""
    data = {"format": SIDECAR_FORMAT, "deck": os.path.basename(str(deck_path)), "deck_md5": _md5_file(str(deck_path)),
            "deck_size": os.path.getsize(str(deck_path)), "slides": list(hashes)}
    path = sidecar_path(deck_path)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1)
    return path


def read_sidecar(deck_path: str) -> List[str]:
    """The deck's slide hashes; raises IncrementalUnavailable when there is no sidecar or the deck changed since."""
    path = sidecar_path(deck_path)
    if not os.path.isfile(path):
        raise IncrementalUnavailable(f"no {os.path.basename(path)} beside the deck (built before incremental builds, "
                                     "or by another tool): seed it or build the deck in full once")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if data.get("format") != SIDECAR_FORMAT:
        raise IncrementalUnavailable(f"{path}: format {data.get('format')}, expected {SIDECAR_FORMAT}")
    if os.path.getsize(str(deck_path)) != data["deck_size"] or _md5_file(str(deck_path)) != data["deck_md5"]:
        raise IncrementalUnavailable(f"{deck_path} changed since its sidecar was written (edited by hand, or "
                                     "restored from a backup): build it in full")
    return data["slides"]


def slide_content_hashes(prs) -> List[str]:
    """One hash per slide: the slide XML in canonical form with every relationship id replaced by the type and md5 of
    the part it points to, plus those parts. Equal hashes = same shapes, text, timing, movies and pictures."""
    from lxml import etree
    blob_md5: Dict[int, str] = {}

    def part_md5(part) -> str:
        if id(part) not in blob_md5:
            blob_md5[id(part)] = hashlib.md5(part.blob).hexdigest()
        return blob_md5[id(part)]

    out = []
    for slide in prs.slides:
        rels = slide.part.rels
        key = {}
        for rId, rel in rels.items():
            if rel.reltype == _RT_NOTES:
                continue
            if rel.reltype == _RT_LAYOUT:   # by name: each deck has its own copy of the layout (_copy_slide maps names)
                key[rId] = "slideLayout:" + slide.slide_layout.name
                continue
            key[rId] = rel.reltype.rsplit("/", 1)[-1] + ":" + (rel.target_ref if rel.is_external
                                                                 else part_md5(rel.target_part))
        el = copy.deepcopy(slide.part._element)
        for node in el.iter():
            for name, value in list(node.attrib.items()):
                if name.startswith("{%s}" % _R_NS) and value in key:
                    node.set(name, key[value])
        h = hashlib.sha1(etree.tostring(el, method="c14n"))
        h.update("|".join(sorted(key.values())).encode("utf-8"))
        out.append(h.hexdigest())
    return out


def _copy_slide(src_slide, dst_prs):
    """Append a copy of src_slide (another presentation) to dst_prs: same XML, its movies and pictures added as new
    parts of dst_prs (renamed so no part name collides), relationship ids remapped. Returns the new sldId element."""
    pkg = dst_prs.part.package
    layout_name = src_slide.slide_layout.name
    layouts = [lo for lo in dst_prs.slide_layouts if lo.name == layout_name]
    if not layouts:
        raise RuntimeError(f"no slide layout named {layout_name!r} in the target deck")
    dst = dst_prs.slides.add_slide(layouts[0])
    dst_layout_rid = next(r for r, rel in dst.part.rels.items() if rel.reltype == _RT_LAYOUT)
    rid_map, renamed = {}, {}
    for rId, rel in src_slide.part.rels.items():
        if rel.reltype == _RT_NOTES:
            continue
        if rel.reltype == _RT_LAYOUT:
            rid_map[rId] = dst_layout_rid
        elif rel.is_external:
            rid_map[rId] = dst.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            part = rel.target_part
            if id(part) not in renamed:   # rename, then relate at once, so the next name sees this one as taken
                tmpl = re.sub(r"\d*(\.[A-Za-z0-9]+)$", r"%d\1", str(part.partname))
                part.partname = pkg.next_partname(tmpl)
                renamed[id(part)] = part
            rid_map[rId] = dst.part.relate_to(part, rel.reltype)
    el = copy.deepcopy(src_slide.part._element)
    for node in el.iter():
        for name, value in list(node.attrib.items()):
            if name.startswith("{%s}" % _R_NS) and value in rid_map:
                node.set(name, rid_map[value])
    dst.part._element = el
    return dst_prs.slides._sldIdLst[-1]


_STORED_EXTS = {"mp4", "m4v", "mov", "wmv", "avi", "mp3", "m4a", "jpeg", "jpg", "png", "gif"}


def save_fast(prs, out_path: str) -> None:
    """Save a python-pptx Presentation like prs.save, but store already-compressed media (movies, JPEG, PNG) without
    deflating them again: same parts, same bytes inside, but much faster for a movie deck (2026-11-23, 481 MB deck:
    31 s for splice and save with prs.save). XML parts are still deflated."""
    import zipfile
    from pptx.opc.oxml import serialize_part_xml
    from pptx.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI
    from pptx.opc.serialized import _ContentTypesItem
    pkg = prs.part.package
    parts = tuple(pkg.iter_parts())
    with zipfile.ZipFile(str(out_path), "w", zipfile.ZIP_DEFLATED, strict_timestamps=False) as z:
        z.writestr(CONTENT_TYPES_URI.membername, serialize_part_xml(_ContentTypesItem.xml_for(parts)))
        z.writestr(PACKAGE_URI.rels_uri.membername, pkg._rels.xml)
        for part in parts:
            stored = part.partname.ext.lower() in _STORED_EXTS
            z.writestr(part.partname.membername, part.blob,
                       compress_type=zipfile.ZIP_STORED if stored else zipfile.ZIP_DEFLATED)
            if part._rels:
                z.writestr(part.partname.rels_uri.membername, part.rels.xml)


def splice(old_path: str, new_slides_path: Optional[str], plan: Sequence[Tuple[str, int]], out_path: str) -> None:
    """Write out_path: the slides of old_path and new_slides_path in plan order. plan: per output slide, ("old", i)
    or ("new", j), 0-based. Old slides not in the plan are dropped; old_path itself is never written."""
    from pptx import Presentation
    prs = Presentation(str(old_path))
    src = Presentation(str(new_slides_path)) if new_slides_path else None
    lst = prs.slides._sldIdLst
    old_ids = list(lst)
    order = []
    for kind, i in plan:
        order.append(old_ids[i] if kind == "old" else _copy_slide(src.slides[i], prs))
    keep = {id(e) for e in order}
    for e in old_ids:
        if id(e) not in keep:
            prs.part.drop_rel(e.rId)
    for e in list(lst):
        lst.remove(e)
    for e in order:
        lst.append(e)
    save_fast(prs, str(out_path))


def check_splice(old_path: str, new_slides_path: Optional[str], out_path: str,
                 plan: Sequence[Tuple[str, int]]) -> List[str]:
    """Problems (empty = PASS): out_path has len(plan) slides, each kept slide's content hash equals its old slide's,
    each new slide's equals its slide in the temporary deck."""
    from pptx import Presentation
    ho = slide_content_hashes(Presentation(str(old_path)))
    ht = slide_content_hashes(Presentation(str(new_slides_path))) if new_slides_path else []
    hn = slide_content_hashes(Presentation(str(out_path)))
    if len(hn) != len(plan):
        return [f"{len(hn)} slides written, {len(plan)} planned"]
    problems = []
    for k, (kind, i) in enumerate(plan):
        want = ho[i] if kind == "old" else ht[i]
        if hn[k] != want:
            problems.append(f"slide {k + 1}: differs from {'old slide' if kind == 'old' else 'new slide'} {i + 1}")
    return problems


def plan_update(old_hashes: Sequence[str], new_hashes: Sequence[str]) -> Tuple[List[Tuple[str, int]], List[int]]:
    """Reuse an old slide for every new slide with the same hash (each old slide once, earliest first); the rest are
    built. Returns (plan, indices of the new slides to build)."""
    pool: Dict[str, deque] = defaultdict(deque)
    for i, h in enumerate(old_hashes):
        pool[h].append(i)
    plan, build = [], []
    for k, h in enumerate(new_hashes):
        if pool[h]:
            plan.append(("old", pool[h].popleft()))
        else:
            plan.append(("new", len(build)))
            build.append(k)
    return plan, build


def update_deck(specs: Sequence[SlideSpec], old_deck: str, out_path: str, slide_w: float, slide_h: float,
                loop: bool = False, extra: str = "", prepare: Optional[Callable[[List[SlideSpec]], None]] = None,
                restore_pictures: bool = False, timer=None) -> Dict:
    """Write out_path = old_deck with only the changed slides rebuilt (see the module docstring), then its sidecar.
    prepare(specs to build): called before the COM build (e.g. write those slides' posters). timer: a StageTimer.
    Returns {"built", "kept", "moved", "dropped" (1-based slide numbers), "autoplay" (built slides whose autoplay
    trigger was set), "problems" (check_splice)}.
    Raises IncrementalUnavailable when old_deck cannot be patched."""
    from contextlib import nullcontext
    stage = timer.stage if timer else (lambda name: nullcontext())
    with stage("read sidecar, check deck md5"):
        old_hashes = read_sidecar(old_deck)
    with stage("hash slide specs"):
        new_hashes = [spec_hash(s, slide_w, slide_h, loop, extra, restore_pictures) for s in specs]
    plan, build = plan_update(old_hashes, new_hashes)
    tmp_dir = tempfile.mkdtemp(prefix="incremental_build_")
    tmp_deck = os.path.join(tmp_dir, "changed_slides.pptx") if build else None
    autoplay = 0
    try:
        if build:
            if prepare:
                with stage(f"posters ({sum(len(specs[k].movies) for k in build)} movies)"):
                    prepare([specs[k] for k in build])
            autoplay = build_movie_deck_via_com([specs[k] for k in build], tmp_deck, slide_w, slide_h, loop=loop,
                                                restore_pictures=restore_pictures, timer=timer)
        with stage("splice into the existing deck and save"):
            splice(old_deck, tmp_deck, plan, out_path)
        with stage("check splice (content hash of every slide)"):
            problems = check_splice(old_deck, tmp_deck, out_path, plan)
        with stage("write sidecar"):
            write_sidecar(out_path, new_hashes)
    finally:
        if tmp_deck and os.path.exists(tmp_deck):
            os.unlink(tmp_deck)
        os.rmdir(tmp_dir)
    used = {i for kind, i in plan if kind == "old"}
    return {"built": [k + 1 for k in build],
            "kept": [k + 1 for k, (kind, i) in enumerate(plan) if kind == "old" and i == k],
            "moved": [(i + 1, k + 1) for k, (kind, i) in enumerate(plan) if kind == "old" and i != k],
            "dropped": [i + 1 for i in range(len(old_hashes)) if i not in used], "autoplay": autoplay,
            "problems": problems}


def check_deck_matches_specs(prs, specs: Sequence[SlideSpec]) -> List[str]:
    """Problems (empty = PASS) comparing a built deck with its specs, for seeding a sidecar on a deck built before
    incremental builds: slide count; per slide, the text boxes in order (text, left, top, width within 0.5 pt, first
    run's size, bold, italic and color), the movies and pictures in order (left, top, width, height within 0.5 pt;
    embedded bytes equal to the file). Not compared: indents, spacing, the poster images."""
    import hashlib as _h
    pt = 12700.0
    problems = []
    if len(prs.slides) != len(specs):
        return [f"deck has {len(prs.slides)} slides, specs {len(specs)}"]
    file_md5: Dict[str, str] = {}

    def fmd5(p):
        if p not in file_md5:
            file_md5[p] = _md5_file(p)
        return file_md5[p]

    def near(sh, *want):
        got = (sh.left / pt, sh.top / pt, sh.width / pt, sh.height / pt)[:len(want)]
        return all(abs(g - w) <= 0.5 for g, w in zip(got, want))

    for k, (slide, spec) in enumerate(zip(prs.slides, specs), 1):
        movies, pics, texts = [], [], []
        for sh in slide.shapes:
            link = sh._element.xpath(".//a:videoFile/@r:link")
            if link:
                movies.append((sh, slide.part.related_part(link[0])))
            elif sh.shape_type == 13:   # MSO_SHAPE_TYPE.PICTURE
                pics.append(sh)
            elif sh.has_text_frame:
                texts.append(sh)
        if len(texts) != len(spec.textboxes) or len(movies) != len(spec.movies) or len(pics) != len(spec.pictures):
            problems.append(f"slide {k}: {len(texts)} text boxes, {len(movies)} movies, {len(pics)} pictures; spec "
                            f"{len(spec.textboxes)}, {len(spec.movies)}, {len(spec.pictures)}")
            continue
        for sh, tb in zip(texts, spec.textboxes):
            run = next((r for p in sh.text_frame.paragraphs for r in p.runs), None)
            color = None
            if run is not None and run.font.color and run.font.color.type == 1:   # MSO_COLOR_TYPE.RGB
                rgb = str(run.font.color.rgb)   # RRGGBB; TextboxSpec.color_rgb is a BGR integer (COM)
                color = int(rgb[4:6] + rgb[2:4] + rgb[0:2], 16)
            want_color = tb.color_rgb if tb.color_rgb is not None else color
            # PowerPoint writes no size when it equals its default text box size, 18 pt
            size = run.font.size.pt if run is not None and run.font.size is not None else 18.0
            if (sh.text_frame.text != tb.text.replace("\r", "\n") or not near(sh, tb.left_pt, tb.top_pt, tb.width_pt)
                    or run is None or abs(size - tb.font_size_pt) > 0.01
                    or bool(run.font.bold) != bool(tb.main_bold if tb.main_bold is not None else tb.bold)
                    or bool(run.font.italic) != bool(tb.italic) or color != want_color):
                problems.append(f"slide {k}: text box {tb.text[:40]!r} differs from its spec")
        for (sh, part), mv in zip(movies, spec.movies):
            if not near(sh, mv.left_pt, mv.top_pt, mv.width_pt, mv.height_pt) or \
                    _h.md5(part.blob).hexdigest() != fmd5(os.path.abspath(mv.movie_path)):
                problems.append(f"slide {k}: movie {os.path.basename(mv.movie_path)} differs from its spec")
        for sh, pic in zip(pics, spec.pictures):
            if not near(sh, pic.left_pt, pic.top_pt, pic.width_pt, pic.height_pt) or \
                    _h.md5(sh.image.blob).hexdigest() != fmd5(os.path.abspath(pic.image_path)):
                problems.append(f"slide {k}: picture {os.path.basename(pic.image_path)} differs from its spec")
    return problems


def seed_sidecar(deck_path: str, specs: Sequence[SlideSpec], slide_w: float, slide_h: float, loop: bool = False,
                 extra: str = "", restore_pictures: bool = False) -> List[str]:
    """Write the sidecar for a deck built before incremental builds, only if check_deck_matches_specs passes.
    Returns the problems (empty = sidecar written)."""
    from pptx import Presentation
    problems = check_deck_matches_specs(Presentation(str(deck_path)), specs)
    if not problems:
        write_sidecar(deck_path, [spec_hash(s, slide_w, slide_h, loop, extra, restore_pictures) for s in specs])
    return problems


def swap_in(new_path: str, live_path: str) -> str:
    """Replace live_path by new_path (and its sidecar): stops when the live deck is open in PowerPoint (its ~$ lock
    file exists; PowerPoint is never closed), backs the live deck up to backups/ (latest, previous, ...) first.
    Returns the new md5."""
    from .backup import backup_presentation
    live_dir, live_name = os.path.split(os.path.abspath(str(live_path)))
    if os.path.exists(os.path.join(live_dir, "~$" + live_name)):
        raise RuntimeError(f"{live_name} is open in PowerPoint: the new deck stays at {new_path}; close the deck, then "
                           "swap again")
    if os.path.exists(str(live_path)):
        backup_presentation(str(live_path))
        old_side = sidecar_path(live_path)
        if os.path.exists(old_side):
            os.remove(old_side)
    os.replace(str(new_path), str(live_path))
    if os.path.exists(sidecar_path(new_path)):
        os.replace(sidecar_path(new_path), sidecar_path(live_path))
        with open(sidecar_path(live_path), encoding="utf-8") as f:
            data = json.load(f)
        data["deck"] = live_name
        with open(sidecar_path(live_path), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=1)
    return _md5_file(str(live_path))


def _main(argv: Sequence[str]) -> int:
    """python -m ppt_image_inserter.incremental swap <new deck> <live deck>
    python -m ppt_image_inserter.incremental check <old deck> <new deck>   (slides whose content hash differs)"""
    if len(argv) == 3 and argv[0] == "swap":
        print(f"swapped in {argv[2]}: md5 {swap_in(argv[1], argv[2]).upper()} (previous deck in backups/latest)")
        return 0
    if len(argv) == 3 and argv[0] == "check":
        from pptx import Presentation
        a, b = (slide_content_hashes(Presentation(p)) for p in argv[1:])
        pos = {h: i for i, h in reversed(list(enumerate(a)))}
        new = [k + 1 for k, h in enumerate(b) if h not in pos]
        print(f"{len(a)} -> {len(b)} slides; new or changed: {new}; old slides not kept: "
              f"{[i + 1 for i, h in enumerate(a) if h not in set(b)]}")
        return 0
    print(_main.__doc__)
    return 2


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
