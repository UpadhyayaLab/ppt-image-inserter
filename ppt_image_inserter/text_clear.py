"""Keep slide text off the data (FF 2026-11-20: "it cannot overlap with any data. this needs a hard coded tst (titles
too)"; "if it overlaps something needs to moved").

Measured as PowerPoint renders the slide (COM, Windows), never estimated: each slide is exported in memory three times,
with only the pictures, only the title box, and only the bottom line (shape "Bottom text", workflow._add_bottom_text)
visible. A text overlaps the data when any non-white figure pixel lies inside the box around that text's letters,
widened by GAP_PT on every side, so text that nearly touches the data counts too (white margins of the picture do not
count).

- check(deck): every overlap, deck opened read-only and never saved.
- fix(deck): for each slide that overlaps, moves the pictures, and scales them (aspect kept) only when they do not fit,
  into the free band between the title's lowest letter and the bottom line's highest letter; the text is never moved or
  shrunk. Saves the deck and returns what moved; a scale below BIG_SHRINK is FF's call and is flagged.
The deck must be closed in PowerPoint.
"""
import os
import tempfile

PICTURE = 13      # msoPicture
WIDTH_PX = 1000   # export width (about 1 px per point on a 13.33 in slide)
DARK = 140        # text: gray level below which a pixel is a letter
FIG_INK = 245     # figure: any pixel darker than this is data (light colors such as cyan included)
GAP_PT = 4        # free space kept between text and data, points
BIG_SHRINK = 0.85


def _box(png, thr):
    from PIL import Image
    return Image.open(png).convert("L").point(lambda v: 255 if v < thr else 0).getbbox()


def _hits(fig_png, box, pad):
    from PIL import Image
    if box is None:
        return False
    im = Image.open(fig_png).convert("L")
    box = (max(box[0] - pad, 0), max(box[1] - pad, 0), min(box[2] + pad, im.width), min(box[3] + pad, im.height))
    return im.crop(box).point(lambda v: 255 if v < FIG_INK else 0).getbbox() is not None


def _open(deck, read_only):
    import win32com.client as w
    deck = os.path.abspath(deck)
    app = w.Dispatch("PowerPoint.Application")
    for k in range(1, app.Presentations.Count + 1):
        try:
            same = os.path.normcase(app.Presentations(k).FullName) == os.path.normcase(deck)
        except Exception:
            continue
        if same:
            raise SystemExit(f"{deck} is open in PowerPoint: close it first")
    return app.Presentations.Open(deck, read_only, False, False)   # ReadOnly, Untitled, WithWindow


def _slide_layers(slide, n, tmp, wpx, hpx):
    """(pictures, {name: png}, {name: ink box in px}) for 'figure', 'title', 'bottom'."""
    shapes = [slide.Shapes(i) for i in range(1, slide.Shapes.Count + 1)]
    pics = [s for s in shapes if s.Type == PICTURE]
    texts = [s for s in shapes if s.Type != PICTURE and s.HasTextFrame and s.TextFrame.HasText
             and not s.TextFrame.TextRange.Text.strip().isdigit()]
    bottom = [s for s in texts if s.Name == "Bottom text"]
    rest = [s for s in texts if s.Name != "Bottom text"]
    layers = {"figure": pics}
    if rest:
        layers["title"] = [min(rest, key=lambda s: s.Top)]
    if bottom:
        layers["bottom"] = bottom
    pngs, boxes = {}, {}
    if not pics or len(layers) == 1:
        return pics, pngs, boxes
    # Skip the exports when no text's rendered extent (plus GAP_PT) reaches any picture's box: then no letter can be
    # near a data pixel (2026-11-20: the three exports per slide made a 79-slide build 60-75 s slower)
    def reach(t):
        r = t.TextFrame.TextRange
        x0, y0 = r.BoundLeft - GAP_PT, r.BoundTop - GAP_PT
        x1, y1 = r.BoundLeft + r.BoundWidth + GAP_PT, r.BoundTop + r.BoundHeight + GAP_PT
        return any(x0 < p.Left + p.Width and p.Left < x1 and y0 < p.Top + p.Height and p.Top < y1 for p in pics)
    if not any(reach(t) for name, ts in layers.items() if name != "figure" for t in ts):
        return pics, pngs, boxes
    vis = {id(s): s.Visible for s in shapes}
    try:
        for name, keep in layers.items():
            for s in shapes:
                s.Visible = -1 if any(s is k for k in keep) else 0
            pngs[name] = os.path.join(tmp, f"s{n}_{name}.png")
            slide.Export(pngs[name], "PNG", wpx, hpx)
            boxes[name] = _box(pngs[name], FIG_INK if name == "figure" else DARK)
    finally:
        for s in shapes:
            s.Visible = vis[id(s)]
    return pics, pngs, boxes


def _overlaps(pngs, boxes, pad):
    return [name for name in ("title", "bottom") if name in pngs and _hits(pngs["figure"], boxes[name], pad)]


def check(deck, slides=None):
    """[{"slide": n, "text": "title" | "bottom"}] for every text whose letters' box holds figure pixels; slides limits
    the check to those slide numbers (e.g. the ones fix() just moved)."""
    pres = _open(deck, True)
    tmp = tempfile.mkdtemp(prefix="text_clear_")
    out = []
    try:
        sw, sh = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        hpx = round(WIDTH_PX * sh / sw)
        for n in (slides or range(1, pres.Slides.Count + 1)):
            _, pngs, boxes = _slide_layers(pres.Slides(n), n, tmp, WIDTH_PX, hpx)
            out += [{"slide": n, "text": t} for t in _overlaps(pngs, boxes, round(GAP_PT * WIDTH_PX / sw))]
            fig = boxes.get("figure")
            if fig and (fig[1] <= 0 or fig[3] >= hpx - 1):   # data cut by the top or bottom slide edge
                out.append({"slide": n, "text": "edge"})
    finally:
        pres.Close()
    return out


def fix(deck):
    """Move (and if needed scale) the pictures of every overlapping slide into the free band; saves the deck.
    Returns [{"slide", "scale", "moved_in", "flag"}]; flag is set when the scale is below BIG_SHRINK or the data
    cannot clear the text (then the slide is left as it was)."""
    pres = _open(deck, False)
    tmp = tempfile.mkdtemp(prefix="text_clear_")
    out = []
    try:
        sw, sh = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        hpx = round(WIDTH_PX * sh / sw)
        pt = sw / WIDTH_PX                                  # points per exported pixel
        for n in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(n)
            pics, pngs, boxes = _slide_layers(slide, n, tmp, WIDTH_PX, hpx)
            if not _overlaps(pngs, boxes, round(GAP_PT / pt)) or not boxes.get("figure"):
                continue
            # + 2 px: clear of the check's own rounding. Without a bottom line the band ends GAP_PT above the slide
            # edge (2026-11-20 reviewer: slides 3-4 were moved until the data ran off the bottom of the slide)
            top = (boxes["title"][3] * pt + GAP_PT + 2 * pt) if boxes.get("title") else GAP_PT
            bot = (boxes["bottom"][1] * pt - GAP_PT - 2 * pt) if boxes.get("bottom") else sh - GAP_PT - 2 * pt
            fx0, fy0, fx1, fy1 = (v * pt for v in boxes["figure"])
            # data cut by the slide edge: the export cannot show its true extent, so use the pictures' own edge
            if boxes["figure"][3] >= hpx - 1:
                fy1 = max(fy1, max(p.Top + p.Height for p in pics))
            if boxes["figure"][1] <= 0:
                fy0 = min(fy0, min(p.Top for p in pics))
            if bot - top <= 0:
                out.append({"slide": n, "scale": None, "moved_in": None, "flag": "no free band between the texts"})
                continue
            s = min(1.0, (bot - top) / (fy1 - fy0))
            cx = (fx0 + fx1) / 2
            new_top = min(max(fy0, top), bot - (fy1 - fy0) * s)   # move as little as possible
            for p in pics:   # scale about the data's top-center, then shift the data's top to new_top
                p.LockAspectRatio = -1
                left, ptop, w, h = p.Left, p.Top, p.Width, p.Height
                p.Width = w * s                                     # aspect locked: Height follows
                p.Left = cx + (left - cx) * s
                p.Top = new_top + (ptop - fy0) * s
            out.append({"slide": n, "scale": round(s, 3), "moved_in": round((new_top - fy0) / 72, 2),
                        "flag": "shrinks a lot" if s < BIG_SHRINK else ""})
        pres.Save()
    finally:
        pres.Close()
    return out


def report(problems):
    """One line per failing slide, for build logs."""
    return [f"slide {p['slide']}: the data runs off the slide" if p["text"] == "edge" else
            f"slide {p['slide']}: the {'title' if p['text'] == 'title' else 'bottom line'} overlaps the data"
            for p in problems]
