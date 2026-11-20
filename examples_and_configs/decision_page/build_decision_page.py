"""Build a decision page: every option shown as its full rendered slide, one card per decision, answered with large
buttons and saved in the page's database (FF 2026-11-19: "I need to see everything in context. you need to show full
slide for each item! establish this workflow"; "I need to be able to clearly indicate answers. on the page").

Look and behavior follow the BIAD2255 approval page FF answered on: one card per decision; the card turns green with
"Your pick: <option>" at the top right when an answer button is clicked, or amber with "Change saved" when text is
written in "Change: write what it should say" and "Save change" is clicked; a counter at the top ("N of M picked,
K with changes") and a sticky summary of every decision's state that jumps to it.

Input: a YAML spec (format in README.md beside this file). Output: one self-contained HTML page (slide PNGs embedded
once each, as base64) for the Artifact tool, published with capabilities {db: {}}.

Answers: collection <spec collection, default "decisions">, one document per decision code, written with set() on
each click, fields
  decision   the code ("1", "3a", ...)
  status     "pick" (an answer button was clicked last) or "change" (Save change was clicked last)
  choice     the option picked ("1A", "keep", "use", ...; "" if none)
  text       the exact string of the picked option ("" when the option carries none)
  note       the text written in the change box
  updatedAt  ISO time of the last click
A document without status but with a choice counts as "pick". Read with ArtifactData action "list".
"Edit this text" under a quoted option copies its exact text into the card's change box (FF 2026-11-19: "I need to
be able to paste text from here so I can change in response box").

Every slide on the page has "Change this slide" (FF 2026-11-19: "I want somehting for every single slide too! so I
can change wahatever easily"): a box holding the slide's text (spec slide_text: {id: {lines, unlisted}}) to edit and
save. Collection <spec slide_collection, default "slides">, one document per slide id: slide, status ("change" or
"cleared"), before (the text the page showed), text (what FF saved), updatedAt. All copies of a slide share it.

With spec all_slides, the page opens on an "All slides" tab: every deck slide in deck order, one per row, numbered,
with its text already open in an edit box saving to the same slides collection (Histone mark data agent presentation,
2026-11-19); "Decisions" is the other tab.

Item sources (file:line) stay in the spec for the record and are not shown: FF reads no file names.

Test without claude.ai: --test-dir <dir> writes test pages with the db replaced by an in-memory stub and screenshots
them with headless Edge at desktop (1400 px) and phone (400 px) width: unanswered, answered (picks, one change, one
slide change), and the answered page scrolled to the changed card and to the changed slide, plus one tall shot per
item. It also runs one answer cycle (edit, pick, change, re-pick, slide change, reload, remove slide change) and
fails unless every save survives the reload.

Usage:  python build_decision_page.py <spec.yaml> <out.html> [--test-dir <dir>]
"""
import base64
import html
import io
import json
import re
import subprocess
import sys
from pathlib import Path

import yaml
from PIL import Image

E = html.escape
EDGE = "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"
_IMG_CACHE = {}   # (path, width, mtime) -> data URI


class Page:
    def __init__(self, spec, spec_dir):
        self.spec = spec
        self.dir = spec_dir
        self.png_dir = Path(spec.get("png_dir", spec_dir))
        self.width = int(spec.get("png_width", 1280))
        self.images = {}      # image key -> (path, data URI); each PNG embedded once
        self.decisions = []   # decision codes in page order
        self.item_of = {}     # decision code -> item number
        self.item = None
        st = spec.get("slide_text")   # {slide id: {"lines": [...], "unlisted": "..."}}; every slide must be in it
        self.slide_text = json.loads(self.resolve(st).read_text(encoding="utf-8")) if st else None
        sf = spec.get("slide_facts")  # {slide id: {"text": "...", "warn": bool}}: shown under every slide
        self.slide_facts = json.loads(self.resolve(sf).read_text(encoding="utf-8")) if sf else None
        self.slide_info = {}          # slide id -> {"text", "unlisted"}, embedded once as JSON
        self.slide_seen = set()
        # under every change box (FF 2026-11-19: "what I say there is general! like: it applies to many slides")
        self.scope = E(spec.get("change_scope_note", "Applies to every slide like this one."))
        self.all = spec.get("all_slides")   # optional All slides view (the default view when given)
        # {slide id: one line}: shown with that slide's saved change, on every copy, only while a change is saved
        self.saved_notes = spec.get("saved_change_notes") or {}
        # Latest saved state is the base for every slide (FF 2026-11-19: "then why is this showing old slide????"):
        # spec saved_state = {dir: the db export (ArtifactData list with out_dir, run right before the build),
        # rebuilt: {slide id: png rendered from that slide's saved change}}. Every copy of a slide id with a saved
        # change must use its rebuilt png, or the build fails.
        self.saved_slides, self.rebuilt = {}, {}
        ss = spec.get("saved_state")
        if ss:
            d = Path(ss["dir"]) / spec.get("slide_collection", "slides")
            if not d.is_dir():
                raise SystemExit(f"saved_state: no export at {d}: export the db first (ArtifactData list, out_dir)")
            for f in d.glob("*.json"):
                doc = json.loads(f.read_text(encoding="utf-8"))
                if doc.get("status") == "change":
                    self.saved_slides[f.stem] = doc
            self.rebuilt = {k: self.resolve(v) for k, v in (ss.get("rebuilt") or {}).items()}
            missing = sorted(set(self.saved_slides) - set(self.rebuilt))
            if missing:
                raise SystemExit(f"slides with a saved change but no rebuilt png in saved_state.rebuilt: {missing}")
        self.saved_decs = {}   # decision code -> saved document (db export)
        if ss:
            dd = Path(ss["dir"]) / spec.get("collection", "decisions")
            for f in (dd.glob("*.json") if dd.is_dir() else []):
                self.saved_decs[f.stem] = json.loads(f.read_text(encoding="utf-8"))
        self.deck_index = json.loads(self.resolve(self.all["index"]).read_text(encoding="utf-8")) if self.all else []
        self.deck_num = {str(x["id"]): x["slide"] for x in self.deck_index}   # deck slide id -> slide number
        self.placed = []     # every slide shown on the decisions view: {id, base, expect, outside, where}
        self.res_info = {}   # decision code -> {labels, settle, extra}: the page's Resolved lines
        self.errors = []     # failed checks, all reported at the end of build()
        self.ctx = None      # (decision code, choice, option text) while an option's slides are drawn
        self.n_deck = 0

    def resolve(self, p):
        p = Path(p)
        return p if p.is_absolute() else self.png_dir / p

    def image(self, path, width=None):
        path = self.resolve(path)
        width = int(width or self.width)
        for k, (p, _) in self.images.items():
            if p == path:
                return k
        if not path.exists():
            raise SystemExit(f"slide PNG not found: {path}")
        ck = (str(path), width, path.stat().st_mtime)
        if ck not in _IMG_CACHE:   # test pages rebuild the page many times: encode each PNG once per run
            im = Image.open(path).convert("RGB")
            if im.width > width:
                im = im.resize((width, round(im.height * width / im.width)), Image.LANCZOS)
            buf = io.BytesIO()
            im.quantize(256).save(buf, "PNG", optimize=True)
            _IMG_CACHE[ck] = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
        key = f"i{len(self.images)}"
        self.images[key] = (path, _IMG_CACHE[ck])
        return key

    def check_saved(self, sid, png):
        """A slide id with a saved change is shown only as rendered from that change."""
        if sid in self.saved_slides and self.resolve(png) != self.rebuilt[sid]:
            raise SystemExit(f"slide {sid} has a saved change: show it as {self.rebuilt[sid].name}, not {png}")

    def slide_parts(self, sid, caption=""):
        """Shared by both views: registers the slide's text and returns (caption + fact html, box label, unlisted
        note html, whether its text is known)."""
        if self.slide_text is not None and sid not in self.slide_text:
            raise SystemExit(f"slide {sid}: not in slide_text")
        info = (self.slide_text or {}).get(sid, {})
        self.slide_info[sid] = {"text": "\n".join(info.get("lines", [])), "unlisted": info.get("unlisted", "")}
        cap = f'<figcaption>{E(caption)}</figcaption>' if caption else ""
        if self.slide_facts is not None:   # a measured fact per slide (e.g. title lines), flagged when warn
            if sid not in self.slide_facts:
                raise SystemExit(f"slide {sid}: not in slide_facts")
            f = self.slide_facts[sid]
            if f.get("text"):
                cap += f'<p class="fact{" warn" if f.get("warn") else ""}">{E(f["text"])}</p>'
        known = bool(self.slide_info[sid]["text"])
        lab = ("Text on this slide: edit any line, then click Save slide change" if known
               else "Write what to change on this slide, then click Save slide change")
        note = (f'<p class="slnote">{E(info["unlisted"])} Write any change to them here too.</p>'
                if info.get("unlisted") else "")
        if sid in self.saved_notes:
            cap += f'<p class="savednote" data-savednote="{E(sid)}" hidden>{E(self.saved_notes[sid])}</p>'
        return cap, lab, note, known

    def slide_buttons(self, sid, known):
        return ('<div class="slbtns">'
                f'<button type="button" class="slsave" data-slide="{E(sid)}">Save slide change</button>'
                + (f'<button type="button" class="slreset" data-slide="{E(sid)}">Back to the slide\'s text</button>'
                   if known else "")
                + f'<button type="button" class="slclear" data-slide="{E(sid)}" hidden>Remove saved change</button>'
                "</div>")

    def deck_rows(self):
        """The All slides view: every deck slide in deck order, one per row, large, numbered, with its text already
        open in an edit box (spec all_slides: index = [{id, png, slide}], captions = {id: caption}, png_width)."""
        a = self.all
        idx = json.loads(self.resolve(a["index"]).read_text(encoding="utf-8"))
        caps = a.get("captions") or {}
        unknown = set(caps) - {x["id"] for x in idx}
        if unknown:
            raise SystemExit(f"all_slides captions for slides not in the index: {sorted(unknown)}")
        nums = [x["slide"] for x in idx]
        if nums != list(range(1, len(idx) + 1)):
            raise SystemExit("all_slides index: slides are not numbered 1..N in order")
        rows = []
        for x in idx:
            sid = str(x["id"])
            self.check_saved(sid, x["png"])
            key = self.image(x["png"], a.get("png_width"))
            cap, lab, note, known = self.slide_parts(sid, caps.get(sid, ""))
            rows.append(
                f'<figure class="slide deck" data-slide="{E(sid)}" id="a-{E(sid)}">'
                f'<div class="dhead"><span class="dnum">Slide {x["slide"]}</span>'
                f'<span class="slstate" data-slstate="{E(sid)}"></span></div>'
                f'<img data-img="{key}" alt="Slide {x["slide"]}" title="Click to enlarge">{cap}'
                f'<div class="sledit"><label for="t-{E(sid)}">{lab}</label>'
                f'<textarea id="t-{E(sid)}" data-sltext="{E(sid)}" rows="3"></textarea>{note}'
                f'<p class="scope">{self.scope}</p>{self.slide_buttons(sid, known)}</div></figure>')
        self.n_deck = len(idx)
        return "".join(rows)

    def lines(self, sid):
        return (self.slide_text or {}).get(sid, {}).get("lines", [])

    def slides(self, items):
        """Each slide with its own change box: "Change this slide" opens its text (from slide_text) to edit and save.
        Every copy of one slide on the page shares one saved change (keyed by the slide id). A slide that is not a
        deck slide gives its base (the deck slide or preview it changes) or outside_deck (why it is not in the deck)."""
        out = []
        for s in items or []:
            png = s.get("png") or f"{s['id']}.png"
            key = self.image(png)
            sid = str(s.get("id") or Path(png).stem)
            self.check_saved(sid, png)
            if sid in self.deck_num and self.deck_index:   # a deck slide: always its deck PNG
                deck_png = next(x["png"] for x in self.deck_index if str(x["id"]) == sid)
                if self.resolve(png) != self.resolve(deck_png):
                    self.errors.append(f"slide {sid} is deck slide {self.deck_num[sid]}: show it as {deck_png}")
            self.placed.append({"id": sid, "base": s.get("base"), "expect": s.get("expect"),
                                "outside": s.get("outside_deck"), "caption": s.get("caption", ""), "ctx": self.ctx})
            cap, lab, note, known = self.slide_parts(sid, s.get("caption", ""))
            first = sid not in self.slide_seen
            self.slide_seen.add(sid)
            out.append(
                f'<figure class="slide" data-slide="{E(sid)}"' + (f' id="s-{E(sid)}"' if first else "") + ">"
                f'<img data-img="{key}" alt="{E(s.get("caption", ""))}" title="Click to enlarge">{cap}'
                f'<div class="slbar"><button type="button" class="slopen" data-slide="{E(sid)}" aria-expanded="false">'
                f'Change this slide</button><span class="slstate" data-slstate="{E(sid)}"></span></div>'
                f'<div class="sledit" hidden><label>{lab}</label>'
                f'<textarea data-sltext="{E(sid)}" rows="3"></textarea>{note}<p class="scope">{self.scope}</p>'
                + self.slide_buttons(sid, known) + "</div></figure>")
        return "".join(out)

    def affected(self, needle):
        """Deck slides whose text (latest state) carries needle, in deck order."""
        return [x for x in self.deck_index if any(needle in ln for ln in self.lines(str(x["id"])))]

    def decision(self, d):
        code = str(d["code"])
        if code in self.decisions:
            raise SystemExit(f"decision code {code} appears twice")
        self.decisions.append(code)
        self.item_of[code] = self.item
        opts = []   # (choice, button label, quoted text or None, plain label, note, slides, not_rendered, raw option)
        cur = d.get("current")
        if cur:
            opts.append(("keep", cur.get("button", "Keep current"), cur.get("text"),   # label: an unquoted description
                         None if cur.get("text") is not None else cur["label"], cur.get("where", ""),
                         cur.get("slides"), cur.get("not_rendered"), cur))
        for o in d["options"]:
            text = o.get("text")
            quoted = text if (text is not None and o.get("quoted", True)) else None
            opts.append((str(o["choice"]), o.get("button", str(o["choice"])), quoted,
                         None if quoted is not None else o.get("label", text or ""), o.get("note", ""),
                         o.get("slides"), o.get("not_rendered"), o))
        if len({o[0] for o in opts}) != len(opts):
            raise SystemExit(f"decision {code}: two options share a choice")
        # Answers are general (FF 2026-11-19): the slides an answer touches are every deck slide carrying the string
        # it replaces; each option must be previewed on all of them, and the card lists them.
        needle = d.get("replaces") or (cur or {}).get("text")
        hit = self.affected(needle) if (needle and self.deck_index) else []
        if needle and self.deck_index and not hit:
            self.errors.append(f"decision {code}: no deck slide carries {needle!r}")
        if hit:
            applies = "Applies to slides " + slide_ranges([x["slide"] for x in hit]) + "."
        elif d.get("applies_to_note"):
            applies = "Applies to: " + d["applies_to_note"]
        else:
            applies = ""
            if self.deck_index:
                self.errors.append(f"decision {code}: give current.text or replaces (to find its slides), "
                                   "or applies_to_note")
        btns, cols = [], []
        for choice, label, quoted, plain, note, slides, nr, raw in opts:
            if note and not raw.get("note_source") and choice != "keep":
                self.errors.append(f"decision {code} option {choice}: note without note_source")
            body = (f'<span class="atext"><q>{E(quoted)}</q></span>' if quoted is not None
                    else f'<span class="atext">{E(plain)}</span>' if plain and plain != label else "")
            edit = (f'<button type="button" class="edit" data-dec="{E(code)}" data-text="{E(quoted)}">'
                    "Edit this text</button>" if quoted is not None else "")
            btns.append(
                '<div class="opt">'
                f'<button type="button" class="ans" role="radio" aria-checked="false" data-dec="{E(code)}" '
                f'data-choice="{E(choice)}" data-label="{E(label)}" data-text="{E(quoted or "")}">'
                f'<span class="ahead"><span class="check" aria-hidden="true">✓</span>'
                f'<span class="alabel">{E(label)}</span></span>{body}'
                + (f'<span class="anote">{E(note)}</span>' if note else "") + f"</button>{edit}</div>")
            if slides or nr:
                self.ctx = (code, choice, raw.get("expect") or quoted)
                cols.append(f'<div class="col" data-dec="{E(code)}" data-choice="{E(choice)}">'
                            f'<div class="colhead">{E(label)}</div>'
                            + (f'<p class="nr">{E(nr)}</p>' if nr else "") + self.slides(slides) + "</div>")
                self.ctx = None
            if hit and slides and not nr:   # previewed on every slide it touches
                shown = {str(x.get("base") or x.get("id")) for x in slides}
                missing = [x["slide"] for x in hit if str(x["id"]) not in shown]
                if missing:
                    self.errors.append(f"decision {code} option {choice}: no preview for slides {missing}")
        se = d.get("settled_by")   # {slide, text, name}: resolved while FF's saved text on that slide carries text
        self.res_info[code] = {"labels": {o[0]: (o[2] if o[2] is not None else o[3] or o[1]) for o in opts},
                               "settle": se, "extra": d.get("resolved_note", "")}
        card = (f'<div class="card" id="d-{E(code)}" data-card="{E(code)}">'
                f'<div class="chead"><span class="ccode">Decision {E(code)}</span>'
                f'<span class="cstate" data-state="{E(code)}">Not answered</span></div>'
                + (f'<p class="applies">{E(applies)}</p>' if applies else "")
                + f'<div class="answers" role="radiogroup" aria-label="Decision {E(code)}">{"".join(btns)}</div>'
                + (f'<div class="cols">{"".join(cols)}</div>' if cols else "")
                + f'<div class="chgbox"><label for="note-{E(code)}">Change: write what it should say</label>'
                f'<textarea id="note-{E(code)}" data-note="{E(code)}" rows="2"></textarea>'
                f'<p class="scope">{self.scope}</p>'
                f'<button type="button" class="savechg" data-dec="{E(code)}">Save change</button></div></div>')
        extra = f'<p class="resnote" hidden>{E(d["resolved_note"])}</p>' if d.get("resolved_note") else ""
        return (f'<div class="dwrap" data-dwrap="{E(code)}"><button type="button" class="resline" data-res="{E(code)}" '
                f'hidden><span class="rt"></span> <span class="rh"></span> <span class="ro"></span></button>{extra}'
                f"{card}</div>")

    def resolution(self, code):
        """How the saved state resolves a decision: "pick", "change", "slide" (settled by FF's saved slide text) or
        None; the latest saved answer wins. The page computes the same live (JS resolution())."""
        cands = []
        doc = self.saved_decs.get(code)
        if doc:
            st = doc.get("status") or ("pick" if doc.get("choice") else "")
            if st in ("pick", "change"):
                cands.append((doc.get("updatedAt", ""), st))
        se = self.res_info[code]["settle"]
        if se:
            sd = self.saved_slides.get(se["slide"])
            if sd and se["text"] in sd.get("text", ""):
                cands.append((sd.get("updatedAt", ""), "slide"))
        return max(cands)[1] if cands else None

    def block(self, b):
        if "p" in b:
            return f"<p>{E(b['p'])}</p>"
        if "h3" in b:
            return f"<h3>{E(b['h3'])}</h3>"
        if "note" in b:
            if not b.get("source"):
                self.errors.append(f"item {self.item}: note without source: {b['note'][:60]!r}")
            return f'<p class="note">{E(b["note"])}</p>'
        if "ul" in b:
            return "<ul>" + "".join(f"<li>{E(x)}</li>" for x in b["ul"]) + "</ul>"
        if "now" in b:
            n = b["now"]
            return (f'<div class="static"><span class="lbl">Now</span><span><q>{E(n["text"])}</q>'
                    f'<span class="anote">{E(n.get("where", ""))}</span></span></div>')
        if "slides" in b:
            return f'<div class="context">{self.slides(b["slides"])}</div>'
        if "decision" in b:
            return self.decision(b["decision"])
        raise SystemExit(f"unknown block: {list(b)}")

    def section(self, it, order):
        num = str(it.get("number", ""))
        self.item = num
        before = len(self.decisions)
        inner = "".join(self.block(b) for b in it.get("blocks", []))
        decs = self.decisions[before:]
        badge = (f'<span class="badge" data-item="{E(num)}">Not answered</span>'
                 if len(decs) > 1 else "")   # one decision: its card already says it
        return (f'<section id="i{E(num)}" data-order="{order}" data-decs="{E(" ".join(decs))}"><h2>'
                f'<span class="num">{E(num)}</span><span class="h2t">{E(it["title"])}</span>{badge}</h2>'
                f'<div class="sbody">{inner}</div></section>')

    def checks(self):
        """Today's corrections as checks (Histone mark data agent presentation relaying FF, 2026-11-19: "take the time
        to improve the infra here!"). Appends to self.errors; build() fails on any."""
        words = re.compile(r"\b(long|longer|short|shorter|shorten\w*|chars?|characters?)\b", re.I)
        for it in self.spec["items"]:
            for b in it.get("blocks", []):
                d = b.get("decision")
                for o in ([d.get("current") or {}] + d["options"]) if d else []:
                    for k in ("button", "label", "note", "not_rendered"):
                        if o.get(k) and words.search(o[k]):
                            self.errors.append(f"decision {d['code']}: length word in {k}: {o[k]!r}")
        for x in self.placed:
            if x["caption"] and words.search(x["caption"]):
                self.errors.append(f"slide {x['id']}: length word in caption {x['caption']!r}")
            sid = x["id"]
            if sid not in self.deck_num and not x["base"] and not x["outside"]:
                self.errors.append(f"slide {sid}: not a deck slide; give base (the slide it changes) or outside_deck")
            if x["base"]:   # differs from its base only by the option's one change
                base = str(x["base"])
                if base not in (self.slide_text or {}):
                    self.errors.append(f"slide {sid}: base {base} has no slide text")
                    continue
                expect = x["expect"] or (x["ctx"] or (None, None, None))[2]
                if not expect:
                    self.errors.append(f"slide {sid}: has a base but no expected change (option text or expect)")
                    continue
                added = [ln for ln in self.lines(sid) if ln not in self.lines(base)]
                if not added:
                    self.errors.append(f"slide {sid}: same text as its base {base}")
                bad = [ln for ln in added if expect not in ln]
                if bad:
                    self.errors.append(f"slide {sid}: differs from {base} beyond {expect!r}: {bad}")
        for sid in set(self.slide_info):   # measured, not estimated
            f = (self.slide_facts or {}).get(sid, {})
            if f.get("text") and not isinstance(f.get("lines"), int):
                self.errors.append(f"slide {sid}: fact without a measured line count")

    def build(self, test=None):
        s = self.spec
        body = "".join(self.section(it, k) for k, it in enumerate(s["items"]))
        n = len(self.decisions)
        chips = "".join(f'<a class="chip" href="#d-{E(c)}" data-chip="{E(c)}"><b>{E(c)}</b>'
                        f'<span class="cv">–</span></a>' for c in self.decisions)
        chips += '<a class="chip slchip" id="slchip" href="#" hidden><b>Slides</b><span class="cv">0 changed</span></a>'
        dec_view = (f'<div id="view-dec" data-viewname="dec"><h1>{E(s.get("heading", s["title"]))}</h1>'
                    f'<p class="lead">{E(s.get("lead", ""))}</p><div id="open-items">{body}</div>'
                    '<h2 class="reshead" id="reshead" hidden>Resolved</h2><div id="resolved-items"></div></div>')
        tabs = all_view = ""
        if self.all:   # All slides first and the default view; decisions one click away
            rows = self.deck_rows()
            a = self.all
            all_view = (f'<div id="view-all" data-viewname="all"><h1>{E(a.get("heading", "All slides"))}</h1>'
                        f'<p class="lead">{E(a.get("lead", ""))}</p><div class="decklist">{rows}</div></div>')
            tabs = ('<div class="tabs" role="tablist" aria-label="View">'
                    f'<button type="button" class="tab on" role="tab" aria-selected="true" data-view="all">All slides ({self.n_deck})</button>'
                    f'<button type="button" class="tab" role="tab" aria-selected="false" data-view="dec">Decisions ({n})</button></div>')
        stray = set(self.saved_notes) - set(self.slide_info)
        if stray:
            self.errors.append(f"saved_change_notes for slides not on the page: {sorted(stray)}")
        self.checks()
        if self.errors:
            raise SystemExit("build checks failed:\n  " + "\n  ".join(self.errors))
        imgs = json.dumps({k: v for k, (_, v) in self.images.items()})
        slides = json.dumps(self.slide_info, ensure_ascii=False).replace("</", "<\\/")
        js = (JS.replace("__DECS__", json.dumps(self.decisions)).replace("__ITEMS__", json.dumps(self.item_of))
              .replace("__COLL__", json.dumps(s.get("collection", "decisions")))
              .replace("__SLCOLL__", json.dumps(s.get("slide_collection", "slides")))
              .replace("__RES__", json.dumps(self.res_info, ensure_ascii=False).replace("</", "<\\/")))
        hide = " hidden" if self.all else ""   # with All slides: decisions view and its summary start hidden
        page = (f'<title>{E(s["title"])}</title><meta name="claude:viewport-fit-sides" content="cover">'
                f"<style>{CSS}</style>"
                f'<div class="topbar"><div class="tb"><div class="tbrow">{tabs}'
                f'<span class="count" id="count">{n} open, 0 resolved</span>'
                '<span class="status" id="status" role="status" aria-live="polite"></span></div>'
                f'<nav class="chips" id="chips" aria-label="Your answers"{hide}>{chips}</nav></div></div><main>'
                + dec_view.replace('<div id="view-dec"', f'<div id="view-dec"{hide}', 1) + all_view
                + '</main><div class="lightbox" id="lightbox" hidden><img alt=""></div>'
                + f'<script type="application/json" id="imgs">{imgs}</script>'
                + f'<script type="application/json" id="slides">{slides}</script>'
                + (f"<script>{test}</script>" if test else "") + f"<script>{js}</script>")
        for ch, name in (("·", "middle dot"), ("µ", "micro sign (use Greek mu U+03BC)")):
            if ch in page:
                raise SystemExit(f"{name} on the page")
        visible = re.sub(r"<script.*?</script>|<style>.*?</style>|<[^>]+>", " ", page, flags=re.S)
        if re.search(r"before your saved change", visible, re.I):
            raise SystemExit("the page shows a slide as it was before a saved change: show the latest state only")
        files = re.findall(r"\b[\w-]+\.(?:m|py|csv|png|yaml|json|mat)\b", visible)
        if files:
            raise SystemExit(f"file names in page text: {sorted(set(files))}")
        return page


def slide_ranges(nums):
    """[2, 35, 36, 37] -> "2, 35–37"."""
    out, run = [], []
    for k in sorted(nums):
        if run and k == run[-1] + 1:
            run.append(k)
            continue
        if run:
            out.append(f"{run[0]}–{run[-1]}" if len(run) > 1 else str(run[0]))
        run = [k]
    if run:
        out.append(f"{run[0]}–{run[-1]}" if len(run) > 1 else str(run[0]))
    return ", ".join(out)


CSS = """
/* Layout: sticky counter and summary; per item a section, per decision a card (gray, green = picked,
   amber = change saved) holding a row of large answer buttons, each option's slide column and the change box. */
:root{--bg:#f6f6f3;--fg:#1d1d1b;--muted:#62625d;--card:#fff;--line:#d9d8d0;--accent:#1f5fbf;--accent-fg:#fff;
--ok:#1b7a3d;--okbg:#e6f4ea;--okline:#2e9e55;--chg:#9a5a00;--chgbg:#fdf1dc;--chgline:#d9921f;--err:#b3261e;--bar:#f6f6f3}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#1b1b19;--fg:#ecebe6;--muted:#a3a29c;--card:#252422;--line:#3a3935;--accent:#7fb0f5;--accent-fg:#10161f;--ok:#7fd39b;--okbg:#1d3326;--okline:#4fbf77;--chg:#f0c172;--chgbg:#3a2d15;--chgline:#d9a03a;--err:#f2867e;--bar:#1b1b19;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#1b1b19;--fg:#ecebe6;--muted:#a3a29c;--card:#252422;--line:#3a3935;--accent:#7fb0f5;--accent-fg:#10161f;--ok:#7fd39b;--okbg:#1d3326;--okline:#4fbf77;--chg:#f0c172;--chgbg:#3a2d15;--chgline:#d9a03a;--err:#f2867e;--bar:#1b1b19;color-scheme:dark}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
body::before{content:"";position:fixed;inset:0 0 auto;height:env(safe-area-inset-top,0px);z-index:10;background:var(--bar)}
.topbar{position:sticky;top:env(safe-area-inset-top,0px);z-index:5;background:var(--bar);border-bottom:1px solid var(--line);
padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
.tb{max-width:1200px;margin:0 auto;padding-block:8px}
.tbrow{display:flex;flex-wrap:wrap;gap:4px 16px;align-items:baseline;margin-bottom:6px}
.count{font-weight:700;font-size:17px;font-variant-numeric:tabular-nums}
.status{font-size:13px;color:var(--muted)}.status.ok{color:var(--ok)}.status.err{color:var(--err)}
.chips{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.chip{display:inline-flex;gap:6px;align-items:baseline;padding:3px 9px;border-radius:999px;border:1px solid var(--line);
background:var(--card);color:var(--muted);text-decoration:none;font-size:13px;white-space:nowrap}
.chip b{color:var(--fg)}
.chip.pick{background:var(--okbg);border-color:var(--okline);color:var(--ok)}
.chip.change{background:var(--chgbg);border-color:var(--chgline);color:var(--chg)}
@media (max-width:640px){.chips{flex-wrap:nowrap;overflow-x:auto;padding-bottom:4px}.count{font-size:14px}.tbrow{margin-bottom:4px}}
main{max-width:1200px;margin:0 auto;padding-block:18px 64px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}
h1{font-size:22px;margin:0 0 6px;text-wrap:balance}
.lead{color:var(--muted);margin:0 0 20px;max-width:80ch}
section{margin:0 0 28px;scroll-margin-top:130px}
h2{font-size:18px;margin:0 0 10px;display:flex;flex-wrap:wrap;gap:8px 10px;align-items:center}
.h2t{flex:1 1 300px;min-width:0;text-wrap:balance}
h3{font-size:15px;margin:20px 0 6px}
.num{display:inline-block;min-width:1.8em;height:1.8em;line-height:1.8em;text-align:center;border-radius:50%;background:var(--fg);color:var(--bg);font-size:13px;flex:none}
.badge{flex:none;font-size:13px;font-weight:700;padding:2px 10px;border-radius:999px;background:var(--line);color:var(--muted)}
.badge.done{background:var(--okbg);color:var(--ok)}
.card{background:var(--card);border:2px solid var(--line);border-radius:12px;padding:14px;margin:10px 0 14px;scroll-margin-top:130px}
.card.pick{border-color:var(--okline)}
.card.change{border-color:var(--chgline)}
.chead{display:flex;gap:10px;align-items:center;justify-content:space-between;margin:0 0 10px}
.ccode{font-weight:700;font-size:14px;color:var(--muted)}
.cstate{font-size:14px;font-weight:700;padding:3px 12px;border-radius:999px;background:var(--line);color:var(--muted)}
.card.pick .cstate{background:var(--okbg);color:var(--ok)}
.card.change .cstate{background:var(--chgbg);color:var(--chg)}
.answers{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(260px,100%),1fr));gap:10px}
.opt{display:flex;flex-direction:column;gap:6px;min-width:0}.opt .ans{flex:1}
.edit{align-self:flex-start;font:600 13px system-ui,-apple-system,"Segoe UI",sans-serif;padding:4px 12px;border-radius:8px;border:1px solid var(--chgline);background:var(--card);color:var(--chg);cursor:pointer}.edit:hover{background:var(--chgbg)}
.ans{display:flex;flex-direction:column;gap:6px;text-align:left;font:inherit;color:var(--fg);background:var(--card);
border:2px solid var(--accent);border-radius:10px;padding:12px 14px;min-height:72px;cursor:pointer}
.ans:hover{background:var(--okbg)}
.ahead{display:flex;align-items:center;gap:8px}
.alabel{font-size:17px;font-weight:700;color:var(--accent)}
.check{display:none;font-size:18px;font-weight:700}
.atext{font-size:14px;line-height:1.4}
.anote{display:block;font-size:12.5px;color:var(--muted)}
.ans.picked{background:var(--okline);border-color:var(--okline);color:#fff}
.ans.picked .alabel,.ans.picked .anote{color:#fff}
.ans.picked .check{display:inline}
.answers.has .ans:not(.picked){opacity:.5}
.ans:focus-visible,.chip:focus-visible,.savechg:focus-visible,textarea:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
q{quotes:"\\201C" "\\201D"}
.cols{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(330px,100%),1fr));gap:14px;margin-top:12px;align-items:start}
.col{border:1px solid var(--line);border-radius:10px;padding:10px;min-width:0}
.col.picked{border:2px solid var(--okline);padding:9px}
@media (min-width:760px){.cols>.col:only-child,.context>.slide:only-child{max-width:calc(50% - 7px)}}
.colhead{font-weight:700;margin-bottom:4px}
.nr{margin:4px 0 6px;font-weight:600;color:var(--chg)}
.slide{margin:6px 0 10px}
.slide img{display:block;width:100%;max-width:100%;height:auto;border:1px solid var(--line);border-radius:6px;background:#fff;cursor:zoom-in}
.slide figcaption{font-size:12.5px;color:var(--muted);margin-top:3px}
.context{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(330px,100%),1fr));gap:12px}
.settled{display:grid;gap:10px;border:2px solid var(--okline);background:var(--okbg);border-radius:12px;padding:12px;margin:8px 0 14px}
.stext p{margin:0 0 4px;font-weight:600;color:var(--fg)}
.static{display:flex;gap:10px;padding:8px 10px;border:1px dashed var(--line);border-radius:8px;margin:8px 0;background:var(--card)}
.lbl{flex:none;font:700 12px system-ui;color:var(--muted);padding:3px 4px}
.note{font-size:13px;color:var(--muted)}
.chgbox{margin:12px 0 0;display:grid;gap:6px}
.chgbox label{font-size:13px;font-weight:600}
textarea{width:100%;font:14px/1.45 system-ui,-apple-system,"Segoe UI",sans-serif;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:6px 8px;resize:vertical}
.savechg{justify-self:start;font:600 14px system-ui,-apple-system,"Segoe UI",sans-serif;padding:6px 16px;border-radius:8px;border:2px solid var(--chgline);background:var(--card);color:var(--chg);cursor:pointer}
.savechg:hover{background:var(--chgbg)}
ul{margin:6px 0;padding-left:20px}
[hidden]{display:none!important}
.slbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-top:6px}
.slopen,.slreset,.slclear{font:600 13px system-ui,-apple-system,"Segoe UI",sans-serif;padding:4px 12px;border-radius:8px;border:1px solid var(--chgline);background:var(--card);color:var(--chg);cursor:pointer}
.slopen:hover,.slreset:hover,.slclear:hover{background:var(--chgbg)}
.slstate{font-size:13px;font-weight:700;color:var(--chg)}
.slide.changed img{border:3px solid var(--chgline)}
.sledit{display:grid;gap:6px;margin-top:6px;padding:10px;border:1px solid var(--line);border-radius:8px;background:var(--bg)}
.slide.changed .sledit{border-color:var(--chgline);background:var(--chgbg)}   /* amber only once a change is saved */
.sledit label{font-size:13px;font-weight:600}
.sledit textarea{background:var(--card);white-space:pre-wrap;overflow:hidden;resize:none}
.slnote{font-size:12.5px;color:var(--muted);margin:0}
.savednote{font-size:13px;font-weight:600;color:var(--chg);margin:2px 0 0}
.scope{font-size:13px;color:var(--fg);margin:0}
.fact{font-size:13px;font-weight:600;margin:2px 0 0;color:var(--ok)}
.fact.warn{color:var(--err)}
.slbtns{display:flex;flex-wrap:wrap;gap:8px}
.slsave{font:600 14px system-ui,-apple-system,"Segoe UI",sans-serif;padding:6px 16px;border-radius:8px;border:2px solid var(--chgline);background:var(--card);color:var(--chg);cursor:pointer}
.slsave:hover{background:var(--chgbg)}
.slopen:focus-visible,.slsave:focus-visible,.slreset:focus-visible,.slclear:focus-visible,.edit:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
.tabs{display:flex;gap:6px;flex:none}
.tab{font:700 15px system-ui,-apple-system,"Segoe UI",sans-serif;padding:6px 16px;border-radius:8px;border:2px solid var(--accent);background:var(--card);color:var(--accent);cursor:pointer}
.tab.on{background:var(--accent);color:var(--accent-fg)}
.tab:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
.decklist{display:grid;gap:22px}
.slide.deck{margin:0;padding:12px;background:var(--card);border:1px solid var(--line);border-radius:12px;scroll-margin-top:90px;max-width:1100px}
.slide.deck.changed{border:2px solid var(--chgline);padding:11px}
.dhead{display:flex;flex-wrap:wrap;gap:6px 14px;align-items:baseline;margin:0 0 6px}
.dnum{font-size:17px;font-weight:700}
.slide.deck .sledit{margin-top:8px}
@media (max-width:640px){.tab{font-size:14px;padding:5px 10px}.slide.deck{padding:8px;scroll-margin-top:140px}.slide.deck.changed{padding:7px}}
.lightbox{position:fixed;inset:0;z-index:20;background:rgba(0,0,0,.85);display:flex;align-items:center;justify-content:center;padding:16px;cursor:zoom-out}
.lightbox[hidden]{display:none}
.lightbox img{max-width:100%;max-height:100%;background:#fff}
"""

JS = r"""
const DECS = __DECS__, ITEM = __ITEMS__, COLL = __COLL__, SLCOLL = __SLCOLL__;
const state = {}, slState = {}, writing = {};
let db = null, dbFailed = false, loaded = false;
const $ = s => document.querySelector(s), $$ = s => document.querySelectorAll(s);
const statusEl = $('#status');
let SLIDES = {};
try {
  const imgs = JSON.parse($('#imgs').textContent);
  $$('img[data-img]').forEach(im => { im.src = imgs[im.dataset.img]; });
  SLIDES = JSON.parse($('#slides').textContent);
} catch (e) {}
// grow a slide box to its text (no inner scrolling); a hidden box is sized when it opens
const fit = t => { if (!t.offsetParent) return; t.style.height = 'auto'; t.style.height = (t.scrollHeight + 4) + 'px'; };
function setStatus(msg, kind) { statusEl.textContent = msg; statusEl.className = 'status' + (kind ? ' ' + kind : ''); }
const statusOf = s => (s && (s.status || (s.choice ? 'pick' : ''))) || '';
const slChanged = id => !!(slState[id] && slState[id].status === 'change');
const before = id => (SLIDES[id] || {}).text || '';
function labelOf(dec, choice) {
  const b = document.querySelector('.ans[data-dec="' + dec + '"][data-choice="' + CSS.escape(choice) + '"]');
  return b ? b.dataset.label : choice;
}
function render() {
  $$('.ans').forEach(b => {
    const s = state[b.dataset.dec], on = !!(s && s.choice === b.dataset.choice);
    b.classList.toggle('picked', on);
    b.setAttribute('aria-checked', on ? 'true' : 'false');
  });
  $$('.col').forEach(c => { const s = state[c.dataset.dec]; c.classList.toggle('picked', !!(s && s.choice === c.dataset.choice)); });
  let picks = 0, changes = 0;
  DECS.forEach(dec => {
    const s = state[dec] || {}, st = statusOf(s);
    if (st === 'pick') picks++;
    if (st === 'change') changes++;
    const card = document.querySelector('[data-card="' + dec + '"]');
    card.classList.toggle('pick', st === 'pick');
    card.classList.toggle('change', st === 'change');
    card.querySelector('.answers').classList.toggle('has', !!s.choice);
    const label = st === 'pick' ? 'Your pick: ' + labelOf(dec, s.choice) : st === 'change' ? 'Change saved' : 'Not answered';
    document.querySelector('[data-state="' + dec + '"]').textContent = label;
    const chip = document.querySelector('[data-chip="' + dec + '"]');
    chip.querySelector('.cv').textContent = st === 'pick' ? labelOf(dec, s.choice) : st === 'change' ? 'Change' : '–';
    chip.classList.toggle('pick', st === 'pick');
    chip.classList.toggle('change', st === 'change');
  });
  $$('.badge[data-item]').forEach(b => {
    const decs = DECS.filter(d => ITEM[d] === b.dataset.item);
    const n = decs.filter(d => statusOf(state[d])).length;
    b.textContent = n === decs.length ? 'Answered' : (n === 0 ? 'Not answered' : n + ' of ' + decs.length + ' answered');
    b.classList.toggle('done', n === decs.length);
  });
  $$('textarea[data-note]').forEach(t => { if (document.activeElement !== t && !t.dataset.dirty) { t.value = (state[t.dataset.note] || {}).note || ''; } });
  // slides: every copy of a slide shows the same saved change
  $$('figure[data-slide]').forEach(f => {
    const id = f.dataset.slide, ch = slChanged(id), open = !f.querySelector('.sledit').hidden;
    f.classList.toggle('changed', ch);
    f.querySelector('.slstate').textContent = ch ? 'Change saved' : '';
    const ob = f.querySelector('.slopen');   // decisions view only; All slides rows keep their box open
    if (ob) {
      ob.textContent = open ? 'Close' : (ch ? 'Show saved change' : 'Change this slide');
      ob.setAttribute('aria-expanded', open ? 'true' : 'false');
    }
    f.querySelector('.slclear').hidden = !ch;
    const sn = f.querySelector('.savednote'); if (sn) sn.hidden = !ch;
    const t = f.querySelector('textarea[data-sltext]');
    if (document.activeElement !== t && !t.dataset.dirty) { t.value = ch ? slState[id].text : before(id); fit(t); }
  });
  const changedSlides = Object.keys(SLIDES).filter(slChanged);
  const sc = $('#slchip');
  sc.hidden = !changedSlides.length;
  sc.classList.toggle('change', !!changedSlides.length);
  sc.querySelector('.cv').textContent = changedSlides.length + ' changed';
  if (changedSlides.length) sc.href = '#' + (document.getElementById('a-' + changedSlides[0]) ? 'a-' : 's-') + changedSlides[0];
  $('#count').textContent = picks + ' of ' + DECS.length + ' picked, ' + changes + ' with changes, ' +
    (DECS.length - picks - changes) + ' not answered' +
    (changedSlides.length ? '; ' + changedSlides.length + ' slide' + (changedSlides.length > 1 ? 's' : '') + ' changed' : '');
}
const idle = () => !Object.keys(writing).length;
// One write at a time per document; a click during a write is sent after it, with the latest state.
async function save(coll, id, data) {
  if (!db) { if (dbFailed) setStatus('Not saved: saving is unavailable in this view. Open the page on claude.ai while signed in.', 'err'); return; }
  const key = coll + '/' + id;
  if (writing[key]) { writing[key].again = true; return; }
  writing[key] = { again: false };
  setStatus('Saving...');
  let failed = null;
  do {
    writing[key].again = false;
    try { await db.collection(coll).doc(id).set(Object.assign(data(), { updatedAt: new Date().toISOString() })); }
    catch (e) { failed = e; }
  } while (writing[key].again && !failed);
  delete writing[key];
  if (failed) setStatus('Not saved (' + (failed.code || 'error') + '). Click again to retry.', 'err');
  else if (idle()) setStatus('Saved', 'ok');
}
const write = dec => save(COLL, dec, () => { const s = state[dec] || {};
  return { decision: dec, status: statusOf(s), choice: s.choice || '', text: s.text || '', note: s.note || '' }; });
const writeSlide = id => save(SLCOLL, id, () => { const s = slState[id] || {};
  return { slide: id, status: s.status || '', before: before(id), text: s.text || '' }; });
function slideSaved(id, st) {
  slState[id] = st;
  $$('textarea[data-sltext="' + id + '"]').forEach(t => { delete t.dataset.dirty; });
  render(); writeSlide(id);
}
// Views: All slides (default when present) and Decisions; any in-page link opens the view that holds its target.
const VIEWS = [...$$('[data-viewname]')];
function show(v) {
  if (!VIEWS.some(x => x.dataset.viewname === v)) return;
  VIEWS.forEach(x => { x.hidden = x.dataset.viewname !== v; });
  $$('.tab').forEach(t => { const on = t.dataset.view === v; t.classList.toggle('on', on); t.setAttribute('aria-selected', on ? 'true' : 'false'); });
  const chips = $('#chips'); if (chips && $('.tab')) chips.hidden = v !== 'dec';
  $$('textarea[data-sltext]').forEach(fit);
}
function jumpTo(id) {
  const el = document.getElementById(id); if (!el) return false;
  const v = el.closest('[data-viewname]'); if (v) show(v.dataset.viewname);
  el.scrollIntoView({ block: 'start' }); return true;
}
window.jumpTo = jumpTo;
document.addEventListener('input', e => { if (e.target.matches('textarea[data-sltext], textarea[data-note]')) { e.target.dataset.dirty = '1'; if (e.target.dataset.sltext) fit(e.target); } });
document.addEventListener('click', e => {
  const im = e.target.closest('.slide img');
  if (im) { const lb = $('#lightbox'); lb.querySelector('img').src = im.src; lb.hidden = false; return; }
  if (e.target.closest('#lightbox')) { $('#lightbox').hidden = true; return; }
  const tab = e.target.closest('.tab');
  if (tab) { show(tab.dataset.view); window.scrollTo(0, 0); return; }
  const link = e.target.closest('a[href^="#"]');
  if (link) { const id = link.getAttribute('href').slice(1); if (id) { e.preventDefault(); jumpTo(id); } return; }
  const fig = e.target.closest('figure[data-slide]');
  if (fig && e.target.closest('.slopen')) {
    const box = fig.querySelector('.sledit'); box.hidden = !box.hidden; render();
    if (!box.hidden) { const t = fig.querySelector('textarea'); fit(t); t.focus(); }
    return;
  }
  if (fig && e.target.closest('.slsave')) {
    const id = fig.dataset.slide, t = fig.querySelector('textarea'), txt = t.value;
    if (txt.trim() === before(id).trim()) {
      if (slChanged(id)) { slideSaved(id, { status: 'cleared', text: '' }); setStatus('Back to the slide\'s text: saved change removed.', 'ok'); }
      else setStatus('Nothing changed yet: edit the text first, then click Save slide change.', 'err');
      return;
    }
    if (!txt.trim()) { setStatus('The box is empty: write what this slide should say.', 'err'); return; }
    slideSaved(id, { status: 'change', text: txt }); return;
  }
  if (fig && e.target.closest('.slreset')) {
    const t = fig.querySelector('textarea'); t.value = before(fig.dataset.slide); t.dataset.dirty = '1'; fit(t);
    setStatus('Slide text restored in the box; click Save slide change to keep it.', 'ok'); return;
  }
  if (fig && e.target.closest('.slclear')) { slideSaved(fig.dataset.slide, { status: 'cleared', text: '' }); return; }
  const ed = e.target.closest('.edit');
  if (ed) {
    const t = document.querySelector('textarea[data-note="' + ed.dataset.dec + '"]');
    t.value = ed.dataset.text; t.dataset.dirty = '1'; t.rows = Math.max(2, Math.ceil(t.value.length / 90));
    t.scrollIntoView({ block: 'center' }); t.focus();
    setStatus('Text copied into the change box: edit it, then click Save change.', 'ok'); return;
  }
  const sv = e.target.closest('.savechg');
  if (sv) {
    const dec = sv.dataset.dec, t = document.querySelector('textarea[data-note="' + dec + '"]');
    if (!t.value.trim()) { setStatus('Write the change in the box first.', 'err'); t.focus(); return; }
    state[dec] = Object.assign({}, state[dec], { note: t.value, status: 'change' });
    delete t.dataset.dirty;
    render(); write(dec); return;
  }
  const b = e.target.closest('.ans');
  if (!b) return;
  const dec = b.dataset.dec, t = document.querySelector('textarea[data-note="' + dec + '"]');
  state[dec] = Object.assign({}, state[dec], { choice: b.dataset.choice, text: b.dataset.text, status: 'pick', note: t.value });
  delete t.dataset.dirty;
  render(); write(dec);
});
document.addEventListener('keydown', e => { if (e.key === 'Escape') $('#lightbox').hidden = true; });
render();
if (!(location.hash.length > 1 && jumpTo(location.hash.slice(1)))) show(VIEWS.some(x => x.dataset.viewname === 'all') ? 'all' : 'dec');
setStatus('Loading saved answers...');
(async () => {
  try { db = window.claude && window.claude.use ? await window.claude.use('db') : null; } catch (e) { db = null; }
  if (!db) { dbFailed = true; setStatus('Saving is unavailable in this view. Open the page on claude.ai while signed in.', 'err'); return; }
  const timer = setTimeout(() => { if (!loaded) setStatus('Saved answers cannot be reached yet; clicks are kept on this page and saved when the connection returns.', 'err'); }, 10000);
  const fail = err => { clearTimeout(timer); setStatus('Saved answers could not be loaded (' + err.code + '). Reload the page.', 'err'); };
  db.collection(COLL).onSnapshot(snap => {
    const seen = {};
    snap.docs.forEach(d => { seen[d.id] = d.data(); });
    DECS.forEach(dec => {
      if (writing[COLL + '/' + dec]) return;
      const r = seen[dec];
      state[dec] = r ? { status: r.status || '', choice: r.choice || '', text: r.text || '', note: r.note || '' } : {};
    });
    render();
    if (!snap.metadata.fromCache && !loaded) { loaded = true; clearTimeout(timer); if (idle()) setStatus('Ready: each click saves', 'ok'); }
  }, fail);
  db.collection(SLCOLL).onSnapshot(snap => {
    const seen = {};
    snap.docs.forEach(d => { seen[d.id] = d.data(); });
    Object.keys(SLIDES).forEach(id => {
      if (writing[SLCOLL + '/' + id]) return;
      const r = seen[id];
      slState[id] = r ? { status: r.status || '', text: r.text || '' } : {};
    });
    render();
  }, fail);
})();
"""

# In-memory stand-in for claude.use("db") for local test pages only (never published). Documents persist in
# sessionStorage (one store per collection) so a reload in the same tab sees them. ACTIONS run after load:
# ["pick", dec, button index], ["change", dec, text] or ["slide", slide id, text to append]; SCROLL is an element id
# to scroll to; CYCLE = [decision, slide id] runs the answer cycle (below), reloads, and prints the result.
STUB = r"""
(() => {
  const KEY = 'stubdb_' + location.pathname;
  const load = c => JSON.parse(sessionStorage.getItem(KEY + '|' + c) || '{}');
  const subs = {};
  const snap = c => { const docs = load(c); return { docs: Object.keys(docs).map(id => ({ id, exists: true, data: () => docs[id] })), metadata: { fromCache: false } }; };
  const db = { collection: c => ({
    doc: id => ({ set: async d => { const docs = load(c); docs[id] = d; sessionStorage.setItem(KEY + '|' + c, JSON.stringify(docs));
      setTimeout(() => (subs[c] || []).forEach(f => f(snap(c))), 0); } }),
    onSnapshot: f => { (subs[c] = subs[c] || []).push(f); setTimeout(() => f(snap(c)), 0); return () => {}; } }) };
  window.claude = { use: async name => (name === 'db' ? db : null) };
  const ACTIONS = __ACTIONS__, SCROLL = __SCROLL__, CYCLE = __CYCLE__;
  const wait = ms => new Promise(r => setTimeout(r, ms));
  const pick = (dec, i) => document.querySelectorAll('.ans[data-dec="' + dec + '"]')[i].click();
  const change = (dec, txt) => { document.querySelector('textarea[data-note="' + dec + '"]').value = txt;
    document.querySelector('.savechg[data-dec="' + dec + '"]').click(); };
  const slideEdit = async (id, add, deck) => {
    const f = document.querySelector((deck ? 'figure.deck' : 'figure.slide:not(.deck)') + '[data-slide="' + id + '"]');
    if (f.querySelector('.slopen')) { f.querySelector('.slopen').click(); await wait(50); }
    const t = f.querySelector('textarea'); t.value = t.value + add; t.dispatchEvent(new Event('input', { bubbles: true }));
    f.querySelector('.slsave').click(); };
  const cardState = dec => { const c = document.querySelector('[data-card="' + dec + '"]');
    return (c.classList.contains('pick') ? 'green' : c.classList.contains('change') ? 'amber' : 'gray') + ' | ' + c.querySelector('.cstate').textContent; };
  const slideState = id => { const fs = [...document.querySelectorAll('figure[data-slide="' + id + '"]')];
    return fs.filter(f => f.classList.contains('changed')).length + '/' + fs.length + ' | ' + fs[0].querySelector('.slstate').textContent; };
  window.addEventListener('load', async () => {
    await wait(300);
    if (CYCLE) {
      const [dec, sl, coll, slColl] = CYCLE;
      const phase = sessionStorage.getItem(KEY + '_phase');
      const out = document.createElement('pre'); out.id = 'cycle'; document.body.prepend(out);
      if (!phase) {
        const log = [];
        const va = document.getElementById('view-all');
        if (va) log.push('opens on: ' + (!va.hidden && document.getElementById('view-dec').hidden ? 'all slides' : 'decisions'));
        const ed = document.querySelector('.edit[data-dec="' + dec + '"]');
        if (ed) { ed.click(); await wait(50);
          log.push('edit copies text: ' + (document.querySelector('textarea[data-note="' + dec + '"]').value === ed.dataset.text)); }
        pick(dec, 0); await wait(150); log.push('after pick: ' + cardState(dec));
        change(dec, 'Test change text'); await wait(150); log.push('after change: ' + cardState(dec));
        pick(dec, 1); await wait(150); log.push('after re-pick: ' + cardState(dec));
        await slideEdit(sl, '\nTEST SLIDE EDIT'); await wait(150); log.push('slide after save: ' + slideState(sl));
        sessionStorage.setItem(KEY + '_log', JSON.stringify(log)); sessionStorage.setItem(KEY + '_phase', 'reloaded');
        await wait(200); location.reload(); return;
      }
      const log = JSON.parse(sessionStorage.getItem(KEY + '_log'));
      log.push('after reload: ' + cardState(dec));
      log.push('stored decision: ' + JSON.stringify(load(coll)[dec]));
      const f = document.querySelector('figure[data-slide="' + sl + '"]');
      log.push('slide after reload: ' + slideState(sl) + ' | box has edit: ' + f.querySelector('textarea').value.includes('TEST SLIDE EDIT'));
      log.push('stored slide: ' + JSON.stringify(load(slColl)[sl]));
      log.push('count: ' + document.getElementById('count').textContent);
      f.querySelector('.slclear').click(); await wait(150);
      log.push('slide after remove: ' + slideState(sl) + ' | stored status ' + load(slColl)[sl].status);
      const dsl = CYCLE[4];   // All slides view: its box is open and filled, and its save reaches every copy
      if (dsl) {
        document.querySelector('.tab[data-view="all"]').click(); await wait(100);
        const df = document.querySelector('figure.deck[data-slide="' + dsl + '"]'), dt = df.querySelector('textarea');
        const want = (JSON.parse(document.getElementById('slides').textContent)[dsl] || {}).text || '';
        log.push('deck row box open and filled: ' + (!df.querySelector('.sledit').hidden && dt.offsetHeight > 0 && want.length > 0 && dt.value === want));
        await slideEdit(dsl, '\nDECK ROW EDIT', true); await wait(150);
        log.push('deck row save: ' + slideState(dsl) + ' | stored ' + (load(slColl)[dsl] || {}).status + ' ' + ((load(slColl)[dsl] || {}).text || '').includes('DECK ROW EDIT'));
        document.querySelector('.tab[data-view="dec"]').click(); await wait(100);
        log.push('decisions tab: ' + (!document.getElementById('view-dec').hidden && document.getElementById('view-all').hidden));
      }
      out.textContent = 'CYCLE\n' + log.join('\n') + '\nEND';
      return;
    }
    for (const a of ACTIONS) {
      if (a[0] === 'pick') pick(a[1], a[2]);
      else if (a[0] === 'slide') await slideEdit(a[1], a[2]);
      else change(a[1], a[2]);
      await wait(100);
    }
    if (SCROLL) { await wait(200); window.jumpTo(SCROLL); }
  });
})();
"""


def edge(src, png=None, size="1400,1000", dump=False):
    cmd = [EDGE, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
           f"--window-size={size}", "--virtual-time-budget=8000"]
    cmd += ["--dump-dom"] if dump else [f"--screenshot={png}"]
    r = subprocess.run(cmd + [Path(src).resolve().as_uri()], capture_output=True, timeout=180,
                       text=dump, encoding="utf-8" if dump else None, errors="replace" if dump else None)
    if r.returncode:
        raise SystemExit(f"Edge failed on {src}")
    return r.stdout if dump else None


def stub(actions=(), scroll=None, cycle=None):
    return (STUB.replace("__ACTIONS__", json.dumps(list(actions))).replace("__SCROLL__", json.dumps(scroll))
            .replace("__CYCLE__", json.dumps(cycle)))


def shoot(page_html, png, w, h):
    """Screenshot through an iframe: headless Edge will not open a window narrower than ~492 px and does not paint a
    scrolled top-level page."""
    frame = page_html.with_name(page_html.stem + f"_{w}x{h}_frame.html")
    frame.write_text(f'<body style="margin:0;overflow:hidden"><iframe src="{page_html.name}" '
                     f'style="width:{w}px;height:{h}px;border:0"></iframe></body>', encoding="utf-8")
    edge(frame, png, f"{max(w, 600)},{h}")
    Image.open(png).crop((0, 0, w, h)).save(png)
    return png


def run_tests(spec, spec_dir, test_dir, decisions, slide_ids, deck_ids=()):
    """Screenshots (desktop 1400 px, phone 400 px) and one answer cycle with a reload; fails on any lost save.
    With an All slides view: shots of it, one tall shot per 3 deck slides, and a save from a deck row in the cycle."""
    test_dir = Path(test_dir)
    (test_dir / "review").mkdir(parents=True, exist_ok=True)
    changed = decisions[3] if len(decisions) > 3 else decisions[-1]
    sl = slide_ids[0]
    acts = [["pick", decisions[0], 0], ["pick", decisions[1], 1], ["pick", decisions[2], 0],
            ["change", changed, "Example change text"], ["slide", sl, "\nExample slide change"]]
    shots = []
    views = [("unanswered", stub()), ("answered", stub(acts)), ("answered_changed_card", stub(acts, "d-" + changed)),
             ("answered_slide_change", stub(acts, "s-" + sl))]
    if deck_ids:
        views += [("answered_decisions", stub(acts, "view-dec")),
                  ("answered_deck_row", stub(acts, "a-" + deck_ids[min(10, len(deck_ids) - 1)]))]
    for name, st in views:
        hp = test_dir / f"test_{name}.html"
        hp.write_text(Page(spec, spec_dir).build(st), encoding="utf-8")
        shots.append(shoot(hp, test_dir / f"test_{name}_desktop.png", 1400, 1000))
        shots.append(shoot(hp, test_dir / f"test_{name}_phone.png", 400, 860))
    for it in spec["items"]:   # one tall desktop shot per item of the answered page, for reviewing every card
        num = str(it.get("number", ""))
        safe = re.sub(r"[^\w-]", "x", num)
        rh = test_dir / "review" / f"item_{safe}.html"
        rh.write_text(Page(spec, spec_dir).build(stub(acts, "i" + num)), encoding="utf-8")
        shots.append(shoot(rh, test_dir / "review" / f"item_{safe}.png", 1400, 2600))
    for old in (test_dir / "review").glob("all_*"):
        old.unlink()
    for k in range(0, len(deck_ids), 3):   # the All slides view, 3 deck slides per tall shot
        rh = test_dir / "review" / f"all_{k + 1:02d}.html"
        rh.write_text(Page(spec, spec_dir).build(stub(acts, "a-" + deck_ids[k])), encoding="utf-8")
        shots.append(shoot(rh, test_dir / "review" / f"all_{k + 1:02d}.png", 1400, 2600))
    print(f"test screenshots (answered = picks on {', '.join(decisions[:3])}, change saved on {changed}, "
          f"slide change saved on {sl}):")
    for p in shots:
        print("  ", p)
    cyc = test_dir / "test_cycle.html"
    dsl = next((d for d in deck_ids if d in slide_ids[:slide_ids.index(deck_ids[0])]), deck_ids[0]) if deck_ids else None
    cyc.write_text(Page(spec, spec_dir).build(stub(cycle=[decisions[0], sl, spec.get("collection", "decisions"),
                                                          spec.get("slide_collection", "slides"), dsl])), encoding="utf-8")
    m = re.search(r"CYCLE\n(.*?)\nEND", html.unescape(edge(cyc, dump=True)), re.S)
    if not m:
        raise SystemExit("answer cycle test produced no result")
    log = m.group(1)
    print(f"answer cycle (decision {decisions[0]}, slide {sl}):\n  " + log.replace("\n", "\n  "))
    expect = [r"edit copies text: true", r"after pick: green", r"after change: amber \| Change saved",
              r"after re-pick: green", r"after reload: green", r'stored decision: .*"status":"pick"',
              r'stored decision: .*"note":"Test change text"', r"slide after save: (\d+)/\1 \| Change saved",
              r"slide after reload: (\d+)/\1 \| Change saved \| box has edit: true",
              r'stored slide: .*"status":"change".*TEST SLIDE EDIT', r"slide after remove: 0/\d+ \|  \| stored status cleared"]
    if deck_ids:
        expect += [r"opens on: all slides", r"deck row box open and filled: true",
                   r"deck row save: (\d+)/\1 \| Change saved \| stored change true", r"decisions tab: true"]
    missing = [x for x in expect if not re.search(x, log)]
    if missing:
        raise SystemExit(f"answer cycle failed: no match for {missing}")
    print("answer cycle: PASS")


def main(argv):
    spec_path, out_html = Path(argv[0]), argv[1]
    test_dir = argv[argv.index("--test-dir") + 1] if "--test-dir" in argv else None
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    p = Page(spec, spec_path.parent)
    page = p.build()
    Path(out_html).write_text(page, encoding="utf-8")
    mb = len(page.encode()) / 1e6
    print(f"{out_html}: {len(p.decisions)} decisions, {len(p.images)} slide images, {mb:.1f} MB")
    if mb > 15:
        raise SystemExit("page over 15 MB (artifact limit 16 MB): lower png_width")
    if test_dir:
        deck = [str(x["id"]) for x in json.loads(p.resolve(spec["all_slides"]["index"]).read_text(encoding="utf-8"))]             if spec.get("all_slides") else []
        run_tests(spec, spec_path.parent, test_dir, p.decisions, list(p.slide_info), deck)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    main(sys.argv[1:])
