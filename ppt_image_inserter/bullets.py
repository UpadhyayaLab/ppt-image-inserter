"""Bulleted text slides: real PowerPoint bullets and sub-bullets with a hanging indent.

(FF 2026-11-24: "we need bullets. apply standard formatting, subbullets in sensible locations".) Formatting as
bullet_list in examples_and_configs/CART/CART_GzmB_ppt/build_gzmb_vs_ctsb_deck.py: level 0 "•", deeper levels "–",
each level 0.5 in further in and 2 pt smaller; a gap after each group, not between a line and its sub lines.

Lines come from a deck text file: a string is a bullet; {text, sub: [...]} nests its sub lines one level down.
"""
from typing import Iterable, List, Sequence, Tuple, Union

from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

Line = Union[str, dict]


def flatten(lines: Iterable[Line], level: int = 0, sub_under: str = None, sub: Sequence[str] = ()) -> List[Tuple[str, int]]:
    """(text, level) for each line, sub lines one level down. The lines `sub` go one level under the line equal to
    `sub_under` (used for option previews)."""
    items = []
    for ln in lines:
        t = ln if isinstance(ln, str) else ln["text"]
        items.append((t, level))
        if t == sub_under:
            items += [(x, level + 1) for x in sub]
        if not isinstance(ln, str):
            items += flatten(ln.get("sub", []), level + 1, sub_under, sub)
    return items


def add_bullets(slide, items: Sequence[Tuple[str, int]], left: float, top: float, width: float, height: float,
                size: int = 24, color: RGBColor = RGBColor(0, 0, 0)):
    """Draw (text, level) items as one text box of real bullets (positions in inches). Returns the text box."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, (t, level) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        nxt = items[i + 1][1] if i + 1 < len(items) else 0
        p.space_after = Pt(16 if nxt == 0 else 6)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(Inches(0.4 + 0.5 * level)))
        pPr.set("indent", str(-Inches(0.4)))
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": "•" if level == 0 else "–"}))
        r = p.add_run()
        r.text = t
        r.font.size = Pt(size - 2 * level)
        r.font.color.rgb = color
    return tb
