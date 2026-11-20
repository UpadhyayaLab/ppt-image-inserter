"""Deck text check run on every build: no walls of text on slides (FF 2026-11-19: "you need clean formatting.
period. no exceptions. no walls of text. no exceptions" / "how do we stop this for good ... for all future ppts?").

Checked on every slide except the title slide (slide 1):
- no middle dot or division sign in slide text (FF: "do not use · anywere"; a "Relative = ÷ ..." note was slop)
No line limit per text box (FF 2026-11-19: "drop it everytwhere") and no character limit (FF 2026-11-19: "there
should not be any character limit").
Text drawn inside images (rendered tables, figure labels) is not seen here; builders that render text images check
their own source text (e.g. fig6_metrics/rebuild_fig6_deck.py).
"""
from typing import List

FORBIDDEN = {"·": "middle dot", "÷": "division sign"}


def lint_presentation(prs) -> List[str]:
    """Return one line per problem ('slide N: ...'); empty when the deck's slide text is clean."""
    problems = []
    for n, slide in enumerate(prs.slides, 1):
        if n == 1:
            continue
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text
            for ch, name in FORBIDDEN.items():
                if ch in text:
                    problems.append(f"slide {n}: {name} in {text[:60]!r}")
    return problems
