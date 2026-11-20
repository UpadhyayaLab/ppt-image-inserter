"""Approved slide text: a deck may carry only text listed in its approved-text file (FF 2026-11-19: "we need to
prevent these mistakes...period", after agents put unchecked wording on slides again and again).

The file (YAML) lists every string a slide may show, each with its source:
    exact:            whole lines of slide text
    title_prefixes:   a title may be "<prefix>: <label>" ...
    title_labels:     ... with the prefix and the label each listed (so 60 titles need not be listed one by one)
Each entry is {text: ..., source: ...}. A source is FF's approval (date and his words), a code-checked definitions
file entry, or "settled before <date>, not re-checked" for text that was already on a finished deck when the file was
seeded. An entry may also carry open: "<why>" when FF may still change it. Agent wording never goes in on its own.

batch_insert_images.py runs the check when the config has `approved_text: <file>` and fails the build on any slide
line not in the file. Other builders call `unapproved(deck_strings(prs), load_approved(path))` before saving.

Command line:
    python examples_and_configs/check_approved_text.py check <deck.pptx> <approved.yaml>
    python examples_and_configs/check_approved_text.py seed  <deck.pptx> <approved.yaml> "<source>"
`seed` adds every line of an existing deck that the file lacks, with the given source (e.g. "settled before
2026-11-19, not re-checked"), so text already settled on a finished deck is not re-verified one string at a time.
"""
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Set, Tuple

import yaml

HEADER = """# Approved slide text for one deck (ppt_image_inserter.approved_text). The build fails on any slide line not
# listed here. Add a line ONLY with its source: FF's approval (date and his words), a code-checked definitions entry,
# or "settled before <date>, not re-checked" (seeded from a finished deck). Never wording an agent wrote itself.
"""
KEYS = ("exact", "title_prefixes", "title_labels")


def load_approved(path) -> Dict[str, Set[str]]:
    """{'exact': set, 'title_prefixes': set, 'title_labels': set} from an approved-text YAML file."""
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    unknown = set(data) - set(KEYS)
    if unknown:
        raise ValueError(f"{path}: unknown keys {sorted(unknown)} (allowed: {', '.join(KEYS)})")
    out = {}
    for key in KEYS:
        rows = data.get(key) or []
        missing = [r.get("text") for r in rows if not str(r.get("source", "")).strip()]
        if missing:
            raise ValueError(f"{path}: {len(missing)} {key} entries have no source, e.g. {missing[0]!r}")
        out[key] = {r["text"] for r in rows}
    return out


def deck_strings(prs) -> List[Tuple[str, str]]:
    """(where, line) for every non-empty line of text on every slide (paragraphs and line breaks split). Lines that
    are only digits (slide numbers) are skipped."""
    out = []
    for n, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for line in re.split(r"[\n\v]", shape.text_frame.text):
                line = line.strip()
                if line and not line.isdigit():
                    out.append((f"slide {n}", line))
    return out


def is_approved(text: str, ok: Dict[str, Set[str]]) -> bool:
    prefix, sep, label = text.partition(": ")
    return (text in ok["exact"] or text in ok["title_labels"]
            or bool(sep and prefix in ok["title_prefixes"] and label in ok["title_labels"]))


def unapproved(strings: Iterable[Tuple[str, str]], ok: Dict[str, Set[str]]) -> List[str]:
    """One problem line per slide string that is not approved; empty when every string is."""
    return [f"{where}: not in the approved-text file: {text!r}" for where, text in strings if not is_approved(text, ok)]


def seed_approved(strings: Iterable[Tuple[str, str]], path, source: str) -> int:
    """Add every string the file lacks as an `exact` entry with `source`; keep existing entries. Returns the count
    added. For text already settled on a finished deck, so it is not re-verified."""
    if not source.strip():
        raise ValueError("a source is required")
    path = Path(path)
    data = yaml.safe_load(path.read_text(encoding="utf-8")) if path.exists() else {}
    data = data or {}
    ok = {k: {r["text"] for r in data.get(k) or []} for k in KEYS}
    added = []
    for _, text in strings:
        if not is_approved(text, ok) and text not in added:
            added.append(text)
    data.setdefault("exact", []).extend({"text": t, "source": source} for t in added)
    body = yaml.safe_dump({k: data[k] for k in KEYS if data.get(k)}, allow_unicode=True, sort_keys=False, width=200)
    old = path.read_text(encoding="utf-8") if path.exists() else ""
    head = "".join(ln + "\n" for ln in old.splitlines() if ln.startswith("#")) or HEADER
    path.write_text(head + body, encoding="utf-8")
    return len(added)


def _main(argv):
    from pptx import Presentation
    if len(argv) < 3 or argv[0] not in ("check", "seed") or (argv[0] == "seed" and len(argv) != 4):
        raise SystemExit(__doc__)
    strings = deck_strings(Presentation(argv[1]))
    if argv[0] == "seed":
        print(f"added {seed_approved(strings, argv[2], argv[3])} lines to {argv[2]}")
        return 0
    problems = unapproved(strings, load_approved(argv[2]))
    for p in problems:
        print(p)
    print(f"{len(strings)} slide lines, {len(problems)} not approved")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
