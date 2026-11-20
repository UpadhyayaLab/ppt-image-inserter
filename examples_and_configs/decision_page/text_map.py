"""One text map for a deck and its decision page (FF 2026-11-19: "it shouldn't be this hard").

An answer is a list of line replacements {old, new, where}. A replacement matches whole lines only, never part of a
line (Histone mark data agent presentation, 2026-11-19: "Relative level near nadir" starts other lines):
- a slide title as a whole, its prefix (before ": ") or its label (after ": "), and a contents row's `titles`
  (so contents ranges follow title changes);
- a subtitle, or any other string in the deck text file (contents rows and notes, table headers and keys).
where: "titles" (titles only), "lines" (everything else), or omitted (both). A replacement that matches nothing fails.

The same function makes the decision page's previews (in memory) and applies FF's answers to the files
(apply_answers.py), so a preview always shows exactly what applying the answer will do.
"""
import copy

from pathlib import Path

import yaml


def _title(t, old, new):
    """(new title, part matched: whole | prefix | label) or (None, None)."""
    if t == old:
        return new, "whole"
    if ": " in t:
        p, label = t.split(": ", 1)
        if p == old:
            return f"{new}: {label}", "prefix"
        if label == old:
            return f"{p}: {new}", "label"
    return None, None


def apply(cfg, text, pairs):
    """Copies of (deck config, deck text) with every pair applied in order, and per pair the places it changed:
    [(place, part)], part = whole | prefix | label for titles, "line" otherwise."""
    cfg, text = copy.deepcopy(cfg), copy.deepcopy(text)
    hits = []
    for pr in pairs:
        old, new, where = pr["old"], pr["new"], pr.get("where")
        h = []
        if where in (None, "titles"):
            for n, e in enumerate(cfg["images"], 2):
                t, part = _title(e["title"], old, new)
                if t is not None:
                    e["title"] = t
                    h.append((f"slide {n} title", part))
            for r in (text.get("contents") or {}).get("rows", []):
                if r["titles"] == old:
                    r["titles"] = new
                    h.append(("contents row titles", "prefix"))
        if where in (None, "lines"):
            for n, e in enumerate(cfg["images"], 2):
                if e.get("subtitle") == old:
                    e["subtitle"] = new
                    h.append((f"slide {n} subtitle", "line"))
            _walk(text, old, new, "", h)
        if not h:
            raise SystemExit(f"replacement matches no whole line: {old!r}")
        hits.append(h)
    return cfg, text, hits


def _walk(node, old, new, path, h):
    items = node.items() if isinstance(node, dict) else enumerate(node) if isinstance(node, list) else []
    for k, v in list(items):
        if k == "titles":
            continue
        if v == old:
            node[k] = new
            h.append((f"{path}{k}", "line"))
        else:
            _walk(v, old, new, f"{path}{k}.", h)


def _scalars(node):
    if isinstance(node, dict):
        for v in node.values():
            yield from _scalars(v)
    elif isinstance(node, list):
        for v in node:
            yield from _scalars(v)
    elif isinstance(node, str):
        yield node


def _yaml_str(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def _str_nodes(node):
    """The string value nodes of a composed YAML tree, in document order (mapping keys skipped)."""
    if isinstance(node, yaml.MappingNode):
        for _, v in node.value:
            yield from _str_nodes(v)
    elif isinstance(node, yaml.SequenceNode):
        for v in node.value:
            yield from _str_nodes(v)
    elif isinstance(node, yaml.ScalarNode) and node.tag == "tag:yaml.org,2002:str":
        yield node


def write_yaml(path, new_data):
    """Write new_data into a YAML file by replacing each changed string where it sits in the file's text (positions
    from the YAML parser), so comments and layout stay; fails unless the file then parses to exactly new_data."""
    path = Path(path)
    src = path.read_text(encoding="utf-8")
    nodes, news = list(_str_nodes(yaml.compose(src))), list(_scalars(new_data))
    if len(nodes) != len(news):
        raise SystemExit(f"{path}: structure changed; only strings may change")
    for node, n in reversed(list(zip(nodes, news))):
        if node.value != n:
            src = src[:node.start_mark.index] + _yaml_str(n) + src[node.end_mark.index:]
    if yaml.safe_load(src) != new_data:
        raise SystemExit(f"{path}: rewrite did not give the expected data (nothing written)")
    path.write_text(src, encoding="utf-8", newline="")


def approved_rows(pairs, hits, source):
    """approved_text.yaml rows (key, text, source) for the new lines: a new title prefix goes in title_prefixes, a
    new title label in title_labels, anything else in exact."""
    rows = []
    for pr, h in zip(pairs, hits):
        parts = {part for _, part in h}
        keys = ((["title_prefixes"] if "prefix" in parts else []) + (["title_labels"] if "label" in parts else [])
                + (["exact"] if parts & {"whole", "line"} else []))
        rows += [(k, pr["new"], source) for k in keys]
    return rows
