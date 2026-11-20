"""Apply FF's answers on a decision page to the deck's own sources (FF 2026-11-19: answers are general, applied to every
slide that carries the text; approved text gets FF's source so the build's text check never fails on his words).

Reads the page's db export (ArtifactData list, collections `decisions` and `slides`, out_dir = <export dir>) and the
decision spec. For every decision FF answered that is not yet applied:
- a pick of an option with `replace`: its replacements (text_map.apply, whole lines only);
- a pick of "Keep current" or of a `request` option: nothing changes in the deck; a request is listed for its owner;
- a saved change: when every option of that decision replaces the same one line, that line becomes FF's text;
  otherwise it is listed for a person to apply.
Saved slide edits (All slides view, keyed by the slide's image name): each changed line becomes a replacement of that
line. A slide edit that adds or removes lines is listed for a person.

Dry run by default: prints each replacement, the places it changes, and the approved-text rows. With --write it writes
the deck config, the deck text file, approved_text.yaml (source "FF <date> decision page <code>") and the applied list
(the page shows those as Resolved, the previews skip them). Changes nothing FF did not answer.

Usage:  python apply_answers.py <decision spec.yaml> <db export dir> [--write]
"""
import json
import re
import sys
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))
import text_map  # noqa: E402
from ppt_image_inserter.approved_text import add_approved  # noqa: E402

RANGE = re.compile(r"^\d+(?:–\d+)?\s{2,}")   # contents row "11–13  text": the range is computed, not text


def load(spec_path):
    spec_path = Path(spec_path)
    spec = yaml.safe_load(spec_path.read_text(encoding="utf-8"))
    d = spec["deck"]
    paths = {k: (spec_path.parent / d[k]).resolve() for k in ("config", "text", "approved")}
    paths["applied"] = (spec_path.parent / spec["applied"]).resolve()
    applied = yaml.safe_load(paths["applied"].read_text(encoding="utf-8")) or [] if paths["applied"].exists() else []
    return spec, paths, applied


def docs(export, coll):
    d = Path(export) / coll
    return {f.stem: json.loads(f.read_text(encoding="utf-8")) for f in sorted(d.glob("*.json"))} if d.is_dir() else {}


def answers(spec, applied, export):
    """(replacement pairs with their source, applied records, things for a person)"""
    done = {str(a.get("code") or "slide " + a.get("slide", "")): a for a in applied}
    pairs, records, manual = [], [], []
    dec = docs(export, spec.get("collection", "decisions"))
    for it in spec["items"]:
        code = str(it["code"])
        doc = dec.get(code)
        if code in done or not doc:
            continue
        st = doc.get("status") or ("pick" if doc.get("choice") else "")
        day = doc.get("updatedAt", "")[:10]
        src = f"FF {day} decision page {code}"
        rec = {"code": code, "title": it["title"], "at": doc.get("updatedAt", "")}
        if st == "pick":
            opt = next((o for o in it.get("options", []) if str(o["choice"]) == doc.get("choice")), None)
            if doc.get("choice") == "keep":
                records.append(dict(rec, text="Keep current", how="your pick"))
            elif opt is None:
                manual.append(f"decision {code}: picked {doc.get('choice')!r}, not an option in the spec")
            elif opt.get("replace"):
                pairs += [dict(p, source=src + f" (pick {opt['choice']})") for p in opt["replace"]]
                records.append(dict(rec, text=opt.get("text") or opt.get("label", ""), how="your pick"))
            else:
                manual.append(f"decision {code}: request for {opt.get('request', 'the data owner')}: "
                              f"{opt.get('text') or opt.get('label')}")
                records.append(dict(rec, text=opt.get("text") or opt.get("label", ""), how="your pick (request)"))
        elif st == "change":
            olds = {(p["old"], p.get("where")) for o in it.get("options", []) for p in o.get("replace", [])}
            if len(olds) == 1 and all(len(o.get("replace", [])) == 1 for o in it["options"] if o.get("replace")):
                old, where = olds.pop()
                pairs.append({"old": old, "new": doc["note"].strip(), "where": where, "source": src + " (his change)"})
                records.append(dict(rec, text=doc["note"].strip(), how="your change"))
            else:
                manual.append(f"decision {code}: change to apply by hand: {doc.get('note', '')!r}")
    for sid, doc in docs(export, spec.get("slide_collection", "slides")).items():
        if doc.get("status") != "change" or ("slide " + sid) in done:
            continue
        before, after = doc.get("before", "").split("\n"), doc.get("text", "").split("\n")
        if len(before) != len(after):
            manual.append(f"slide {sid}: edit adds or removes lines, apply by hand")
            continue
        src = f"FF {doc.get('updatedAt', '')[:10]} decision page, slide edit"
        for b, a in zip(before, after):
            if b.strip() != a.strip():
                pairs.append({"old": RANGE.sub("", b.strip()), "new": RANGE.sub("", a.strip()), "source": src})
        records.append({"slide": sid, "at": doc.get("updatedAt", ""), "how": "your slide edit"})
    return pairs, records, manual


def main(argv):
    spec, paths, applied = load(argv[0])
    write = "--write" in argv
    pairs, records, manual = answers(spec, applied, argv[1])
    cfg = yaml.safe_load(paths["config"].read_text(encoding="utf-8"))
    text = yaml.safe_load(paths["text"].read_text(encoding="utf-8"))
    cfg2, text2, hits = text_map.apply(cfg, text, pairs)
    rows = []
    for p, h in zip(pairs, hits):
        print(f"{p['old']!r}\n  -> {p['new']!r}\n  at: {', '.join(place for place, _ in h)}")
        rows += text_map.approved_rows([p], [h], p["source"])
    for m in manual:
        print("FOR A PERSON:", m)
    print(f"{len(pairs)} replacements, {len(records)} answers recorded, {len(manual)} for a person"
          + ("" if write else " (dry run: nothing written; add --write)"))
    if write and (pairs or records):
        text_map.write_yaml(paths["config"], cfg2)
        text_map.write_yaml(paths["text"], text2)
        print(f"approved text: {add_approved(paths['approved'], rows)} rows added")
        head = "".join(ln + "\n" for ln in (paths["applied"].read_text(encoding="utf-8").splitlines()
                                            if paths["applied"].exists() else []) if ln.startswith("#"))
        paths["applied"].write_text(head + yaml.safe_dump(applied + records, allow_unicode=True, sort_keys=False,
                                                          width=200), encoding="utf-8")
    return 1 if manual else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
