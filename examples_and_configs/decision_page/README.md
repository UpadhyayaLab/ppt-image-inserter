# Decision pages: full slides, one-click answers

(FF 2026-11-19: "I need to see everything in context. you need to show full slide for each item!"; "I should not be
expected to go back there - and type like 10 things"; "clean up the code here dramatically - it shouldn't be this hard".)

Any decision about slide content goes to FF as one page. Each option shows every slide it changes, rendered by the
deck's own pipeline; FF answers by clicking, and answers save to the page's database. Look and behavior follow the
BIAD2255 approval page (green pick, amber change).

## The one idea: an answer is a line replacement

`text_map.py` holds the only rule. An option is a list of replacements `{old, new, where}`; each matches **whole lines
only**: a slide title (whole, its prefix before ": ", or its label after it), a subtitle, or a line in the deck's text
file (contents rows and notes, table headers and keys). `where: titles | lines` limits it; a replacement that matches
nothing fails. The same function makes the previews and applies FF's answers, so a preview shows exactly what applying
it will do. An option that needs a re-render by the data owner is a `request` instead: the page shows the current
slides and says who must re-render it; the answer is recorded for that owner.

A deck that uses this keeps every line drawn into its rendered images in one YAML text file (Fig 6:
`fig6_metrics/deck_text.yaml`), and its titles and subtitles in its config. Structure (which slides, their order,
title size) lives in the config, never on the page.

## Steps (one deck: Fig 6)

1. **Export the db** right before a build: ArtifactData `list`, collections `decisions` and `slides`,
   `out_dir` = the spec's `db_export`. The page build refuses an export older than 15 minutes, and refuses while any
   saved answer in it is not yet applied (step 2), so no answer FF saved can be left off the page.
2. **Apply answers**: `python decision_page/apply_answers.py <spec> <db export>` (dry run: prints every replacement,
   where it lands and the approved-text rows), then again with `--write`. It writes the deck config, the deck text
   file, `approved_text.yaml` (source "FF <date> decision page <code>") and the spec's applied list (shown as
   Resolved). It changes nothing FF did not answer and lists what needs a person (a change that is not one line, a
   slide edit that adds lines, a request).
3. **Previews**: `python fig6_metrics/fig6_decision_previews.py <spec>`. Two decks, each built, exported and measured
   once: the deck as it is (All slides view; built with the approved-text check on) and one options deck (only the
   slides each open option changes, one copy per option). Each slide's text comes from the sources, and each title's
   line count at the config's title size (32 pt) from PowerPoint (`measure_title_lines.py`). Writes `page/page_data.json` and the PNGs; prints
   the time per step.
4. **figure-reviewer** checks the slides before FF sees anything; its verdict blocks.
5. **Build and test**: `python decision_page/build_decision_page.py <spec> <out.html> --test-dir <data>/tests`.
6. **Publish** with the Artifact tool to the same file path or `url` (capabilities `{db: {}}` on first publish).

## Spec

```yaml
title / heading / lead / all_heading / all_lead
collection: decisions            # db collections
slide_collection: slides
deck: {config: ..., text: ..., approved: ...}   # paths relative to the spec
applied: <spec>_applied.yaml     # answers already applied (apply_answers.py writes it)
db_export: K:/.../db_export
page_data: K:/.../page/page_data.json
png_width: 1100
items:
  - code: "12a"                  # one db document per code
    title: "..."
    source: "..."                # for the record, not shown
    about: "..." | ["...", ...]  # optional paragraph or bullets
    current: "..."               # optional: the current text, quoted on "Keep current" (default: the one line all
                                 # options replace); keep_label: "..." instead; keep: false drops "Keep current"
    slides: [image name, ...]    # optional: deck slides to show as "Now" (for request-only items)
    options:
      - choice: "use"
        button: "Use this"       # default: choice
        text: "..."              # exact string, quoted   (or label: "..." unquoted)
        note: "..."              # needs note_source
        note_source: "..."
        wording_source: "..."    # shown as "Source: ..."; set it when the text is not copied from the data owner's reply
        replace: [{old: "...", new: "...", where: titles}]   # previewed and applied automatically
        request: "Histone mark data"                         # or: needs a re-render by this owner
        not_rendered: "..."                                  # optional wording for a request
  - code: "seen"                 # no options: an info note
    title: "..."
    about: [...]
```

## What FF sees

- Two tabs: **All slides** (opens first): every deck slide in order, its text open in a box to edit and save
  ("Save slide change", amber when saved). **Decisions**: open items first, then a **Resolved** list.
- Each open decision: its slides ("Applies to slides 2, 35–37", linked), answer buttons ("Keep current" and the
  options, with "Edit this text" under quoted ones), a "Now" column and one column per option with its full slides
  and measured title lines (red at 3+ lines or a title running into the figure), and a change box. Under every change
  box: "Applies to every slide like this one."
- A pick or a saved change marks the item "Resolved: <final text> (your pick | your change)", collapsed to one line
  with links to its slides; clicking it reopens the card. An item just answered stays open until closed or reloaded.
- Counter: "N open, M resolved; K slides changed".

## Checks (each fails the build or test)

- Build: a replacement that matches no whole line; an open option with neither a preview nor a request; a note
  without `note_source`; a slide name in `slides` that is not in the deck; a resolved text on no deck slide; a middle
  dot, a micro sign (use μ, U+03BC), or a file name in page text; page over 15 MB.
- Previews: an option that changes no slide; a contents row that is not one unbroken block of slides; a measured
  title that differs from the configured one; text cut off at the bottom of the contents image; the deck's
  approved-text check.
- Test (`test_decision_page.py`, db replaced by a stand-in preloaded from the export): every item with a saved answer
  shows as Resolved; a pick, change and reload on the first open decision (it moves to Resolved with the change
  text, and reopens); a slide edit saved, reloaded and removed; the page opens on All slides. Screenshots of the three
  views at 1400 and 400 px, taken in parallel, and the page's whole text in `page_text.txt`.

Before publishing, have a reviewer who has not seen your reasoning check the shots, `page_text.txt` and the slide PNGs
beside `page_data.json` (deck_NN.png, opt_<code>_<choice>_NN.png) against the deck sources.
Not `publish-figures/gallery.py`: it has no answer controls.
