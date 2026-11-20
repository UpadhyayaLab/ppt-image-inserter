# Decision pages: full slides, one-click answers

(FF 2026-11-19: "I need to see everything in context. you need to show full slide for each item! establish this
workflow"; "I should not be expected to go back there - and type like 10 things".)

Any decision about slide content goes to FF as one page. Each option is shown as its full rendered slide, made by the
deck's own pipeline, and FF answers by clicking. Picks save to the page's database, so FF never types codes in chat.
Look and behavior follow the BIAD2255 approval page FF answered on (green pick, amber change).

## Rule: every slide shows the latest saved state

(FF 2026-11-19, on an option shown on the old contents slide: "then why is this showing old slide????")
- FF's saved slide changes and picks are the base for every option. A slide that a saved change has replaced appears
  nowhere on the page, not even labeled "before".
- An item FF's saved answers settle collapses to one line per answer ("Settled by your slide 2 text: ..."), with the
  slide as it now is (spec block `settled`). No old options, no old slides.
- Every option on a changed slide is rendered from FF's saved text plus that option's one change (example: item 12 of
  the Fig 6 page, `FIX12` in `render_decision_previews.py`).
- A slide that is in the deck is shown with its deck slide id and PNG everywhere, so a change saved on either view
  shows on both, and a deck slide carrying a saved change keeps the id the change was saved under.
- Enforced by the builder: export the db right before every build (ArtifactData `list`, collections `slides` and
  `decisions`, `out_dir` = spec `saved_state.dir`). Every slide id with a saved change must be listed in
  `saved_state.rebuilt` with the PNG rendered from that change, and every copy of it on the page must use that PNG;
  the build fails otherwise, and also on the words "before your saved change" anywhere on the page.

## Steps

1. **Render every option as a full slide** with the deck's own tools. Do not crop figures and do not show text alone.
   - Images that carry text (contents, tables, keys): use the deck's renderer with the option's string.
   - Titles and subtitles: put the image on the deck's template with `batch_insert_images.py`. Use a preview config
     with no `approved_text`, so the build prints a note instead of failing.
   - Export with `export_deck_slides.py`.
   - If an option needs a re-render that only the data owner can make, show the current slide and set
     `not_rendered: "Not rendered: needs a re-render by <owner>"`. Never fake it.
   - Write each slide's text to a `slide_text.json` beside the PNGs: `{slide id: {"lines": [...], "unlisted":
     "..."}}`. Take it from the built deck (`approved_text.deck_strings`) plus the text drawn into rendered images,
     and say in `unlisted` what is not listed (table numbers, labels inside figures). The page's per-slide change box
     starts from these lines.
   - Measured facts per slide go in `slide_facts.json`: `{slide id: {"text": "...", "warn": bool}}`, printed under
     each slide (red when `warn`). For titles use `examples_and_configs/measure_title_lines.py <deck.pptx> <out.json>`:
     PowerPoint's own line count for each title, and whether the title's letters run into the figure (compared in
     two in-memory exports, title only and pictures only). Never estimate line counts from characters.
   - Example: `fig6_metrics/render_decision_previews.py`. It writes `named/<id>.png`, `named/slide_text.json`,
     `named/slide_facts.json` ("Title: 2 lines at 36 pt") and `preview_ids.json` (the slide index).
   - **The whole deck, for the All slides view** (Histone mark data agent presentation, 2026-11-19: every slide of the
     deck in deck order, numbered as the deck will be, each with its text open to edit; the default view). Build the
     deck as it will be with FF's answers so far (its own config with the new slides inserted, text FF saved rendered
     by the deck's own renderer), export it, and write `deck_slides.json` (`[{id, png, slide}]`, slides 1..N in
     order) plus the deck slides' entries in `slide_text.json` and `slide_facts.json`. A slide that also appears on
     the decisions view keeps the same id (for example the contents slide), so a change saved on either view shows
     on both. Example: `render_decision_previews.py --from full` (79 slides: build 10 s, export 17 s, measure 31 s).
2. **Have figure-reviewer check the slides** before FF sees anything. Its verdict blocks.
3. **Write the spec** (format below), with every string exact and each item's source (kept in the spec, not shown:
   FF reads no file names).
4. **Build:** `python examples_and_configs/decision_page/build_decision_page.py <spec.yaml> <out.html>`
5. **Publish** with the Artifact tool, passing `capabilities: {db: {}}`.
   - To keep a URL, publish to the same file path or pass `url`.
   - Check that the store is empty before FF starts (ArtifactData `list`).
6. **Read the answers:** ArtifactData `list`, collection `decisions` (or the spec's `collection`). There is one
   document per decision code, with these fields:
   - `decision`: the decision code.
   - `status`: `pick` (an answer button was clicked last) or `change` (Save change was clicked last). A document with
     a `choice` and no `status` (written by the earlier page) counts as `pick`.
   - `choice`: the option picked (`""` if none).
   - `text`: the exact string of the picked option.
   - `note`: the text FF wrote in "Change: write what it should say".
   - `updatedAt`: when it was last changed.

   `status: change` means FF wants different text: use `note`, not `choice`.

   Slide changes: collection `slides` (or the spec's `slide_collection`), one document per slide id:
   - `slide`: the slide id.
   - `status`: `change` (FF saved new text) or `cleared` (FF removed the change).
   - `before`: the slide's text as the page showed it.
   - `text`: the text FF saved (all lines; compare with `before` to see what changed).

   Only the strings FF picked go into the deck's `approved_text` file.
7. **Apply each answer to every slide it covers** (FF 2026-11-19: "what I say there is general! like: it applies to
   many slides in some cases"). The page says so under every change box ("Applies to every slide like this one.").
   - Apply a pick, a change or a slide change to every slide that has the same text or is of the same kind (for
     example all near-centrosome violin slides, not only the one FF wrote on).
   - Before rebuilding, list the affected slides (deck, slide number, title) for each answer.

## Spec format

```yaml
title: "Fig 6 Deck Decisions"          # page <title>
heading: "Fig 6 deck: decisions"       # h1 (default: title)
lead: "..."                            # one paragraph under the heading
collection: decisions                  # db collection (default decisions)
png_dir: "K:/FF/PPT/.../named"         # relative PNG paths and slide ids (<id>.png) resolve here
png_width: 1280                        # embed width (page must stay under 16 MB)
slide_text: slide_text.json            # per-slide text (resolves like png_dir); every slide id must be in it
slide_collection: slides               # db collection for slide changes (default slides)
slide_facts: slide_facts.json          # optional measured line under every slide; every slide id must be in it
saved_state:                           # latest saved state (see the rule above)
  dir: "K:/.../db_export"              # ArtifactData list out_dir, exported right before the build
  rebuilt: {contents_base: deck_02.png} # slide id with a saved change -> PNG rendered from it
saved_change_notes: {slide id: "..."}  # optional: one plain line shown with that slide's saved change
all_slides:                            # optional: the All slides view, shown first and on open
  index: deck_slides.json              # [{id, png, slide}] in deck order, slides numbered 1..N
  png_width: 1100                      # embed width for these images (the page must stay under 15 MB)
  heading: "..."
  lead: "..."
  captions: {contents_base: "..."}     # optional caption per slide id
items:
  - number: "1"
    title: "..."
    source: "file.m:45-62"             # for the record; not shown on the page
    blocks:                            # in page order; each block is one of:
      - p: "text"
      - h3: "text"
      - note: "text"                   # small gray paragraph
      - ul: ["...", "..."]
      - now: {text: "...", where: "..."}            # current string, shown only
      - slides: [{id: x, caption: "..."}]           # context slides (or png: file.png)
      - settled: {text: "Settled by ...", slides: [...]}   # an answer FF already gave (lines: [...] for several)
      - decision:
          code: "1"                                 # one db document per code, unique on the page
          current: {text: "...", where: "...", slides: [...]}   # optional: adds a "Keep current" answer (choice keep)
                                                    # (label: "..." instead of text: shown unquoted)
          options:
            - choice: "1A"                          # stored as choice
              button: "1A"                          # big label on the answer button and in the top bar (default: choice)
              text: "..."                           # exact string, shown quoted on the button, stored as text
              label: "..."                          # instead of text: an action, shown unquoted, text stored ""
              note: "..."                           # small line on the button
              slides: [{id: ..., caption: ...}]     # the option's full slides, in its column below the buttons
              not_rendered: "Not rendered: ..."     # shown in amber at the top of its column
```

What the viewer sees:
- With `all_slides`: two tabs at the top, "All slides (N)" (open on arrival) and "Decisions (M)". All slides shows
  every deck slide in order, one per row and large, headed "Slide n", with its text already open in an edit box,
  "Save slide change", the applies line and the measured title line. Changes save to the same slides collection as
  the decisions view and show on every copy of that slide in both views. Any link on the page (summary chips, the
  Slides chip) opens the view that holds its target.
- At the top, a counter ("3 of 13 picked, 1 with changes, 9 not answered") and a sticky summary of every decision: green with the
  pick, amber "Change", or "–"; each entry jumps to its decision.
- One card per decision, with a row of large answer buttons, one per option plus "Keep current" where it applies,
  each carrying the option's text.
- Clicking a button turns the card green with "Your pick: <option>" at the top right; the picked button turns solid
  with a check mark and the others dim.
- Writing in "Change: write what it should say" and clicking "Save change" turns the card amber with "Change saved".
- Under every change box (decision and slide): "Applies to every slide like this one." (spec `change_scope_note`).
  Clicking an answer button afterwards makes it a pick again (the change text is kept in `note`).
- Under each quoted option, "Edit this text" copies the option's exact text into the card's change box, ready to
  edit and save.
- Below the buttons, each option's slides sit in a column of their own. Clicking a slide enlarges it.
- Under every slide, "Change this slide" opens a box with that slide's text, one line per line. FF edits any line
  and clicks "Save slide change": the slide gets an amber border and "Change saved", on every copy of that slide on
  the page, and the counter adds "N slides changed". "Back to the slide's text" restores the box; "Remove saved
  change" removes it.

The build fails on a middle dot or a micro sign (U+00B5), on a file name in the page text, on a missing PNG, or on a
repeated decision code. Each PNG is embedded once even when several options show it.

**Check before publishing:** add `--test-dir <data folder>/tests` to the build. With the database replaced by a
stand-in, it
- screenshots the page unanswered, answered (picks on the first 3 decisions, a saved change on the 4th, a saved
  change on the first slide), and answered scrolled to the changed card and to the changed slide, at desktop (1400 px) and phone (400 px) width, through an iframe (headless Edge will
  not open a window narrower than about 492 px, and does not paint a scrolled top-level page);
- writes one tall desktop shot per item to `review/item_<n>.png`, for reviewing every card and slide;
- runs one answer cycle on the first decision and the first slide (Edit this text, pick, change, re-pick, slide
  change, reload, remove slide change) and fails unless every save survives the reload and every copy of the slide
  shows its change;
- with `all_slides`, also: shots of the answered page on the Decisions tab and scrolled to deck slide 11, one tall
  shot per 3 deck slides (`review/all_<n>.png`), and in the cycle: the page opens on All slides, a deck row's box is
  open and holds that slide's text, a save from it marks every copy of the slide, and the Decisions tab switches view.

Look at the shots, and have a reviewer who has not seen your reasoning check them against FF's list, before
publishing.

Not `publish-figures/gallery.py`: that has no answer controls, and its subtitle uses a middle dot.
