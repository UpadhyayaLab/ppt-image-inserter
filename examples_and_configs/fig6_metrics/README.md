# Fig 6 metrics deck

The deck shows histone marks and RNAP II at nuclear invaginations in T cells (Figure 6 panels C–F) for every
experiment. It contains:
- a title slide
- a contents slide
- 6 summary-table slides
- one figure per slide after that

## Rebuild: one command

```bash
python examples_and_configs/fig6_metrics/rebuild_fig6_deck.py examples_and_configs/configs/fixed/fig6_metrics/config_fig6_metrics_by_metric_20261118.yaml <page.html>
```

The command runs each step and prints its time. Measured on this PC on Nov 19, 2026, with 113 slides, the total was
56 s (52 s with 115 slides). Any step slower than 3× its measured time is flagged `[SLOW]`.

| Step | What it does | Time |
|---|---|---|
| tables | `render_fig6_table_slides.py`: the summary tables and the contents slide, rendered as HTML and screenshotted with headless Edge, into `table_png/` | 12 s |
| deck | `batch_insert_images.py <config>`: opens the deck once, adds every slide, saves once | 11 s |
| check | Deterministic. Each slide has 1 picture, its config title, and alt text naming its image file. The title slide text matches the config. No "·" anywhere. The contents-slide ranges match the slides. Every slide string is in `approved_text.yaml`. | 0.3 s |
| export | `export_deck_slides.py`: slide PNGs via PowerPoint COM (read-only; never closes FF's decks), into `slides_png/` | 29 s |
| page | A self-contained HTML page of all slides. Table slides stay exact; figure slides are reduced to 256 colors to fit the 16 MB page limit. | 4 s |

`--from deck|export|page` skips the earlier steps; the check always runs.

Two steps stay outside the command:
1. **figure-reviewer:** run it on the changed slides. Its verdict blocks showing the deck to FF.
2. **Publish:** republish the page to its existing URL with the Artifact tool.

## Where things are

| What | Where |
|---|---|
| Config: slide order and titles, title-slide text | `examples_and_configs/configs/fixed/fig6_metrics/config_fig6_metrics_by_metric_20261118.yaml` |
| Table values | Imported unchanged from the Histone mark data session's builder: `TCell-3D-Morphodynamics/dev/reporting/fig6_metrics_page/build_fig6_summary_table.py` (`collect()`, `cell()`) |
| Near-centrosome values (paper Fig 6C) | `<figs>/C_cent_mean_sem.csv` |
| Slides 3–4, defining the nadir | Figures copied from `Noco, Jurkats, 20261026 and 20261029.pptx` slides 49–50, in the output folder's `definition_png/` |
| Approved slide text | `approved_text.yaml`: every string a slide may carry (title slide, titles, subtitles, contents, table headers and keys), each with its source (FF's approval or a `fig6_definitions.yaml` term). The check fails on any slide string not listed. A new string goes in only with that source, never agent wording. |
| Contents-slide text and ranges | `CONTENTS` in `render_fig6_table_slides.py`. The check fails if a range no longer matches the slide titles. |
| Figures, `page_order.csv`, `manifest.json` | `Q:/FF/Nucleus_Data/3D_Nucleus/Fixed/fig6_metrics_page/figs_20261118/` |
| Outputs: deck, `build.log`, `table_png/`, `slides_png/` | `K:/FF/PPT/PPT_autogeneration/Fig6_metrics_across_experiments/` |
| Review page | https://claude.ai/artifact/GRgcKgBcwNTVNtvpj31krq |

## Who does what

- **Histone mark data** (sidebar title) owns the data: the figures, the table builder, and the centrosome CSV. Ask it
  about any number. Never edit its code from here.
- **Histone mark data agent presentation** (sidebar title) builds this deck and its page.

## Rules for this deck (from FF)

- The summary tables come first. They are kept simple: the relative level near the nadir, mark and DNA. Near-centrosome
  values (as in Fig 6C) appear only where the centrosome was imaged; elsewhere they are blank.
- No "·" anywhere (Nov 19, 2026). Use μ (U+03BC), not µ.
- No wall of text under a slide.
- Layout and wording are FF's decisions: ask in one line first.
