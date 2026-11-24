# PPT Batch Image Inserter

A Python toolkit for batch-inserting images into PowerPoint presentations. Designed for researchers who need to create presentation decks with large numbers of analysis plots following a consistent template.

## Quick Start

### Installation

```bash
# Clone and navigate
git clone https://github.com/UpadhyayaLab/ppt-image-inserter.git
cd ppt-image-inserter

# Create environment
conda create -n ppt_inserter python=3.9
conda activate ppt_inserter

# Install dependencies
pip install -r requirements.txt
```

### Basic Usage

1. **Copy and edit an example config**:
   ```cmd
   copy examples_and_configs\configs\example_config.yaml my_config.yaml
   ```
   Or for the scientific microscopy example:
   ```cmd
   copy examples_and_configs\configs\example_config_PLL_aCD3_vim_fixed_cells.yaml my_config.yaml
   ```
   Then edit `my_config.yaml` with your presentation path, image directory, and image list.

2. **Run the batch script**:
   ```bash
   python examples_and_configs/batch_insert_images.py my_config.yaml
   ```

## How It Works

1. **Template slide**: Create a slide with your desired layout and one or more placeholder images. Specify which slide to use as the template in your config (0-indexed: slide 2 in PowerPoint = index 1)
2. **Config file**: List all images you want to insert and which slides to preserve
3. **Batch process**: Script copies the template slide for each image set, replacing the placeholder(s)
4. **Result**: Presentation with consistent formatting across all slides

## Prerequisites

- Python 3.9+
- PowerPoint files (.pptx format)
- Images in common formats (PNG, JPG, TIFF, GIF, BMP)
- Windows OS

**Important**: Close the file in PowerPoint before running the script. The script needs exclusive access to modify the file.

## Key Features

- **Template-based** - Consistent formatting across all slides
- **Multi-image per slide** - Side-by-side comparisons using list syntax (e.g., `[control.png, treated.png]`)
- **Config-driven** - YAML configuration for easy batch processing
- **Template preservation** - Use `output_path` to create a new file, leaving the original untouched
- **Approved slide text** - Set `approved_text: <file>` in the config and the build lists every slide line not in that file (reported, never blocking) (each line carries its source). `examples_and_configs/check_approved_text.py` checks a deck or seeds the file from a finished one. See CLAUDE.md, "Approved slide text".
- **Per-slide title, subtitle and bottom line** - In the dict form of an entry, `title:` and `subtitle:` go at the top. `bottom_text: "..."` adds one line at the bottom of the slide, for example an experiment's dates, which stay out of titles. Its style is set in the config: `bottom_text_size` (pt, default 16), `bottom_text_italic` (default false), `bottom_text_align` (`left` (default), `center` or `right`) and `bottom_text_margin` (inches from the slide edges, default 0.15). The figure is never moved; `measure_title_lines.py` reports a bottom line whose letters touch the figure.
- **Text never on the data** - After every build, `batch_insert_images.py` runs `ppt_image_inserter.text_clear` (Windows, PowerPoint). Each slide's title and bottom line are measured as rendered; if a non-white figure pixel lies inside the box around their letters, the pictures are moved, and scaled only if they must be (aspect kept), into the band between the two texts. Text is never moved or shrunk. Every move is printed as `[TEXT CLEAR]`, a scale below 0.85 is flagged, and any overlap left is printed as `[TEXT]`. `text_clear: false` in the config turns it off. `examples_and_configs/test_text_clear.py` builds a slide that overlaps and shows the check failing before the fix and passing after it.
- **Incremental rebuilds of movie decks** - Change a few slides of a big movie deck without rebuilding it: see "Fast rebuilds of movie decks" below.
- **Stage times** - Builders that use `ppt_image_inserter.stage_timer.StageTimer` print one `STAGE` line per stage and a `STAGE TIMES` table, and flag a stage far over its expected time.
- **Metadata tracking** - Displays image filenames on slides and stores paths in alt text
- **Automatic backups** - Creates timestamped backups before modifications
- **Image pre-validation** - Checks all images exist before making any changes

## Fast rebuilds of movie decks

A full build of a movie deck (`ppt_image_inserter.movies.build_movie_deck_via_com`) embeds every movie again. For the
119-slide TGT movie-pairs deck (398 movies, 481 MB) that took 11 to 18 min (2026-11-23), even when one text slide
changed; 5.8 min after the same day's speed-ups (posters and movie checks read in 12 threads, one rewrite of the file
instead of three), of which PowerPoint itself takes 4.6 min. `ppt_image_inserter.incremental.update_deck` rebuilds only
what changed:

1. Each slide's spec is hashed: text, positions and formatting, each movie file (path, size, modification time), each
   picture's bytes, the slide size, and the drawing code itself.
2. The hashes of the deck's current slides come from the sidecar `<deck>.slides.json`, which every full build writes
   next to the deck. The sidecar holds the deck's md5, so a deck edited by hand or restored from a backup is refused.
3. Only slides with a new hash are built, through the same PowerPoint code as a full build, into a small temporary
   deck, and spliced into a copy of the existing deck in the new order. Slides no longer wanted are dropped.
4. A check proves nothing else changed. Every kept slide's content hash (its XML, with each link replaced by the md5 of
   the movie, picture or poster it points to) must equal the old deck's slide; every new slide's must equal its fresh
   build.

Text edits, new values or notes on movie slides, and slides added, removed or moved each cost only the slides that
changed. A later full build gives the same slides. Measured on a copy of the TGT movie-pairs deck (2026-11-23):

| Change | Slides built | Time |
|---|---|---|
| None (splice and save only) | 0 | 40 s |
| Two text slides merged into one, 115 later slides shifted | 1 | 43 s |
| A new note on 20 movie slides (80 movies) | 20 | 88 s |

About 20 s of each is writing the 481 MB file to K:. In each case the result was checked against the deck it came
from, and PowerPoint opened it.

**TGT decks** (`examples_and_configs/tgt/build_cat_fmc_slides.py`):

```bash
# incremental: only changed slides; the live deck is not touched
python examples_and_configs/tgt/build_cat_fmc_slides.py --config <deck config> --pairs --from <live deck> --out <live deck>_building.pptx
# after the checks pass: back up the live deck to backups/, then swap the new deck and its sidecar in (stops if the deck is open in PowerPoint)
python -m ppt_image_inserter.incremental swap <live deck>_building.pptx <live deck>
# a deck built before 2026-11-23 has no sidecar: check it against its config, and write the sidecar if it matches
python examples_and_configs/tgt/build_cat_fmc_slides.py --config <deck config> --pairs --seed <live deck>
```

**Other movie builders** can use it the same way: compute the deck's `SlideSpec` list as for a full build, then call
`update_deck(specs, old_deck, out, slide_w, slide_h, loop=..., prepare=<writes the posters of the slides to build>)`
instead of `build_movie_deck_via_com`. After a full build, call `write_sidecar(out, [spec_hash(s, ...) for s in specs])`.

When it does not apply: a deck with no sidecar, or one changed since (seed it or build it in full once); a changed
slide size, or a change to the drawing code, changes every hash, so every slide is rebuilt. Notes pages are not copied.

## Best Practices

### Template Design
- Create a template slide with your desired layout and placeholder image(s)
- Specify the template slide index in your config file (0-indexed)
- For single-image slides: one placeholder; position is auto-detected
- For multi-image slides (e.g., side-by-side): add N placeholder images; positions are auto-detected from each placeholder in turn
- Use `preserve_slides` in config to specify which slides to keep (default: title slide and template)

### File Organization
```
project/
├── examples_and_configs/
│   ├── batch_insert_images.py      # Generic batch script
│   ├── example_multi_image.py      # Multi-image API example
│   └── configs/
│       ├── example_config.yaml     # Basic example
│       └── my_experiment.yaml      # Your research configs
├── presentations/
│   └── analysis_results.pptx
└── data/
    └── analysis/
        ├── control_result.png
        ├── treated_result.png
        └── ...
```

## Troubleshooting

**File not found errors?**
- Check `base_dir` path is correct
- Verify filenames match exactly (case-sensitive)
- Use forward slashes `/` in paths

**Presentation corrupted or slides lost**
- Check backup files in the `backups/` subfolder next to your presentation
- **Close the file in PowerPoint before running the script** - the file must not be open
- Verify write permissions

**Images not in the right position?**
- Check that the template slide has the correct number of placeholder images
- If images appear at the wrong placeholder, reorder them in your config list to match

## Documentation

- [Detailed Usage Guide](docs/DETAILED_USAGE.md) - Complete function reference and examples
- [Examples and Configs](examples_and_configs/) - Sample configs and scripts

## Support

- **Issues**: [GitHub Issues](https://github.com/UpadhyayaLab/ppt-image-inserter/issues)
- **Discussions**: [GitHub Discussions](https://github.com/UpadhyayaLab/ppt-image-inserter/discussions)
