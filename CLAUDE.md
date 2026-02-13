# Claude Code Instructions for PPT Batch Image Inserter

## Project Overview

This is a Python toolkit for programmatically inserting images into PowerPoint presentations at scale. The main use case is batch-generating presentation slides from large collections of images (e.g., scientific plots, data visualizations) using a template-based approach.

**Core value proposition**: Instead of manually inserting hundreds of images into PowerPoint, users define a YAML config and run a script.

## Platform & Environment

- **Platform**: Cross-platform (Windows, macOS, Linux)
  - Note: Most users are on Windows
  - Avoid Unix-specific commands (e.g., `/dev/null` redirects)
- **Python**: 3.7+
- **Main dependency**: `python-pptx`
- **Package manager**: conda/conda (recommended) or pip

## Architecture

### Package Structure: `ppt_image_inserter/`

The package is organized into modular components. All functions can be imported directly from the package:

```python
from ppt_image_inserter import insert_image, copy_slide_replace_image
```

**Core Principles:**
- **General-purpose**: No experiment-specific logic
- **Well-documented**: Clear docstrings for all public functions
- **Backwards-compatible**: Don't break existing APIs without good reason
- **Standalone**: Minimal external dependencies (just python-pptx, PyYAML)

**Module Organization:**

- **`core.py`** - Basic image insertion functions
  - `insert_image()` - Insert image at exact position on existing slide
  - `insert_image_preserve_aspect()` - Insert image with aspect ratio preserved
  - `list_slides()` - Get slide information

- **`position.py`** - Position and unit conversion utilities
  - `get_image_position()` - Extract position from template image (auto-detection)
  - `cm_to_inches()` - Unit conversion utility

- **`slide_utils.py`** - Slide manipulation utilities
  - `duplicate_slide()` - Copy a slide
  - `remove_pictures_from_slide()` - Remove all pictures from slide
  - `remove_all_text_from_slide()` - Remove all text elements
  - `delete_slide()` - Delete slide with automatic backup

- **`workflow.py`** - High-level workflow functions
  - `copy_slide_replace_image()` - Main workflow for batch processing (duplicate template, replace image)
  - `replace_image_on_existing_slide()` - Update existing slide with new image

- **`backup.py`** - Backup system
  - `backup_presentation()` - Create timestamped backups with multiple time intervals

- **`metadata.py`** - Metadata extraction
  - `extract_image_metadata()` - Extract image source info from all slides

### User-Facing Components

1. **Config files (YAML)**: Define presentation path, template slide, image list
2. **Scripts**: User-written Python scripts that use the core module
3. **Examples**: Template configs and scripts in `examples/`

## Important Constraints

### PowerPoint File Handling

**CRITICAL**: This tool is for IMAGE INSERTION ONLY. You are NOT allowed to:
- ❌ Edit existing slide content (text, shapes, formatting)
- ❌ Modify existing elements beyond images
- ❌ Rearrange slides
- ❌ Change slide layouts or templates (beyond copying them)
- ❌ Modify any other aspects of the presentation

**You ARE allowed to:**
- ✅ Create new slides by copying templates
- ✅ Insert/replace images in slides
- ✅ Delete slides (with backup)
- ✅ Store metadata in slide notes
- ✅ Add text labels (image filenames)

**Rationale**: Users manage their PowerPoint files manually. This tool just handles the tedious batch image insertion. Overstepping these boundaries creates confusion and potential data loss.

### Backup System

- **ALWAYS** create backups before destructive operations (slide deletion, image replacement)
- Backups are stored in `.backups/` with timestamps
- Multiple backup tiers: latest, 5min, 10min, 30min, hourly, daily
- Users trust the backup system - don't skip it to "save time"

### Error Handling

- **Graceful degradation**: If one image fails, continue processing others
- **Clear error messages**: Tell users WHAT failed and WHY
- **Path validation**: Check that image files exist before processing
- **Non-zero exit codes**: Scripts should exit with error code if any failures occur

## Development Guidelines

### Code Style

- **PEP 8 compliance**: Follow Python style guidelines
- **Type hints**: Use type hints for function signatures (Python 3.7+ compatible)
- **Docstrings**: Google-style docstrings for all public functions
- **Clear variable names**: `template_slide_index` not `tsi`

### Testing Philosophy

- **Real-world testing**: Test with actual PowerPoint files, not mocks
- **Cross-platform**: Consider Windows path handling (backslashes)
- **Large batches**: Test with 100+ images to catch performance issues
- **Edge cases**: Empty presentations, slides without images, corrupt files

### Documentation

- **Keep README.md updated**: Any new feature should be documented
- **Example-driven**: Show concrete examples, not just API docs
- **Beginner-friendly**: Assume users know Python basics but not python-pptx

## Common User Workflows

### Workflow 1: Scientific Data Visualization

**Scenario**: Researcher has 50 analysis plots from an experiment, needs them in a presentation

**User approach**:
1. Creates PowerPoint with title slide and template slide
2. Manually places one image in template slide (slide 2) at desired position
3. Creates YAML config listing all 50 images
4. Runs script to delete old slides and regenerate from config
5. Tweaks template, re-runs script until satisfied

**Claude's role**:
- Help create/modify YAML configs
- Write/debug batch processing scripts
- Troubleshoot image position issues
- Optimize performance for large batches

### Workflow 2: Report Automation

**Scenario**: Weekly/monthly reports with updated charts

**User approach**:
1. Sets up template presentation
2. Creates script that fetches latest data, generates plots, inserts into PPT
3. Runs script automatically (cron/scheduled task)

**Claude's role**:
- Build end-to-end automation scripts
- Handle error cases (missing data, corrupted files)
- Set up logging and notifications

### Workflow 3: A/B Testing Presentations

**Scenario**: Multiple versions of a presentation with different images

**User approach**:
1. Creates different configs for different image sets
2. Generates multiple presentation variants
3. Compares side-by-side

**Claude's role**:
- Create config templates
- Build batch generation scripts
- Organize output files

## Anti-Patterns to Avoid

### ❌ Don't Over-Engineer

**Bad**:
```python
class PresentationImageInserterFactoryBuilder:
    def __init__(self):
        self.strategies = []
        self.observers = []
    # ... 500 lines of abstraction
```

**Good**:
```python
def insert_image(ppt_path, slide_index, image_path, left, top, width, height):
    """Insert image at specified position."""
    # ... simple, direct implementation
```

**Rationale**: This is a utility library. Keep it simple and focused.

### ❌ Don't Break Existing Configs

**Bad**: Changing YAML field names without backwards compatibility
```yaml
# Old config breaks after your change
presentation: "my.pptx"  # You renamed this to "ppt_file"
```

**Good**: Support both old and new field names, or provide clear migration guide

### ❌ Don't Assume Directory Structure

**Bad**:
```python
config_path = "configs/config.yaml"  # Assumes specific location
```

**Good**:
```python
config_path = os.path.join(os.path.dirname(__file__), "configs", "config.yaml")
# Or better: take config path as argument
```

### ❌ Don't Silently Fail

**Bad**:
```python
try:
    insert_image(...)
except:
    pass  # User has no idea what went wrong
```

**Good**:
```python
try:
    insert_image(...)
except FileNotFoundError as e:
    print(f"[ERROR] Image file not found: {e}")
    error_count += 1
```

## When Users Ask For Help

### Debugging Checklist

1. **Check paths**: Are all file paths correct? (absolute vs relative)
2. **Check indices**: Slide indices are 0-based (slide 1 = index 0)
3. **Check template**: Does template slide have exactly one image?
4. **Check permissions**: Can we write to the PowerPoint file?
5. **Check PowerPoint version**: Is it .pptx (not .ppt)?

### Common Questions & Answers

**Q: "Images aren't in the right position"**
- A: Check that `auto_position: true` and template slide has one image where you want new images

**Q: "Script is slow with 200 images"**
- A: This is normal - PowerPoint files are complex XML. Consider batching.

**Q: "Can I insert multiple images per slide?"**
- A: Not directly supported. Workaround: create multiple template slides or use direct API.

**Q: "Can I modify text in slides?"**
- A: No - outside scope of this tool. Use python-pptx directly for that.

**Q: "Does it work with .ppt files?"**
- A: No - only .pptx (Office Open XML format). Convert old files first.

## Contributing Guidelines for Claude

When helping users extend this tool:

### New Features

1. **Start with user story**: Why do they need this?
2. **Check if possible**: Can python-pptx do this?
3. **Maintain simplicity**: Does it fit the core mission?
4. **Add tests**: Provide test case or example
5. **Update docs**: Add to README and examples

### Bug Fixes

1. **Reproduce first**: Understand the issue fully
2. **Minimal change**: Fix the bug, don't refactor unrelated code
3. **Add regression test**: Prevent it from happening again
4. **Document**: Note the fix in comments if non-obvious

### Performance Improvements

1. **Measure first**: Profile before optimizing
2. **Document tradeoffs**: Faster but more memory? Note it.
3. **Backwards compatible**: Don't break existing usage

## Example Interactions

### Good Interaction Example

**User**: "I need to add 100 images to a presentation. They're in /data/plots/"

**Claude**:
```
Let me help you set up a config and script for that.

First, let's create a config file:
```
[Creates clear, well-commented YAML config]
```

Now here's a script to process your images:
```
[Creates simple, robust script with error handling]
```

To run:
1. Make sure your PowerPoint has a template slide (slide 2) with one image
2. Save the config as config.yaml
3. Run: python insert_images.py

The script will:
- Delete old slides (keeping title & template)
- Create 100 new slides from your images
- Create backups before any changes

Questions?
```

### Bad Interaction Example

**User**: "I need to add 100 images to a presentation."

**Claude**:
```python
# Here's the code
from ppt import *
do_the_thing("ppt.pptx", "/data")
```

[No explanation, no config, no error handling, invalid code]

## Technical Gotchas

### Windows Path Handling

```python
# ✅ GOOD - works on all platforms
base_dir = "C:/Users/Name/Data"  # Forward slashes work on Windows!

# ⚠️ OK - but annoying
base_dir = "C:\\Users\\Name\\Data"  # Need double backslashes

# ❌ BAD - breaks on Windows
base_dir = r"C:\Users\Name\Data"  # Raw string, but backslash issues
```

### Slide Indices (0-based)

```python
# PowerPoint UI: Slide 1, Slide 2, Slide 3
# Python indices:     0,       1,       2

template_slide: 1  # This is "Slide 2" in PowerPoint UI
```

### Image Position Units

```python
# python-pptx uses inches (not pixels, not centimeters)
left = 1.5    # 1.5 inches from left edge
top = 2.0     # 2.0 inches from top edge
width = 6.0   # 6 inches wide
height = 4.0  # 4 inches tall

# Use cm_to_inches() if needed
from ppt_image_inserter import cm_to_inches
width = cm_to_inches(15.24)  # 15.24 cm = 6 inches
```

### PowerPoint Object Model

```python
from pptx import Presentation

prs = Presentation("file.pptx")
slide = prs.slides[0]  # First slide

# Slides have:
# - shapes (text boxes, images, etc.)
# - notes
# - slide_layout (template)
# - placeholders

# Find images:
for shape in slide.shapes:
    if hasattr(shape, "image"):
        # It's an image!
        print(shape.left, shape.top, shape.width, shape.height)
```

## Testing Checklist

Before suggesting code to users, verify:

- [ ] **Paths**: Are all paths valid for user's OS?
- [ ] **Error handling**: What if image doesn't exist?
- [ ] **Indices**: Are slide indices correct (0-based)?
- [ ] **Dependencies**: Only python-pptx and PyYAML?
- [ ] **Backups**: Are destructive operations backed up?
- [ ] **Output**: Clear success/failure messages?
- [ ] **Documentation**: Is the code commented?

## Resources

- **python-pptx docs**: https://python-pptx.readthedocs.io/
- **PowerPoint Open XML format**: Understanding helps debug issues
- **Example configs**: See `examples/` directory

## Philosophy

> "Automate the tedious, empower the creative"

This tool exists because manually inserting hundreds of images into PowerPoint is:
- Tedious and error-prone
- Time-consuming
- Not creative work

Good contributions make batch operations:
- Easier to set up
- More reliable
- Faster to execute
- Clearer when things go wrong

Bad contributions:
- Add complexity without clear benefit
- Break existing workflows
- Assume specific use cases
- Ignore cross-platform concerns

---

**Remember**: Users are researchers, data scientists, analysts - not software engineers. Keep things simple, clear, and robust. When in doubt, ask clarifying questions rather than making assumptions.
