"""Check a deck's slide text against its approved-text file, or seed that file from a finished deck.

    python examples_and_configs/check_approved_text.py check <deck.pptx> <approved.yaml>
    python examples_and_configs/check_approved_text.py seed  <deck.pptx> <approved.yaml> "<source>"

See ppt_image_inserter/approved_text.py. `seed` is for text already settled on a finished deck (source e.g.
"settled before 2026-11-19, not re-checked"); new wording goes in by hand, each line with its source.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from ppt_image_inserter.approved_text import _main  # noqa: E402

if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
