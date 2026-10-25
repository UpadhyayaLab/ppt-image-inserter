"""
Backup system for PowerPoint presentations.
"""

import os
import re
import shutil
import datetime
from pathlib import Path
from typing import Dict


def subset_output_path(output_path, **subset) -> Path:
    """Output path for a SUBSET run (--cells / --since / --fovs / --only ...) when no
    --output was given: '<stem> (subset <name> <value>; ...)<suffix>' beside the full deck.

    Subset runs used to write the full deck's default path, so a quick '--cells 1,2'
    check replaced the whole deck (bug hunt 2026-10-24). Returns output_path unchanged
    when every subset value is empty.
    """
    output_path = Path(output_path)
    parts = [f"{k} {v}" for k, v in subset.items() if v not in (None, "", [], (), set())]
    if not parts:
        return output_path
    tag = re.sub(r"[^A-Za-z0-9,;._ -]+", "-", "; ".join(parts))[:60].strip()
    out = output_path.with_name(f"{output_path.stem} (subset {tag}){output_path.suffix}")
    print(f"Subset run -> {out} (the full deck {output_path.name} is left as is)")
    return out


def backup_presentation(ppt_path: str, backup_base: str = 'PPT/backups') -> Dict[str, str]:
    """
    Create backups in multiple time-interval categories (one per category).

    This function implements a smart backup strategy that maintains ONE backup
    per presentation in each time interval category. When a backup is created,
    it overwrites the existing backup for that presentation in that category.
    Interval backups are only created if the time threshold has been exceeded
    for that presentation.

    Args:
        ppt_path (str): Path to the PowerPoint file to backup
        backup_base (str): Base directory for backups (default: 'PPT/backups')

    Returns:
        dict: Dictionary mapping category names to backup file paths for backups created
              Example: {'latest': 'PPT/backups/latest/file.pptx'}

    Categories and time intervals:
        - 'latest': Always creates a backup (0 seconds threshold)
        - '5min': Creates backup if >5 minutes since last backup in this category
        - '10min': Creates backup if >10 minutes since last backup in this category
        - '30min': Creates backup if >30 minutes since last backup in this category
        - 'hourly': Creates backup if >1 hour since last backup in this category

    Directory structure:
        PPT/backups/
        ├── latest/
        ├── 5min/
        ├── 10min/
        ├── 30min/
        └── hourly/

    Example:
        >>> backups = backup_presentation('presentation.pptx')
        >>> print(f"Created backups: {list(backups.keys())}")
        Created backups: ['latest', '5min', '10min', '30min', 'hourly']
    """
    # Validate input file
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    # Time intervals in seconds
    intervals = {
        'latest': 0,      # Always backup
        '5min': 300,      # 5 minutes = 300 seconds
        '10min': 600,     # 10 minutes = 600 seconds
        '30min': 1800,    # 30 minutes = 1800 seconds
        'hourly': 3600    # 1 hour = 3600 seconds
    }

    created_backups = {}
    timestamp = datetime.datetime.now()
    filename = os.path.basename(ppt_path)

    for category, threshold_seconds in intervals.items():
        # Create category directory
        category_dir = os.path.join(backup_base, category)
        os.makedirs(category_dir, exist_ok=True)
        backup_path = os.path.join(category_dir, filename)

        # Check the existing backup for this presentation in this category
        existing_backups = [backup_path] if os.path.exists(backup_path) else []
        should_backup = True

        if existing_backups and threshold_seconds > 0:
            # Get most recent backup by modification time
            latest_backup = max(existing_backups, key=os.path.getmtime)
            backup_time = datetime.datetime.fromtimestamp(os.path.getmtime(latest_backup))
            time_diff = (timestamp - backup_time).total_seconds()

            # Only backup if threshold exceeded
            should_backup = time_diff >= threshold_seconds

        if should_backup:
            shutil.copy2(ppt_path, backup_path)
            # Stamp the backup with the time it was MADE. copy2 keeps the source's mtime,
            # so the tier test above measured the deck's content age instead: after two
            # quick reruns of a deck last saved >1 h ago every tier looked stale and all
            # were overwritten with the newest copy (bug hunt 2026-10-24).
            now_ts = timestamp.timestamp()
            os.utime(backup_path, (now_ts, now_ts))
            created_backups[category] = backup_path

    return created_backups
