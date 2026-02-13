"""
Backup system for PowerPoint presentations.
"""

import os
import shutil
import datetime
import glob
from typing import Dict


def backup_presentation(ppt_path: str, backup_base: str = 'PPT/backups') -> Dict[str, str]:
    """
    Create backups in multiple time-interval categories (one per category).

    This function implements a smart backup strategy that maintains ONE backup
    per time interval category. When a backup is created, it overwrites any
    existing backup in that category. Backups are only created if the time
    threshold has been exceeded.

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

    for category, threshold_seconds in intervals.items():
        # Create category directory
        category_dir = os.path.join(backup_base, category)
        os.makedirs(category_dir, exist_ok=True)

        # Check most recent backup in this category
        existing_backups = glob.glob(os.path.join(category_dir, '*.pptx'))
        should_backup = True

        if existing_backups and threshold_seconds > 0:
            # Get most recent backup by modification time
            latest_backup = max(existing_backups, key=os.path.getmtime)
            backup_time = datetime.datetime.fromtimestamp(os.path.getmtime(latest_backup))
            time_diff = (timestamp - backup_time).total_seconds()

            # Only backup if threshold exceeded
            should_backup = time_diff >= threshold_seconds

        if should_backup:
            # Use original filename (will overwrite existing backup)
            filename = os.path.basename(ppt_path)
            backup_path = os.path.join(category_dir, filename)

            # Copy the file (preserving metadata)
            shutil.copy2(ppt_path, backup_path)
            created_backups[category] = backup_path

    return created_backups
