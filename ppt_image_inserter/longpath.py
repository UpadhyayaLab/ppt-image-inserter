"""
Windows extended-length path helpers.

On Windows, paths longer than 260 characters (MAX_PATH) cause os.path.exists(),
open(), and python-pptx's add_picture() to silently fail or raise misleading
FileNotFoundError. The fix is the \\?\ extended-length prefix, which lifts the
limit to ~32 767 characters.

Usage::

    from ppt_image_inserter import safe_path, path_exists

    # Before passing a long path to add_picture / Image.open / open():
    pic = slide.shapes.add_picture(safe_path(image_path), ...)

    # Instead of Path(p).exists() or os.path.exists(p):
    if path_exists(p):
        ...
"""

import os


def safe_path(p):
    """Return *p* in Windows extended-length (``\\\\?\\``) form.

    Resolves the path to an absolute path first (so drive-relative and
    ``./``-relative paths work), then prepends the ``\\\\?\\`` prefix on
    Windows if it is not already present.  On non-Windows platforms the
    path is returned unchanged (as a ``str``).

    Parameters
    ----------
    p : str or os.PathLike
        File or directory path (may use forward slashes).

    Returns
    -------
    str
        The extended-length path string.
    """
    ap = os.path.abspath(str(p))
    if os.name == "nt" and not ap.startswith("\\\\?\\"):
        ap = "\\\\?\\" + ap
    return ap


def path_exists(p):
    """Like ``os.path.exists`` but safe for paths exceeding MAX_PATH on Windows."""
    return os.path.exists(safe_path(p))
