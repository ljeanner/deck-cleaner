"""
file_utils.py
-------------
Utility helpers for file handling: unique filename generation,
temporary directory management, and safe path operations.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import uuid
from pathlib import Path


def generate_unique_filename(original_name: str, suffix: str = "_optimized") -> str:
    """
    Generate a unique output filename derived from the original name.

    Example::

        generate_unique_filename("deck.pptx")
        # → "deck_optimized_3f2a1b.pptx"
    """
    stem = Path(original_name).stem
    short_id = uuid.uuid4().hex[:6]
    return f"{stem}{suffix}_{short_id}.pptx"


def make_temp_dir() -> str:
    """Create and return the path of a fresh temporary directory."""
    return tempfile.mkdtemp(prefix="deck_cleaner_")


def remove_temp_dir(path: str) -> None:
    """Remove a temporary directory and all its contents (best-effort)."""
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


def ensure_dir(path: str | Path) -> None:
    """Create *path* (and any missing parents) if it does not exist."""
    os.makedirs(path, exist_ok=True)


def file_size_bytes(path: str | Path) -> int:
    """Return the size of *path* in bytes."""
    return os.path.getsize(path)


def safe_filename(name: str) -> str:
    """
    Strip any directory components from *name* and replace characters
    that are unsafe in filenames.  This guards against path-traversal
    when using user-supplied filenames.
    """
    # Remove any directory separator the client might have injected
    name = os.path.basename(name)
    # Replace characters that are not alphanumeric, dot, hyphen, or underscore
    safe = "".join(c if c.isalnum() or c in "._- " else "_" for c in name)
    return safe.strip() or "upload.pptx"
