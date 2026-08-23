"""Desktop integration for opening output folders and revealing files in Explorer/Finder."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path
from platform_adapter.capabilities import CAPABILITIES


def open_folder(path: str | os.PathLike) -> bool:
    """Open a folder in the native file manager (Explorer, Finder, or file manager)."""
    p = str(Path(path).resolve())
    if not os.path.exists(p):
        return False

    try:
        if CAPABILITIES.is_windows:
            os.startfile(p)  # type: ignore
            return True
        elif CAPABILITIES.is_macos:
            subprocess.Popen(["open", p])
            return True
        else:
            subprocess.Popen(["xdg-open", p])
            return True
    except Exception:
        return False


def reveal_in_file_manager(path: str | os.PathLike) -> bool:
    """Reveal a specific file selected in Explorer or Finder."""
    p = str(Path(path).resolve())
    if not os.path.exists(p):
        return False

    try:
        if CAPABILITIES.is_windows:
            subprocess.Popen(["explorer", "/select,", p])
            return True
        elif CAPABILITIES.is_macos:
            subprocess.Popen(["open", "-R", p])
            return True
        else:
            return open_folder(os.path.dirname(p))
    except Exception:
        return False
