"""System appearance and theme detection (Dark mode, Accent colors, High contrast)."""

from __future__ import annotations

import sys
from platform_adapter.capabilities import CAPABILITIES


def detect_system_dark_mode() -> bool:
    """Detect if the operating system is currently using a dark theme."""
    if CAPABILITIES.is_windows:
        try:
            import winreg

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                return val == 0
        except Exception:
            return False
    elif CAPABILITIES.is_macos:
        try:
            import subprocess

            cmd = ["defaults", "read", "-g", "AppleInterfaceStyle"]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=1)
            return "Dark" in res.stdout
        except Exception:
            return False
    return False


def get_system_accent_color() -> str | None:
    """Return system accent color hex if readable on Windows/macOS, else None."""
    if CAPABILITIES.is_windows:
        try:
            import winreg

            key_path = r"Software\Microsoft\Windows\DWM"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                val, _ = winreg.QueryValueEx(key, "AccentColor")
                # Format is ABGR in 0xAABBGGRR
                r = val & 0xFF
                g = (val >> 8) & 0xFF
                b = (val >> 16) & 0xFF
                return f"#{r:02X}{g:02X}{b:02X}"
        except Exception:
            return None
    return None
