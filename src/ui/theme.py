"""Windows 11 Fluent Design tokens and theme management."""

from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Literal


@dataclass
class ThemeColors:
    window_bg: str
    card_bg: str
    card_secondary: str
    text_primary: str
    text_secondary: str
    text_placeholder: str
    border: str
    border_strong: str
    accent: str
    accent_hover: str
    accent_pressed: str
    accent_text: str
    button_bg: str
    button_hover: str
    button_pressed: str
    button_border: str
    entry_bg: str
    entry_border: str
    badge_bg: str
    error: str
    error_bg: str
    success: str
    success_bg: str
    warning: str
    warning_bg: str
    focus_border: str
    table_header_bg: str
    table_row_alt: str
    log_bg: str
    log_fg: str


LIGHT_THEME = ThemeColors(
    window_bg="#F3F3F3",
    card_bg="#FFFFFF",
    card_secondary="#F9F9F9",
    text_primary="#1A1A1A",
    text_secondary="#5D5D5D",
    text_placeholder="#8A8A8A",
    border="#E5E5E5",
    border_strong="#CCCCCC",
    accent="#005FB8",
    accent_hover="#004E98",
    accent_pressed="#003D78",
    accent_text="#FFFFFF",
    button_bg="#FAFAFA",
    button_hover="#EAEAEA",
    button_pressed="#DFDFDF",
    button_border="#D1D1D1",
    entry_bg="#FFFFFF",
    entry_border="#CECECE",
    badge_bg="#E8EDF2",
    error="#C42B1C",
    error_bg="#FDF3F2",
    success="#0F7B0F",
    success_bg="#F1F9F1",
    warning="#9D5D00",
    warning_bg="#FFF9E6",
    focus_border="#005FB8",
    table_header_bg="#F0F0F0",
    table_row_alt="#FAFAFA",
    log_bg="#FFFFFF",
    log_fg="#1E1E1E",
)

DARK_THEME = ThemeColors(
    window_bg="#202020",
    card_bg="#2B2B2B",
    card_secondary="#323232",
    text_primary="#FFFFFF",
    text_secondary="#C7C7C7",
    text_placeholder="#808080",
    border="#3E3E3E",
    border_strong="#555555",
    accent="#60CDFF",
    accent_hover="#4EB8EB",
    accent_pressed="#3DA3D6",
    accent_text="#000000",
    button_bg="#383838",
    button_hover="#454545",
    button_pressed="#2E2E2E",
    button_border="#4E4E4E",
    entry_bg="#202020",
    entry_border="#4E4E4E",
    badge_bg="#383838",
    error="#FF99A4",
    error_bg="#3D2022",
    success="#6CCB5F",
    success_bg="#1F3620",
    warning="#FCE100",
    warning_bg="#3B3615",
    focus_border="#60CDFF",
    table_header_bg="#333333",
    table_row_alt="#2F2F2F",
    log_bg="#181818",
    log_fg="#D4D4D4",
)


def get_font_family() -> str:
    """Return the ideal typography font family based on OS."""
    if sys.platform.startswith("win"):
        return "Segoe UI Variable, Segoe UI, Microsoft YaHei UI, sans-serif"
    elif sys.platform == "darwin":
        return ".AppleSystemUIFont, PingFang SC, Helvetica Neue, sans-serif"
    else:
        return "DejaVu Sans, Ubuntu, Noto Sans CJK SC, sans-serif"


def primary_font_name() -> str:
    """Return the single primary font family name for Tkinter."""
    if sys.platform.startswith("win"):
        return "Microsoft YaHei UI"
    elif sys.platform == "darwin":
        return "PingFang SC"
    else:
        return "Helvetica"


class ThemeManager:
    """Manages active theme state and change listeners."""

    def __init__(self, mode: Literal["light", "dark"] = "light"):
        self.mode = mode
        self.colors = LIGHT_THEME if mode == "light" else DARK_THEME
        self._font_name = primary_font_name()
        self._listeners = []

    @property
    def is_dark(self) -> bool:
        return self.mode == "dark"

    def toggle(self) -> str:
        self.set_mode("dark" if self.mode == "light" else "light")
        return self.mode

    def set_mode(self, mode: Literal["light", "dark"]) -> None:
        self.mode = mode
        self.colors = DARK_THEME if mode == "dark" else LIGHT_THEME
        for listener in self._listeners:
            listener(self.colors)

    def add_listener(self, callback) -> None:
        self._listeners.append(callback)

    def font(self, size: int = 10, weight: str = "normal", slant: str = "roman"):
        return (self._font_name, size, weight, slant)


# Global Theme Manager instance
THEME = ThemeManager("light")
