"""Theme manager orchestrating Qt palettes, styles, and dynamic light/dark toggling."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Callable, Literal

from platform_adapter.appearance import detect_system_dark_mode, get_system_accent_color
from platform_adapter.capabilities import CAPABILITIES
from ui.theme.tokens import ColorTokens, DARK_TOKENS, LIGHT_TOKENS

try:
    from PySide6.QtGui import QColor, QPalette
    from PySide6.QtWidgets import QApplication

    HAS_QT = True
except ImportError:
    HAS_QT = False


class ThemeManager:
    """Manages active theme state, system color synchronization, and QSS compilation."""

    def __init__(self, mode: Literal["light", "dark", "system"] = "system"):
        self.mode_preference = mode
        self.is_dark = self._resolve_dark_mode(mode)
        self.tokens: ColorTokens = DARK_TOKENS if self.is_dark else LIGHT_TOKENS
        self._listeners: list[Callable[[ColorTokens], None]] = []

    def _resolve_dark_mode(self, mode: str) -> bool:
        if mode == "dark":
            return True
        elif mode == "light":
            return False
        return detect_system_dark_mode()

    def toggle(self) -> bool:
        self.is_dark = not self.is_dark
        self.mode_preference = "dark" if self.is_dark else "light"
        self.tokens = DARK_TOKENS if self.is_dark else LIGHT_TOKENS
        self.apply_theme()
        for cb in self._listeners:
            cb(self.tokens)
        return self.is_dark

    def check_system_theme_update(self) -> bool:
        """If mode_preference is 'system', check if OS appearance changed and reload if needed."""
        if self.mode_preference != "system":
            return False

        current_system_dark = detect_system_dark_mode()
        if current_system_dark != self.is_dark:
            self.is_dark = current_system_dark
            self.tokens = DARK_TOKENS if self.is_dark else LIGHT_TOKENS
            self.apply_theme()
            for cb in self._listeners:
                cb(self.tokens)
            return True
        return False

    def add_listener(self, callback: Callable[[ColorTokens], None]) -> None:
        self._listeners.append(callback)

    def compile_qss(self, tokens: ColorTokens | None = None) -> str:
        """Compile QSS template by replacing @token_name@ placeholders safely without string.format."""
        t = tokens or self.tokens

        # Choose appropriate QSS template based on OS
        qss_filename = "windows.qss" if CAPABILITIES.is_windows else "macos.qss"
        qss_path = Path(__file__).parent / qss_filename
        if not qss_path.exists():
            qss_path = Path(__file__).parent / "windows.qss"

        raw_qss = qss_path.read_text(encoding="utf-8")

        # Map all token fields
        replacements = {
            "@window_bg@": t.window_bg,
            "@card_bg@": t.card_bg,
            "@card_secondary@": t.card_secondary,
            "@text_primary@": t.text_primary,
            "@text_secondary@": t.text_secondary,
            "@text_placeholder@": t.text_placeholder,
            "@border@": t.border,
            "@border_subtle@": t.border_subtle,
            "@border_strong@": t.border_strong,
            "@accent@": t.accent,
            "@accent_hover@": t.accent_hover,
            "@accent_pressed@": t.accent_pressed,
            "@accent_text@": t.accent_text,
            "@button_bg@": t.button_bg,
            "@button_hover@": t.button_hover,
            "@button_pressed@": t.button_pressed,
            "@button_border@": t.button_border,
            "@entry_bg@": t.entry_bg,
            "@entry_border@": t.entry_border,
            "@badge_bg@": t.badge_bg,
            "@error@": t.error,
            "@error_bg@": t.error_bg,
            "@success@": t.success,
            "@success_bg@": t.success_bg,
            "@warning@": t.warning,
            "@warning_bg@": t.warning_bg,
            "@table_header_bg@": t.table_header_bg,
            "@table_row_alt@": t.table_row_alt,
            "@table_selected_bg@": t.table_selected_bg,
            "@log_bg@": t.log_bg,
            "@log_fg@": t.log_fg,
        }

        # Override accent with Windows DWM system accent if available on Windows
        if CAPABILITIES.is_windows:
            system_accent = get_system_accent_color()
            if system_accent:
                replacements["@accent@"] = system_accent

        compiled = raw_qss
        for placeholder, value in replacements.items():
            compiled = compiled.replace(placeholder, value)

        return compiled

    def apply_theme(self, app=None) -> None:
        if not HAS_QT:
            return

        target_app = app or QApplication.instance()
        if not target_app:
            return

        qss = self.compile_qss()
        target_app.setStyleSheet(qss)


THEME = ThemeManager("system")
