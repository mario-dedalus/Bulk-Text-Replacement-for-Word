"""Unit tests for ThemeManager and QSS compilation without Qt dependency."""

from pathlib import Path
import sys
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from ui.theme.theme_manager import THEME, ThemeManager
from ui.theme.tokens import DARK_TOKENS, LIGHT_TOKENS


# Prevent Windows system accent color from overriding token values in tests
_NO_SYSTEM_ACCENT = patch(
    "ui.theme.theme_manager.get_system_accent_color", return_value=None
)



class ThemeManagerTests(unittest.TestCase):
    @_NO_SYSTEM_ACCENT
    def test_qss_compilation_non_empty(self, _mock):
        manager = ThemeManager("light")
        qss_light = manager.compile_qss(LIGHT_TOKENS)
        self.assertTrue(len(qss_light) > 100)
        self.assertNotIn("@window_bg@", qss_light)
        self.assertNotIn("@accent@", qss_light)
        self.assertIn(LIGHT_TOKENS.window_bg, qss_light)
        self.assertIn(LIGHT_TOKENS.accent, qss_light)

    @_NO_SYSTEM_ACCENT
    def test_dark_mode_qss_tokens(self, _mock):
        manager = ThemeManager("dark")
        qss_dark = manager.compile_qss(DARK_TOKENS)
        self.assertTrue(len(qss_dark) > 100)
        self.assertNotIn("@window_bg@", qss_dark)
        self.assertNotIn("@accent@", qss_dark)
        self.assertIn(DARK_TOKENS.window_bg, qss_dark)
        self.assertIn(DARK_TOKENS.accent, qss_dark)

        # Light and dark outputs must be distinct
        qss_light = manager.compile_qss(LIGHT_TOKENS)
        self.assertNotEqual(qss_light, qss_dark)

    def test_qss_contains_disabled_and_focus_rules(self):
        manager = ThemeManager("light")
        qss = manager.compile_qss()
        self.assertIn(":disabled", qss)
        self.assertIn(":focus", qss)

    def test_qss_contains_viewport_and_content_background(self):
        manager = ThemeManager("dark")
        qss = manager.compile_qss(DARK_TOKENS)
        self.assertIn("QScrollArea::viewport", qss)
        self.assertIn('isPageContent="true"', qss)
        self.assertIn(DARK_TOKENS.window_bg, qss)

    def test_toggle_theme(self):
        manager = ThemeManager("light")
        self.assertFalse(manager.is_dark)
        manager.toggle()
        self.assertTrue(manager.is_dark)
        self.assertEqual(manager.tokens.window_bg, DARK_TOKENS.window_bg)
        manager.toggle()
        self.assertFalse(manager.is_dark)
        self.assertEqual(manager.tokens.window_bg, LIGHT_TOKENS.window_bg)


if __name__ == "__main__":
    unittest.main()
