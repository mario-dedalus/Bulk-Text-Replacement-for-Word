"""GUI regression tests for the field-mapping combo-box delegate."""

from __future__ import annotations

import os
from pathlib import Path
import sys
import unittest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

try:
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication, QComboBox

    from ui.models.field_mapping_model import FieldMappingModel
    from ui.pages.merge_page import MergePage

    HAS_QT = True
except ImportError:
    HAS_QT = False


@unittest.skipUnless(HAS_QT, "PySide6 is required for GUI delegate tests")
class MappingComboDelegateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.page = MergePage()
        self.page.resize(900, 650)
        self.page.mapping_model.set_data(["Client"], {"Client": "Column A"})
        self.page.combo_delegate.set_headers(["Column A", "Column B"])
        self.page.card_step2.setEnabled(True)
        self.page.show()
        self.app.processEvents()

    def tearDown(self):
        self.page.close()
        self.app.processEvents()

    def _mapping_index(self):
        return self.page.mapping_model.index(0, FieldMappingModel.COL_EXCEL)

    def _empty_field_index(self):
        return self.page.mapping_model.index(0, FieldMappingModel.COL_DEFAULT)

    def test_single_click_opens_editor_without_premature_commit(self):
        index = self._mapping_index()
        rect = self.page.tbl_mapping.visualRect(index)
        QTest.mouseClick(self.page.tbl_mapping.viewport(), Qt.LeftButton, pos=rect.center())
        self.app.processEvents()

        combo = self.page.tbl_mapping.findChild(QComboBox)
        self.assertIsNotNone(combo)
        self.assertTrue(combo.isVisible())
        self.assertEqual(combo.currentData(), "Column A")
        self.assertEqual(self.page.mapping_model.get_mapping()["Client"], "Column A")

    def test_only_user_activation_commits_and_closes_editor(self):
        index = self._mapping_index()
        self.page._open_mapping_editor(index)
        self.app.processEvents()

        combo = self.page.tbl_mapping.findChild(QComboBox)
        self.assertIsNotNone(combo)
        target = combo.findData("Column B")

        # Programmatic initialisation/change must not commit or close the editor.
        combo.setCurrentIndex(target)
        self.app.processEvents()
        self.assertEqual(self.page.mapping_model.get_mapping()["Client"], "Column A")
        self.assertTrue(combo.isVisible())

        # A real user activation commits once and removes the editor cleanly.
        combo.activated.emit(target)
        self.app.processEvents()
        self.assertEqual(self.page.mapping_model.get_mapping()["Client"], "Column B")
        self.assertFalse(combo.isVisible())

    def test_empty_field_policy_defaults_to_keep_and_offers_three_choices(self):
        index = self._empty_field_index()
        self.assertEqual(index.data(Qt.DisplayRole), "保持原变量（默认）")
        self.assertEqual(
            self.page.mapping_model.get_empty_behaviors()["Client"],
            FieldMappingModel.EMPTY_KEEP,
        )

        self.page._open_mapping_editor(index)
        self.app.processEvents()
        combo = self.page.tbl_mapping.findChild(QComboBox)
        self.assertIsNotNone(combo)
        self.assertEqual(
            [combo.itemText(i) for i in range(combo.count())],
            ["保持原变量（默认）", "替换为空", "自定义…"],
        )

    def test_empty_field_policy_is_stored_per_template_variable(self):
        model = self.page.mapping_model
        index = self._empty_field_index()

        self.assertTrue(
            model.setData(index, (FieldMappingModel.EMPTY_REPLACE, ""), Qt.EditRole)
        )
        self.assertEqual(index.data(Qt.DisplayRole), "替换为空")
        self.assertEqual(model.get_defaults(), {})

        self.assertTrue(
            model.setData(index, (FieldMappingModel.EMPTY_CUSTOM, "待补充"), Qt.EditRole)
        )
        self.assertEqual(index.data(Qt.DisplayRole), "自定义：待补充")
        self.assertEqual(model.get_defaults(), {"Client": "待补充"})

    def test_step_three_no_longer_contains_global_empty_field_selector(self):
        self.assertFalse(hasattr(self.page, "cmb_empty_behavior"))


if __name__ == "__main__":
    unittest.main()
