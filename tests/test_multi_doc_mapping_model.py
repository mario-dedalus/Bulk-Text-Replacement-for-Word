import os
from pathlib import Path
import sys
import unittest

from PySide6.QtCore import QModelIndex, Qt
from PySide6.QtWidgets import QApplication

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from core.models import ExcelData
from ui.models.multi_doc_mapping_model import MultiDocMappingModel

# Ensure QApplication is initialized for Qt model tests
app = QApplication.instance() or QApplication(sys.argv)


class MultiDocMappingModelTests(unittest.TestCase):
    def setUp(self):
        self.model = MultiDocMappingModel()
        self.doc_a = os.path.abspath("/path/to/DocA.docx")
        self.doc_b = os.path.abspath("/path/to/DocB.docx")

    def test_empty_model_geometry_and_headers(self):
        self.assertEqual(self.model.rowCount(), 0)
        self.assertEqual(self.model.columnCount(), 3)  # Index, Filename, Status
        self.assertEqual(self.model.headerData(0, Qt.Horizontal), "序号")
        self.assertEqual(self.model.headerData(1, Qt.Horizontal), "文档名称")
        self.assertEqual(self.model.headerData(2, Qt.Horizontal), "状态")

    def test_add_documents_and_set_variables(self):
        added = self.model.add_documents([self.doc_a, self.doc_b])
        self.assertEqual(added, 2)
        self.assertEqual(self.model.rowCount(), 2)

        # Set variables
        all_vars = ["姓名", "身份证号"]
        doc_vars_map = {
            self.doc_a: ["姓名", "身份证号"],
            self.doc_b: ["姓名"],
        }
        self.model.set_variables(all_vars, doc_vars_map)

        self.assertEqual(self.model.columnCount(), 5)  # Index, Filename, {{姓名}}, {{身份证号}}, Status
        self.assertEqual(self.model.headerData(2, Qt.Horizontal), "{{姓名}}")
        self.assertEqual(self.model.headerData(3, Qt.Horizontal), "{{身份证号}}")
        self.assertEqual(self.model.headerData(4, Qt.Horizontal), "状态")

        # Check initial status (no values entered yet)
        status_a = self.model.data(self.model.index(0, 4), Qt.DisplayRole)
        self.assertIn("待补充数据", status_a)

    def test_cell_editing_and_status_update(self):
        self.model.add_documents([self.doc_a])
        self.model.set_variables(["姓名", "身份证号"], {self.doc_a: ["姓名", "身份证号"]})

        idx_name = self.model.index(0, 2)
        idx_id = self.model.index(0, 3)

        # Check editable flag
        flags = self.model.flags(idx_name)
        self.assertTrue(bool(flags & Qt.ItemIsEditable))

        # Edit name
        self.model.setData(idx_name, "张三", Qt.EditRole)
        self.assertEqual(self.model.data(idx_name, Qt.DisplayRole), "张三")

        # Status should now be partial
        status_part = self.model.data(self.model.index(0, 4), Qt.DisplayRole)
        self.assertIn("待补齐", status_part)

        # Edit ID
        self.model.setData(idx_id, "110101199001011234", Qt.EditRole)
        status_ready = self.model.data(self.model.index(0, 4), Qt.DisplayRole)
        self.assertEqual(status_ready, "✓ 数据就绪")

        items = self.model.get_items()
        self.assertEqual(items[0].replacements["姓名"], "张三")
        self.assertEqual(items[0].replacements["身份证号"], "110101199001011234")

    def test_import_table_data_and_manual_override(self):
        doc1 = os.path.abspath("/path/to/Doc1.docx")
        doc2 = os.path.abspath("/path/to/Doc2.docx")
        self.model.add_documents([doc1, doc2])
        self.model.set_variables(["姓名", "金额"], {
            doc1: ["姓名", "金额"],
            doc2: ["姓名", "金额"],
        })


        excel_data = ExcelData(
            headers=["姓名", "金额", "备注"],
            rows=[
                {"姓名": "李四", "金额": "5000", "备注": "一等奖"},
                {"姓名": "王五", "金额": "8000", "备注": "特等奖"},
            ],
            excel_rows=[2, 3],
        )

        filled = self.model.import_table_data(excel_data, match_by_header=True)
        self.assertEqual(filled, 2)

        # Check values imported
        self.assertEqual(self.model.data(self.model.index(0, 2), Qt.DisplayRole), "李四")
        self.assertEqual(self.model.data(self.model.index(0, 3), Qt.DisplayRole), "5000")
        self.assertEqual(self.model.data(self.model.index(1, 2), Qt.DisplayRole), "王五")
        self.assertEqual(self.model.data(self.model.index(1, 3), Qt.DisplayRole), "8000")

        # Manual override row 2
        self.model.setData(self.model.index(1, 3), "9999", Qt.EditRole)
        self.assertEqual(self.model.data(self.model.index(1, 3), Qt.DisplayRole), "9999")

        items = self.model.get_items()
        self.assertEqual(items[1].replacements["金额"], "9999")

    def test_move_rows_up_and_down(self):
        self.model.add_documents(["/path/to/First.docx", "/path/to/Second.docx", "/path/to/Third.docx"])

        # Move second row up
        moved = self.model.move_row_up(1)
        self.assertTrue(moved)
        self.assertEqual(self.model.data(self.model.index(0, 1), Qt.DisplayRole), "Second.docx")
        self.assertEqual(self.model.data(self.model.index(1, 1), Qt.DisplayRole), "First.docx")

        # Move top row up (should fail safely)
        self.assertFalse(self.model.move_row_up(0))

        # Move first row down
        moved_down = self.model.move_row_down(0)
        self.assertTrue(moved_down)
        self.assertEqual(self.model.data(self.model.index(0, 1), Qt.DisplayRole), "First.docx")


if __name__ == "__main__":
    unittest.main()
