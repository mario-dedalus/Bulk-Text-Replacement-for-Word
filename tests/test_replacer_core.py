"""Unit tests for the headless replacer_core module."""

from pathlib import Path
import shutil
import sys
import tempfile
import unittest

from docx import Document
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from replacer_core import (  # noqa: E402
    count_occurrences,
    find_match_contexts,
    get_document_text,
    perform_standard_preview,
    perform_standard_replace,
    preprocess_text_with_nbsp,
    replace_in_paragraph_advanced,
    replace_in_table,
    scan_hyperlinks,
    strip_invisible_chars,
)


class ReplacerCoreTests(unittest.TestCase):
    def test_invisible_chars_and_nbsp_cleaning(self):
        dirty = "甲\u00ad方\u200b名\u200c称\u200d\u2060\ufeff"
        cleaned = strip_invisible_chars(dirty)
        self.assertEqual(cleaned, "甲方名称")

        nbsp_text = "合同 金额 [NBSP] 100 &nbsp; 元"
        processed = preprocess_text_with_nbsp(nbsp_text)
        self.assertIn("\u00a0", processed)
        self.assertNotIn("[NBSP]", processed)
        self.assertNotIn("&nbsp;", processed)

    def test_count_occurrences_modes(self):
        text = "Hello world, hello WORLD. Hello123 World."
        # Case insensitive
        self.assertEqual(count_occurrences(text, "hello", case_sensitive=False), 3)
        # Case sensitive
        self.assertEqual(count_occurrences(text, "Hello", case_sensitive=True), 2)
        # Whole word
        self.assertEqual(count_occurrences(text, "Hello", case_sensitive=False, whole_word=True), 2)
        # Regex
        self.assertEqual(count_occurrences(text, r"Hello\d+", use_regex=True), 1)

    def test_find_match_contexts(self):
        text = "这是甲方的保密协议，甲方承诺保护相关数据安全。甲方签字生效。"
        contexts = find_match_contexts(text, "甲方", context_chars=6, max_matches=2)
        self.assertEqual(len(contexts), 2)
        self.assertTrue(all("【甲方】" in c for c in contexts))

    def test_replace_in_paragraph_single_and_multi_run(self):
        doc = Document()
        # 1. Single run
        p1 = doc.add_paragraph("甲方名称：山西ABC有限公司")
        count1 = replace_in_paragraph_advanced(p1, "山西ABC有限公司", "山西XYZ有限公司")
        self.assertEqual(count1, 1)
        self.assertEqual(p1.text, "甲方名称：山西XYZ有限公司")

        # 2. Multi-run (split run) with formatting preservation
        p2 = doc.add_paragraph()
        r1 = p2.add_run("合同")
        r1.bold = True
        r2 = p2.add_run("签署")
        r2.font.name = "Arial"
        r2.font.size = Pt(14)
        r3 = p2.add_run("地点：太原市")

        self.assertEqual(p2.text, "合同签署地点：太原市")
        count2 = replace_in_paragraph_advanced(p2, "签署地点", "签订城市")
        self.assertEqual(count2, 1)
        self.assertEqual(p2.text, "合同签订城市：太原市")

    def test_replace_in_table_and_nested_table(self):
        doc = Document()
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "原告：张三"
        table.cell(0, 1).text = "被告：李四"
        table.cell(1, 0).text = "案由：借款"
        table.cell(1, 1).text = "原告住所地：北京"

        count = replace_in_table(table, "原告", "申请人")
        self.assertEqual(count, 2)
        self.assertIn("申请人：张三", table.cell(0, 0).text)
        self.assertIn("申请人住所地：北京", table.cell(1, 1).text)

    def test_standard_replace_and_backup(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        with tempfile.TemporaryDirectory() as temp_dir:
            test_doc = Path(temp_dir) / "test_contract.docx"
            shutil.copy2(sample_template, test_doc)

            # Perform standard replace with backup
            res = perform_standard_replace(
                file_paths=[str(test_doc)],
                search_for="甲方名称",
                replace_with="北京顶级企业",
                case_sensitive=False,
                use_regex=False,
                whole_word=False,
                create_backup=True,
            )

            self.assertEqual(res.successful_files, 1)
            self.assertGreater(res.total_count, 0)
            self.assertEqual(len(res.backup_files), 1)
            self.assertTrue(Path(res.backup_files[0]).exists())

            # Verify replaced document content
            updated_doc = Document(test_doc)
            updated_text = get_document_text(updated_doc)
            self.assertIn("北京顶级企业", updated_text)

    def test_standard_preview(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        with tempfile.TemporaryDirectory() as temp_dir:
            test_doc = Path(temp_dir) / "test_contract.docx"
            shutil.copy2(sample_template, test_doc)

            preview_res = perform_standard_preview(
                file_paths=[str(test_doc)],
                search_for="甲方名称",
                case_sensitive=False,
            )

            self.assertEqual(preview_res.files_processed, 1)
            self.assertGreater(preview_res.total_count, 0)
            self.assertTrue(len(preview_res.details[0].contexts) > 0)

    def test_scan_hyperlinks(self):
        sample_template = ROOT / "tests" / "samples" / "合同模板.docx"
        links = scan_hyperlinks([str(sample_template)])
        self.assertEqual(len(links), 1)
        self.assertEqual(links[0]["count"], 1)
        self.assertEqual(links[0]["urls"], ["https://example.com"])


if __name__ == "__main__":
    unittest.main()
