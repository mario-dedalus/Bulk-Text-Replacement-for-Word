"""Unit tests for application services, platform adapters, cancellation, and worker signatures."""

from pathlib import Path
import shutil
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from application.merge_service import MergeService
from application.replace_service import ReplaceService
from application.task_models import CancellationToken, TaskProgress, TaskState
from application.workers import TaskWorker
from core.models import ExcelData, MergeResult
from platform_adapter.appearance import detect_system_dark_mode, get_system_accent_color
from platform_adapter.capabilities import CAPABILITIES
from platform_adapter.file_manager import open_folder, reveal_in_file_manager

SAMPLES = ROOT / "tests" / "samples"


class ApplicationServiceTests(unittest.TestCase):
    def test_cancellation_token(self):
        token = CancellationToken()
        self.assertFalse(token.is_cancelled)
        token.cancel()
        self.assertTrue(token.is_cancelled)
        token.reset()
        self.assertFalse(token.is_cancelled)

    def test_task_progress_calculation(self):
        p = TaskProgress.calculate(5, 10, "测试中")
        self.assertEqual(p.percentage, 50)
        self.assertEqual(p.current, 5)
        self.assertEqual(p.total, 10)
        self.assertEqual(p.message, "测试中")

    def test_replace_service_validation(self):
        # Empty files
        err1 = ReplaceService.validate_inputs([], "test")
        self.assertIsNotNone(err1)

        # Empty search text
        err2 = ReplaceService.validate_inputs(["/fake/path.docx"], "")
        self.assertIsNotNone(err2)

        # Non-existent file
        err3 = ReplaceService.validate_inputs(["/non/existent/doc.docx"], "search")
        self.assertIn("文件不存在", err3)

    def test_replace_service_preview_and_execute(self):
        with tempfile.TemporaryDirectory() as folder:
            doc_path = Path(folder) / "doc.docx"
            shutil.copy2(SAMPLES / "合同模板.docx", doc_path)

            # Preview
            prev_res = ReplaceService.execute_preview(
                file_paths=[str(doc_path)],
                search_text="甲方名称",
            )
            self.assertTrue(prev_res.success)
            self.assertEqual(prev_res.state, TaskState.SUCCESS)
            self.assertGreater(prev_res.data.total_count, 0)

            # Replace
            rep_res = ReplaceService.execute_replace(
                file_paths=[str(doc_path)],
                search_text="甲方名称",
                replace_text="上海测试企业",
                create_backup=True,
            )
            self.assertTrue(rep_res.success)
            self.assertEqual(rep_res.data.successful_files, 1)
            self.assertEqual(len(rep_res.data.backup_files), 1)

    def test_scan_links_cancellation(self):
        tmpl = SAMPLES / "合同模板.docx"
        token = CancellationToken()
        token.cancel()

        res = ReplaceService.scan_links(
            file_paths=[str(tmpl)],
            cancel_token=token,
        )
        self.assertFalse(res.success)
        self.assertEqual(res.state, TaskState.CANCELLED)

    def test_merge_service_scan_preview_and_generate(self):
        tmpl = SAMPLES / "合同模板.docx"
        excel = SAMPLES / "合同数据.xlsx"

        # Scan
        scan_res = MergeService.scan_template_and_excel(str(tmpl), str(excel))
        self.assertTrue(scan_res.success)
        fields, data, mapping = scan_res.data
        self.assertEqual(len(fields), 6)
        self.assertEqual(len(data.rows), 3)

        with tempfile.TemporaryDirectory() as out_folder:
            # Preview snippets with default fallback
            defaults = {"管辖法院": "上海市第一中级人民法院"}
            snippets = MergeService.generate_preview_snippets(
                template_path=str(tmpl),
                excel_data=data,
                fields=fields,
                mapping=mapping,
                output_folder=out_folder,
                filename_rule="{{甲方名称}}-合同.docx",
                default_values=defaults,
            )
            self.assertEqual(len(snippets), 3)
            self.assertEqual(snippets[0]["filename"], "A公司-合同.docx")

            # Execute batch merge directly
            merge_res = MergeService.execute_batch_merge(
                template_path=str(tmpl),
                excel_data=data,
                mapping=mapping,
                output_folder=out_folder,
                filename_rule="{{甲方名称}}-合同.docx",
                default_values=defaults,
            )
            self.assertTrue(merge_res.success)
            self.assertEqual(len(merge_res.data), 3)
            self.assertTrue(all(r.success for r in merge_res.data))

    def test_merge_service_scan_cancellation(self):
        tmpl = SAMPLES / "合同模板.docx"
        excel = SAMPLES / "合同数据.xlsx"
        token = CancellationToken()
        token.cancel()

        scan_res = MergeService.scan_template_and_excel(str(tmpl), str(excel), cancel_token=token)
        self.assertFalse(scan_res.success)
        self.assertEqual(scan_res.state, TaskState.CANCELLED)

    def test_merge_preview_respects_per_field_empty_behaviors(self):
        data = ExcelData(
            headers=["保留", "清空", "自定义"],
            rows=[{"保留": "", "清空": "", "自定义": ""}],
            excel_rows=[2],
        )
        fields = ["保留", "清空", "自定义"]
        snippets = MergeService.generate_preview_snippets(
            template_path="template.docx",
            excel_data=data,
            fields=fields,
            mapping={field: field for field in fields},
            output_folder=".",
            filename_rule="preview.docx",
            default_values={"自定义": "待补充"},
            empty_field_behaviors={
                "保留": "keep_variable",
                "清空": "replace_empty",
                "自定义": "custom",
            },
        )

        preview_fields = {item["field"]: item for item in snippets[0]["fields"]}
        self.assertEqual(preview_fields["保留"]["empty_behavior"], "keep_variable")
        self.assertEqual(preview_fields["清空"]["empty_behavior"], "replace_empty")
        self.assertEqual(preview_fields["清空"]["value"], "")
        self.assertEqual(preview_fields["自定义"]["empty_behavior"], "custom")
        self.assertEqual(preview_fields["自定义"]["value"], "待补充")

    def test_merge_worker_progress_callback_signature(self):
        """Verify TaskWorker executes MergeService without TypeError on progress callback."""
        tmpl = SAMPLES / "合同模板.docx"
        excel = SAMPLES / "合同数据.xlsx"
        scan_res = MergeService.scan_template_and_excel(str(tmpl), str(excel))
        fields, data, mapping = scan_res.data

        with tempfile.TemporaryDirectory() as out_folder:
            progress_events: list[TaskProgress] = []

            def progress_sink(p: TaskProgress):
                progress_events.append(p)

            res = MergeService.execute_batch_merge(
                template_path=str(tmpl),
                excel_data=data,
                mapping=mapping,
                output_folder=out_folder,
                filename_rule="{{甲方名称}}-test.docx",
                progress_cb=progress_sink,
            )

            self.assertTrue(res.success)
            self.assertEqual(len(progress_events), 3)
            self.assertEqual(progress_events[-1].percentage, 100)
            self.assertIsInstance(progress_events[0].data, MergeResult)

    def test_task_worker_signature_inspection(self):
        """Test TaskWorker parameter pre-inspection without broad runtime TypeError retry."""
        def simple_fn(a, b):
            return a + b

        worker = TaskWorker(simple_fn, 3, 5)
        self.assertFalse(worker._inject_progress)
        self.assertFalse(worker._inject_cancel)

        def cancellable_fn(x, cancel_token=None):
            return x * 2

        worker_c = TaskWorker(cancellable_fn, 10)
        self.assertFalse(worker_c._inject_progress)
        self.assertTrue(worker_c._inject_cancel)

    def test_platform_capabilities_and_appearance(self):
        self.assertIsNotNone(CAPABILITIES.platform_name)
        self.assertIsInstance(CAPABILITIES.is_windows, bool)
        self.assertIsInstance(CAPABILITIES.is_macos, bool)

        dark = detect_system_dark_mode()
        self.assertIsInstance(dark, bool)

    def test_merge_service_scan_with_sheet_name(self):
        """Test scan_template_and_excel with a specific sheet_name."""
        from openpyxl import Workbook

        tmpl = SAMPLES / "合同模板.docx"
        with tempfile.TemporaryDirectory() as folder:
            excel_path = Path(folder) / "multi.xlsx"
            wb = Workbook()

            ws1 = wb.active
            ws1.title = "常规信息"
            ws1.append(["无用字段1", "无用字段2"])
            ws1.append(["值1", "值2"])

            ws2 = wb.create_sheet(title="合同要素")
            ws2.append(["甲方名称", "乙方名称", "合同金额", "签订日期", "管辖法院", "Contract ID"])
            ws2.append(["北京公司", "天津公司", "800万", "2026-08-22", "北京朝阳法院", "HT-2026"])

            wb.save(excel_path)
            wb.close()

            # Scan sheet 2
            scan_res = MergeService.scan_template_and_excel(
                str(tmpl), str(excel_path), sheet_name="合同要素"
            )
            self.assertTrue(scan_res.success)
            fields, data, mapping = scan_res.data
            self.assertEqual(len(fields), 6)
            self.assertEqual(len(data.rows), 1)
            self.assertEqual(data.rows[0]["甲方名称"], "北京公司")
            self.assertEqual(mapping["甲方名称"], "甲方名称")


if __name__ == "__main__":
    unittest.main()
