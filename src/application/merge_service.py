"""Application service orchestrating Word template + Excel batch generation."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Callable, Mapping, Optional

from application.task_models import CancellationToken, ServiceResult, TaskProgress, TaskState
from core.models import ExcelData, MergeResult
from core.template_merge import (
    build_field_mapping,
    build_output_filename,
    extract_template_fields,
    extract_template_fields_com,
    generate_batch,
    load_excel_data,
    mapped_row_values,
    resolve_empty_field,
)
from platform_adapter.capabilities import CAPABILITIES


class MergeService:
    """Encapsulates validation and execution of template scan and batch merge tasks."""

    @staticmethod
    def validate_sources(template_path: str, excel_path: str) -> str | None:
        if not template_path or not os.path.isfile(template_path):
            return "请选择存在的 Word 模板文件（.docx/.docm/.doc）。"
        if not excel_path or not os.path.isfile(excel_path):
            return "请选择存在的表格数据文件（.xlsx/.csv/.xlsm）。"

        suffix = Path(template_path).suffix.lower()
        if CAPABILITIES.is_macos and suffix == ".doc":
            return "macOS 暂不支持旧版 .doc 格式，请将模板转换为 .docx 后使用。"

        table_suffix = Path(excel_path).suffix.lower()
        if table_suffix not in {".xlsx", ".xlsm", ".csv", ".xltx", ".xltm"}:
            return "数据源文件必须为 .xlsx、.xlsm 或 .csv 格式表格。"

        return None

    @staticmethod
    def scan_template_and_excel(
        template_path: str,
        excel_path: str,
        use_com: bool = False,
        sheet_name: Optional[str] = None,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[tuple[list[str], ExcelData, dict[str, str]]]:
        err = MergeService.validate_sources(template_path, excel_path)
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        if cancel_token and cancel_token.is_cancelled:
            return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

        try:
            if progress_cb:
                progress_cb(TaskProgress.calculate(1, 2, "正在提取 Word 模板变量..."))

            suffix = Path(template_path).suffix.lower()
            if suffix == ".doc":
                if not CAPABILITIES.has_word_com:
                    return ServiceResult(
                        success=False,
                        error=".doc 模板扫描需要 Windows、Microsoft Word 和 pywin32 支持。",
                        state=TaskState.FAILED,
                    )
                fields = extract_template_fields_com(template_path)
            else:
                fields = extract_template_fields(template_path)

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

            if not fields:
                return ServiceResult(
                    success=False,
                    error="Word 模板中未检测到 {{字段名}} 格式的变量。",
                    state=TaskState.FAILED,
                )

            if progress_cb:
                progress_cb(TaskProgress.calculate(2, 2, "正在读取 Excel 数据表头..."))

            excel_data = load_excel_data(excel_path, sheet_name=sheet_name)
            mapping = build_field_mapping(fields, excel_data.headers)

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, error="任务已取消", state=TaskState.CANCELLED)

            return ServiceResult(
                success=True,
                data=(fields, excel_data, mapping),
                state=TaskState.SUCCESS,
            )
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def generate_preview_snippets(
        template_path: str,
        excel_data: ExcelData,
        fields: list[str],
        mapping: Mapping[str, str],
        output_folder: str,
        filename_rule: str,
        default_values: Optional[Mapping[str, str]] = None,
        empty_field_behaviors: Optional[Mapping[str, str]] = None,
        max_rows: int = 5,
    ) -> list[dict]:
        extension = Path(template_path).suffix.lower() or ".docx"
        reserved: set[str] = set()
        snippets = []
        defaults = default_values or {}

        for excel_row, row in list(zip(excel_data.excel_rows, excel_data.rows))[:max_rows]:
            mapped = mapped_row_values(row, mapping, defaults)
            filename = build_output_filename(filename_rule, mapped, output_folder, extension, reserved)

            field_pairs = []
            for field in fields:
                header = mapping.get(field, "")
                raw_val = row.get(header, "") if header else ""
                _replace, effective_val, behavior = resolve_empty_field(
                    field,
                    raw_val,
                    default_values=defaults,
                    empty_field_behaviors=empty_field_behaviors,
                )

                field_pairs.append({
                    "field": field,
                    "header": header,
                    "value": str(effective_val) if effective_val is not None else "",
                    "is_mapped": bool(header),
                    "is_default": behavior == "custom",
                    "empty_behavior": behavior,
                })

            snippets.append({
                "excel_row": excel_row,
                "filename": filename,
                "fields": field_pairs,
            })

        return snippets

    @staticmethod
    def execute_batch_merge(
        template_path: str,
        excel_data: ExcelData,
        mapping: Mapping[str, str],
        output_folder: str,
        filename_rule: str,
        use_com: bool = False,
        default_values: Optional[Mapping[str, str]] = None,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
        replace_empty: bool = True,
        empty_field_behaviors: Optional[Mapping[str, str]] = None,
    ) -> ServiceResult[list[MergeResult]]:
        if not output_folder or not output_folder.strip():
            return ServiceResult(success=False, error="请指定输出文件夹路径。", state=TaskState.FAILED)
        if not filename_rule or not filename_rule.strip():
            return ServiceResult(success=False, error="请输入输出文件名规则。", state=TaskState.FAILED)

        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, result: MergeResult):
            if progress_cb:
                prog = TaskProgress.calculate(
                    current, total, f"正在生成：{result.filename}", data=result
                )
                try:
                    progress_cb(prog)
                except TypeError:
                    progress_cb(prog, result)  # type: ignore

        try:
            results = generate_batch(
                template_path=template_path,
                data=excel_data,
                mapping=mapping,
                output_folder=output_folder,
                filename_rule=filename_rule,
                use_com=use_com,
                default_values=default_values,
                progress=_inner_progress,
                is_cancelled=is_cancelled_fn,
                replace_empty=replace_empty,
                empty_field_behaviors=empty_field_behaviors,
            )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=results, error="任务已取消", state=TaskState.CANCELLED)

            failed = sum(1 for r in results if not r.success)
            state = TaskState.WARNING if failed > 0 else TaskState.SUCCESS
            return ServiceResult(success=True, data=results, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)
