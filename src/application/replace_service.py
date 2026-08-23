"""Application service orchestrating text replacement and preview jobs."""

from __future__ import annotations

import os
from typing import Callable, Optional

from application.task_models import CancellationToken, ServiceResult, TaskProgress, TaskState
from core.models import BatchProcessResult
from core.replacer_core import (
    get_document_text,
    perform_com_preview,
    perform_com_replace,
    perform_standard_preview,
    perform_standard_replace,
    preprocess_text_with_nbsp,
    scan_hyperlinks,
)
from platform_adapter.capabilities import CAPABILITIES


class ReplaceService:
    """Encapsulates validation and execution of document search and replace tasks."""

    @staticmethod
    def validate_inputs(file_paths: list[str], search_text: str) -> str | None:
        if not file_paths:
            return "请先添加至少一个 Word 文档。"
        if not search_text or not search_text.strip():
            return "请输入要查找的内容。"
        for path in file_paths:
            if not os.path.isfile(path):
                return f"文件不存在：{path}"
            suffix = os.path.splitext(path)[1].lower()
            if CAPABILITIES.is_macos and suffix == ".doc":
                return f"macOS 暂不支持旧版 .doc 格式文件：{os.path.basename(path)}"
        return None

    @staticmethod
    def execute_preview(
        file_paths: list[str],
        search_text: str,
        mode: str = "fast",
        case_sensitive: bool = False,
        use_regex: bool = False,
        whole_word: bool = False,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[BatchProcessResult]:
        err = ReplaceService.validate_inputs(file_paths, search_text)
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        search_for = preprocess_text_with_nbsp(search_text)
        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, filename: str):
            if progress_cb:
                progress_cb(TaskProgress.calculate(current, total, f"正在分析：{filename}"))

        try:
            if mode == "full":
                if not CAPABILITIES.has_word_com:
                    return ServiceResult(
                        success=False,
                        error="完整模式需要 Windows、Microsoft Word 和 pywin32 支持。",
                        state=TaskState.FAILED,
                    )
                res = perform_com_preview(
                    file_paths=file_paths,
                    search_for=search_for,
                    case_sensitive=case_sensitive,
                    whole_word=whole_word,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )
            else:
                res = perform_standard_preview(
                    file_paths=file_paths,
                    search_for=search_for,
                    case_sensitive=case_sensitive,
                    use_regex=use_regex,
                    whole_word=whole_word,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=res, error="任务已取消", state=TaskState.CANCELLED)

            state = TaskState.WARNING if res.errors else TaskState.SUCCESS
            return ServiceResult(success=True, data=res, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def execute_replace(
        file_paths: list[str],
        search_text: str,
        replace_text: str,
        mode: str = "fast",
        case_sensitive: bool = False,
        use_regex: bool = False,
        whole_word: bool = False,
        create_backup: bool = True,
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[BatchProcessResult]:
        err = ReplaceService.validate_inputs(file_paths, search_text)
        if err:
            return ServiceResult(success=False, error=err, state=TaskState.FAILED)

        search_for = preprocess_text_with_nbsp(search_text)
        replace_with = preprocess_text_with_nbsp(replace_text)
        is_cancelled_fn = cancel_token.is_cancelled if cancel_token else None

        def _inner_progress(current: int, total: int, filename: str):
            if progress_cb:
                progress_cb(TaskProgress.calculate(current, total, f"正在处理：{filename}"))

        try:
            if mode == "full":
                if not CAPABILITIES.has_word_com:
                    return ServiceResult(
                        success=False,
                        error="完整模式需要 Windows、Microsoft Word 和 pywin32 支持。",
                        state=TaskState.FAILED,
                    )
                res = perform_com_replace(
                    file_paths=file_paths,
                    search_for=search_for,
                    replace_with=replace_with,
                    case_sensitive=case_sensitive,
                    whole_word=whole_word,
                    create_backup=create_backup,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )
            else:
                res = perform_standard_replace(
                    file_paths=file_paths,
                    search_for=search_for,
                    replace_with=replace_with,
                    case_sensitive=case_sensitive,
                    use_regex=use_regex,
                    whole_word=whole_word,
                    create_backup=create_backup,
                    progress_callback=_inner_progress,
                    is_cancelled=is_cancelled_fn,
                )

            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=res, error="任务已取消", state=TaskState.CANCELLED)

            state = TaskState.WARNING if res.errors else TaskState.SUCCESS
            return ServiceResult(success=True, data=res, state=state)
        except Exception as exc:
            return ServiceResult(success=False, error=str(exc), state=TaskState.FAILED)

    @staticmethod
    def scan_links(
        file_paths: list[str],
        progress_cb: Optional[Callable[[TaskProgress], None]] = None,
        cancel_token: Optional[CancellationToken] = None,
    ) -> ServiceResult[list[dict]]:
        results = []
        total = len(file_paths)
        for i, path in enumerate(file_paths, start=1):
            if cancel_token and cancel_token.is_cancelled:
                return ServiceResult(success=False, data=results, error="超链接扫描已取消", state=TaskState.CANCELLED)

            filename = os.path.basename(path)
            if progress_cb:
                progress_cb(TaskProgress.calculate(i, total, f"正在扫描链接：{filename}"))

            res = scan_hyperlinks([path])
            if res:
                results.extend(res)

        return ServiceResult(success=True, data=results, state=TaskState.SUCCESS)
