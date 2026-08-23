"""Shared data structures for core document processing and template merging."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ExcelData:
    headers: list[str]
    rows: list[dict[str, str]]
    excel_rows: list[int]


@dataclass
class MergeResult:
    excel_row: int
    filename: str
    success: bool
    replacements: int = 0
    error: str = ""


@dataclass
class FileProcessDetail:
    filename: str
    file_path: str
    total: int = 0
    details: list[str] = field(default_factory=list)
    contexts: list[str] = field(default_factory=list)
    error: str = ""
    backup_path: str = ""


@dataclass
class BatchProcessResult:
    files_processed: int
    total_count: int
    files_with_matches: int
    successful_files: int
    details: list[FileProcessDetail] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    backup_files: list[str] = field(default_factory=list)


@dataclass
class MultiDocItem:
    """Represents a single document and its mapped variable replacements."""

    file_path: str
    filename: str
    detected_variables: list[str] = field(default_factory=list)
    replacements: dict[str, str] = field(default_factory=dict)
    status: str = "就绪"


@dataclass
class MultiDocBatchResult:
    """Outcome of a multi-document batch replacement operation."""

    total_docs: int
    success_docs: int
    failed_docs: int
    errors: list[str] = field(default_factory=list)
    details: list[dict] = field(default_factory=list)
    backup_files: list[str] = field(default_factory=list)

