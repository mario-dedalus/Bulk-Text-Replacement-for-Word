"""Core document replacement engine and text processing services.

Deliberately decoupled from any GUI framework (zero Qt / Tkinter dependencies)
so all replacement, counting, preview, and formatting algorithms are 100% headless and testable.
"""

from __future__ import annotations

import os
import re
import shutil
from pathlib import Path
from typing import Callable, Optional

from docx import Document
from docx.text.paragraph import Paragraph

try:
    from core.models import BatchProcessResult, FileProcessDetail
except ImportError:
    from replacer_core import BatchProcessResult, FileProcessDetail  # type: ignore

# Soft hyphens, zero-width spaces, etc., inserted by Word for formatting
INVISIBLE_CHARS_TABLE = str.maketrans("", "", "\u00ad\u200b\u200c\u200d\u2060\ufeff")


def strip_invisible_chars(text: str) -> str:
    """Remove invisible formatting characters like soft hyphens and zero-width spaces."""
    if not text:
        return ""
    return text.translate(INVISIBLE_CHARS_TABLE)


def preprocess_text_with_nbsp(text: str) -> str:
    """Convert literal [NBSP] and non-breaking space characters into standard space for comparison."""
    if not text:
        return ""
    return text.replace("[NBSP]", "\u00a0").replace("&nbsp;", "\u00a0")


def count_occurrences(
    text: str,
    search_for: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
) -> int:
    """Count occurrences of search_for in text with options."""
    if not text or not search_for:
        return 0

    clean_text = strip_invisible_chars(text)
    clean_search = strip_invisible_chars(search_for)

    if use_regex:
        try:
            flags = 0 if case_sensitive else re.IGNORECASE
            return len(re.findall(clean_search, clean_text, flags))
        except re.error:
            return 0
    elif whole_word:
        pattern = r"\b" + re.escape(clean_search) + r"\b"
        flags = 0 if case_sensitive else re.IGNORECASE
        try:
            return len(re.findall(pattern, clean_text, flags))
        except re.error:
            return 0
    else:
        if case_sensitive:
            return clean_text.count(clean_search)
        else:
            return clean_text.lower().count(clean_search.lower())


def find_match_contexts(
    text: str,
    search_for: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    context_chars: int = 40,
    max_matches: int = 5,
) -> list[str]:
    """Find contextual snippets for matching occurrences."""
    if not text or not search_for:
        return []

    clean_text = strip_invisible_chars(text)
    clean_search = strip_invisible_chars(search_for)
    flags = 0 if case_sensitive else re.IGNORECASE

    if use_regex:
        try:
            pattern = re.compile(clean_search, flags)
        except re.error:
            return []
    elif whole_word:
        pattern = re.compile(r"\b" + re.escape(clean_search) + r"\b", flags)
    else:
        pattern = re.compile(re.escape(clean_search), flags)

    contexts = []
    for match in pattern.finditer(clean_text):
        if len(contexts) >= max_matches:
            break
        start = max(0, match.start() - context_chars)
        end = min(len(clean_text), match.end() + context_chars)

        prefix = ("..." if start > 0 else "") + clean_text[start:match.start()].replace("\n", " ")
        match_str = clean_text[match.start():match.end()].replace("\n", " ")
        suffix = clean_text[match.end():end].replace("\n", " ") + ("..." if end < len(clean_text) else "")

        contexts.append(f"{prefix}【{match_str}】{suffix}")

    return contexts


def _collect_table_text(table, text_list: list[str]) -> None:
    """Recursively collect text from table cells, including nested tables."""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                if p.text:
                    text_list.append(p.text)
            for nested in cell.tables:
                _collect_table_text(nested, text_list)


def get_document_text(doc: Document) -> str:
    """Extract all text from paragraphs and tables in a python-docx Document."""
    text_chunks: list[str] = []
    for p in doc.paragraphs:
        if p.text:
            text_chunks.append(p.text)
    for t in doc.tables:
        _collect_table_text(t, text_chunks)
    return "\n".join(text_chunks)


def replace_in_paragraph_advanced(
    paragraph: Paragraph,
    search_text: str,
    replace_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
) -> int:
    """Replace text in a paragraph that may span multiple runs, preserving formatting."""
    # Clean invisible characters from runs first
    for run in paragraph.runs:
        cleaned = strip_invisible_chars(run.text)
        if cleaned != run.text:
            run.text = cleaned

    paragraph_text = paragraph.text
    if not paragraph_text:
        return 0

    effective_regex = use_regex
    effective_search = search_text
    if whole_word and not use_regex:
        effective_search = r"\b" + re.escape(search_text) + r"\b"
        effective_regex = True

    # Count matches
    if effective_regex:
        try:
            flags = 0 if case_sensitive else re.IGNORECASE
            replacements = len(re.findall(effective_search, paragraph_text, flags))
        except re.error:
            return 0
        if replacements == 0:
            return 0
    else:
        if case_sensitive:
            if search_text not in paragraph_text:
                return 0
            replacements = paragraph_text.count(search_text)
        else:
            if search_text.lower() not in paragraph_text.lower():
                return 0
            replacements = paragraph_text.lower().count(search_text.lower())

    # Try simple replacement if contained within a single run
    for run in paragraph.runs:
        run_text = run.text
        if not run_text:
            continue
        if effective_regex:
            try:
                flags = 0 if case_sensitive else re.IGNORECASE
                if re.search(effective_search, run_text, flags):
                    run.text = re.sub(effective_search, replace_text, run_text, flags=flags)
                    return replacements
            except re.error:
                return 0
        elif case_sensitive:
            if search_text in run_text:
                run.text = run_text.replace(search_text, replace_text)
                return replacements
        else:
            if search_text.lower() in run_text.lower():
                pattern = re.escape(search_text)
                run.text = re.sub(pattern, replace_text, run_text, flags=re.IGNORECASE)
                return replacements

    # Text spans multiple runs — rebuild paragraph across runs
    if effective_regex:
        try:
            flags = 0 if case_sensitive else re.IGNORECASE
            new_paragraph_text = re.sub(effective_search, replace_text, paragraph_text, flags=flags)
        except re.error:
            return 0
    elif case_sensitive:
        new_paragraph_text = paragraph_text.replace(search_text, replace_text)
    else:
        pattern = re.escape(search_text)
        new_paragraph_text = re.sub(pattern, replace_text, paragraph_text, flags=re.IGNORECASE)

    # Clear all runs and assign rebuilt text to the first run
    for run in paragraph.runs:
        run.text = ""

    if paragraph.runs:
        paragraph.runs[0].text = new_paragraph_text
    else:
        paragraph.text = new_paragraph_text

    return replacements


def replace_in_table(
    table,
    search_text: str,
    replace_text: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
) -> int:
    """Replace text in a table recursively."""
    count = 0
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                count += replace_in_paragraph_advanced(
                    paragraph, search_text, replace_text, case_sensitive, use_regex, whole_word
                )
            for nested_table in cell.tables:
                count += replace_in_table(
                    nested_table, search_text, replace_text, case_sensitive, use_regex, whole_word
                )
    return count


def perform_standard_preview(
    file_paths: list[str],
    search_for: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Fast preview of main content across files using python-docx."""
    total_matches = 0
    files_with_matches = 0
    details: list[FileProcessDetail] = []
    errors: list[str] = []

    total_files = len(file_paths)
    for i, file_path in enumerate(file_paths):
        if is_cancelled and is_cancelled():
            break

        filename = os.path.basename(file_path)
        if progress_callback:
            progress_callback(i + 1, total_files, filename)

        detail = FileProcessDetail(filename=filename, file_path=file_path)
        try:
            doc = Document(file_path)
            doc_text = get_document_text(doc)
            matches = count_occurrences(doc_text, search_for, case_sensitive, use_regex, whole_word)
            detail.total = matches
            if matches > 0:
                files_with_matches += 1
                total_matches += matches
                detail.details.append(f"正文与表格：找到 {matches} 处匹配")
                detail.contexts = find_match_contexts(
                    doc_text, search_for, case_sensitive, use_regex, whole_word
                )
            else:
                detail.details.append("正文与表格：未找到匹配")
        except Exception as exc:
            err_msg = str(exc)
            detail.error = err_msg
            detail.details.append(f"错误：{err_msg}")
            errors.append(f"{filename}: {err_msg}")

        details.append(detail)

    return BatchProcessResult(
        files_processed=len(details),
        total_count=total_matches,
        files_with_matches=files_with_matches,
        successful_files=len(details) - len(errors),
        details=details,
        errors=errors,
    )


def perform_standard_replace(
    file_paths: list[str],
    search_for: str,
    replace_with: str,
    case_sensitive: bool = False,
    use_regex: bool = False,
    whole_word: bool = False,
    create_backup: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Execute fast replacement in main content across files using python-docx."""
    total_replacements = 0
    successful_files = 0
    details: list[FileProcessDetail] = []
    errors: list[str] = []
    backup_files: list[str] = []

    total_files = len(file_paths)
    for i, file_path in enumerate(file_paths):
        if is_cancelled and is_cancelled():
            break

        filename = os.path.basename(file_path)
        if progress_callback:
            progress_callback(i + 1, total_files, filename)

        detail = FileProcessDetail(filename=filename, file_path=file_path)
        try:
            if create_backup:
                backup_path = file_path + ".backup"
                shutil.copy2(file_path, backup_path)
                backup_files.append(backup_path)
                detail.backup_path = backup_path

            doc = Document(file_path)
            file_replacements = 0

            for paragraph in doc.paragraphs:
                file_replacements += replace_in_paragraph_advanced(
                    paragraph, search_for, replace_with, case_sensitive, use_regex, whole_word
                )

            for table in doc.tables:
                file_replacements += replace_in_table(
                    table, search_for, replace_with, case_sensitive, use_regex, whole_word
                )

            doc.save(file_path)

            detail.total = file_replacements
            if file_replacements > 0:
                detail.details.append(f"正文与表格替换完成：共替换 {file_replacements} 处")
            else:
                detail.details.append("无需替换")

            total_replacements += file_replacements
            successful_files += 1
        except Exception as exc:
            err_msg = str(exc)
            detail.error = err_msg
            detail.details.append(f"错误：{err_msg}")
            errors.append(f"{filename}: {err_msg}")

        details.append(detail)

    return BatchProcessResult(
        files_processed=len(details),
        total_count=total_replacements,
        files_with_matches=sum(1 for d in details if d.total > 0),
        successful_files=successful_files,
        details=details,
        errors=errors,
        backup_files=backup_files,
    )


# ---------------- COM Automation Helpers (Windows Only) ----------------


def _build_shape_ranges(doc) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []

    def _collect(shapes):
        for shape in shapes:
            try:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    ranges.append((shape.TextFrame.TextRange.Start, shape.TextFrame.TextRange.End))
            except Exception:
                pass
            try:
                if shape.Type == 6:  # msoGroup
                    _collect(shape.GroupItems)
            except Exception:
                pass

    try:
        _collect(doc.Shapes)
    except Exception:
        pass
    return ranges


def _find_replace_count_com(
    rng, search_text: str, replace_text: str, case_sensitive: bool, whole_word: bool = False
) -> int:
    count = 0
    rng.Find.ClearFormatting()
    rng.Find.Replacement.ClearFormatting()
    while True:
        found = rng.Find.Execute(
            FindText=search_text,
            ReplaceWith=replace_text,
            Replace=1,
            Forward=True,
            Wrap=0,
            MatchCase=case_sensitive,
            MatchWholeWord=whole_word,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Format=False,
        )
        if not found:
            break
        count += 1
    return count


def perform_com_preview(
    file_paths: list[str],
    search_for: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Comprehensive preview across all areas via Word COM automation."""
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("完整模式预览需要 Windows、Microsoft Word 和 pywin32。") from exc

    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False
    word_app.ScreenUpdating = False

    details: list[FileProcessDetail] = []
    errors: list[str] = []
    total_matches = 0
    files_with_matches = 0
    total_files = len(file_paths)

    try:
        for i, file_path in enumerate(file_paths):
            if is_cancelled and is_cancelled():
                break

            filename = os.path.basename(file_path)
            if progress_callback:
                progress_callback(i + 1, total_files, filename)

            detail = FileProcessDetail(filename=filename, file_path=file_path)
            doc = None
            try:
                full_path = os.path.abspath(file_path)
                doc = word_app.Documents.Open(full_path, ReadOnly=True)

                shapes_count = 0
                headers_count = 0
                footers_count = 0
                footnotes_count = 0
                endnotes_count = 0
                form_fields_count = 0
                hyperlinks_count = 0

                # 1. Text boxes / Shapes
                try:
                    story = doc.StoryRanges(5)  # wdTextFrameStory
                    while story:
                        if len(story.Text or "") > 1:
                            shapes_count += count_occurrences(
                                story.Text, search_for, case_sensitive, False, whole_word
                            )
                        try:
                            story = story.NextStoryRange
                        except Exception:
                            break
                except Exception:
                    pass

                # 2. Headers & Footers
                for section in doc.Sections:
                    for s_idx in range(1, 4):
                        try:
                            if section.Headers(s_idx).Exists:
                                h_text = section.Headers(s_idx).Range.Text or ""
                                if len(h_text) > 1:
                                    headers_count += count_occurrences(
                                        h_text, search_for, case_sensitive, False, whole_word
                                    )
                        except Exception:
                            pass
                        try:
                            if section.Footers(s_idx).Exists:
                                f_text = section.Footers(s_idx).Range.Text or ""
                                if len(f_text) > 1:
                                    footers_count += count_occurrences(
                                        f_text, search_for, case_sensitive, False, whole_word
                                    )
                        except Exception:
                            pass

                # 3. Footnotes & Endnotes
                for footnote in doc.Footnotes:
                    try:
                        fn_text = footnote.Range.Text or ""
                        if len(fn_text) > 1:
                            footnotes_count += count_occurrences(
                                fn_text, search_for, case_sensitive, False, whole_word
                            )
                    except Exception:
                        pass

                for endnote in doc.Endnotes:
                    try:
                        en_text = endnote.Range.Text or ""
                        if len(en_text) > 1:
                            endnotes_count += count_occurrences(
                                en_text, search_for, case_sensitive, False, whole_word
                            )
                    except Exception:
                        pass

                # 4. Form fields
                for ff in doc.FormFields:
                    try:
                        if ff.Type == 70:
                            ff_text = ff.Result or ""
                            if len(ff_text) > 0:
                                form_fields_count += count_occurrences(
                                    ff_text, search_for, case_sensitive, False, whole_word
                                )
                    except Exception:
                        pass

                # 5. Hyperlinks & Main Content
                shape_ranges = _build_shape_ranges(doc)
                for hl in doc.Hyperlinks:
                    try:
                        disp = hl.TextToDisplay or ""
                        if len(disp) > 0:
                            hl_range = hl.Range
                            if not (
                                hl_range
                                and any(
                                    hl_range.Start >= s and hl_range.End <= e for s, e in shape_ranges
                                )
                            ):
                                hyperlinks_count += count_occurrences(
                                    disp, search_for, case_sensitive, False, whole_word
                                )
                    except Exception:
                        pass

                main_count = 0
                try:
                    main_text = doc.Content.Text or ""
                    if len(main_text) > 1:
                        raw_main = count_occurrences(main_text, search_for, case_sensitive, False, whole_word)
                        main_count = max(0, raw_main - hyperlinks_count)
                except Exception:
                    pass

                file_total = (
                    shapes_count
                    + headers_count
                    + footers_count
                    + footnotes_count
                    + endnotes_count
                    + form_fields_count
                    + hyperlinks_count
                    + main_count
                )
                detail.total = file_total

                if file_total > 0:
                    files_with_matches += 1
                    total_matches += file_total
                    if main_count > 0:
                        detail.details.append(f"📄 正文：{main_count} 处匹配")
                    if shapes_count > 0:
                        detail.details.append(f"📦 文本框/形状：{shapes_count} 处匹配")
                    if headers_count > 0:
                        detail.details.append(f"📄 页眉：{headers_count} 处匹配")
                    if footers_count > 0:
                        detail.details.append(f"📄 页脚：{footers_count} 处匹配")
                    if footnotes_count > 0:
                        detail.details.append(f"📝 脚注：{footnotes_count} 处匹配")
                    if endnotes_count > 0:
                        detail.details.append(f"📝 尾注：{endnotes_count} 处匹配")
                    if form_fields_count > 0:
                        detail.details.append(f"📋 表单域：{form_fields_count} 处匹配")
                    if hyperlinks_count > 0:
                        detail.details.append(f"🔗 超链接：{hyperlinks_count} 处匹配")
                else:
                    detail.details.append("全部区域未找到匹配")
            except Exception as exc:
                err_msg = str(exc)
                detail.error = err_msg
                detail.details.append(f"错误：{err_msg}")
                errors.append(f"{filename}: {err_msg}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
            details.append(detail)
    finally:
        try:
            word_app.ScreenUpdating = True
            word_app.Quit()
        except Exception:
            pass

    return BatchProcessResult(
        files_processed=len(details),
        total_count=total_matches,
        files_with_matches=files_with_matches,
        successful_files=len(details) - len(errors),
        details=details,
        errors=errors,
    )


def perform_com_replace(
    file_paths: list[str],
    search_for: str,
    replace_with: str,
    case_sensitive: bool = False,
    whole_word: bool = False,
    create_backup: bool = False,
    progress_callback: Optional[Callable[[int, int, str], None]] = None,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> BatchProcessResult:
    """Comprehensive replacement across all document areas via Word COM automation."""
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("完整模式替换需要 Windows、Microsoft Word 和 pywin32。") from exc

    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False
    word_app.ScreenUpdating = False

    details: list[FileProcessDetail] = []
    errors: list[str] = []
    backup_files: list[str] = []
    total_replacements = 0
    successful_files = 0
    total_files = len(file_paths)

    try:
        for i, file_path in enumerate(file_paths):
            if is_cancelled and is_cancelled():
                break

            filename = os.path.basename(file_path)
            if progress_callback:
                progress_callback(i + 1, total_files, filename)

            detail = FileProcessDetail(filename=filename, file_path=file_path)
            doc = None
            try:
                if create_backup:
                    backup_path = file_path + ".backup"
                    shutil.copy2(file_path, backup_path)
                    backup_files.append(backup_path)
                    detail.backup_path = backup_path

                full_path = os.path.abspath(file_path)
                doc = word_app.Documents.Open(full_path)
                file_replacements = 0
                shape_ranges = _build_shape_ranges(doc)

                # 1. Text boxes (wdTextFrameStory = 5)
                shape_count = 0
                try:
                    story = doc.StoryRanges(5)
                    while story:
                        if len(story.Text or "") > 1:
                            shape_count += _find_replace_count_com(
                                story, search_for, replace_with, case_sensitive, whole_word
                            )
                        try:
                            story = story.NextStoryRange
                        except Exception:
                            break
                except Exception:
                    pass
                if shape_count > 0:
                    detail.details.append(f"📦 文本框/形状：替换 {shape_count} 处")
                    file_replacements += shape_count

                # 2. Headers & Footers
                header_count = 0
                footer_count = 0
                for section in doc.Sections:
                    for idx in range(1, 4):
                        try:
                            if section.Headers(idx).Exists and len(section.Headers(idx).Range.Text or "") > 1:
                                header_count += _find_replace_count_com(
                                    section.Headers(idx).Range, search_for, replace_with, case_sensitive, whole_word
                                )
                        except Exception:
                            pass
                        try:
                            if section.Footers(idx).Exists and len(section.Footers(idx).Range.Text or "") > 1:
                                footer_count += _find_replace_count_com(
                                    section.Footers(idx).Range, search_for, replace_with, case_sensitive, whole_word
                                )
                        except Exception:
                            pass
                if header_count > 0:
                    detail.details.append(f"📄 页眉：替换 {header_count} 处")
                    file_replacements += header_count
                if footer_count > 0:
                    detail.details.append(f"📄 页脚：替换 {footer_count} 处")
                    file_replacements += footer_count

                # 3. Footnotes & Endnotes
                fn_count = 0
                for fn in doc.Footnotes:
                    try:
                        if len(fn.Range.Text or "") > 1:
                            fn_count += _find_replace_count_com(
                                fn.Range, search_for, replace_with, case_sensitive, whole_word
                            )
                    except Exception:
                        pass
                if fn_count > 0:
                    detail.details.append(f"📝 脚注：替换 {fn_count} 处")
                    file_replacements += fn_count

                en_count = 0
                for en in doc.Endnotes:
                    try:
                        if len(en.Range.Text or "") > 1:
                            en_count += _find_replace_count_com(
                                en.Range, search_for, replace_with, case_sensitive, whole_word
                            )
                    except Exception:
                        pass
                if en_count > 0:
                    detail.details.append(f"📝 尾注：替换 {en_count} 处")
                    file_replacements += en_count

                # 4. Form fields
                ff_count = 0
                for field in doc.FormFields:
                    try:
                        if field.Type == 70:
                            orig = field.Result or ""
                            if len(orig) > 0:
                                clean = strip_invisible_chars(orig)
                                occ = count_occurrences(clean, search_for, case_sensitive, False, whole_word)
                                if occ > 0:
                                    if whole_word:
                                        pat = r"\b" + re.escape(search_for) + r"\b"
                                        fl = 0 if case_sensitive else re.IGNORECASE
                                        new_t = re.sub(pat, replace_with, clean, flags=fl)
                                    elif case_sensitive:
                                        new_t = clean.replace(search_for, replace_with)
                                    else:
                                        pat = re.escape(search_for)
                                        new_t = re.sub(pat, replace_with, clean, flags=re.IGNORECASE)
                                    field.Result = new_t
                                    ff_count += occ
                    except Exception:
                        pass
                if ff_count > 0:
                    detail.details.append(f"📋 表单域：替换 {ff_count} 处")
                    file_replacements += ff_count

                # 5. Main content
                main_count = _find_replace_count_com(
                    doc.Content, search_for, replace_with, case_sensitive, whole_word
                )

                # Report breakdown
                hl_count = 0
                for hl in doc.Hyperlinks:
                    try:
                        hl_r = hl.Range
                        if hl_r and not any(hl_r.Start >= s and hl_r.End <= e for s, e in shape_ranges):
                            disp = hl_r.Text or ""
                            if disp:
                                occ = count_occurrences(disp, replace_with, case_sensitive, False, whole_word) if replace_with else 0
                                hl_count += occ
                    except Exception:
                        pass

                non_hl_main = max(0, main_count - hl_count)
                if non_hl_main > 0:
                    detail.details.append(f"📄 正文：替换 {non_hl_main} 处")
                    file_replacements += non_hl_main
                if hl_count > 0:
                    detail.details.append(f"🔗 超链接：替换 {hl_count} 处")
                    file_replacements += hl_count

                doc.Save()
                doc.Close()
                doc = None

                detail.total = file_replacements
                if file_replacements == 0:
                    detail.details.append("无需替换")

                total_replacements += file_replacements
                successful_files += 1
            except Exception as exc:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
                    doc = None
                err_msg = str(exc)
                detail.error = err_msg
                detail.details.append(f"错误：{err_msg}")
                errors.append(f"{filename}: {err_msg}")

            details.append(detail)
    finally:
        try:
            word_app.ScreenUpdating = True
            word_app.Quit()
        except Exception:
            pass

    return BatchProcessResult(
        files_processed=len(details),
        total_count=total_replacements,
        files_with_matches=sum(1 for d in details if d.total > 0),
        successful_files=successful_files,
        details=details,
        errors=errors,
        backup_files=backup_files,
    )


def scan_hyperlinks(file_paths: list[str]) -> list[dict]:
    """Scan documents for hyperlinks and return analysis."""
    results = []
    for path in file_paths:
        try:
            doc = Document(path)
            urls = []
            for rel in doc.part.rels.values():
                if "hyperlink" in rel.reltype:
                    urls.append(rel.target_ref)
            results.append({
                "filename": os.path.basename(path),
                "count": len(urls),
                "urls": urls,
                "error": None,
            })
        except Exception as exc:
            results.append({
                "filename": os.path.basename(path),
                "count": 0,
                "urls": [],
                "error": str(exc),
            })
    return results
