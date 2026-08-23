"""Text search and replace page with Fluent Design controls, responsive grid, and zero horizontal overflow."""

from __future__ import annotations

import os

from application.replace_service import ReplaceService
from application.task_models import ServiceResult, TaskProgress, TaskState
from application.workers import TaskWorker
from core.models import BatchProcessResult
from core.replacer_core import count_occurrences, get_document_text
from platform_adapter.capabilities import CAPABILITIES
from ui.dialogs.result_dialog import ResultDialog
from ui.models.file_list_model import FileListModel
from ui.theme.theme_manager import THEME
from ui.widgets.card import FluentCard

try:
    from PySide6.QtCore import Qt, QThreadPool, Signal
    from PySide6.QtWidgets import (
        QApplication,
        QButtonGroup,
        QCheckBox,
        QFileDialog,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QListView,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QRadioButton,
        QScrollArea,
        QVBoxLayout,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class ReplacePage(QWidget):
        """View for batch finding and replacing text across multiple Word documents."""

        statusMessage = Signal(str, str)  # (text, state)
        statusMetrics = Signal(str)        # metrics string
        statusProgress = Signal(object)    # TaskProgress
        statusHideProgress = Signal()

        def __init__(self, parent: QWidget | None = None):
            super().__init__(parent)
            self._text_cache: dict[str, str] = {}
            self._cache_generation_id: int = 0
            self.file_model = FileListModel(parent=self)
            self._thread_pool = QThreadPool.globalInstance()
            self._active_worker: TaskWorker | None = None

            self._build_ui()
            self.file_model.rowsInserted.connect(self._on_files_changed)
            self.file_model.rowsRemoved.connect(self._on_files_changed)
            self.file_model.modelReset.connect(self._on_files_changed)

        def _build_ui(self):
            root_layout = QVBoxLayout(self)
            root_layout.setContentsMargins(0, 0, 0, 0)
            root_layout.setSpacing(0)

            # ---------------- Scroll Area for Content Cards ----------------
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setFrameShape(QScrollArea.NoFrame)
            self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.scroll_area.setAccessibleName("文本替换内容滚动区")

            self.content_widget = QWidget()
            self.content_widget.setProperty("isPageContent", True)
            layout = QVBoxLayout(self.content_widget)
            layout.setContentsMargins(18, 10, 18, 12)
            layout.setSpacing(12)

            # ---------------- Card 1: Selected Word Documents ----------------
            subtitle = "支持批量添加 .docx / .docm 文件" if CAPABILITIES.is_macos else "支持批量添加 .docx / .doc / .docm 文件"
            self.card_files = FluentCard(title="已选 Word 文档", subtitle=subtitle)
            self.card_files.setAccessibleName("已选 Word 文档卡片")

            list_row = QHBoxLayout()
            self.file_view = QListView()
            self.file_view.setModel(self.file_model)
            self.file_view.setSelectionMode(QListView.ExtendedSelection)
            self.file_view.setFixedHeight(105)
            self.file_view.setAccessibleName("Word 文档列表")
            self.file_view.setAccessibleDescription("展示当前已添加的待处理 Word 文件列表，按 Delete 键可移除所选")
            list_row.addWidget(self.file_view)
            self.card_files.addLayout(list_row)

            btn_row = QHBoxLayout()
            btn_row.setSpacing(8)

            btn_add = QPushButton("添加文件...")
            btn_add.setAccessibleName("添加文件按钮")
            btn_add.setAccessibleDescription("打开文件选择对话框添加 Word 文档")
            btn_add.clicked.connect(self._browse_files)
            btn_row.addWidget(btn_add)

            btn_remove = QPushButton("移除所选")
            btn_remove.setAccessibleName("移除所选文档按钮")
            btn_remove.setAccessibleDescription("从列表中移除当前选中的文档")
            btn_remove.clicked.connect(self._remove_selected_files)
            btn_row.addWidget(btn_remove)

            btn_clear = QPushButton("清空列表")
            btn_clear.setAccessibleName("清空列表按钮")
            btn_clear.setAccessibleDescription("清空当前列表中所有添加的文档")
            btn_clear.clicked.connect(self._clear_files)
            btn_row.addWidget(btn_clear)

            btn_links = QPushButton("检查超链接")
            btn_links.setAccessibleName("检查超链接按钮")
            btn_links.setAccessibleDescription("扫描并报告所选文档中的所有超链接地址")
            btn_links.clicked.connect(self._check_hyperlinks)
            btn_row.addWidget(btn_links)

            btn_row.addStretch()

            self.file_count_lbl = QLabel("已选 0 个文件")
            self.file_count_lbl.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
            self.file_count_lbl.setAccessibleName("已选文件计数")
            btn_row.addWidget(self.file_count_lbl)

            self.card_files.addLayout(btn_row)
            layout.addWidget(self.card_files)

            # ---------------- Card 2: Find & Replace Inputs ----------------
            self.card_inputs = FluentCard(title="查找与替换")
            self.card_inputs.setAccessibleName("查找与替换输入卡片")

            # Search Header
            search_hdr = QHBoxLayout()
            lbl_search = QLabel("查找内容：")
            lbl_search.setAccessibleName("查找内容标签")
            search_hdr.addWidget(lbl_search)
            search_hdr.addStretch()

            btn_search_nbsp = QPushButton("NBSP 检查")
            btn_search_nbsp.setProperty("isGhost", True)
            btn_search_nbsp.setAccessibleName("不间断空格检测按钮")
            btn_search_nbsp.clicked.connect(self._check_nbsp)
            search_hdr.addWidget(btn_search_nbsp)

            btn_search_paste = QPushButton("粘贴")
            btn_search_paste.setProperty("isGhost", True)
            btn_search_paste.setAccessibleName("粘贴到查找输入框按钮")
            btn_search_paste.clicked.connect(self._paste_search)
            search_hdr.addWidget(btn_search_paste)

            self.card_inputs.addLayout(search_hdr)

            self.txt_search = QPlainTextEdit()
            self.txt_search.setFixedHeight(60)
            self.txt_search.setPlaceholderText("输入要查找的文本或正则表达式...")
            self.txt_search.setAccessibleName("查找内容输入框")
            self.txt_search.setAccessibleDescription("输入要查找的关键词或正则表达式")
            lbl_search.setBuddy(self.txt_search)
            self.txt_search.textChanged.connect(self._on_search_text_changed)
            self.card_inputs.addWidget(self.txt_search)

            # Replace Header
            replace_hdr = QHBoxLayout()
            lbl_replace = QLabel("替换为：")
            lbl_replace.setAccessibleName("替换为标签")
            replace_hdr.addWidget(lbl_replace)
            replace_hdr.addStretch()

            btn_replace_paste = QPushButton("粘贴")
            btn_replace_paste.setProperty("isGhost", True)
            btn_replace_paste.setAccessibleName("粘贴到替换输入框按钮")
            btn_replace_paste.clicked.connect(self._paste_replace)
            replace_hdr.addWidget(btn_replace_paste)

            self.card_inputs.addLayout(replace_hdr)

            self.txt_replace = QPlainTextEdit()
            self.txt_replace.setFixedHeight(60)
            self.txt_replace.setPlaceholderText("输入替换后的新文本...")
            self.txt_replace.setAccessibleName("替换内容输入框")
            self.txt_replace.setAccessibleDescription("输入替换后的目标文本")
            lbl_replace.setBuddy(self.txt_replace)
            self.card_inputs.addWidget(self.txt_replace)

            self.match_count_lbl = QLabel("")
            self.match_count_lbl.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-style: italic; font-size: 12px;")
            self.match_count_lbl.setAccessibleName("实时匹配统计标签")
            self.card_inputs.addWidget(self.match_count_lbl)

            layout.addWidget(self.card_inputs)

            # ---------------- Card 3: Processing Options (Responsive Grid) ----------------
            self.card_options = FluentCard(title="处理选项")
            self.card_options.setAccessibleName("处理选项卡片")

            # Modes: Stacked Vertically for clean responsive width
            mode_box = QVBoxLayout()
            mode_box.setSpacing(6)

            self.rb_fast = QRadioButton("快速模式（推荐 — 基于 python-docx，速度快）")
            self.rb_fast.setChecked(True)
            self.rb_fast.setAccessibleName("快速模式单选按钮")
            self.rb_fast.toggled.connect(self._on_mode_toggled)
            mode_box.addWidget(self.rb_fast)

            com_state_hint = "" if CAPABILITIES.has_word_com else "（需 Windows 与 Microsoft Word）"
            self.rb_full = QRadioButton(f"完整模式{com_state_hint}（保留超链接/形状/页眉页脚）")
            self.rb_full.setEnabled(CAPABILITIES.has_word_com)
            self.rb_full.setAccessibleName("完整模式单选按钮")
            mode_box.addWidget(self.rb_full)

            self.card_options.addLayout(mode_box)

            # Checkbox options in 2x2 Grid to fit narrow widths comfortably
            grid_opts = QGridLayout()
            grid_opts.setContentsMargins(0, 4, 0, 0)
            grid_opts.setHorizontalSpacing(24)
            grid_opts.setVerticalSpacing(8)

            self.cb_backup = QCheckBox("创建备份文件 (.backup)")
            self.cb_backup.setChecked(True)
            self.cb_backup.setAccessibleName("创建备份文件复选框")
            grid_opts.addWidget(self.cb_backup, 0, 0)

            self.cb_case = QCheckBox("区分大小写")
            self.cb_case.setAccessibleName("区分大小写复选框")
            self.cb_case.toggled.connect(self._on_search_text_changed)
            grid_opts.addWidget(self.cb_case, 0, 1)

            self.cb_whole_word = QCheckBox("全字匹配")
            self.cb_whole_word.setAccessibleName("全字匹配复选框")
            self.cb_whole_word.toggled.connect(self._on_search_text_changed)
            grid_opts.addWidget(self.cb_whole_word, 1, 0)

            self.cb_regex = QCheckBox("正则表达式（仅限快速模式）")
            self.cb_regex.setAccessibleName("正则表达式复选框")
            self.cb_regex.toggled.connect(self._on_search_text_changed)
            grid_opts.addWidget(self.cb_regex, 1, 1)

            self.card_options.addLayout(grid_opts)
            layout.addWidget(self.card_options)

            self.scroll_area.setWidget(self.content_widget)
            root_layout.addWidget(self.scroll_area, 1)

            # ---------------- Fixed Bottom Action Bar (Outside ScrollArea) ----------------
            action_container = QFrame(self)
            action_container.setStyleSheet(f"""
                QFrame {{
                    background-color: {THEME.tokens.card_bg};
                    border-top: 1px solid {THEME.tokens.border_subtle};
                    padding: 8px 18px;
                }}
            """)
            action_bar = QHBoxLayout(action_container)
            action_bar.setContentsMargins(0, 0, 0, 0)
            action_bar.setSpacing(10)

            self.btn_preview = QPushButton("预览更改")
            self.btn_preview.setAccessibleName("预览更改按钮")
            self.btn_preview.clicked.connect(self.preview_changes)
            action_bar.addWidget(self.btn_preview)

            action_bar.addStretch()

            self.btn_replace = QPushButton("开始替换")
            self.btn_replace.setProperty("isPrimary", True)
            self.btn_replace.setAccessibleName("开始替换主按钮")
            self.btn_replace.clicked.connect(self.start_replace)
            action_bar.addWidget(self.btn_replace)

            root_layout.addWidget(action_container)

        # ---------------- Actions & Event Handlers ----------------

        def _on_mode_toggled(self, checked: bool):
            if not checked:  # Full mode selected
                self.cb_regex.setChecked(False)
                self.cb_regex.setEnabled(False)
            else:  # Fast mode selected
                self.cb_regex.setEnabled(True)
            self._on_search_text_changed()

        def _browse_files(self):
            filter_str = (
                "Word 文档 (*.docx *.docm);;所有文件 (*.*)"
                if CAPABILITIES.is_macos
                else "Word 文档 (*.docx *.docm *.doc);;所有文件 (*.*)"
            )
            files, _ = QFileDialog.getOpenFileNames(self, "选择 Word 文档", "", filter_str)
            if files:
                valid_files = []
                for f in files:
                    if CAPABILITIES.is_macos and f.lower().endswith(".doc"):
                        QMessageBox.warning(self, "格式不支持", f"macOS 首版暂不支持旧版 .doc 格式文件：{os.path.basename(f)}")
                        continue
                    valid_files.append(f)
                if valid_files:
                    self.file_model.add_files(valid_files)

        def _remove_selected_files(self):
            indexes = self.file_view.selectionModel().selectedIndexes()
            if indexes:
                rows = [idx.row() for idx in indexes]
                self.file_model.remove_indices(rows)

        def _clear_files(self):
            if self.file_model.count() == 0:
                return
            reply = QMessageBox.question(
                self, "确认清空", f"确定要清空已选的 {self.file_model.count()} 个文档吗？"
            )
            if reply == QMessageBox.Yes:
                self.file_model.clear()

        def _on_files_changed(self):
            count = self.file_model.count()
            self.file_count_lbl.setText(f"已选 {count} 个文件")
            self._async_refresh_text_cache()
            self.statusMetrics.emit(f"已加载 {count} 个文件" if count > 0 else "")

        def _async_refresh_text_cache(self):
            self._cache_generation_id += 1
            current_gen = self._cache_generation_id
            paths = self.file_model.get_all_paths()

            def _worker_fn():
                cache = {}
                for path in paths:
                    if path.lower().endswith((".docx", ".docm")):
                        try:
                            from docx import Document

                            doc = Document(path)
                            cache[path] = get_document_text(doc)
                        except Exception:
                            cache[path] = ""
                return (cache, current_gen)

            worker = TaskWorker(_worker_fn)
            worker.signals.result.connect(self._on_cache_refreshed)
            self._thread_pool.start(worker)

        def _on_cache_refreshed(self, result: tuple[dict[str, str], int]):
            cache, gen_id = result
            if gen_id == self._cache_generation_id:
                self._text_cache = cache
                self._on_search_text_changed()

        def _on_search_text_changed(self):
            search_text = self.txt_search.toPlainText().rstrip("\n")
            if not search_text or self.file_model.count() == 0:
                self.match_count_lbl.setText("")
                return

            case_sensitive = self.cb_case.isChecked()
            use_regex = self.cb_regex.isChecked() and self.rb_fast.isChecked()
            whole_word = self.cb_whole_word.isChecked()

            total = 0
            matching_files = 0
            for path in self.file_model.get_all_paths():
                text = self._text_cache.get(path, "")
                if text:
                    occ = count_occurrences(text, search_text, case_sensitive, use_regex, whole_word)
                    if occ > 0:
                        total += occ
                        matching_files += 1

            msg = f"实时统计：在 {matching_files} 个文件中找到 {total} 处匹配"
            self.match_count_lbl.setText(msg)
            self.statusMetrics.emit(f"{self.file_model.count()} 个文件 | {total} 处匹配")

        def _paste_search(self):
            clip = QApplication.clipboard()
            if clip:
                self.txt_search.setPlainText(clip.text())

        def _paste_replace(self):
            clip = QApplication.clipboard()
            if clip:
                self.txt_replace.setPlainText(clip.text())

        def _check_nbsp(self):
            text = self.txt_search.toPlainText()
            nbsp_count = text.count("\u00a0")
            literal_count = text.count("[NBSP]") + text.count("&nbsp;")
            QMessageBox.information(
                self,
                "NBSP 空格检测",
                f"当前查找字符数：{len(text)}\n"
                f"不间断空格 (U+00A0)：{nbsp_count} 个\n"
                f"NBSP 占位符：{literal_count} 个\n\n"
                "提示：程序在查找和替换时会自动规范化处理不间断空格。",
            )

        def _check_hyperlinks(self):
            paths = self.file_model.get_all_paths()
            if not paths:
                QMessageBox.warning(self, "提示", "请先添加至少一个 Word 文档。")
                return

            self._set_ui_busy(True)
            self.statusMessage.emit("正在扫描超链接...", "running")

            worker = TaskWorker(ReplaceService.scan_links, file_paths=paths)
            self._active_worker = worker

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_links_scanned)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_ui_busy(False))
            self._thread_pool.start(worker)

        def _on_links_scanned(self, res: ServiceResult[list[dict]]):
            self.statusHideProgress.emit()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("超链接扫描已取消", "warning")
                    return
                self.statusMessage.emit("扫描异常", "error")
                QMessageBox.critical(self, "扫描异常", res.error)
                return

            paths = self.file_model.get_all_paths()
            results = res.data or []
            self.statusMessage.emit("超链接扫描完成", "success")

            total_links = sum(r["count"] for r in results)
            files_with_links = sum(1 for r in results if r["count"] > 0)

            lines = [
                "文档超链接扫描报告",
                "=" * 50,
                f"扫描文件总数：{len(paths)}",
                f"含超链接文件：{files_with_links}",
                f"超链接总数量：{total_links}",
                "=" * 50,
                "",
            ]
            for r in results:
                lines.append(f"📄 {r['filename']}：{r['count']} 个链接")
                for url in r["urls"][:5]:
                    lines.append(f"   • {url}")
                if len(r["urls"]) > 5:
                    lines.append(f"   • ... 以及其他 {len(r['urls']) - 5} 个链接")
                lines.append("")

            if total_links > 0:
                lines.append("提示：若需替换超链接显示文字且保留链接地址，请在处理选项中选择【完整模式】。")

            dialog = ResultDialog(
                title="超链接检查结果",
                summary_text=f"共扫描 {len(paths)} 个文档，发现 {total_links} 处超链接。",
                details_text="\n".join(lines),
                is_success=True,
                parent=self,
            )
            dialog.exec()

        # ---------------- Asynchronous Execution ----------------

        def preview_changes(self):
            paths = self.file_model.get_all_paths()
            search_text = self.txt_search.toPlainText().rstrip("\n")
            err = ReplaceService.validate_inputs(paths, search_text)
            if err:
                QMessageBox.warning(self, "提示", err)
                return

            mode = "full" if self.rb_full.isChecked() else "fast"
            case_sensitive = self.cb_case.isChecked()
            use_regex = self.cb_regex.isChecked() and mode == "fast"
            whole_word = self.cb_whole_word.isChecked()

            self._set_ui_busy(True)
            self.statusMessage.emit("正在生成预览...", "running")

            worker = TaskWorker(
                ReplaceService.execute_preview,
                file_paths=paths,
                search_text=search_text,
                mode=mode,
                case_sensitive=case_sensitive,
                use_regex=use_regex,
                whole_word=whole_word,
            )
            self._active_worker = worker

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_preview_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_ui_busy(False))

            self._thread_pool.start(worker)

        def _on_preview_finished(self, res: ServiceResult[BatchProcessResult]):
            self.statusHideProgress.emit()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("预览已取消", "warning")
                    return
                self.statusMessage.emit("预览失败", "error")
                QMessageBox.critical(self, "预览失败", res.error)
                return

            data = res.data
            mode_name = "完整模式" if self.rb_full.isChecked() else "快速模式"
            lines = [
                f"查找替换预览报告（{mode_name}）",
                "=" * 60,
                f"查找内容：{self.txt_search.toPlainText()}",
                f"替换为：{self.txt_replace.toPlainText()}",
                f"匹配总数：{data.total_count} 处（分布于 {data.files_with_matches} 个文件）",
                "=" * 60,
                "",
            ]
            for detail in data.details:
                lines.append(f"📁 文件：{detail.filename}")
                lines.append(f"   匹配数量：{detail.total}")
                for d in detail.details:
                    lines.append(f"   {d}")
                if detail.contexts:
                    lines.append("   📝 匹配上下文样例：")
                    for ctx in detail.contexts:
                        lines.append(f"      {ctx}")
                lines.append("")

            self.statusMessage.emit("预览完成", "success")
            dialog = ResultDialog(
                title="查找替换预览结果",
                summary_text=f"在 {data.files_processed} 个文件中找到 {data.total_count} 处匹配（涉及 {data.files_with_matches} 个文件）。",
                details_text="\n".join(lines),
                is_success=True,
                parent=self,
            )
            dialog.exec()

        def start_replace(self):
            paths = self.file_model.get_all_paths()
            search_text = self.txt_search.toPlainText().rstrip("\n")
            replace_text = self.txt_replace.toPlainText().rstrip("\n")

            err = ReplaceService.validate_inputs(paths, search_text)
            if err:
                QMessageBox.warning(self, "提示", err)
                return

            mode = "full" if self.rb_full.isChecked() else "fast"
            mode_name = "完整模式" if mode == "full" else "快速模式"
            create_backup = self.cb_backup.isChecked()
            backup_hint = "（将创建 .backup 备份）" if create_backup else "（未勾选备份）"

            reply = QMessageBox.question(
                self,
                "确认执行替换",
                f"即将使用【{mode_name}】在 {len(paths)} 个文档中执行查找替换 {backup_hint}。\n\n"
                f"查找：{search_text[:40]}\n"
                f"替换：{replace_text[:40]}\n\n"
                "是否继续？",
            )
            if reply != QMessageBox.Yes:
                return

            case_sensitive = self.cb_case.isChecked()
            use_regex = self.cb_regex.isChecked() and mode == "fast"
            whole_word = self.cb_whole_word.isChecked()

            self._set_ui_busy(True)
            self.statusMessage.emit("正在执行替换...", "running")

            worker = TaskWorker(
                ReplaceService.execute_replace,
                file_paths=paths,
                search_text=search_text,
                replace_text=replace_text,
                mode=mode,
                case_sensitive=case_sensitive,
                use_regex=use_regex,
                whole_word=whole_word,
                create_backup=create_backup,
            )
            self._active_worker = worker

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_replace_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_ui_busy(False))

            self._thread_pool.start(worker)

        def _on_replace_finished(self, res: ServiceResult[BatchProcessResult]):
            self.statusHideProgress.emit()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("替换任务已取消", "warning")
                    return
                self.statusMessage.emit("替换失败", "error")
                QMessageBox.critical(self, "替换失败", res.error)
                return

            data = res.data
            self._async_refresh_text_cache()

            mode_name = "完整模式" if self.rb_full.isChecked() else "快速模式"
            lines = [
                f"查找替换执行报告（{mode_name}）",
                "=" * 60,
                f"成功处理文件：{data.successful_files} / {data.files_processed}",
                f"替换总次数：{data.total_count}",
                f"已创建备份：{len(data.backup_files)} 个文件",
                "=" * 60,
                "",
            ]
            for detail in data.details:
                lines.append(f"📁 文件：{detail.filename}")
                for d in detail.details:
                    lines.append(f"   {d}")
                lines.append("")

            if data.errors:
                lines.append("❌ 错误列表：")
                for err in data.errors:
                    lines.append(f"• {err}")
                self.statusMessage.emit("替换完成（存在部分错误）", "warning")
            else:
                self.statusMessage.emit("替换完成", "success")

            dialog = ResultDialog(
                title="替换完成",
                summary_text=f"处理完成！成功替换 {data.total_count} 处，涉及 {data.successful_files} 个文件。",
                details_text="\n".join(lines),
                is_success=not bool(data.errors),
                parent=self,
            )
            dialog.exec()

        def _on_task_error(self, err_msg: str):
            self.statusHideProgress.emit()
            self.statusMessage.emit("任务异常", "error")
            QMessageBox.critical(self, "执行异常", err_msg)

        def cancel_active_task(self):
            if self._active_worker:
                self._active_worker.cancel()
                self.statusMessage.emit("正在请求取消任务...", "warning")

        def _set_ui_busy(self, is_busy: bool):
            self.btn_preview.setEnabled(not is_busy)
            self.btn_replace.setEnabled(not is_busy)

else:

    class ReplacePage:  # type: ignore
        pass
