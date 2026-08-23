"""PySide6 MainWindow for Bulk Text Replacement for Word with safe graceful shutdown."""

from __future__ import annotations

import sys
from platform_adapter.capabilities import CAPABILITIES
from ui.dialogs.help_dialog import HelpDialog
from ui.pages.merge_page import MergePage
from ui.pages.multi_doc_page import MultiDocPage
from ui.pages.replace_page import ReplacePage
from ui.theme.theme_manager import THEME
from ui.widgets.segmented_nav import SegmentedNav
from ui.widgets.status_bar import FluentStatusBar

try:
    from PySide6.QtCore import Qt, QThreadPool, QTimer
    from PySide6.QtGui import QAction, QCloseEvent, QKeySequence
    from PySide6.QtWidgets import (
        QApplication,
        QHBoxLayout,
        QLabel,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QStackedWidget,
        QVBoxLayout,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class MainWindow(QMainWindow):
        """Unified cross-platform desktop window combining text replacement, template merge, and multi-doc replace."""

        def __init__(self, initial_file: str | None = None):
            super().__init__()
            self.setWindowTitle("Word 批量处理工具 — Bulk Text Replacement for Word")
            self.resize(920, 800)
            self.setMinimumSize(760, 620)
            self.setAccessibleName("Word 批量处理工具主窗口")

            self._build_ui()
            self._setup_shortcuts()
            self._setup_system_theme_timer()

            THEME.apply_theme()

            if initial_file:
                if CAPABILITIES.is_macos and initial_file.lower().endswith(".doc"):
                    QMessageBox.warning(self, "格式不支持", "macOS 首版暂不支持旧版 .doc 格式文件。")
                else:
                    self.page_replace.file_model.add_files([initial_file])

        def _build_ui(self):
            central_widget = QWidget(self)
            self.setCentralWidget(central_widget)

            root_layout = QVBoxLayout(central_widget)
            root_layout.setContentsMargins(0, 0, 0, 0)
            root_layout.setSpacing(0)

            # ---------------- Top Bar ----------------
            top_bar = QWidget()
            top_bar_layout = QHBoxLayout(top_bar)
            top_bar_layout.setContentsMargins(20, 12, 20, 10)
            top_bar_layout.setSpacing(20)

            # Left: App title
            self.app_title = QLabel("Word 批量工具")
            self.app_title.setObjectName("AppTitle")
            self.app_title.setAccessibleName("应用程序标题")
            top_bar_layout.addWidget(self.app_title)

            # Middle: Segmented Navigation
            self.nav = SegmentedNav(
                tabs=[
                    ("replace", "文本查找替换"),
                    ("merge", "模板批量生成"),
                    ("multi_doc", "多文档匹配替换"),
                ],
                on_tab_change=self._on_tab_change,
            )
            top_bar_layout.addWidget(self.nav)

            top_bar_layout.addStretch()

            # Right: Theme Switcher & Help
            self.btn_theme = QPushButton("深色" if not THEME.is_dark else "浅色")
            self.btn_theme.setProperty("isGhost", True)
            self.btn_theme.setAccessibleName("切换深浅色主题按钮")
            self.btn_theme.clicked.connect(self._toggle_theme)
            top_bar_layout.addWidget(self.btn_theme)

            self.btn_help = QPushButton("使用帮助")
            self.btn_help.setProperty("isGhost", True)
            self.btn_help.setAccessibleName("使用指南与帮助说明按钮")
            self.btn_help.clicked.connect(self._open_help)
            top_bar_layout.addWidget(self.btn_help)

            root_layout.addWidget(top_bar)

            # ---------------- Central Pages ----------------
            self.stacked_widget = QStackedWidget()
            self.page_replace = ReplacePage()
            self.page_merge = MergePage()
            self.page_multi_doc = MultiDocPage()

            self.stacked_widget.addWidget(self.page_replace)
            self.stacked_widget.addWidget(self.page_merge)
            self.stacked_widget.addWidget(self.page_multi_doc)

            root_layout.addWidget(self.stacked_widget, 1)

            # ---------------- Bottom Status Bar ----------------
            self.status_bar = FluentStatusBar(self)
            self.status_bar.cancelRequested.connect(self._cancel_active_task)
            root_layout.addWidget(self.status_bar)

            # Connect status signals from Replace Page
            self.page_replace.statusMessage.connect(self.status_bar.set_status)
            self.page_replace.statusMetrics.connect(self.status_bar.set_metrics)
            self.page_replace.statusProgress.connect(self.status_bar.update_progress)
            self.page_replace.statusHideProgress.connect(self.status_bar.hide_progress)

            # Connect status signals from Merge Page
            self.page_merge.statusMessage.connect(self.status_bar.set_status)
            self.page_merge.statusMetrics.connect(self.status_bar.set_metrics)
            self.page_merge.statusProgress.connect(self.status_bar.update_progress)
            self.page_merge.statusHideProgress.connect(self.status_bar.hide_progress)

            # Connect status signals from MultiDoc Page
            self.page_multi_doc.statusMessage.connect(self.status_bar.set_status)
            self.page_multi_doc.statusMetrics.connect(self.status_bar.set_metrics)
            self.page_multi_doc.statusProgress.connect(self.status_bar.update_progress)
            self.page_multi_doc.statusHideProgress.connect(self.status_bar.hide_progress)

        def _setup_system_theme_timer(self):
            self._theme_timer = QTimer(self)
            self._theme_timer.setInterval(2000)
            self._theme_timer.timeout.connect(self._check_system_theme)
            self._theme_timer.start()

        def _check_system_theme(self):
            changed = THEME.check_system_theme_update()
            if changed:
                self.btn_theme.setText("浅色" if THEME.is_dark else "深色")

        def _on_tab_change(self, tab_id: str):
            if tab_id == "replace":
                self.stacked_widget.setCurrentWidget(self.page_replace)
                count = self.page_replace.file_model.count()
                self.status_bar.set_status("就绪", "normal")
                self.status_bar.set_metrics(f"已加载 {count} 个文件" if count > 0 else "")
            elif tab_id == "merge":
                self.stacked_widget.setCurrentWidget(self.page_merge)
                self.status_bar.set_status("就绪", "normal")
                if self.page_merge.excel_data:
                    self.status_bar.set_metrics(f"已加载 {len(self.page_merge.excel_data.rows)} 行数据")
                else:
                    self.status_bar.set_metrics("")
            elif tab_id == "multi_doc":
                self.stacked_widget.setCurrentWidget(self.page_multi_doc)
                count = self.page_multi_doc.mapping_model.rowCount()
                self.status_bar.set_status("就绪", "normal")
                self.status_bar.set_metrics(f"已加载 {count} 个文档" if count > 0 else "")

        def _toggle_theme(self):
            is_dark = THEME.toggle()
            self.btn_theme.setText("浅色" if is_dark else "深色")

        def _open_help(self):
            dialog = HelpDialog(self)
            dialog.exec()

        def _cancel_active_task(self):
            self.page_replace.cancel_active_task()
            self.page_merge.cancel_active_task()
            self.page_multi_doc.cancel_active_task()

        def closeEvent(self, event: QCloseEvent):
            thread_pool = QThreadPool.globalInstance()
            if thread_pool.activeThreadCount() > 0:
                reply = QMessageBox.question(
                    self,
                    "任务正在运行",
                    "当前有后台批量任务正在运行中。关闭窗口将请求取消并等待当前文件安全完成，确定要退出吗？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    self.status_bar.set_status("正在安全停止后台任务...", "warning")
                    self._cancel_active_task()

                    # Wait up to 3 seconds for safe file write completion
                    finished_cleanly = thread_pool.waitForDone(3000)
                    if finished_cleanly:
                        event.accept()
                    else:
                        force_reply = QMessageBox.warning(
                            self,
                            "任务尚未完全停止",
                            "后台任务正在完成当前文件的磁盘保存操作。是否强制立即退出？",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.No,
                        )
                        if force_reply == QMessageBox.Yes:
                            event.accept()
                        else:
                            event.ignore()
                else:
                    event.ignore()
            else:
                event.accept()



        def _setup_shortcuts(self):
            action_preview = QAction(self)
            action_preview.setShortcut(QKeySequence("Ctrl+P" if CAPABILITIES.is_windows else "Cmd+P"))
            action_preview.triggered.connect(self._on_shortcut_preview)
            self.addAction(action_preview)

            action_run = QAction(self)
            action_run.setShortcut(QKeySequence("Ctrl+R" if CAPABILITIES.is_windows else "Cmd+R"))
            action_run.triggered.connect(self._on_shortcut_run)
            self.addAction(action_run)

        def _on_shortcut_preview(self):
            current = self.stacked_widget.currentWidget()
            if current == self.page_replace:
                self.page_replace.preview_changes()
            elif current == self.page_merge:
                self.page_merge.preview_merge()
            elif current == self.page_multi_doc:
                self.page_multi_doc.preview_changes()

        def _on_shortcut_run(self):
            current = self.stacked_widget.currentWidget()
            if current == self.page_replace:
                self.page_replace.start_replace()
            elif current == self.page_merge:
                self.page_merge.start_batch_merge()
            elif current == self.page_multi_doc:
                self.page_multi_doc.start_replace()


else:

    class MainWindow:  # type: ignore
        pass
