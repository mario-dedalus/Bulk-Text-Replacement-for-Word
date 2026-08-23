"""Fluent status bar widget with status message, file metrics, progress, and cancellation."""

from __future__ import annotations

from application.task_models import TaskProgress
from ui.theme.theme_manager import THEME
from ui.theme.tokens import ColorTokens

try:
    from PySide6.QtCore import Qt, Signal
    from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QProgressBar, QPushButton, QWidget

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FluentStatusBar(QFrame):
        """Unified status bar attached at the bottom of the main window."""

        cancelRequested = Signal()

        def __init__(self, parent: QWidget | None = None):
            super().__init__(parent)
            self.setObjectName("FluentStatusBar")
            self.setAccessibleName("应用程序状态栏")
            self.setAccessibleDescription("显示当前运行状态、统计数据与后台任务进度")

            layout = QHBoxLayout(self)
            layout.setContentsMargins(16, 6, 16, 6)
            layout.setSpacing(12)

            # Status dot + text
            self.dot_label = QLabel("●")
            self.dot_label.setStyleSheet(f"color: {THEME.tokens.success}; font-size: 11px;")
            self.dot_label.setAccessibleName("状态指示灯")
            layout.addWidget(self.dot_label)

            self.status_label = QLabel("就绪")
            self.status_label.setStyleSheet("font-weight: 500;")
            self.status_label.setAccessibleName("当前状态文本")
            layout.addWidget(self.status_label)

            # Metrics
            self.metrics_label = QLabel("")
            self.metrics_label.setStyleSheet(f"color: {THEME.tokens.text_secondary};")
            self.metrics_label.setAccessibleName("统计指标")
            layout.addWidget(self.metrics_label)

            layout.addStretch()

            # Progress Bar and progress message
            self.progress_msg = QLabel("")
            self.progress_msg.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 11px;")
            self.progress_msg.setVisible(False)
            self.progress_msg.setAccessibleName("任务进度说明")
            layout.addWidget(self.progress_msg)

            self.progress_bar = QProgressBar()
            self.progress_bar.setFixedWidth(160)
            self.progress_bar.setVisible(False)
            self.progress_bar.setAccessibleName("任务进度条")
            layout.addWidget(self.progress_bar)

            # Cancel button for running tasks
            self.btn_cancel = QPushButton("取消")
            self.btn_cancel.setVisible(False)
            self.btn_cancel.setCursor(Qt.PointingHandCursor)
            self._update_cancel_btn_style(THEME.tokens)
            self.btn_cancel.setAccessibleName("取消当前任务按钮")
            self.btn_cancel.clicked.connect(self.cancelRequested.emit)
            layout.addWidget(self.btn_cancel)

            # Listen to theme changes
            THEME.add_listener(self._on_theme_changed)

        def _update_cancel_btn_style(self, tokens: ColorTokens):
            self.btn_cancel.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    color: {tokens.error};
                    border: 1px solid {tokens.error};
                    border-radius: 4px;
                    padding: 2px 10px;
                    font-size: 11px;
                }}
                QPushButton:hover {{
                    background-color: rgba(196, 43, 28, 0.15);
                }}
            """)

        def _on_theme_changed(self, tokens: ColorTokens):
            self._update_cancel_btn_style(tokens)
            self.metrics_label.setStyleSheet(f"color: {tokens.text_secondary};")
            self.progress_msg.setStyleSheet(f"color: {tokens.text_secondary}; font-size: 11px;")

        def set_status(self, text: str, state: str = "normal") -> None:
            """state: 'normal', 'running', 'success', 'warning', 'error'."""
            self.status_label.setText(text)
            t = THEME.tokens
            if state == "running":
                self.dot_label.setText("●")
                self.dot_label.setStyleSheet(f"color: {t.accent}; font-size: 11px;")
                self.btn_cancel.setVisible(True)
            elif state == "success":
                self.dot_label.setText("●")
                self.dot_label.setStyleSheet(f"color: {t.success}; font-size: 11px;")
                self.btn_cancel.setVisible(False)
            elif state == "warning":
                self.dot_label.setText("▲")
                self.dot_label.setStyleSheet(f"color: {t.warning}; font-size: 11px;")
                self.btn_cancel.setVisible(False)
            elif state == "error":
                self.dot_label.setText("✕")
                self.dot_label.setStyleSheet(f"color: {t.error}; font-size: 11px;")
                self.btn_cancel.setVisible(False)
            else:
                self.dot_label.setText("●")
                self.dot_label.setStyleSheet(f"color: {t.success}; font-size: 11px;")
                self.btn_cancel.setVisible(False)

        def set_metrics(self, text: str) -> None:
            self.metrics_label.setText(text)

        def update_progress(self, progress: TaskProgress) -> None:
            self.progress_msg.setVisible(True)
            self.progress_bar.setVisible(True)
            self.btn_cancel.setVisible(True)
            self.progress_bar.setValue(progress.percentage)
            self.progress_msg.setText(
                f"{progress.current}/{progress.total} ({progress.percentage}%) {progress.message}"
            )

        def hide_progress(self) -> None:
            self.progress_msg.setVisible(False)
            self.progress_bar.setVisible(False)
            self.btn_cancel.setVisible(False)

else:

    class FluentStatusBar:  # type: ignore
        pass
