"""Structured task result dialog displaying overview cards, detailed logs, and quick actions."""

from __future__ import annotations

from platform_adapter.file_manager import open_folder

try:
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont, QFontDatabase
    from PySide6.QtWidgets import (
        QApplication,
        QDialog,
        QHBoxLayout,
        QLabel,
        QPlainTextEdit,
        QPushButton,
        QVBoxLayout,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class ResultDialog(QDialog):
        """Clean modal dialog displaying formatted results with copy and folder reveal actions."""

        def __init__(
            self,
            title: str,
            summary_text: str,
            details_text: str,
            is_success: bool = True,
            output_folder: str | None = None,
            parent: QWidget | None = None,
        ):
            super().__init__(parent)
            self.setWindowTitle(title)
            self.resize(680, 520)
            self.setMinimumSize(520, 380)
            self.setAccessibleName("任务处理结果对话框")

            self.output_folder = output_folder
            self.details_text = details_text

            layout = QVBoxLayout(self)
            layout.setContentsMargins(18, 18, 18, 18)
            layout.setSpacing(14)

            # Summary Card
            summary_card = QWidget()
            summary_card.setStyleSheet("""
                background-color: rgba(128, 128, 128, 0.08);
                border-radius: 8px;
                padding: 12px;
            """)
            summary_layout = QHBoxLayout(summary_card)
            summary_layout.setContentsMargins(12, 10, 12, 10)
            summary_layout.setSpacing(12)

            status_icon = "✓" if is_success else "✕"
            icon_color = "#0F7B0F" if is_success else "#C42B1C"
            icon_lbl = QLabel(status_icon)
            icon_lbl.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {icon_color};")
            summary_layout.addWidget(icon_lbl)

            summary_lbl = QLabel(summary_text)
            summary_lbl.setWordWrap(True)
            summary_lbl.setStyleSheet("font-size: 13px; font-weight: 500;")
            summary_layout.addWidget(summary_lbl, 1)

            layout.addWidget(summary_card)

            # Log / Details View with Platform Monospace Font
            self.text_edit = QPlainTextEdit()
            self.text_edit.setPlainText(details_text)
            self.text_edit.setReadOnly(True)
            self.text_edit.setAccessibleName("任务详情日志文本区")

            fixed_font = QFontDatabase.systemFont(QFontDatabase.FixedFont)
            fixed_font.setPointSize(12)
            self.text_edit.setFont(fixed_font)
            layout.addWidget(self.text_edit, 1)

            # Bottom Action Bar
            btn_layout = QHBoxLayout()
            btn_layout.setSpacing(10)

            btn_copy = QPushButton("复制详情")
            btn_copy.setAccessibleName("复制详情文本按钮")
            btn_copy.clicked.connect(self._copy_details)
            btn_layout.addWidget(btn_copy)

            if output_folder:
                btn_folder = QPushButton("打开输出目录")
                btn_folder.setAccessibleName("打开输出文件夹按钮")
                btn_folder.clicked.connect(lambda: open_folder(output_folder))
                btn_layout.addWidget(btn_folder)

            btn_layout.addStretch()

            btn_ok = QPushButton("确定")
            btn_ok.setProperty("isPrimary", True)
            btn_ok.setAccessibleName("确定并关闭对话框按钮")
            btn_ok.clicked.connect(self.accept)
            btn_layout.addWidget(btn_ok)

            layout.addLayout(btn_layout)

        def _copy_details(self):
            clipboard = QApplication.clipboard()
            if clipboard:
                clipboard.setText(self.details_text)

else:

    class ResultDialog:  # type: ignore
        pass
