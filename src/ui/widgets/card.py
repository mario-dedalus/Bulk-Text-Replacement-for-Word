"""Fluent card surface widget."""

from __future__ import annotations

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FluentCard(QFrame):
        """Card surface container with 8px radius, subtle border, title, and content layout."""

        def __init__(self, title: str = "", subtitle: str = "", parent: QWidget | None = None):
            super().__init__(parent)
            self.setProperty("isCard", True)

            self._main_layout = QVBoxLayout(self)
            self._main_layout.setContentsMargins(16, 14, 16, 14)
            self._main_layout.setSpacing(10)

            if title or subtitle:
                header_layout = QHBoxLayout()
                header_layout.setSpacing(8)

                if title:
                    self.title_label = QLabel(title)
                    self.title_label.setObjectName("CardTitle")
                    header_layout.addWidget(self.title_label)

                if subtitle:
                    self.subtitle_label = QLabel(subtitle)
                    self.subtitle_label.setObjectName("CardSubtitle")
                    header_layout.addWidget(self.subtitle_label)

                header_layout.addStretch()
                self._main_layout.addLayout(header_layout)

        def addLayout(self, layout, stretch: int = 0) -> None:
            self._main_layout.addLayout(layout, stretch)

        def addWidget(self, widget: QWidget, stretch: int = 0) -> None:
            self._main_layout.addWidget(widget, stretch)

else:

    class FluentCard:  # type: ignore
        pass
