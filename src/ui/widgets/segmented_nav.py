"""Segmented navigation widget for switching top-level pages with keyboard and a11y support."""

from __future__ import annotations

from typing import Callable

try:
    from PySide6.QtCore import Qt, Signal
    from PySide6.QtGui import QKeyEvent
    from PySide6.QtWidgets import QButtonGroup, QFrame, QHBoxLayout, QPushButton, QWidget

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class SegmentedNav(QFrame):
        """Top bar segmented navigation control with accessible tab switching."""

        tabChanged = Signal(str)

        def __init__(
            self,
            tabs: list[tuple[str, str]],
            on_tab_change: Callable[[str], None] | None = None,
            parent: QWidget | None = None,
        ):
            super().__init__(parent)
            self.tabs = tabs
            self.active_tab = tabs[0][0] if tabs else ""
            self._buttons: dict[str, QPushButton] = {}
            self._tab_order: list[str] = [t[0] for t in tabs]

            self.setAccessibleName("主功能分类导航")
            self.setAccessibleDescription("使用左右方向键或点击切换文本替换与模板生成功能")
            self.setFocusPolicy(Qt.StrongFocus)

            layout = QHBoxLayout(self)
            layout.setContentsMargins(0, 0, 0, 0)
            layout.setSpacing(4)

            self.btn_group = QButtonGroup(self)
            self.btn_group.setExclusive(True)

            for i, (tab_id, label) in enumerate(tabs):
                btn = QPushButton(label)
                btn.setCheckable(True)
                btn.setCursor(Qt.PointingHandCursor)
                btn.setObjectName("NavTabButton")
                btn.setAccessibleName(f"切换至{label}页面")
                btn.setAccessibleDescription(f"点击或使用快捷键切换到{label}功能模块")
                btn.setStyleSheet("""
                    QPushButton#NavTabButton {
                        background-color: transparent;
                        border: none;
                        border-radius: 4px;
                        padding: 6px 14px;
                        font-weight: 500;
                        font-size: 13px;
                    }
                    QPushButton#NavTabButton:hover {
                        background-color: rgba(128, 128, 128, 0.15);
                    }
                    QPushButton#NavTabButton:checked {
                        background-color: rgba(128, 128, 128, 0.22);
                        font-weight: 600;
                    }
                    QPushButton#NavTabButton:focus {
                        outline: none;
                        border: none;
                    }
                """)
                if i == 0:
                    btn.setChecked(True)

                self.btn_group.addButton(btn, i)
                layout.addWidget(btn)
                self._buttons[tab_id] = btn

                btn.clicked.connect(lambda _, tid=tab_id: self._on_btn_clicked(tid))

            if on_tab_change:
                self.tabChanged.connect(on_tab_change)

        def _on_btn_clicked(self, tab_id: str):
            if self.active_tab != tab_id:
                self.active_tab = tab_id
                self.tabChanged.emit(tab_id)

        def select_tab(self, tab_id: str):
            if tab_id in self._buttons:
                self._buttons[tab_id].setChecked(True)
                self._on_btn_clicked(tab_id)

        def keyPressEvent(self, event: QKeyEvent):
            if not self._tab_order:
                super().keyPressEvent(event)
                return

            current_idx = self._tab_order.index(self.active_tab) if self.active_tab in self._tab_order else 0

            if event.key() in (Qt.Key_Left, Qt.Key_Up):
                new_idx = (current_idx - 1) % len(self._tab_order)
                new_tab = self._tab_order[new_idx]
                self.select_tab(new_tab)
                if new_tab in self._buttons:
                    self._buttons[new_tab].setFocus()
                event.accept()
            elif event.key() in (Qt.Key_Right, Qt.Key_Down):
                new_idx = (current_idx + 1) % len(self._tab_order)
                new_tab = self._tab_order[new_idx]
                self.select_tab(new_tab)
                if new_tab in self._buttons:
                    self._buttons[new_tab].setFocus()
                event.accept()
            else:
                super().keyPressEvent(event)

else:

    class SegmentedNav:  # type: ignore
        pass
