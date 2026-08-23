"""Item delegate rendering QComboBox for Excel column selection inside QTableView."""

from __future__ import annotations

try:
    from PySide6.QtCore import QSignalBlocker, QTimer, Qt
    from PySide6.QtWidgets import QComboBox, QStyledItemDelegate, QWidget

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class MappingComboDelegate(QStyledItemDelegate):
        """Allows direct dropdown selection of Excel headers in the mapping table."""

        def __init__(self, headers: list[str] | None = None, parent=None):
            super().__init__(parent)
            self._headers = list(headers) if headers else []

        def set_headers(self, headers: list[str]) -> None:
            self._headers = list(headers)

        def createEditor(self, parent: QWidget, option, index) -> QWidget:
            combo = QComboBox(parent)
            combo.setAutoFillBackground(True)
            combo.addItem("（未指定 / 留空）", "")
            for h in self._headers:
                combo.addItem(h, h)

            # `currentIndexChanged` is also emitted by setEditorData().  Using it
            # here closes the editor while Qt is still initialising it, which in
            # turn causes missed clicks and stale/overlaid text.  `activated` is
            # emitted only for an explicit user selection.
            combo.activated.connect(lambda _index: self._commit_and_close(combo))

            # The table explicitly starts editing on a single click.  Open the
            # popup after Qt has installed and positioned the editor.
            QTimer.singleShot(0, combo.showPopup)
            return combo

        def _commit_and_close(self, editor: QComboBox) -> None:
            viewport = editor.parentWidget()
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QStyledItemDelegate.NoHint)
            if viewport is not None:
                # Repaint after the editor is removed so the model text cannot
                # remain visually stacked with the former combo-box text.
                QTimer.singleShot(0, viewport.update)

        def setEditorData(self, editor: QWidget, index) -> None:
            if isinstance(editor, QComboBox):
                blocker = QSignalBlocker(editor)
                current_val = index.model().data(index, Qt.EditRole) or ""
                idx = editor.findData(current_val)
                if idx >= 0:
                    editor.setCurrentIndex(idx)
                else:
                    editor.setCurrentIndex(0)
                del blocker

        def setModelData(self, editor: QWidget, model, index) -> None:
            if isinstance(editor, QComboBox):
                selected_data = editor.currentData()
                model.setData(index, selected_data, Qt.EditRole)

        def updateEditorGeometry(self, editor: QWidget, option, index) -> None:
            editor.setGeometry(option.rect)

else:

    class MappingComboDelegate:  # type: ignore
        pass
