"""Delegate for selecting how an empty mapped value should be handled."""

from __future__ import annotations

try:
    from PySide6.QtCore import QSignalBlocker, QTimer, Qt
    from PySide6.QtWidgets import QComboBox, QInputDialog, QStyledItemDelegate, QWidget

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class EmptyFieldDelegate(QStyledItemDelegate):
        """Three-way editor: keep placeholder, replace empty, or custom text."""

        KEEP = "keep_variable"
        EMPTY = "replace_empty"
        CUSTOM = "custom"

        def createEditor(self, parent: QWidget, option, index) -> QWidget:
            combo = QComboBox(parent)
            combo.setAutoFillBackground(True)
            combo.addItem("保持原变量（默认）", self.KEEP)
            combo.addItem("替换为空", self.EMPTY)
            combo.addItem("自定义…", self.CUSTOM)
            combo.activated.connect(lambda _row: self._activate(editor=combo))
            QTimer.singleShot(0, combo.showPopup)
            return combo

        def _activate(self, editor: QComboBox) -> None:
            behavior = editor.currentData()
            if behavior == self.CUSTOM:
                current = str(editor.property("customValue") or "")
                text, accepted = QInputDialog.getText(
                    editor,
                    "自定义空字段替换",
                    "当该数据列为空时，替换为：",
                    text=current,
                )
                if not accepted:
                    self.closeEditor.emit(editor, QStyledItemDelegate.RevertModelCache)
                    return
                editor.setProperty("customValue", text)

            viewport = editor.parentWidget()
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QStyledItemDelegate.NoHint)
            if viewport is not None:
                QTimer.singleShot(0, viewport.update)

        def setEditorData(self, editor: QWidget, index) -> None:
            if not isinstance(editor, QComboBox):
                return
            blocker = QSignalBlocker(editor)
            value = index.model().data(index, Qt.EditRole)
            if isinstance(value, (tuple, list)) and len(value) >= 2:
                behavior, custom_value = str(value[0]), str(value[1])
            else:
                behavior, custom_value = self.KEEP, ""
            selected = editor.findData(behavior)
            editor.setCurrentIndex(selected if selected >= 0 else 0)
            editor.setProperty("customValue", custom_value)
            del blocker

        def setModelData(self, editor: QWidget, model, index) -> None:
            if isinstance(editor, QComboBox):
                model.setData(
                    index,
                    (str(editor.currentData()), str(editor.property("customValue") or "")),
                    Qt.EditRole,
                )

        def updateEditorGeometry(self, editor: QWidget, option, index) -> None:
            editor.setGeometry(option.rect)

else:

    class EmptyFieldDelegate:  # type: ignore
        pass
