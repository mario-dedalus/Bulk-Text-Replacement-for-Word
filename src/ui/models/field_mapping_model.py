"""Qt model for template mappings and per-field empty-value behavior."""

from __future__ import annotations

from typing import Mapping

try:
    from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
    from PySide6.QtGui import QBrush, QColor

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FieldMappingModel(QAbstractTableModel):
        """Manage Word mappings and the empty-value policy for each field."""

        COL_VARIABLE = 0
        COL_EXCEL = 1
        COL_DEFAULT = 2
        COL_STATUS = 3

        EMPTY_KEEP = "keep_variable"
        EMPTY_REPLACE = "replace_empty"
        EMPTY_CUSTOM = "custom"
        VALID_EMPTY_BEHAVIORS = {EMPTY_KEEP, EMPTY_REPLACE, EMPTY_CUSTOM}

        HEADERS = ["Word 模板变量", "对应 Excel 列（可下拉修改）", "空字段处理", "匹配状态"]

        def __init__(
            self,
            fields: list[str] | None = None,
            mapping: dict[str, str] | None = None,
            defaults: dict[str, str] | None = None,
            parent=None,
        ):
            super().__init__(parent)
            self._fields: list[str] = list(fields) if fields else []
            self._mapping: dict[str, str] = dict(mapping) if mapping else {}
            self._defaults: dict[str, str] = dict(defaults) if defaults else {}
            self._empty_behaviors: dict[str, str] = {
                field: self.EMPTY_CUSTOM if field in self._defaults else self.EMPTY_KEEP
                for field in self._fields
            }

        def rowCount(self, parent=QModelIndex()) -> int:
            return len(self._fields)

        def columnCount(self, parent=QModelIndex()) -> int:
            return len(self.HEADERS)

        def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
            if orientation == Qt.Horizontal and role == Qt.DisplayRole:
                return self.HEADERS[section]
            return None

        def flags(self, index: QModelIndex) -> Qt.ItemFlags:
            if not index.isValid():
                return Qt.NoItemFlags

            base_flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
            if index.column() in (self.COL_EXCEL, self.COL_DEFAULT):
                return base_flags | Qt.ItemIsEditable
            return base_flags

        def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
            if not index.isValid() or not (0 <= index.row() < len(self._fields)):
                return None

            field = self._fields[index.row()]
            excel_col = self._mapping.get(field, "")
            is_matched = bool(excel_col)
            default_val = self._defaults.get(field, "")
            empty_behavior = self._empty_behaviors.get(field, self.EMPTY_KEEP)

            if role == Qt.DisplayRole:
                if index.column() == self.COL_VARIABLE:
                    return f"{{{{{field}}}}}"
                elif index.column() == self.COL_EXCEL:
                    return excel_col if excel_col else "（未指定）"
                elif index.column() == self.COL_DEFAULT:
                    if empty_behavior == self.EMPTY_REPLACE:
                        return "替换为空"
                    if empty_behavior == self.EMPTY_CUSTOM:
                        return f"自定义：{default_val}" if default_val else "自定义…"
                    return "保持原变量（默认）"
                elif index.column() == self.COL_STATUS:
                    return "✓ 已匹配" if is_matched else "未匹配"

            elif role == Qt.EditRole:
                if index.column() == self.COL_EXCEL:
                    return excel_col
                elif index.column() == self.COL_DEFAULT:
                    return (empty_behavior, default_val)

            elif role == Qt.ForegroundRole:
                if index.column() == self.COL_STATUS:
                    return QBrush(QColor("#0F7B0F") if is_matched else QColor("#C42B1C"))

            elif role == Qt.TextAlignmentRole:
                if index.column() == self.COL_STATUS:
                    return int(Qt.AlignCenter)

            return None

        def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
            if not index.isValid() or role != Qt.EditRole:
                return False

            field = self._fields[index.row()]
            if index.column() == self.COL_EXCEL:
                self._mapping[field] = str(value).strip()
                status_idx = self.index(index.row(), self.COL_STATUS)
                self.dataChanged.emit(index, status_idx)
                return True
            elif index.column() == self.COL_DEFAULT:
                if isinstance(value, (tuple, list)) and len(value) >= 2:
                    behavior = str(value[0])
                    custom_value = str(value[1])
                else:
                    # Backward-compatible direct text edit means a custom value.
                    behavior = self.EMPTY_CUSTOM
                    custom_value = str(value)
                if behavior not in self.VALID_EMPTY_BEHAVIORS:
                    behavior = self.EMPTY_KEEP
                self._empty_behaviors[field] = behavior
                if behavior == self.EMPTY_CUSTOM:
                    self._defaults[field] = custom_value
                self.dataChanged.emit(index, index)
                return True

            return False

        def set_data(
            self,
            fields: list[str],
            mapping: dict[str, str],
            defaults: dict[str, str] | None = None,
            empty_behaviors: dict[str, str] | None = None,
        ) -> None:
            self.beginResetModel()
            self._fields = list(fields)
            self._mapping = dict(mapping)
            self._defaults = dict(defaults) if defaults else {}
            supplied_behaviors = empty_behaviors or {}
            self._empty_behaviors = {}
            for field in self._fields:
                behavior = supplied_behaviors.get(field)
                if behavior not in self.VALID_EMPTY_BEHAVIORS:
                    behavior = self.EMPTY_CUSTOM if field in self._defaults else self.EMPTY_KEEP
                self._empty_behaviors[field] = behavior
            self.endResetModel()

        def get_mapping(self) -> dict[str, str]:
            return dict(self._mapping)

        def get_defaults(self) -> dict[str, str]:
            return {
                field: self._defaults.get(field, "")
                for field in self._fields
                if self._empty_behaviors.get(field) == self.EMPTY_CUSTOM
            }

        def get_empty_behaviors(self) -> dict[str, str]:
            return {
                field: self._empty_behaviors.get(field, self.EMPTY_KEEP)
                for field in self._fields
            }

        def get_missing_count(self) -> int:
            return sum(1 for f in self._fields if not self._mapping.get(f, ""))

        def get_fields(self) -> list[str]:
            return list(self._fields)

else:

    class FieldMappingModel:  # type: ignore
        pass
