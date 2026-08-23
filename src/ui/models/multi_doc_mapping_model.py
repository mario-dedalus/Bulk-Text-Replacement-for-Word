"""Qt Table Model managing multi-document lists, dynamic variable columns, Excel data importing, and cell editing."""

from __future__ import annotations

import os
from typing import Mapping

from core.models import ExcelData, MultiDocItem

try:
    from PySide6.QtCore import QAbstractTableModel, QModelIndex, Qt
    from PySide6.QtGui import QBrush, QColor

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class MultiDocMappingModel(QAbstractTableModel):
        """Dynamic table model with editable variable columns and per-document status tracking."""

        COL_INDEX = 0
        COL_FILENAME = 1

        def __init__(self, parent=None):
            super().__init__(parent)
            self._items: list[MultiDocItem] = []
            self._variables: list[str] = []

        def _status_col(self) -> int:
            return 2 + len(self._variables)

        def rowCount(self, parent=QModelIndex()) -> int:
            return len(self._items)

        def columnCount(self, parent=QModelIndex()) -> int:
            return 3 + len(self._variables)

        def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
            if orientation == Qt.Horizontal and role == Qt.DisplayRole:
                if section == self.COL_INDEX:
                    return "序号"
                elif section == self.COL_FILENAME:
                    return "文档名称"
                elif section == self._status_col():
                    return "状态"
                else:
                    var_idx = section - 2
                    if 0 <= var_idx < len(self._variables):
                        return f"{{{{{self._variables[var_idx]}}}}}"
            return None

        def flags(self, index: QModelIndex) -> Qt.ItemFlags:
            if not index.isValid():
                return Qt.NoItemFlags

            base_flags = Qt.ItemIsEnabled | Qt.ItemIsSelectable
            col = index.column()
            if 2 <= col < self._status_col():
                return base_flags | Qt.ItemIsEditable
            return base_flags

        def _compute_status(self, item: MultiDocItem) -> str:
            if not item.detected_variables:
                return "✓ 数据就绪"

            missing = [v for v in item.detected_variables if not str(item.replacements.get(v, "")).strip()]
            if not missing:
                return "✓ 数据就绪"
            elif len(missing) == len(item.detected_variables):
                return "⚠️ 待补充数据"
            else:
                return f"⚠️ 待补齐 ({len(missing)} 项)"

        def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
            if not index.isValid() or not (0 <= index.row() < len(self._items)):
                return None

            item = self._items[index.row()]
            col = index.column()
            status_col = self._status_col()

            if role in (Qt.DisplayRole, Qt.EditRole):
                if col == self.COL_INDEX:
                    return str(index.row() + 1)
                elif col == self.COL_FILENAME:
                    return item.filename
                elif col == status_col:
                    return self._compute_status(item)
                else:
                    var_idx = col - 2
                    if 0 <= var_idx < len(self._variables):
                        var_name = self._variables[var_idx]
                        val = item.replacements.get(var_name, "")
                        if role == Qt.DisplayRole and not val:
                            if var_name not in item.detected_variables:
                                return "—"
                        return val

            elif role == Qt.ForegroundRole:
                if col == status_col:
                    st = self._compute_status(item)
                    is_ok = "✓" in st
                    return QBrush(QColor("#0F7B0F") if is_ok else QColor("#C42B1C"))
                elif 2 <= col < status_col:
                    var_idx = col - 2
                    if 0 <= var_idx < len(self._variables):
                        var_name = self._variables[var_idx]
                        val = item.replacements.get(var_name, "")
                        if not val and var_name not in item.detected_variables:
                            return QBrush(QColor("#8A8A8A"))

            elif role == Qt.TextAlignmentRole:
                if col in (self.COL_INDEX, status_col):
                    return int(Qt.AlignCenter)

            elif role == Qt.ToolTipRole:
                if col == self.COL_FILENAME:
                    return item.file_path
                elif 2 <= col < status_col:
                    var_idx = col - 2
                    if 0 <= var_idx < len(self._variables):
                        var_name = self._variables[var_idx]
                        has_v = var_name in item.detected_variables
                        status_str = "包含此变量" if has_v else "不含此变量（无需填写）"
                        return f"文档：{item.filename}\n变量：{{{{{var_name}}}}}\n状态：{status_str}\n提示：双击单元格可就地修改替换值"

            return None

        def setData(self, index: QModelIndex, value, role: int = Qt.EditRole) -> bool:
            if not index.isValid() or role != Qt.EditRole:
                return False

            col = index.column()
            if 2 <= col < self._status_col():
                var_idx = col - 2
                var_name = self._variables[var_idx]
                item = self._items[index.row()]
                item.replacements[var_name] = str(value).strip() if value is not None else ""
                item.status = self._compute_status(item)

                status_idx = self.index(index.row(), self._status_col())
                self.dataChanged.emit(index, status_idx)
                return True

            return False

        def add_documents(self, paths: list[str]) -> int:
            existing_paths = {item.file_path for item in self._items}
            new_items = []
            for path in paths:
                abs_path = os.path.abspath(path)
                if abs_path not in existing_paths:
                    item = MultiDocItem(
                        file_path=abs_path,
                        filename=os.path.basename(abs_path),
                    )
                    new_items.append(item)
                    existing_paths.add(abs_path)

            if not new_items:
                return 0

            start_row = len(self._items)
            end_row = start_row + len(new_items) - 1
            self.beginInsertRows(QModelIndex(), start_row, end_row)
            self._items.extend(new_items)
            self.endInsertRows()
            return len(new_items)

        def remove_indices(self, rows: list[int]) -> None:
            for row in sorted(rows, reverse=True):
                if 0 <= row < len(self._items):
                    self.beginRemoveRows(QModelIndex(), row, row)
                    del self._items[row]
                    self.endRemoveRows()

        def clear(self) -> None:
            self.beginResetModel()
            self._items.clear()
            self._variables.clear()
            self.endResetModel()

        def set_variables(
            self,
            all_vars: list[str],
            doc_vars_map: dict[str, list[str]],
        ) -> None:
            self.beginResetModel()
            self._variables = list(all_vars)

            # Build a cross-platform lookup dictionary for path keys
            lookup: dict[str, list[str]] = {}
            for k, v in doc_vars_map.items():
                lookup[k] = v
                lookup[os.path.abspath(k)] = v
                lookup[os.path.normpath(k)] = v
                lookup[os.path.normcase(os.path.abspath(k))] = v

            for item in self._items:
                norm_fp = os.path.normcase(os.path.abspath(item.file_path))
                vars_for_item = (
                    lookup.get(norm_fp)
                    or lookup.get(item.file_path)
                    or lookup.get(os.path.abspath(item.file_path))
                    or lookup.get(os.path.normpath(item.file_path))
                    or []
                )
                item.detected_variables = list(vars_for_item)
                item.status = self._compute_status(item)
            self.endResetModel()


        def import_table_data(self, excel_data: ExcelData, match_by_header: bool = True) -> int:
            """Import Excel/CSV rows into multi-document items."""
            if not self._items or not excel_data.rows:
                return 0

            filled_count = 0
            headers_clean = [h.strip() for h in excel_data.headers]

            # Build a lookup for matching: normalized header -> actual header
            header_map = {}
            for h in headers_clean:
                header_map[h] = h
                # Also strip {{ and }} if present
                if h.startswith("{{") and h.endswith("}}"):
                    header_map[h[2:-2].strip()] = h

            for doc_idx, item in enumerate(self._items):
                if doc_idx >= len(excel_data.rows):
                    break

                row_dict = excel_data.rows[doc_idx]
                for var in self._variables:
                    target_header = None
                    if match_by_header:
                        if var in header_map:
                            target_header = header_map[var]
                        elif var in row_dict:
                            target_header = var

                    if target_header is not None and target_header in row_dict:
                        val = str(row_dict[target_header]).strip()
                        item.replacements[var] = val
                    elif not match_by_header:
                        # Sequential alignment fallback
                        var_pos = self._variables.index(var)
                        if var_pos < len(headers_clean):
                            col_header = headers_clean[var_pos]
                            item.replacements[var] = str(row_dict.get(col_header, "")).strip()

                item.status = self._compute_status(item)
                filled_count += 1

            top_left = self.index(0, 2)
            bottom_right = self.index(len(self._items) - 1, self._status_col())
            self.dataChanged.emit(top_left, bottom_right)
            return filled_count

        def clear_table_data(self) -> None:
            """Clear all entered replacement values while retaining documents and variables."""
            for item in self._items:
                item.replacements.clear()
                item.status = self._compute_status(item)

            if self._items and self._variables:
                top_left = self.index(0, 2)
                bottom_right = self.index(len(self._items) - 1, self._status_col())
                self.dataChanged.emit(top_left, bottom_right)

        def move_row_up(self, row: int) -> bool:
            if row <= 0 or row >= len(self._items):
                return False

            self.beginMoveRows(QModelIndex(), row, row, QModelIndex(), row - 1)
            self._items[row - 1], self._items[row] = self._items[row], self._items[row - 1]
            self.endMoveRows()
            return True

        def move_row_down(self, row: int) -> bool:
            if row < 0 or row >= len(self._items) - 1:
                return False

            self.beginMoveRows(QModelIndex(), row + 1, row + 1, QModelIndex(), row)
            self._items[row], self._items[row + 1] = self._items[row + 1], self._items[row]
            self.endMoveRows()
            return True

        def get_items(self) -> list[MultiDocItem]:
            return [
                MultiDocItem(
                    file_path=item.file_path,
                    filename=item.filename,
                    detected_variables=list(item.detected_variables),
                    replacements=dict(item.replacements),
                    status=item.status,
                )
                for item in self._items
            ]

        def get_variables(self) -> list[str]:
            return list(self._variables)

        def get_ready_metrics(self) -> tuple[int, int]:
            """Return (ready_count, total_count)."""
            ready = sum(1 for item in self._items if "✓" in self._compute_status(item))
            return ready, len(self._items)

else:

    class MultiDocMappingModel:  # type: ignore
        pass
