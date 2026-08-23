"""Qt Model for managing the list of selected Word documents."""

from __future__ import annotations

import os

try:
    from PySide6.QtCore import QAbstractListModel, QModelIndex, Qt

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class FileListModel(QAbstractListModel):
        """Model for displaying selected Word documents with size and path info."""

        def __init__(self, file_paths: list[str] | None = None, parent=None):
            super().__init__(parent)
            self._files: list[str] = list(file_paths) if file_paths else []

        def rowCount(self, parent=QModelIndex()) -> int:
            return len(self._files)

        def data(self, index: QModelIndex, role: int = Qt.DisplayRole):
            if not index.isValid() or not (0 <= index.row() < len(self._files)):
                return None

            path = self._files[index.row()]
            filename = os.path.basename(path)

            if role == Qt.DisplayRole:
                try:
                    size_kb = os.path.getsize(path) / 1024
                    size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
                except OSError:
                    size_str = "未知大小"
                return f"{filename}  ({size_str}) — {path}"

            elif role == Qt.ToolTipRole:
                return path

            elif role == Qt.UserRole:
                return path

            return None

        def add_files(self, paths: list[str]) -> int:
            added = 0
            for path in paths:
                abs_path = os.path.abspath(path)
                if abs_path not in self._files:
                    self.beginInsertRows(QModelIndex(), len(self._files), len(self._files))
                    self._files.append(abs_path)
                    self.endInsertRows()
                    added += 1
            return added

        def remove_row(self, row: int) -> None:
            if 0 <= row < len(self._files):
                self.beginRemoveRows(QModelIndex(), row, row)
                del self._files[row]
                self.endRemoveRows()

        def remove_indices(self, rows: list[int]) -> None:
            for row in sorted(rows, reverse=True):
                self.remove_row(row)

        def clear(self) -> None:
            self.beginResetModel()
            self._files.clear()
            self.endResetModel()

        def get_all_paths(self) -> list[str]:
            return list(self._files)

        def count(self) -> int:
            return len(self._files)

else:

    class FileListModel:  # type: ignore
        pass
