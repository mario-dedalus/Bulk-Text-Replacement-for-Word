"""Primary application entry point for PySide6 cross-platform desktop tool."""

from __future__ import annotations

import os
import sys

from platform_adapter.capabilities import CAPABILITIES

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication
    from ui.main_window import MainWindow

    HAS_PYSIDE6 = True
except ImportError:
    HAS_PYSIDE6 = False


def main() -> int:
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    if initial_file and not os.path.exists(initial_file):
        initial_file = None

    if initial_file:
        suffix = os.path.splitext(initial_file)[1].lower()
        if suffix not in {".docx", ".docm", ".doc"}:
            initial_file = None
        elif CAPABILITIES.is_macos and suffix == ".doc":
            print("提示：macOS 首版暂不支持旧版 .doc 格式文件，请使用 .docx 格式。", file=sys.stderr)
            initial_file = None

    if not HAS_PYSIDE6:
        print("错误：未找到 PySide6 库。请通过 'pip install -r requirements.txt' 安装依赖。", file=sys.stderr)
        return 1

    # High DPI screen scale factor setup
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

    app = QApplication(sys.argv)
    app.setApplicationName("WordTextReplacer")
    app.setOrganizationName("BulkTextReplacement")

    window = MainWindow(initial_file=initial_file)
    window.show()

    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
