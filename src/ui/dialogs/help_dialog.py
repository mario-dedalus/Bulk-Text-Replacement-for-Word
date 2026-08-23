"""Help and shortcut keys documentation dialog."""

from __future__ import annotations

from platform_adapter.capabilities import CAPABILITIES

try:
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QTextBrowser, QVBoxLayout, QWidget

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class HelpDialog(QDialog):
        """User manual and keyboard shortcuts viewer."""

        def __init__(self, parent: QWidget | None = None):
            super().__init__(parent)
            self.setWindowTitle("使用指南与快捷键 — Word 批量处理工具")
            self.resize(600, 480)

            layout = QVBoxLayout(self)
            layout.setContentsMargins(18, 18, 18, 18)
            layout.setSpacing(12)

            mod_key = "Cmd" if CAPABILITIES.is_macos else "Ctrl"

            content = f"""
            <h2>Word 批量处理工具使用指南</h2>
            <p>本软件是一款面向 Windows 11 与 macOS 的跨平台 Word 文档批量处理工具。</p>

            <h3>功能模块</h3>
            <ul>
                <li><b>文本查找替换：</b>在多个 Word 文档中快速查找并替换指定文本，支持全字匹配、区分大小写及正则表达式。</li>
                <li><b>模板批量生成：</b>根据单份 Word 模板与 Excel 数据表格，按行自动派生多份填充后的个性化文档。支持在命名规则中直接使用 Excel 表头字段名（如 <code>{{序号}}</code>、<code>{{合同编号}}</code>）或 Word 变量名动态命名。</li>
                <li><b>多文档匹配替换：</b>支持将多份不同 Word 文档与多行变量数据 1 对 1 绑定与差异化填充。既可在界面表格中直接双击手动录入，也可从 Excel 导入后就地编辑覆盖，支持直接修改原文件（带备份）或另存为到新目录。</li>
            </ul>

            <h3>处理模式说明</h3>
            <ul>
                <li><b>快速模式（推荐）：</b>基于 <code>python-docx</code>，速度极快，适合正文、表格及正则替换。</li>
                <li><b>完整模式：</b>基于 Microsoft Word COM 自动化，完整覆盖页眉页脚、文本框/形状、脚注尾注与超链接保护。</li>
            </ul>

            <h3>键盘快捷键</h3>
            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%;">
                <tr style="background-color: rgba(128,128,128,0.1);">
                    <th>快捷键</th>
                    <th>操作说明</th>
                </tr>
                <tr>
                    <td><b>{mod_key} + O</b></td>
                    <td>添加 Word 文档 / 浏览模板文件</td>
                </tr>
                <tr>
                    <td><b>{mod_key} + R</b></td>
                    <td>执行查找替换 / 开始批量生成 / 开始多文档替换</td>
                </tr>
                <tr>
                    <td><b>{mod_key} + P</b></td>
                    <td>生成并查看预览更改</td>
                </tr>

                <tr>
                    <td><b>Delete</b></td>
                    <td>在文档列表中移除选中的文件</td>
                </tr>
                <tr>
                    <td><b>F1</b></td>
                    <td>打开本帮助说明窗口</td>
                </tr>
            </table>
            """

            browser = QTextBrowser()
            browser.setHtml(content)
            browser.setOpenExternalLinks(True)
            layout.addWidget(browser, 1)

            btn_layout = QHBoxLayout()
            btn_layout.addStretch()
            btn_close = QPushButton("关闭")
            btn_close.setProperty("isPrimary", True)
            btn_close.clicked.connect(self.accept)
            btn_layout.addWidget(btn_close)

            layout.addLayout(btn_layout)

else:

    class HelpDialog:  # type: ignore
        pass
