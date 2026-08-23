# 模板批量生成：实现与架构说明

## 原项目架构

项目的运行时代码集中在 `src/word_text_replacer_single_with_add.py`：

- `WordTextReplacerSingle` 负责 Tkinter GUI、文件列表、预览、计数、快捷键与结果窗口。
- `get_document_text`、`_collect_table_text` 使用 python-docx 读取正文和（含嵌套）表格。
- `replace_in_paragraph_advanced`、`_replace_in_table` 是 Standard Replace 的搜索/替换核心。
- `replace_text_in_documents` 用 python-docx 依次处理文件；速度快，支持正则，但复杂 Word 区域有限。
- `preview_advanced_areas`、`advanced_replace_with_vba` 用 Word COM Automation。Word 只启动一次，处理正文、表格、文本框、页眉、页脚、脚注、尾注、表单字段和超链接。
- `_find_replace_count` 使用 Word Range.Find，格式和超链接由 Word 自身维护。
- `_build_shape_ranges`、`_collect_shape_ranges`、`_replace_in_shapes` 负责文本框/Shape 及避免超链接重复计数。
- `check_hyperlinks` 检查 OOXML 超链接关系，辅助用户选择 Standard 或 Advanced。

可直接复用的稳定部分包括 Tkinter 主窗口结构、结果/进度窗口、Word COM 生命周期、Range.Find、Shape/Story Range 处理以及原 Standard/Advanced 工作流。原功能没有删除或改写。

## 新增结构

- `src/template_merge.py`：模板字段扫描、Excel 读取与类型转换、自动映射、文件名清理/去重、格式保留替换和批次错误隔离。
- `src/template_merge_gui.py`：独立模板合并窗口、Treeview 字段映射、预览、校验、进度及日志。
- `tests/create_samples.py`：生成包含正文、多变量段落、表格、页眉、页脚、粗体、不同字体、跨 run、中英文变量和超链接的样例。
- `tests/test_template_merge.py`：核心自动化回归测试。

## 格式与复杂内容策略

Standard 模板替换不使用 `paragraph.text = ...`。它定位占用变量的 run，原地清空跨越部分，并把替换值写入变量首 run，因此保留段落属性和首 run 字符格式。扫描与替换使用底层段落 run 元素，因此也能覆盖嵌套在超链接中的 run；遍历主文档、页眉和页脚 XML 时也包含表格与文本框段落。

Advanced 模板替换让 Word COM 遍历 StoryRanges，并对每个变量调用 Word 原生 Find。一个批次只创建一个 Word.Application；每个 Excel 行只打开它自己的模板副本。该模式是文本框、脚注、尾注、域及其他复杂 Word 内容的优先方案。

## 已知限制

- COM 高级模式只能在安装 Microsoft Word 与 pywin32 的 Windows 环境验证/运行。
- Standard 模式不能可靠编辑脚注、尾注和某些 Word 域；这些区域需使用高级模式。
- `.doc` 是旧二进制格式，只支持高级模式。
- 第一版只做纯变量替换，不支持 if、for、表达式或脚本。
- 模板变量不能包含换行或嵌套花括号。
