# 全面代码审计与改进方案

> 审计日期：2026-08-23  
> 审计范围：当前工作区（包括未提交和未跟踪代码），不以最后一次 Git 提交为基准  
> 审计原则：保持现有功能与用户行为；优先修复正确性、文件安全和 Windows/Word COM 稳定性；不建议大规模重写

## 1. 结论摘要

项目当前采用 `UI -> application service -> core/platform adapter` 的分层，三个业务工作流边界清楚，方向是合理的，不需要推倒重写。模板批量生成所使用的占位符替换算法、逐行错误隔离、输出文件名避让、Qt Model/View 分离、扫描代次保护等实现值得保留。

但当前版本仍不适合直接作为“可靠修改用户 Word 原件”的发布版本。审计确认了 3 项 P0：

1. 快速查找替换引擎会漏替换、误报替换数、删除未匹配文本中的不可见字符，并在跨 run 匹配时丢失未匹配内容的字符格式。
2. 后台任务没有真正的互斥状态；`Ctrl+R` 等快捷键可以绕过按钮禁用，重复启动多个写任务并发保存同一文档。
3. Word COM 从 `QThreadPool` 工作线程调用，但没有在线程内初始化/释放 COM apartment；完整模式在 Windows 上存在直接初始化失败的高概率风险，且 Word 实例生命周期分散、缺少隔离。

建议先完成 P0 和文件原子写入，再扩充 Windows Word 集成测试。现有架构可以渐进修复，没有充分理由进行大规模重写。

## 2. 审计方法与验证结果

本轮阅读和关联检查了：

- 正式入口、PyInstaller spec、Windows/macOS 构建脚本和 GitHub Actions；
- `src/core`、`src/application`、`src/platform_adapter` 全部实现；
- 三个 PySide6 页面、Qt 数据模型、代理、主题和结果对话框；
- 兼容转发模块、旧 Tkinter 实现、`_legacy`、上下文菜单文件；
- 全部自动化测试及现有设计/审计文档。

执行结果：

- `.venv/bin/python -m unittest discover -s tests -v`：68 项全部通过；
- `.venv/bin/python -m compileall -q src tests`：通过；
- 系统 `python3` 直接运行测试失败，原因是系统解释器未安装 `python-docx`/`PySide6`，不是代码测试失败；项目虚拟环境可正常运行；
- 审计环境为 macOS、Python 3.14.3，因此没有实机执行 Microsoft Word COM；Windows COM 结论来自代码调用线程和生命周期检查，必须在 Windows + Word 环境补充验收。

额外最小复现确认：

- 一个段落由 `"x"`、`" / "`、`"x"` 三个 run 组成时，快速替换 `x -> y` 返回计数 2，结果却是 `"y / x"`；
- 跨三个不同格式 run 替换时，结果文本被写入第一个 run，粗体/斜体/下划线边界丢失；
- 对不含查找词、但包含软连字符的段落执行替换，返回 0，但软连字符仍被删除。

现有测试通过说明主流程已有良好基础，但不代表上述文件保真、并发和 COM 场景安全。

## 3. 项目结构与核心流程理解

### 3.1 当前分层

- `src/app.py`：正式 PySide6 入口。
- `src/ui`：主窗口、三个业务页面、Qt Model/Delegate、主题和对话框。
- `src/application`：输入校验、任务结果、进度、取消和 core 调度。
- `src/core`：Word 文本替换、模板生成、多文档差异替换和数据模型。
- `src/platform_adapter`：平台能力、外观、Explorer/Finder 集成。
- 根目录 `src/replacer_core.py`、`src/template_merge.py`、`src/platform_capabilities.py`：兼容转发层。
- `src/word_text_replacer_single_with_add.py`、`src/template_merge_gui.py`、`src/ui/fluent_widgets.py`、`src/ui/theme.py`：旧 Tkinter 路径，已不是正式入口。

### 3.2 三条业务链路

1. **统一查找替换**：`ReplacePage -> ReplaceService -> perform_standard_* / perform_com_*`，直接修改原文件，可创建 `.backup`。
2. **模板批量生成**：`MergePage -> MergeService -> generate_batch`，按 Excel/CSV 行复制模板并填充变量。
3. **多文档差异替换**：`MultiDocPage -> MultiDocService -> execute_multi_doc_replace`，每个文档使用独立变量值，原地修改或导出副本。

三条链路都通过 `TaskWorker` 和全局 `QThreadPool` 执行；取消为协作式取消，只在文件/数据行边界检查。

## 4. 做得合理、当前无需修改的部分

以下部分已足够清晰，除非后续需求变化，不建议为了“更现代”而重构：

- 三个顶级业务页面分别对应“多文档同规则”“单模板多行”“多文档多行”，产品心智和模块职责合理。
- `core` 不依赖 Qt，`application` 负责调度，`ui` 不直接实现 Word XML 处理，分层方向正确。
- `ExcelData`、`MergeResult`、`BatchProcessResult`、`MultiDocItem` 等数据类简单直接，无需引入复杂领域框架。
- `CancellationToken` 基于 `threading.Event`，线程安全且足够成熟。
- 模板扫描和文本缓存使用 generation id 丢弃旧结果，适合当前规模。
- 模板生成逐行捕获错误、失败时删除该行输出副本、避免同批次文件名冲突，策略合理。
- `FieldMappingModel` 和 `MultiDocMappingModel` 使用 Qt Model/View，而不是把状态散落在单元格控件中，便于测试和维护。
- 平台能力、主题探测和文件管理器调用已集中到 `platform_adapter`，不建议重新引入更重的平台抽象。
- PyInstaller 使用 `src/app.py` 作为唯一正式入口，正式 UI 不需要运行时回退到 Tkinter。

## 5. 问题总览

| 编号 | 优先级 | 主题 | 主要影响 |
|---|---|---|---|
| CR-01 | P0 | 快速替换正确性与格式保真 | 漏替换、误报、非目标内容被改、格式丢失 |
| CR-02 | P0 | 后台任务可重入和并发写文件 | 同一文件可被多个线程同时保存，可能损坏文档 |
| CR-03 | P0 | Windows COM 线程与实例生命周期 | 完整模式可能在线程中初始化失败或影响 Word 会话 |
| CR-04 | P1 | 原文件写入和备份不是事务式 | 中断/磁盘错误时可能留下损坏文件，旧备份被覆盖 |
| CR-05 | P1 | 文件格式、引擎模式和能力校验不一致 | `.doc` 扫描成功但执行失败，复杂区域扫描与执行不一致 |
| CR-06 | P1 | 多文档扫描吞掉错误 | 损坏/不支持文件被标记为“数据就绪” |
| CR-07 | P1 | 取消、关闭与部分结果不透明 | 用户不知道哪些文件已修改；强退仍可能中断保存 |
| CR-08 | P1 | 正则表达式校验和执行成本 | 非法正则静默变成 0；复杂正则可能长时间占用线程 |
| CR-09 | P1 | Windows/Word 发布验收不足 | CI 不覆盖真实 Word COM、文件锁和 Explorer 场景 |
| CR-10 | P1 | 异常诊断和批处理成功语义偏弱 | 全部失败仍可表现为“成功但警告”，无持久日志/堆栈 |
| CR-11 | P1 | 输出文件扩展名和路径约束 | 可生成扩展名与实际 Word 格式不符的文件 |
| CR-12 | P2 | 旧 Tkinter 与兼容代码仍在正式源树 | 重复实现、导入名冲突、打包噪声 |
| CR-13 | P2 | COM/文档处理函数过长且重复 | 资源释放和区域处理规则难以保持一致 |
| CR-14 | P2 | 回调兼容逻辑会误捕获 `TypeError` | 可能重复回调或掩盖回调内部缺陷 |
| CR-15 | P2 | 超链接扫描错误未展示 | 损坏文件被显示为 0 个链接而不是失败 |
| CR-16 | P2 | 依赖和构建职责混合 | 运行依赖与构建依赖未分离，构建可复现性有限 |
| CR-17 | P2 | 若干明显性能浪费 | 正则重复编译、无匹配也保存、表格顺序导入为二次查找 |
| CR-18 | P3 | 类型和兼容导入可收紧 | `dict`/`Any` 过多，fallback import 易产生循环语义 |
| CR-19 | P3 | 文档与元数据有小幅漂移 | README 测试数、版本信息和入口说明不完全同步 |
| CR-20 | P3 | 少量平台体验可选优化 | 系统主题轮询、长路径提示、启动诊断可改善 |

## 6. 详细审计发现与整改任务

### CR-01（P0）：重写“快速替换”的段落内编辑算法，但保留现有 API

**当前问题**

`src/core/replacer_core.py:140-229` 的 `replace_in_paragraph_advanced` 存在三类已复现问题：

- 在检查是否匹配前，先把所有 run 中的软连字符/零宽字符写回删除（149-153 行）；即使查找词不存在，文档内容也会变化。
- 只要任一单独 run 含匹配，函数就替换该 run 并返回整个段落的匹配总数（184-205 行）；其他 run 中的匹配未处理。
- 只要匹配跨 run，就清空所有 run，并把整个新段落写入第一个 run（207-227 行）；未匹配前后文的字符格式也会丢失。

`perform_standard_replace` 随后无论替换数是否为 0 都保存文件（`src/core/replacer_core.py:347-360`），使非目标变更落盘。现有 `tests/test_replacer_core.py` 只验证结果文本，没有验证多个独立 run、未匹配不可见字符和格式边界。

**为什么需要修改**

这是直接修改用户原件的正确性问题，且与 README 中“preserves run formatting”的行为承诺冲突。计数错误还会让审计报告不可相信。

**推荐修改方式**

- 保持公开函数签名，内部增加一个小型、纯 Python 的“逻辑文本到 run/XML 位置映射”步骤。
- 在不修改原 run 的前提下构建规范化文本，并保留每个规范化字符对应的原始 run/offset。
- 正则只编译一次，取得所有非重叠 match，按从后向前的顺序编辑受影响 run；仅替换匹配跨度，保持所有未匹配文本和格式。
- 正则替换值使用 `match.expand(replace_text)`，明确保留 Python 回溯引用行为。
- 对跨 hyperlink/复杂 XML 容器、无法可靠保持关系的匹配，选择明确警告并建议完整模式，而不是静默破坏链接。
- 无实际替换时不要保存文件。

**必须新增的测试**

- 两个不同 run 各有一个匹配，实际结果和计数均为 2；
- 单个 run 多匹配、跨 run 多匹配、大小写、全字、正则捕获组；
- 查找词不存在时软连字符和零宽字符原样保留；
- 跨粗体/斜体/下划线 run 后，未匹配区域格式保持；
- hyperlink 内、hyperlink 边界、表格和嵌套表格；
- 替换为空、替换文本包含反斜杠和换行。

**收益**：恢复替换计数、内容和格式的可信度。  
**风险/影响范围**：快速模式的核心行为会变化，需用真实复杂 Word 样本做 XML 和视觉回归；不影响 COM 模式和模板字段替换 API。

### CR-02（P0）：增加唯一任务所有权，禁止写任务重入

**当前问题**

三个页面只通过禁用主按钮表示 busy，没有 `_busy` 入口守卫。主窗口快捷键直接调用页面方法（`src/ui/main_window.py:222-249`），因此按钮被禁用后继续按 `Ctrl+R` 仍可再次创建 worker。文本页的“检查超链接”、文件列表和多个输入也未随任务锁定（`src/ui/pages/replace_page.py:407-423, 549-660`）。每页只有一个 `_active_worker` 指针，新任务会覆盖旧任务，取消只能取消最后一个。

不同页面也可同时启动任务；若选择同一文件，多个 `python-docx`/COM 任务可能并发保存。

**推荐修改方式**

- 增加一个简单的 `TaskCoordinator` 或主窗口级写任务锁；不需要引入复杂任务框架。
- 区分只读任务和写任务：同一时刻只允许一个写任务；同一页面最多一个前台任务。
- 所有 public action 方法开头检查任务状态，快捷键和按钮走同一守卫。
- coordinator 保存 worker 集合和任务 ID；`finished` 只结束对应 ID，旧 worker 的完成信号不能解锁新任务。
- 任务期间锁定会改变执行语义的输入，或冻结不可变参数快照并在结果报告中使用同一快照。
- `_active_worker` 在对应任务结束后清空。

**验证**

- 连续触发两次 `Ctrl+R` 只创建一个 worker；
- 按钮、快捷键、跨页面启动和超链接扫描都遵循互斥规则；
- 较早任务的 `finished` 不会解除较晚任务的 busy；
- 取消命中正确任务。

**收益**：消除并发写同一文档的损坏风险，简化关闭和取消。  
**风险/影响范围**：涉及三个页面和主窗口状态连接，但不需要改 core 业务 API。

### CR-03（P0）：统一 Word COM 会话并在线程内初始化 apartment

**当前问题**

COM 调用由 `TaskWorker` 在 `QThreadPool` 线程执行（`src/application/workers.py:70-84`），但以下位置直接 `Dispatch("Word.Application")`，没有 `pythoncom.CoInitialize/CoUninitialize`：

- `src/core/replacer_core.py:441-646, 649-859`；
- `src/core/template_merge.py:220-249, 532-625`；
- `src/core/multi_doc_replacer.py:113-234`。

能力检测只验证 `win32com.client` 能否 import（`src/platform_adapter/capabilities.py:24-31`），不能证明 Microsoft Word 已安装或可启动。各函数又分别管理 `Visible`、`DisplayAlerts`、文档关闭和 `Quit`，并使用 `Dispatch`，缺少“本进程独占实例”的明确保证。

大量区域级 `except Exception: pass` 还会把 COM 区域访问失败伪装成“0 处匹配/替换成功”。

**推荐修改方式**

- 在 `platform_adapter` 或独立 `core/word_com.py` 中实现轻量 `WordAutomationSession` context manager。
- context manager 在实际执行 COM 的同一工作线程调用 `pythoncom.CoInitialize()`，退出时按相反顺序关闭文档、退出自建 Word、释放引用、`CoUninitialize()`。
- 使用 `DispatchEx` 创建隔离实例；不要把 COM proxy 通过 Qt signal 或跨线程缓存。
- 把“pywin32 已安装”和“Word 可启动”分为两个能力；启动失败应给出可操作错误。
- 区域访问失败记录为 warning，关键打开/保存/关闭失败则判定该文件失败。
- 为每个 Word 会话设置可恢复的 alerts/visibility/screen updating，并确保所有异常路径执行清理。

**验证**

- Windows 11 + Word 的 QThreadPool 实际预览、替换、模板生成、多文档替换；
- Word 未安装、Word 正在打开其他用户文档、目标文档只读/被锁、用户取消和异常退出；
- 任务后不存在残留 `WINWORD.EXE`，也不关闭用户原有 Word 窗口。

**收益**：完整模式在 Windows 后台线程中可预测运行，资源释放集中。  
**风险/影响范围**：只影响 COM 路径；必须实机验证 Office 不同版本和受保护视图。

### CR-04（P1）：采用可恢复写入和不覆盖旧备份的策略

**当前问题**

快速模式直接 `doc.save(file_path)`；COM 模式直接 `doc.Save()`。备份固定为 `原路径 + ".backup"`，会覆盖上一次备份。进程崩溃、强制关闭、磁盘写满、网络盘断开或并发写入时，原文件和旧备份都可能不可恢复。

涉及 `src/core/replacer_core.py:312-384, 649-859`、`src/core/multi_doc_replacer.py:113-234`。模板生成虽然写副本，但同样应验证临时结果后再对用户可见。

**推荐修改方式**

- 原地 python-docx 修改：保存到同目录唯一临时文件，重新打开/校验 ZIP 和主文档部件，再使用平台合适的原子替换；失败时保留原件。
- COM 修改：优先保存到临时副本并校验，再关闭 Word 后替换目标；需要单独验证宏文档和 ACL/时间戳。
- 备份使用不冲突命名（例如 `.backup`、`.backup (2)`）或显式“覆盖旧备份”策略；默认不覆盖。
- 写入成功后再把文件计入成功结果；临时文件清理失败要记录 warning。
- 不要用宽泛删除清理掩盖原始异常。

**收益**：显著降低文档损坏和备份失效风险。  
**风险/影响范围**：网络共享、文件锁、NTFS ACL、`.docm` 宏和 Word SaveAs 行为需 Windows 验证。

### CR-05（P1）：集中统一文件格式、处理模式和能力规则

**当前问题**

- `ReplaceService.validate_inputs` 和 `MultiDocService.validate_inputs` 没有白名单校验扩展名，也没有规定快速模式不能处理 `.doc`。
- Windows UI 允许在快速模式选择 `.doc`；查找替换会逐文件失败，多文档扫描则在 `use_com=False` 时把 `.doc` 当作无变量。
- 模板 `.doc` 扫描无论复选框是否勾选都会使用 COM，但批量生成仍使用复选框值；因此可能“扫描成功、全部生成失败”。
- `use_com=True` 时，`.docx/.docm` 的模板/多文档扫描仍走标准扫描器，完整模式执行能处理而扫描器未覆盖的 footnote/endnote 变量可能遗漏。

涉及 `src/application/replace_service.py:25-37`、`merge_service.py:24-101`、`multi_doc_service.py:22-91`、`src/core/multi_doc_replacer.py:43-64`。

**推荐修改方式**

- 新建一个很小的格式策略函数（而非新框架），输入 `operation + suffix + mode + capabilities`，返回允许/拒绝及原因。
- `.doc` 必须明确要求 COM；混合文件列表要么统一切换完整模式，要么在执行前列出拒绝文件。
- 完整模式扫描和完整模式执行使用同一内容覆盖范围。
- 校验 `.docx/.docm/.doc` 白名单、真实文件、输出目录和模式能力；不要依赖文件对话框过滤器。
- `has_word_com` 不仅检查 import，还要把 Word 实际启动失败转成明确错误。

**收益**：消除扫描、预览和执行之间的行为分叉。  
**风险/影响范围**：某些过去“先尝试再报错”的路径会改为提前拒绝，这是更可靠的错误时机，不改变成功场景。

### CR-06（P1）：多文档扫描不得吞掉每个文件的异常

**当前问题**

`scan_documents_variables` 在每个文件上捕获所有异常并返回空字段（`src/core/multi_doc_replacer.py:43-64`）。`MultiDocMappingModel._compute_status` 又把“没有检测到变量”视为“✓ 数据就绪”（`src/ui/models/multi_doc_mapping_model.py:55-66`）。损坏文件、不支持格式、权限错误和真实无变量文档无法区分。

**推荐修改方式**

- 扫描结果增加每文件状态：`success/no_fields/error/cancelled` 和错误文本。
- 真正“扫描成功但无变量”可以就绪；扫描失败必须显示失败并阻止默认执行，除非用户明确跳过。
- service 在部分失败时返回 `WARNING`，全部失败时返回 `FAILED`。
- UI 状态列和结果报告显示失败原因。

**收益**：避免把不可处理文件伪装成可执行文件。  
**风险/影响范围**：需扩展结果模型和测试，但可以保持现有 `doc_vars_map/all_vars` 作为兼容字段。

### CR-07（P1）：让取消和关闭具有可核验的部分结果

**当前问题**

取消只在文件/行之间检查。service 会携带部分结果返回 `CANCELLED`，但替换页面收到取消后直接返回，不展示已修改文件、备份或输出。`MultiDocBatchResult.total_docs` 始终是输入总数，取消后未处理项也没有独立状态。

主窗口在 GUI 线程调用 `waitForDone(3000)`，随后允许用户强制关闭（`src/ui/main_window.py:185-218`）。这既阻塞 UI，也不能保证正在执行的保存已完成；全局线程池还包含只读缓存任务，`activeThreadCount()` 不能准确表达“存在危险写任务”。

**推荐修改方式**

- 结果模型增加 `processed/skipped/cancelled`，取消对话框必须列出已成功修改、失败、未处理及备份路径。
- coordinator 跟踪危险写任务；关闭进入非阻塞 `closing` 状态，发送取消，等待对应任务 `finished` 后真正退出。
- 正在保存单个文件时不强杀；如必须提供强退，明确警告可能损坏文件，并优先依赖 CR-04 的临时写入保证原件。
- 只读缓存任务不应触发与写任务相同的高风险关闭提示。

**收益**：取消后用户知道磁盘上发生了什么，关闭流程与实际任务状态一致。  
**风险/影响范围**：涉及结果模型、状态栏和 close event；核心替换逻辑只需补充状态记录。

### CR-08（P1）：在任务启动前验证并预编译正则

**当前问题**

非法正则在 `count_occurrences`、`find_match_contexts`、`replace_in_paragraph_advanced` 中被捕获后静默返回 0。pattern 在多个文件/段落中反复编译。复杂回溯正则对大段文档可能长时间占用工作线程，期间取消无法生效。

**推荐修改方式**

- service 层在预览/执行前编译一次，非法正则直接显示具体错误位置。
- 将编译后的 pattern 或统一 matcher 传给 core 内循环；literal/whole-word 也走同一匹配计划。
- 文档中说明 Python 正则；对明显危险模式可给提示，但不要自制不完整的“安全正则引擎”。
- 给预览增加可配置的文本/匹配上限和耗时监控；极端场景可考虑独立进程超时，当前阶段不必先引入。

**收益**：错误可见，减少重复编译和无响应。  
**风险/影响范围**：正则边界行为需与现有 Python `re` 保持一致。

### CR-09（P1）：建立真实 Windows + Word 验收矩阵

**当前问题**

GitHub Actions 的 Windows job 会跑标准测试并检查 EXE 是否生成，但 CI runner 通常不能提供可自动化的桌面 Word。现有测试没有 `pythoncom`、Word 生命周期、COM 区域遍历、文件锁、`.doc` 或 EXE 真启动覆盖。当前审计也未在 Windows 实机运行。

**推荐修改方式**

- 保留现有跨平台单元测试；新增 mock/contract 测试验证 COM session 的初始化、清理顺序和异常路径。
- 建立可手动或自托管执行的 Windows 11 + Microsoft 365/Office 2021 集成套件。
- 样本覆盖正文、表格、页眉页脚、文本框、分组形状、超链接、脚注尾注、表单域、`.doc/.docx/.docm`、只读和锁定文件。
- 对打包 EXE 做真实启动、打开文件参数、Explorer 上下文菜单和非 ASCII/长路径冒烟。

**收益**：Windows 核心卖点不再只靠代码推断。  
**风险/影响范围**：需要 Windows/Office 环境维护；不应把不稳定的桌面 Office 测试强塞进每次普通 PR。

### CR-10（P1）：改进错误诊断和批处理终态

**当前问题**

service 广泛只返回 `str(exc)`，`TaskWorker` 也只发字符串，没有持久日志或堆栈。COM 区域错误大量忽略。批处理中即使所有文件/行失败，部分 service 仍返回 `success=True, state=WARNING`。无控制台 PyInstaller 产物发生启动/后台异常时很难定位。

**推荐修改方式**

- 明确终态：至少一个成功且有失败为 `WARNING`；全部失败为 `FAILED`；取消为 `CANCELLED`。
- 用户消息保持简洁，同时写入本地滚动日志（异常类型、堆栈、操作、文件；不要记录文档正文/替换值等敏感内容）。
- `TaskWorker.error` 传结构化错误或保留 traceback；启动入口安装异常处理器，在 GUI 中提示日志路径。
- 对可忽略的 COM 区域错误记录 warning，不再完全 `pass`。

**收益**：现场故障可诊断，UI 状态更可信。  
**风险/影响范围**：日志位置和隐私需要明确；避免记录合同内容和 Excel 数据。

### CR-11（P1）：约束输出扩展名与实际文档格式一致

**当前问题**

`sanitize_filename` 只在规则没有后缀时添加模板后缀；用户可输入 `{{名称}}.pdf`，系统实际复制/保存 Word OOXML，却生成 `.pdf` 文件名。`.docm` 模板也可能被命名为 `.docx`，造成宏/格式认知错误。涉及 `src/core/template_merge.py:417-459`。

**推荐修改方式**

- 默认强制输出后缀与模板一致；如果规则含其他后缀，执行前明确拒绝或替换为模板后缀。
- 如果未来支持 PDF，作为独立导出功能实现，不能只改文件名。
- 增加 Windows 保留名、尾随点/空格、长路径和大小写冲突测试。

**收益**：输出文件类型和内容一致，避免用户得到“打不开的 PDF”。  
**风险/影响范围**：过去允许的错误命名规则会被提前纠正。

### CR-12（P2）：把旧 Tkinter 实现完整移入 `_legacy`，收紧打包清单

**当前问题**

旧实现仍位于正式 `src`，约 2000 行：

- `src/word_text_replacer_single_with_add.py`；
- `src/template_merge_gui.py`；
- `src/ui/fluent_widgets.py`；
- `src/ui/theme.py`。

同时存在 `src/ui/theme.py` 和 `src/ui/theme/` 包；Python 当前优先导入后者，旧入口执行 `from ui.theme import ThemeColors` 会失败。PyInstaller hiddenimports 仍包含旧兼容模块和 `ui.fluent_widgets`，但 spec 又排除 tkinter，增加分析噪声和产物不确定性。

**推荐修改方式**

- 确认不再支持旧入口后，将整套 Tkinter 文件移到 `_legacy/v2_tkinter` 或删除；更新上下文菜单说明。
- 仅在外部调用方确有需要时保留三个薄兼容转发模块，并添加弃用注释/测试。
- 从三个 spec 移除不再使用的 hiddenimports；验证 onedir/onefile 产物。

**收益**：减少重复代码、导入冲突和打包体积/告警。  
**风险/影响范围**：可能影响仍直接运行旧脚本的用户；迁移前应在发行说明中声明正式入口。

### CR-13（P2）：在行为稳定后拆分 COM 区域处理，不拆页面框架

**当前问题**

`perform_com_preview` 和 `perform_com_replace` 各约 200 行，Word 启动/关闭、StoryRanges、页眉页脚、脚注尾注等逻辑重复。`generate_batch` 和 `execute_multi_doc_replace` 又各自复制一套 Word 生命周期。

**推荐修改方式**

- 在完成 CR-03 后，把 COM 生命周期和 story range 枚举集中到 `word_com.py`。
- 预览与替换共享“区域枚举 + 结构化 warning”逻辑，操作本身通过简单回调区分。
- 不建议把三个 UI 页面强行合并，也不建议引入通用工作流 DSL。

**收益**：资源释放和区域覆盖规则只维护一份。  
**风险/影响范围**：COM 回归面大，应在 Windows 测试建立后实施。

### CR-14（P2）：删除 service 中的运行时 `TypeError` 回调重试

**当前问题**

`TaskWorker` 已通过 `inspect.signature` 决定参数注入，但 `MergeService` 和 `MultiDocService` 的内部进度回调仍捕获任意 `TypeError` 后用两个参数重试（`src/application/merge_service.py:173-181`、`multi_doc_service.py:128-136`）。如果单参数回调内部本身抛出 `TypeError`，会被误判为签名不兼容并再次调用。

**推荐修改方式**

- 统一进度回调协议为 `Callable[[TaskProgress], None]`；旧双参数适配只在边界处基于签名预检查一次。
- 回调内部异常正常进入 worker 错误处理，不要重试。

**收益**：避免重复副作用和掩盖真实缺陷。  
**风险/影响范围**：检查是否存在外部调用方依赖双参数回调；仓库内现有 UI 均使用单参数。

### CR-15（P2）：超链接扫描应报告逐文件错误

**当前问题**

`scan_hyperlinks` 会返回 `error` 字段，但 `ReplaceService.scan_links` 总是 `SUCCESS`，UI 也不显示该错误，只把文件统计为 0 个链接。损坏/加密/不支持文件会误导用户选择快速模式。

**推荐修改方式**

- 汇总 error：部分失败为 warning、全部失败为 failed。
- 报告中展示失败文件；不要把失败等同于“无链接”。
- 复用 CR-05 的格式校验。

**收益**：模式选择依据更可靠。  
**风险/影响范围**：仅改变错误展示。

### CR-16（P2）：分离运行依赖和构建依赖

**当前问题**

`requirements.txt` 同时包含最终用户运行库和 PyInstaller。版本只有范围，没有一份已验证锁定组合；三个 spec 的 hiddenimports 大段重复，并通过 `collect_submodules` 广泛收集。

**推荐修改方式**

- `requirements.txt` 保留运行依赖，增加 `requirements-build.txt` 或明确的 dev extra 放 PyInstaller/测试工具。
- 为发布流程记录一份锁定版本或构建 constraints，普通开发仍可保留兼容范围。
- 抽出三个 spec 共用的 hiddenimports/data 常量，或至少用小脚本验证三份清单一致。
- 以实际 import 分析结果缩小 `collect_submodules`，修改后比较产物启动和体积。

**收益**：构建更可复现，运行安装更轻。  
**风险/影响范围**：构建脚本和 CI 要同步更新。

### CR-17（P2）：完成低风险性能整理

**当前问题**

- 正则按文件/段落反复编译；
- 快速替换即使 0 个匹配也保存整个文档；
- `MultiDocMappingModel.import_table_data` 在变量循环中调用 `self._variables.index(var)`，形成不必要的二次查找；
- 实时计数在 UI 线程遍历全部缓存文本，大文档/复杂正则时会卡顿；
- 合并单元格可能让 `row.cells` 重复访问同一底层 cell，预览计数有重复风险。

**推荐修改方式**

- 与 CR-01/CR-08 一起复用编译 matcher，0 替换不保存；
- 用 `enumerate(self._variables)`；
- 实时计数增加 150-300ms debounce，并在大数据时移到带 generation id 的只读 worker；
- 表格遍历按底层 XML cell identity 去重，并增加合并单元格测试。

**收益**：大批量文档下响应更稳定。  
**风险/影响范围**：计数去重需确认 Word 合并单元格的期望语义。

### CR-18（P3）：逐步收紧类型，不引入重型模型库

**当前问题**

`MultiDocBatchResult.details: list[dict]`、多个 `ServiceResult[list[dict]]` 和 UI 解包依赖字符串键；core 的 fallback import 在包内导入失败时可能绕回兼容模块，语义不清。

**推荐修改方式**

- 为文件扫描/多文档详情增加少量 dataclass 或 `TypedDict`；
- 包内统一绝对导入 `from core.models ...`，删除不必要的循环 fallback；
- 不需要引入 Pydantic、依赖注入容器或 repository 模式。

**收益**：重构时更容易发现字段拼写错误。  
**风险/影响范围**：低，分批完成即可。

### CR-19（P3）：同步文档、版本和发行元数据

**当前问题**

README 声称测试为 54 项，当前实际为 68 项；macOS spec 硬编码 `2.0.0`，Windows spec 没有统一版本资源；README 的“clean graceful window closing”在 CR-07 完成前表述过强。

**推荐修改方式**

- 建立单一版本源，并由构建脚本/spec 读取；
- README 不写易漂移的测试总数，或由 CI 自动更新；
- 在完成相应验收前收紧安全关闭和预构建产物描述。

**收益**：对外说明与实际能力一致。  
**风险/影响范围**：无功能风险。

### CR-20（P3）：可选的平台体验改进

可在核心可靠性完成后考虑：

- 系统主题模式目前每 2 秒检查一次；macOS 会启动一次 `defaults` 子进程。可改为更长间隔或平台事件，但当前规模下无需优先优化。
- 在 Windows 对接长路径和网络共享时给出清晰错误，不必自行实现路径虚拟层。
- 为 EXE 增加统一图标、版本资源和崩溃日志入口。

这些均不应排在文件正确性、COM 和并发安全之前。

## 7. 测试补强清单

后续 AI Agent 每完成一阶段，应至少补齐以下测试：

### 7.1 纯单元测试

- run 映射、格式保持、不可见字符、正则捕获组、hyperlink 边界；
- 无匹配不落盘，可通过文件 hash/mtime 或 save spy 验证；
- 格式/模式策略矩阵：`.doc/.docx/.docm` × fast/full × Windows/macOS；
- 扫描成功无变量与扫描失败的区别；
- 全失败、部分失败、取消的 `ServiceResult` 终态；
- 备份冲突命名和临时文件清理；
- 输出扩展名强制一致。

### 7.2 Qt 集成测试

- 快捷键重复触发、按钮触发和跨页面触发均不能重入；
- worker 身份与 busy 生命周期；
- 取消后部分结果对话框；
- 关闭窗口不阻塞 GUI，等待危险写任务完成；
- 任务期间输入快照与报告一致。

### 7.3 Windows + Word 集成测试

- COM apartment 初始化/释放、隔离 Word 实例、无残留进程；
- 所有 Word story 区域的预览与替换计数一致；
- `.doc/.docx/.docm`、宏保留、超链接保留、锁定/只读/受保护文件；
- Explorer 传入含中文、空格、长路径文件；
- PyInstaller onefile/onedir EXE 真启动和最小业务冒烟。

## 8. 推荐实施顺序

每个阶段都可以独立提交和验证；不要把所有问题合并成一次大重构。

### 阶段 0：冻结回归基线

- 将本审计中的三个最小复现写成失败测试；
- 增加重复 `Ctrl+R` 的 Qt 失败测试；
- 记录一组真实 `.docx/.docm` 样本的 XML/hash/截图基线；
- 保持现有 68 项测试继续通过。

**完成标准**：新测试能够稳定暴露 CR-01/CR-02，而不是先修改实现。

### 阶段 1：修复快速替换正确性（CR-01、CR-08 的校验部分）

- 实现逻辑文本到 run 的位置映射；
- 修复多 run、多匹配、不可见字符和格式保持；
- 正则预编译并对非法表达式报错；
- 0 替换不保存。

**完成标准**：新增保真测试和现有 core/service 测试全部通过，复杂样本 XML/视觉回归无非目标变化。

### 阶段 2：阻止任务重入并规范取消（CR-02、CR-07）

- 引入简单 coordinator/任务 ID；
- 所有按钮和快捷键共用入口守卫；
- 冻结任务参数和结果报告快照；
- 展示取消后的部分结果。

**完成标准**：重复快捷键、跨页写任务、取消目标和旧 worker 完成竞态测试通过。

### 阶段 3：建立安全写入与备份（CR-04、CR-11）

- python-docx 临时保存、校验、替换；
- 备份不覆盖；
- 输出扩展名与模板一致；
- 模拟磁盘错误/文件锁/清理失败。

**完成标准**：任何注入失败点都保留可打开的原文件，结果准确列出备份和临时文件状态。

### 阶段 4：统一格式策略与扫描结果（CR-05、CR-06、CR-15）

- 集中模式/格式校验；
- `.doc` 强制 COM；
- 扫描错误结构化，超链接扫描显示错误；
- 完整模式扫描与执行覆盖范围一致。

**完成标准**：格式矩阵测试通过，UI 不再出现“扫描成功、执行必然失败”。

### 阶段 5：统一 COM 会话并做 Windows 实机验收（CR-03、CR-09、CR-13）

- 增加 `WordAutomationSession`；
- 在线程内初始化/释放 COM，使用隔离实例；
- 统一 story range 遍历和 warning；
- 在 Windows + Word 上运行完整测试矩阵。

**完成标准**：所有完整模式业务通过，无残留 `WINWORD.EXE`、不影响用户 Word 会话、异常路径可恢复。

### 阶段 6：诊断、终态和性能整理（CR-10、CR-14、CR-17）

- 统一批处理终态；
- 增加隐私安全的错误日志和 traceback；
- 固定单参数进度协议；
- 完成 debounce、去重和低风险循环优化。

**完成标准**：全部失败不再显示成功，现场错误有日志，性能基准不退化。

### 阶段 7：清理和发布工程（CR-12、CR-16、CR-18 至 CR-20）

- 迁移旧 Tkinter 源码，缩小 PyInstaller 清单；
- 分离构建依赖、统一版本源；
- 收紧类型和 README；
- 完成 EXE/portable 构建与启动冒烟。

**完成标准**：正式源树只保留一套 UI；Windows onefile/onedir 和 macOS app 均能构建、启动并通过核心冒烟。

## 9. 最终建议

当前最有价值的工作不是继续扩展功能或重做界面，而是把“写用户文档”变成可证明正确、互斥、可恢复的操作。建议在阶段 1 至阶段 5 完成前冻结新的文档处理功能。

总体架构无需重写。应保留现有三层结构和三个独立工作流，通过小而明确的组件补齐：run 位置映射、任务 coordinator、原子文件写入、格式策略和 COM session。这样能以最低风险获得最大的可靠性与维护性收益。
