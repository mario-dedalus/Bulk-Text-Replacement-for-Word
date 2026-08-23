"""Template batch generation page with strict 3-step workflow, full a11y, naming history, and expanding table."""

from __future__ import annotations

import os
from pathlib import Path

from application.merge_service import MergeService
from application.task_models import ServiceResult, TaskProgress, TaskState
from application.workers import TaskWorker
from core.models import ExcelData, MergeResult
from core.template_merge import get_table_sheet_names
from platform_adapter.capabilities import CAPABILITIES
from ui.delegates.empty_field_delegate import EmptyFieldDelegate
from ui.delegates.mapping_combo_delegate import MappingComboDelegate
from ui.models.field_mapping_model import FieldMappingModel
from ui.theme.theme_manager import THEME
from ui.widgets.card import FluentCard

try:
    from PySide6.QtCore import QSettings, Qt, QThreadPool, Signal
    from PySide6.QtGui import QFont, QFontDatabase
    from PySide6.QtWidgets import (
        QCheckBox,
        QComboBox,
        QFileDialog,
        QFrame,
        QGridLayout,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMessageBox,
        QPlainTextEdit,
        QPushButton,
        QScrollArea,
        QSizePolicy,
        QTableView,
        QVBoxLayout,
        QWidget,
    )

    HAS_QT = True
except ImportError:
    HAS_QT = False


if HAS_QT:

    class MergePage(QWidget):
        """View for Word template + Excel/CSV mail-merge batch generation with strict 3-step workflow."""

        statusMessage = Signal(str, str)
        statusMetrics = Signal(str)
        statusProgress = Signal(object)
        statusHideProgress = Signal()

        def __init__(self, parent: QWidget | None = None):
            super().__init__(parent)
            self._thread_pool = QThreadPool.globalInstance()
            self._active_worker: TaskWorker | None = None
            self._scan_generation_id: int = 0

            # State data
            self.excel_data: ExcelData | None = None
            self.mapping_model = FieldMappingModel(parent=self)
            self.combo_delegate = MappingComboDelegate(parent=self)
            self.empty_field_delegate = EmptyFieldDelegate(parent=self)

            self._build_ui()
            self._load_naming_rule_history()
            self._update_step_states()

        def _build_ui(self):
            root_layout = QVBoxLayout(self)
            root_layout.setContentsMargins(0, 0, 0, 0)
            root_layout.setSpacing(0)

            # ---------------- Scroll Area for Content Cards ----------------
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setFrameShape(QScrollArea.NoFrame)
            self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.scroll_area.setAccessibleName("模板批量生成内容滚动区")

            self.content_widget = QWidget()
            self.content_widget.setProperty("isPageContent", True)
            layout = QVBoxLayout(self.content_widget)
            layout.setContentsMargins(18, 10, 18, 12)
            layout.setSpacing(12)

            # ---------------- Step 1: Data Sources & Filename ----------------
            self.card_step1 = FluentCard(
                title="步骤 1：选择数据源与命名规则",
                subtitle="指定 Word 模板、数据表格与输出命名规则（支持直接使用 Excel/CSV 列名如 {{序号}}）",
            )
            self.card_step1.setAccessibleName("步骤一：数据源与命名规则卡片")
            self.card_step1.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            s1_grid = QGridLayout()
            s1_grid.setSpacing(8)

            # Row 0: Word Template
            lbl_tmpl = QLabel("Word 模板：")
            s1_grid.addWidget(lbl_tmpl, 0, 0)
            self.edt_template = QLineEdit()
            self.edt_template.setPlaceholderText("选择 .docx / .docm 模板文件..." if CAPABILITIES.is_macos else "选择 .docx / .docm / .doc 模板文件...")
            self.edt_template.setAccessibleName("Word 模板文件路径输入框")
            self.edt_template.setAccessibleDescription("输入或浏览选择包含 {{变量}} 的 Word 模板文档")
            lbl_tmpl.setBuddy(self.edt_template)
            self.edt_template.textChanged.connect(self._on_source_files_changed)
            s1_grid.addWidget(self.edt_template, 0, 1)

            self.btn_browse_tmpl = QPushButton("浏览…")
            self.btn_browse_tmpl.setAccessibleName("浏览 Word 模板按钮")
            self.btn_browse_tmpl.clicked.connect(self._browse_template)
            s1_grid.addWidget(self.btn_browse_tmpl, 0, 2)

            # Row 1: Data Table (Excel / CSV)
            lbl_excel = QLabel("数据表格：")
            s1_grid.addWidget(lbl_excel, 1, 0)
            self.edt_excel = QLineEdit()
            self.edt_excel.setPlaceholderText("选择 .xlsx / .csv 数据源表格...")
            self.edt_excel.setAccessibleName("数据源表格路径输入框")
            self.edt_excel.setAccessibleDescription("输入或浏览选择包含表头与待填充行的 Excel (.xlsx) 或 CSV (.csv) 表格")
            lbl_excel.setBuddy(self.edt_excel)
            self.edt_excel.textChanged.connect(self._on_source_files_changed)
            s1_grid.addWidget(self.edt_excel, 1, 1)

            self.btn_browse_excel = QPushButton("浏览…")
            self.btn_browse_excel.setAccessibleName("浏览数据源表格按钮")
            self.btn_browse_excel.clicked.connect(self._browse_excel)
            s1_grid.addWidget(self.btn_browse_excel, 1, 2)

            # Row 2: Sheet Selector (Enabled only when multiple sheets exist)
            self.lbl_sheet = QLabel("工作表：")
            self.lbl_sheet.setEnabled(False)
            s1_grid.addWidget(self.lbl_sheet, 2, 0)

            self.combo_sheet = QComboBox()
            self.combo_sheet.setEnabled(False)
            self.combo_sheet.addItem("默认表格")
            self.combo_sheet.setAccessibleName("工作表选择下拉框")
            self.combo_sheet.setAccessibleDescription("包含多个工作表时可在此切换目标工作表")
            self.lbl_sheet.setBuddy(self.combo_sheet)
            self.combo_sheet.currentIndexChanged.connect(self._on_sheet_changed)
            s1_grid.addWidget(self.combo_sheet, 2, 1)

            # Row 3: Output Directory
            lbl_out = QLabel("输出目录：")
            s1_grid.addWidget(lbl_out, 3, 0)
            self.edt_output = QLineEdit()
            self.edt_output.setPlaceholderText("指定生成文档的保存目录...")
            self.edt_output.setAccessibleName("输出目录路径输入框")
            self.edt_output.setAccessibleDescription("输入或浏览选择生成结果文档的保存目录")
            lbl_out.setBuddy(self.edt_output)
            self.edt_output.textChanged.connect(self._on_output_config_changed)
            s1_grid.addWidget(self.edt_output, 3, 1)

            btn_browse_out = QPushButton("浏览…")
            btn_browse_out.setAccessibleName("浏览输出目录按钮")
            btn_browse_out.clicked.connect(self._browse_output)
            s1_grid.addWidget(btn_browse_out, 3, 2)

            # Row 4: Filename rule with History ComboBox (10 items)
            lbl_rule = QLabel("命名规则：")
            s1_grid.addWidget(lbl_rule, 4, 0)

            self.edt_fn_rule = QComboBox()
            self.edt_fn_rule.setEditable(True)
            self.edt_fn_rule.setMaxCount(10)
            self.edt_fn_rule.setInsertPolicy(QComboBox.NoInsert)
            self.edt_fn_rule.setToolTip("支持使用 Excel/CSV 表头字段名（如 {{序号}}、{{签约人}}）或 Word 变量名动态命名。")
            self.edt_fn_rule.setAccessibleName("输出文件名规则下拉框")
            self.edt_fn_rule.setAccessibleDescription("支持选择历史记录或直接输入 {{字段名}} 动态构建输出文档名称")
            lbl_rule.setBuddy(self.edt_fn_rule)
            self.edt_fn_rule.editTextChanged.connect(self._on_output_config_changed)
            s1_grid.addWidget(self.edt_fn_rule, 4, 1)

            self.btn_scan = QPushButton("扫描并匹配")
            self.btn_scan.setAccessibleName("扫描并匹配字段按钮")
            self.btn_scan.setAccessibleDescription("扫描 Word 模板中的所有变量并自动与数据表头进行对应匹配")
            self.btn_scan.clicked.connect(self.scan_template)
            s1_grid.addWidget(self.btn_scan, 4, 2)

            self.card_step1.addLayout(s1_grid)
            layout.addWidget(self.card_step1, 0)

            # ---------------- Step 2: Field Mapping Table (Expands Vertically) ----------------
            self.card_step2 = FluentCard(
                title="步骤 2：字段映射与空字段处理",
                subtitle="设置数据列映射，并为每个空字段选择保持原变量、替换为空或自定义",
            )
            self.card_step2.setAccessibleName("步骤二：字段映射与空字段处理卡片")
            self.card_step2.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

            self.tbl_mapping = QTableView()
            self.tbl_mapping.setModel(self.mapping_model)
            self.tbl_mapping.setItemDelegateForColumn(FieldMappingModel.COL_EXCEL, self.combo_delegate)
            self.tbl_mapping.setItemDelegateForColumn(FieldMappingModel.COL_DEFAULT, self.empty_field_delegate)
            self.tbl_mapping.setMinimumHeight(120)
            self.tbl_mapping.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.tbl_mapping.setAccessibleName("字段映射关系表格")
            self.tbl_mapping.setAccessibleDescription("展示 Word 变量与数据列的映射关系，第三列可选择空字段处理方式")
            
            # Configure 4 columns
            header = self.tbl_mapping.horizontalHeader()
            header.setSectionResizeMode(FieldMappingModel.COL_VARIABLE, QHeaderView.Stretch)
            header.setSectionResizeMode(FieldMappingModel.COL_EXCEL, QHeaderView.Stretch)
            header.setSectionResizeMode(FieldMappingModel.COL_DEFAULT, QHeaderView.Stretch)
            header.setSectionResizeMode(FieldMappingModel.COL_STATUS, QHeaderView.Fixed)
            header.resizeSection(FieldMappingModel.COL_STATUS, 95)
            self.tbl_mapping.clicked.connect(self._on_table_clicked)
            
            self.card_step2.addWidget(self.tbl_mapping, 1)

            self.lbl_mapping_summary = QLabel("请先选择有效 Word 模板与数据表格并点击“扫描并匹配”。")
            self.lbl_mapping_summary.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
            self.lbl_mapping_summary.setAccessibleName("字段映射统计说明")
            self.card_step2.addWidget(self.lbl_mapping_summary, 0)

            layout.addWidget(self.card_step2, 1)

            # ---------------- Step 3: Options, Preview & Log (Compact Fixed Height) ----------------
            self.card_step3 = FluentCard(
                title="步骤 3：选项与生成日志",
                subtitle="设置选项并在下方查看生成日志",
            )
            self.card_step3.setAccessibleName("步骤三：选项与生成日志卡片")
            self.card_step3.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            opts_row = QHBoxLayout()
            opts_row.setSpacing(16)

            com_state_hint = "" if CAPABILITIES.has_word_com else "（需 Windows 与 Microsoft Word）"
            self.cb_use_com = QCheckBox(f"启用 Word COM 完整模式{com_state_hint}")
            self.cb_use_com.setEnabled(CAPABILITIES.has_word_com)
            self.cb_use_com.setAccessibleName("启用 Word COM 完整模式复选框")
            opts_row.addWidget(self.cb_use_com)

            opts_row.addStretch()

            self.card_step3.addLayout(opts_row)

            # Log / Output Viewer with Platform Fixed Font
            self.txt_log = QPlainTextEdit()
            self.txt_log.setFixedHeight(95)
            self.txt_log.setReadOnly(True)
            self.txt_log.setPlaceholderText("预览和批量生成的日志将在此处显示...")
            self.txt_log.setAccessibleName("批量生成执行日志区域")

            fixed_font = QFontDatabase.systemFont(QFontDatabase.FixedFont)
            fixed_font.setPointSize(12)
            self.txt_log.setFont(fixed_font)
            self.card_step3.addWidget(self.txt_log)

            layout.addWidget(self.card_step3, 0)

            self.scroll_area.setWidget(self.content_widget)
            root_layout.addWidget(self.scroll_area, 1)

            # ---------------- Fixed Bottom Action Bar (Outside ScrollArea) ----------------
            action_container = QFrame(self)
            action_container.setStyleSheet(f"""
                QFrame {{
                    background-color: {THEME.tokens.card_bg};
                    border-top: 1px solid {THEME.tokens.border_subtle};
                    padding: 8px 18px;
                }}
            """)
            action_bar = QHBoxLayout(action_container)
            action_bar.setContentsMargins(0, 0, 0, 0)
            action_bar.setSpacing(10)

            self.btn_preview = QPushButton("预览前 5 行")
            self.btn_preview.setAccessibleName("预览前五行日志按钮")
            self.btn_preview.clicked.connect(self.preview_merge)
            action_bar.addWidget(self.btn_preview)

            action_bar.addStretch()

            self.btn_generate = QPushButton("开始批量生成")
            self.btn_generate.setProperty("isPrimary", True)
            self.btn_generate.setAccessibleName("开始批量生成主按钮")
            self.btn_generate.clicked.connect(self.start_batch_merge)
            action_bar.addWidget(self.btn_generate)

            root_layout.addWidget(action_container)

        # ---------------- State & Event Handlers ----------------

        def _load_naming_rule_history(self):
            settings = QSettings("BulkReplacementForWord", "WordTextReplacer")
            history = settings.value("merge/naming_rule_history", None)
            if not history or not isinstance(history, list):
                history = [
                    "{{甲方名称}}-{{乙方名称}}-合同.docx",
                    "{{序号}}-{{甲方名称}}-合同.docx",
                    "{{序号}}-{{甲方名称}}.docx",
                    "{{甲方名称}}-合同.docx",
                ]
            self.edt_fn_rule.blockSignals(True)
            self.edt_fn_rule.clear()
            for item in history[:10]:
                if item:
                    self.edt_fn_rule.addItem(str(item))
            if self.edt_fn_rule.count() > 0:
                self.edt_fn_rule.setCurrentIndex(0)
            self.edt_fn_rule.blockSignals(False)

        def _save_naming_rule_to_history(self, rule: str):
            rule = rule.strip()
            if not rule:
                return
            settings = QSettings("BulkReplacementForWord", "WordTextReplacer")
            history = settings.value("merge/naming_rule_history", None)
            if not history or not isinstance(history, list):
                history = []
            
            # Deduplicate and place latest rule at the beginning
            new_history = [rule] + [h for h in history if h and h != rule]
            new_history = new_history[:10]
            settings.setValue("merge/naming_rule_history", new_history)

            # Reload combo box items
            self.edt_fn_rule.blockSignals(True)
            self.edt_fn_rule.clear()
            for item in new_history:
                self.edt_fn_rule.addItem(str(item))
            self.edt_fn_rule.setEditText(rule)
            self.edt_fn_rule.blockSignals(False)

        def _on_table_clicked(self, index):
            if index.isValid() and index.column() in (
                FieldMappingModel.COL_EXCEL,
                FieldMappingModel.COL_DEFAULT,
            ):
                self._open_mapping_editor(index)

        def _open_mapping_editor(self, index):
            self.tbl_mapping.edit(index)

        def _update_sheet_selector(self):
            e_path = self.edt_excel.text().strip()
            if getattr(self, "_last_excel_path", None) == e_path:
                return
            self._last_excel_path = e_path

            sheet_names = get_table_sheet_names(e_path)
            self.combo_sheet.blockSignals(True)
            self.combo_sheet.clear()
            if len(sheet_names) > 1:
                self.combo_sheet.addItems(sheet_names)
                self.combo_sheet.setEnabled(True)
                self.lbl_sheet.setEnabled(True)
            elif len(sheet_names) == 1:
                self.combo_sheet.addItem(sheet_names[0])
                self.combo_sheet.setEnabled(False)
                self.lbl_sheet.setEnabled(False)
            else:
                self.combo_sheet.addItem("默认表格")
                self.combo_sheet.setEnabled(False)
                self.lbl_sheet.setEnabled(False)
            self.combo_sheet.blockSignals(False)

        def _on_sheet_changed(self, index: int = 0):
            """Triggered when the user switches the active Excel sheet."""
            if self.excel_data is not None:
                self.excel_data = None
                self.mapping_model.set_data([], {})
                self.lbl_mapping_summary.setText("工作表已切换，请重新点击“扫描并匹配”。")
                self.lbl_mapping_summary.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
                self._update_step_states()

        def _on_source_files_changed(self):
            """Triggered when Word template or data table paths change (invalidates mapping)."""
            self._update_sheet_selector()
            t_path = self.edt_template.text().strip()
            e_path = self.edt_excel.text().strip()

            # Enable scan button only when both source files exist on disk
            can_scan = bool(t_path and e_path and os.path.isfile(t_path) and os.path.isfile(e_path))
            self.btn_scan.setEnabled(can_scan)

            # Invalidate previous scan results only when template or excel changes
            if self.excel_data is not None:
                self.excel_data = None
                self.mapping_model.set_data([], {})
                self.lbl_mapping_summary.setText("模板或数据源路径已更改，请重新点击“扫描并匹配”。")
                self.lbl_mapping_summary.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
                self._update_step_states()

        def _on_output_config_changed(self):
            """Triggered when output directory or filename rule changes (does NOT invalidate mapping)."""
            pass

        def _update_step_states(self):
            t_path = self.edt_template.text().strip()
            e_path = self.edt_excel.text().strip()
            can_scan = bool(t_path and e_path and os.path.isfile(t_path) and os.path.isfile(e_path))
            self.btn_scan.setEnabled(can_scan)

            has_scanned = self.excel_data is not None and len(self.mapping_model.get_fields()) > 0
            self.card_step2.setEnabled(has_scanned)
            self.btn_preview.setEnabled(has_scanned)
            self.btn_generate.setEnabled(has_scanned)

        def _browse_template(self):
            filter_str = (
                "Word 模板 (*.docx *.docm);;所有文件 (*.*)"
                if CAPABILITIES.is_macos
                else "Word 模板 (*.docx *.docm *.doc);;所有文件 (*.*)"
            )
            path, _ = QFileDialog.getOpenFileName(self, "选择 Word 模板", "", filter_str)
            if path:
                if CAPABILITIES.is_macos and path.lower().endswith(".doc"):
                    QMessageBox.warning(self, "格式不支持", "macOS 首版暂不支持旧版 .doc 格式，请使用 .docx 格式模板。")
                    return
                self.edt_template.setText(os.path.abspath(path))
                if not self.edt_output.text().strip():
                    self.edt_output.setText(str(Path(path).parent / "Generated"))

        def _browse_excel(self):
            filter_str = (
                "表格数据文件 (*.xlsx *.csv *.xlsm);;"
                "Excel 表格 (*.xlsx *.xlsm);;"
                "CSV 文本表格 (*.csv);;"
                "所有文件 (*.*)"
            )
            path, _ = QFileDialog.getOpenFileName(
                self, "选择数据源表格 (Excel / CSV)", "", filter_str
            )
            if path:
                self.edt_excel.setText(os.path.abspath(path))

        def _browse_output(self):
            path = QFileDialog.getExistingDirectory(self, "选择生成文档输出目录")
            if path:
                self.edt_output.setText(os.path.abspath(path))

        # ---------------- Actions ----------------

        def scan_template(self):
            t_path = self.edt_template.text().strip()
            e_path = self.edt_excel.text().strip()
            selected_sheet = self.combo_sheet.currentText().strip() if self.combo_sheet.count() > 1 else None

            err = MergeService.validate_sources(t_path, e_path)
            if err:
                QMessageBox.warning(self, "提示", err)
                return

            self._scan_generation_id += 1
            current_scan_gen = self._scan_generation_id

            self.statusMessage.emit("正在扫描模板与数据表...", "running")
            self._set_scan_inputs_locked(True)

            def _scan_worker_fn(
                template_path: str,
                excel_path: str,
                use_com: bool,
                sheet_name: str | None = None,
                progress_cb=None,
                cancel_token=None,
            ):
                res = MergeService.scan_template_and_excel(
                    template_path=template_path,
                    excel_path=excel_path,
                    use_com=use_com,
                    sheet_name=sheet_name,
                    progress_cb=progress_cb,
                    cancel_token=cancel_token,
                )
                return (res, current_scan_gen, template_path, excel_path, sheet_name)

            worker = TaskWorker(
                _scan_worker_fn,
                template_path=t_path,
                excel_path=e_path,
                use_com=self.cb_use_com.isChecked(),
                sheet_name=selected_sheet,
            )
            self._active_worker = worker

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_scan_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_scan_inputs_locked(False))

            self._thread_pool.start(worker)

        def _on_scan_finished(self, result_tuple: tuple):
            self.statusHideProgress.emit()
            res, gen_id, scanned_tmpl, scanned_excel, *extra = result_tuple
            scanned_sheet = extra[0] if extra else None

            # Discard stale result if inputs or generation changed during background scan
            if gen_id != self._scan_generation_id:
                return
            current_sheet = self.combo_sheet.currentText().strip() if self.combo_sheet.count() > 1 else None
            if (
                self.edt_template.text().strip() != scanned_tmpl
                or self.edt_excel.text().strip() != scanned_excel
                or current_sheet != scanned_sheet
            ):
                return

            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("扫描已取消", "warning")
                    return
                self.statusMessage.emit("扫描失败", "error")
                QMessageBox.critical(self, "扫描失败", res.error)
                return

            fields, excel_data, mapping = res.data
            self.excel_data = excel_data
            self.combo_delegate.set_headers(excel_data.headers)
            self.mapping_model.set_data(fields, mapping)

            missing = self.mapping_model.get_missing_count()
            excel_cols_str = "、".join(excel_data.headers[:8])
            if len(excel_data.headers) > 8:
                excel_cols_str += " 等"
            sheet_info = f"（工作表：{scanned_sheet}）" if scanned_sheet else ""
            summary = (
                f"扫描成功{sheet_info}：检测到 {len(fields)} 个模板变量，{len(excel_data.rows)} 行数据；"
                f"已匹配 {len(fields) - missing} 个，缺失 {missing} 个。\n"
                f"💡 可用命名变量（表格字段）：{excel_cols_str}"
            )
            self.lbl_mapping_summary.setText(summary)
            status_color = THEME.tokens.error if missing > 0 else THEME.tokens.success
            self.lbl_mapping_summary.setStyleSheet(f"color: {status_color}; font-weight: 500;")

            self.statusMessage.emit("扫描完成", "success")
            self.statusMetrics.emit(f"已加载 {len(excel_data.rows)} 行数据")
            self._update_step_states()

        def _set_scan_inputs_locked(self, locked: bool):
            self.btn_scan.setEnabled(not locked)
            self.edt_template.setEnabled(not locked)
            self.edt_excel.setEnabled(not locked)
            self.btn_browse_tmpl.setEnabled(not locked)
            self.btn_browse_excel.setEnabled(not locked)
            has_multi_sheets = self.combo_sheet.count() > 1
            self.combo_sheet.setEnabled(not locked and has_multi_sheets)
            self.lbl_sheet.setEnabled(not locked and has_multi_sheets)

        def preview_merge(self):
            if not self.excel_data:
                return

            t_path = self.edt_template.text().strip()
            out_dir = self.edt_output.text().strip() or str(Path.cwd() / "Generated")
            fn_rule = self.edt_fn_rule.currentText().strip() or "{{甲方名称}}-合同.docx"
            fields = self.mapping_model.get_fields()
            mapping = self.mapping_model.get_mapping()
            defaults = self.mapping_model.get_defaults()
            empty_behaviors = self.mapping_model.get_empty_behaviors()

            snippets = MergeService.generate_preview_snippets(
                template_path=t_path,
                excel_data=self.excel_data,
                fields=fields,
                mapping=mapping,
                output_folder=out_dir,
                filename_rule=fn_rule,
                default_values=defaults,
                empty_field_behaviors=empty_behaviors,
                max_rows=5,
            )

            lines = [
                "模板批量生成预览（前 5 行）",
                "=" * 60,
                f"Word 模板：{os.path.basename(t_path)}",
                f"数据表格：{os.path.basename(self.edt_excel.text().strip())}",
                f"输出目录：{out_dir}",
                f"命名规则：{fn_rule}",
                "=" * 60,
                "",
            ]

            for snip in snippets:
                lines.append(f"【表格第 {snip['excel_row']} 行】→ 输出文件：{snip['filename']}")
                for f_info in snip["fields"]:
                    behavior = f_info["empty_behavior"]
                    if behavior == "custom":
                        tag = f"{f_info['value']} （自定义替换）"
                    elif behavior == "keep_variable":
                        tag = f"{{{{{f_info['field']}}}}}（保持原变量）"
                    elif behavior == "replace_empty":
                        tag = "（替换为空）"
                    else:
                        tag = f_info["value"]
                    lines.append(f"   {{{{{f_info['field']}}}}} => {tag}")
                lines.append("")

            self.txt_log.setPlainText("\n".join(lines))
            self.statusMessage.emit("预览生成就绪", "normal")

        def start_batch_merge(self):
            if not self.excel_data:
                return

            t_path = self.edt_template.text().strip()
            out_dir = self.edt_output.text().strip()
            fn_rule = self.edt_fn_rule.currentText().strip()

            if not out_dir:
                QMessageBox.warning(self, "提示", "请指定生成文档的输出目录。")
                return
            if not fn_rule:
                QMessageBox.warning(self, "提示", "请输入输出文件名规则。")
                return

            # Save rule to history (top 10 items)
            self._save_naming_rule_to_history(fn_rule)

            mapping = self.mapping_model.get_mapping()
            defaults = self.mapping_model.get_defaults()
            empty_behaviors = self.mapping_model.get_empty_behaviors()
            missing = [
                f"{{{{{f}}}}}"
                for f, h in mapping.items()
                if not h and empty_behaviors.get(f) == FieldMappingModel.EMPTY_KEEP
            ]
            if missing:
                reply = QMessageBox.question(
                    self,
                    "存在未匹配变量",
                    "以下模板变量未匹配数据列，并设置为保持原变量：\n\n"
                    + "\n".join(missing)
                    + "\n\n是否继续批量生成？",
                )
                if reply != QMessageBox.Yes:
                    return

            self._set_ui_busy(True)
            self.statusMessage.emit("正在批量生成文档...", "running")
            self.txt_log.setPlainText("正在开始批量生成任务...\n")

            worker = TaskWorker(
                MergeService.execute_batch_merge,
                template_path=t_path,
                excel_data=self.excel_data,
                mapping=mapping,
                output_folder=out_dir,
                filename_rule=fn_rule,
                use_com=self.cb_use_com.isChecked(),
                default_values=defaults,
                empty_field_behaviors=empty_behaviors,
            )
            self._active_worker = worker

            worker.signals.progress.connect(self._on_merge_progress)
            worker.signals.result.connect(self._on_merge_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_ui_busy(False))

            self._thread_pool.start(worker)

        def _on_merge_progress(self, progress: TaskProgress):
            self.statusProgress.emit(progress)
            if progress.data and isinstance(progress.data, MergeResult):
                r = progress.data
                status_str = "✓ 成功" if r.success else f"✕ 失败 ({r.error})"
                self.txt_log.appendPlainText(f"第 {r.excel_row} 行 | {r.filename} | {status_str}")

        def _on_merge_finished(self, res: ServiceResult):
            self.statusHideProgress.emit()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("批量生成已取消", "warning")
                    self.txt_log.appendPlainText("\n任务已由用户取消。")
                    return
                self.statusMessage.emit("生成失败", "error")
                QMessageBox.critical(self, "生成失败", res.error)
                return

            results: list[MergeResult] = res.data or []
            success_count = sum(1 for r in results if r.success)
            fail_count = len(results) - success_count
            out_dir = self.edt_output.text().strip()

            summary = f"批量生成完成：成功 {success_count} 份，失败 {fail_count} 份；保存在：{out_dir}"
            self.statusMessage.emit("生成完成", "success" if fail_count == 0 else "warning")
            QMessageBox.information(self, "生成完成", summary)

        def _on_task_error(self, err_msg: str):
            self.statusHideProgress.emit()
            self.statusMessage.emit("任务异常", "error")
            QMessageBox.critical(self, "执行异常", err_msg)

        def cancel_active_task(self):
            if self._active_worker:
                self._active_worker.cancel()
                self.statusMessage.emit("正在请求取消任务...", "warning")

        def _set_ui_busy(self, is_busy: bool):
            self.btn_scan.setEnabled(not is_busy)
            self.btn_preview.setEnabled(not is_busy)
            self.btn_generate.setEnabled(not is_busy)

else:

    class MergePage:  # type: ignore
        pass
