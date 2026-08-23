"""Multi-document differentiated matching and replacement page with Fluent Design controls."""

from __future__ import annotations

import os
from pathlib import Path

from application.multi_doc_service import MultiDocService
from application.task_models import ServiceResult, TaskProgress, TaskState
from application.workers import TaskWorker
from core.models import ExcelData, MultiDocBatchResult, MultiDocItem
from core.template_merge import get_table_sheet_names, load_table_data
from platform_adapter.capabilities import CAPABILITIES
from ui.dialogs.result_dialog import ResultDialog
from ui.models.multi_doc_mapping_model import MultiDocMappingModel
from ui.theme.theme_manager import THEME
from ui.widgets.card import FluentCard

try:
    from PySide6.QtCore import Qt, QThreadPool, Signal
    from PySide6.QtWidgets import (
        QButtonGroup,
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
        QPushButton,
        QRadioButton,
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

    class MultiDocPage(QWidget):
        """View for multi-document differentiated variable scanning, cell editing, and batch replacement."""

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
            self.mapping_model = MultiDocMappingModel(parent=self)

            self._build_ui()
            self.mapping_model.dataChanged.connect(self._on_table_data_changed)
            self.mapping_model.rowsInserted.connect(self._on_docs_changed)
            self.mapping_model.rowsRemoved.connect(self._on_docs_changed)
            self.mapping_model.modelReset.connect(self._on_docs_changed)

        def _build_ui(self):
            root_layout = QVBoxLayout(self)
            root_layout.setContentsMargins(0, 0, 0, 0)
            root_layout.setSpacing(0)

            # ---------------- Scroll Area for Content Cards ----------------
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setFrameShape(QScrollArea.NoFrame)
            self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.scroll_area.setAccessibleName("多文档匹配替换内容滚动区")

            self.content_widget = QWidget()
            self.content_widget.setProperty("isPageContent", True)
            layout = QVBoxLayout(self.content_widget)
            layout.setContentsMargins(18, 10, 18, 12)
            layout.setSpacing(12)

            # ---------------- Card 1: Documents & Optional Excel ----------------
            self.card_step1 = FluentCard(
                title="步骤 1：添加待处理文档与数据源",
                subtitle="支持批量添加多份 Word 文档，可直接在下方表格手动录入或导入 Excel/CSV 辅助填报",
            )
            self.card_step1.setAccessibleName("步骤一：添加文档与数据源卡片")
            self.card_step1.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            s1_layout = QVBoxLayout()
            s1_layout.setSpacing(10)

            # Row 1: Document action buttons
            doc_btn_row = QHBoxLayout()
            doc_btn_row.setSpacing(8)

            self.btn_add_docs = QPushButton("添加 Word 文档...")
            self.btn_add_docs.setAccessibleName("添加待处理 Word 文档按钮")
            self.btn_add_docs.clicked.connect(self._browse_docs)
            doc_btn_row.addWidget(self.btn_add_docs)

            self.btn_remove_docs = QPushButton("移除所选")
            self.btn_remove_docs.setAccessibleName("移除所选文档按钮")
            self.btn_remove_docs.clicked.connect(self._remove_selected_docs)
            doc_btn_row.addWidget(self.btn_remove_docs)

            self.btn_clear_docs = QPushButton("清空文档")
            self.btn_clear_docs.setAccessibleName("清空所有已添加文档按钮")
            self.btn_clear_docs.clicked.connect(self._clear_docs)
            doc_btn_row.addWidget(self.btn_clear_docs)

            doc_btn_row.addStretch()

            self.lbl_doc_count = QLabel("已选 0 个文档")
            self.lbl_doc_count.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
            self.lbl_doc_count.setAccessibleName("已选文档数量统计")
            doc_btn_row.addWidget(self.lbl_doc_count)

            s1_layout.addLayout(doc_btn_row)

            # Row 2: Excel / CSV data table importer
            excel_row = QHBoxLayout()
            excel_row.setSpacing(8)

            lbl_excel = QLabel("数据表格（选填）：")
            lbl_excel.setAccessibleName("数据源表格标签")
            excel_row.addWidget(lbl_excel)

            self.edt_excel_path = QLineEdit()
            self.edt_excel_path.setPlaceholderText("选择 .xlsx / .csv 数据源表格（用于自动填充每行对应数据）...")
            self.edt_excel_path.setAccessibleName("数据源表格文件路径输入框")
            lbl_excel.setBuddy(self.edt_excel_path)
            self.edt_excel_path.textChanged.connect(self._on_excel_path_changed)
            excel_row.addWidget(self.edt_excel_path, 1)

            self.btn_browse_excel = QPushButton("浏览…")
            self.btn_browse_excel.setAccessibleName("浏览选择数据源表格按钮")
            self.btn_browse_excel.clicked.connect(self._browse_excel)
            excel_row.addWidget(self.btn_browse_excel)

            self.lbl_sheet = QLabel("工作表：")
            self.lbl_sheet.setEnabled(False)
            self.lbl_sheet.setAccessibleName("工作表标签")
            excel_row.addWidget(self.lbl_sheet)

            self.combo_sheet = QComboBox()
            self.combo_sheet.setEnabled(False)
            self.combo_sheet.addItem("默认表格")
            self.combo_sheet.setAccessibleName("工作表选择下拉框")
            self.combo_sheet.setAccessibleDescription("包含多个工作表时可在此切换目标工作表")
            self.combo_sheet.currentIndexChanged.connect(self._on_sheet_changed)
            excel_row.addWidget(self.combo_sheet)

            self.btn_fill_excel = QPushButton("从表格填充")
            self.btn_fill_excel.setAccessibleName("从表格自动填充数据按钮")
            self.btn_fill_excel.clicked.connect(self._fill_from_excel)
            excel_row.addWidget(self.btn_fill_excel)

            self.btn_clear_data = QPushButton("清空填报数据")
            self.btn_clear_data.setAccessibleName("清空表格中已录入的替换值按钮")
            self.btn_clear_data.clicked.connect(self._clear_table_data)
            excel_row.addWidget(self.btn_clear_data)

            s1_layout.addLayout(excel_row)

            self.card_step1.addLayout(s1_layout)
            layout.addWidget(self.card_step1, 0)

            # ---------------- Card 2: Mapping Table (Expands Vertically) ----------------
            self.card_step2 = FluentCard(
                title="步骤 2：文档与变量值对应关系（支持直接编辑单元格手动录入 / 覆盖）",
                subtitle="双击任意单元格即可直接手动输入或修改替换值；也可通过上方或下方工具从 Excel 导入",
            )
            self.card_step2.setAccessibleName("步骤二：文档与变量对应表格卡片")
            self.card_step2.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

            self.tbl_view = QTableView()
            self.tbl_view.setModel(self.mapping_model)
            self.tbl_view.setMinimumHeight(180)
            self.tbl_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            self.tbl_view.setAccessibleName("多文档变量值对应表格")
            self.tbl_view.setAccessibleDescription("展示各文档及对应变量列，双击单元格可就地修改替换值")

            # Table Header
            h_header = self.tbl_view.horizontalHeader()
            h_header.setStretchLastSection(False)
            self.tbl_view.clicked.connect(self._on_table_cell_clicked)
            self.card_step2.addWidget(self.tbl_view, 1)

            # Tool buttons below table
            tbl_tools_row = QHBoxLayout()
            tbl_tools_row.setSpacing(8)

            self.btn_scan_vars = QPushButton("自动扫描文档变量")
            self.btn_scan_vars.setAccessibleName("自动扫描所有文档中的占位变量按钮")
            self.btn_scan_vars.clicked.connect(self.scan_variables)
            tbl_tools_row.addWidget(self.btn_scan_vars)

            self.btn_align_smart = QPushButton("按列名智能对齐")
            self.btn_align_smart.setAccessibleName("按 Excel 列名智能匹配填入数据按钮")
            self.btn_align_smart.clicked.connect(lambda: self._apply_excel_import(match_by_header=True))
            tbl_tools_row.addWidget(self.btn_align_smart)

            self.btn_align_seq = QPushButton("按表格列顺序填充")
            self.btn_align_seq.setAccessibleName("按表格列顺序依次填入数据按钮")
            self.btn_align_seq.clicked.connect(lambda: self._apply_excel_import(match_by_header=False))
            tbl_tools_row.addWidget(self.btn_align_seq)

            tbl_tools_row.addStretch()

            self.btn_move_up = QPushButton("上移")
            self.btn_move_up.setAccessibleName("上移选中行按钮")
            self.btn_move_up.clicked.connect(self._move_row_up)
            tbl_tools_row.addWidget(self.btn_move_up)

            self.btn_move_down = QPushButton("下移")
            self.btn_move_down.setAccessibleName("下移选中行按钮")
            self.btn_move_down.clicked.connect(self._move_row_down)
            tbl_tools_row.addWidget(self.btn_move_down)

            self.card_step2.addLayout(tbl_tools_row)

            self.lbl_table_summary = QLabel("请添加 Word 文档并点击“自动扫描文档变量”以发现各文档中的变量。")
            self.lbl_table_summary.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
            self.lbl_table_summary.setAccessibleName("多文档对应状态统计说明")
            self.card_step2.addWidget(self.lbl_table_summary)

            layout.addWidget(self.card_step2, 1)

            # ---------------- Card 3: Options ----------------
            self.card_step3 = FluentCard(
                title="步骤 3：处理选项与执行",
                subtitle="设置文件保存方式与替换引擎模式",
            )
            self.card_step3.setAccessibleName("步骤三：处理选项卡片")
            self.card_step3.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Maximum)

            opts_grid = QGridLayout()
            opts_grid.setSpacing(10)

            # Destination
            self.rb_inplace = QRadioButton("直接修改原文件")
            self.rb_inplace.setChecked(True)
            self.rb_inplace.setAccessibleName("直接修改原文件单选按钮")
            self.rb_inplace.toggled.connect(self._on_dest_mode_toggled)
            opts_grid.addWidget(self.rb_inplace, 0, 0)

            self.cb_backup = QCheckBox("创建备份文件 (.backup)")
            self.cb_backup.setChecked(True)
            self.cb_backup.setAccessibleName("创建备份文件复选框")
            opts_grid.addWidget(self.cb_backup, 0, 1)

            export_box = QHBoxLayout()
            self.rb_export = QRadioButton("导出至新目录：")
            self.rb_export.setAccessibleName("导出至新目录单选按钮")
            export_box.addWidget(self.rb_export)

            self.edt_out_dir = QLineEdit()
            self.edt_out_dir.setEnabled(False)
            self.edt_out_dir.setPlaceholderText("选择结果文档导出的目标文件夹...")
            self.edt_out_dir.setAccessibleName("导出保存目录路径输入框")
            export_box.addWidget(self.edt_out_dir, 1)

            self.btn_browse_out = QPushButton("浏览…")
            self.btn_browse_out.setEnabled(False)
            self.btn_browse_out.setAccessibleName("浏览导出保存目录按钮")
            self.btn_browse_out.clicked.connect(self._browse_out_dir)
            export_box.addWidget(self.btn_browse_out)

            opts_grid.addLayout(export_box, 1, 0, 1, 2)

            # Mode
            mode_box = QHBoxLayout()
            mode_box.setSpacing(16)

            self.rb_fast = QRadioButton("快速模式（推荐 — 基于 python-docx）")
            self.rb_fast.setChecked(True)
            self.rb_fast.setAccessibleName("快速模式单选按钮")
            mode_box.addWidget(self.rb_fast)

            com_hint = "" if CAPABILITIES.has_word_com else "（需 Windows 与 Microsoft Word）"
            self.rb_full = QRadioButton(f"完整模式{com_hint}（保留超链接/形状/页眉页脚）")
            self.rb_full.setEnabled(CAPABILITIES.has_word_com)
            self.rb_full.setAccessibleName("完整模式单选按钮")
            mode_box.addWidget(self.rb_full)

            opts_grid.addLayout(mode_box, 2, 0, 1, 2)

            self.card_step3.addLayout(opts_grid)
            layout.addWidget(self.card_step3, 0)

            self.scroll_area.setWidget(self.content_widget)
            root_layout.addWidget(self.scroll_area, 1)

            # ---------------- Fixed Bottom Action Bar ----------------
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

            self.btn_preview = QPushButton("预览更改")
            self.btn_preview.setAccessibleName("预览更改按钮")
            self.btn_preview.clicked.connect(self.preview_changes)
            action_bar.addWidget(self.btn_preview)

            action_bar.addStretch()

            self.btn_replace = QPushButton("开始批量替换")
            self.btn_replace.setProperty("isPrimary", True)
            self.btn_replace.setAccessibleName("开始批量替换主按钮")
            self.btn_replace.clicked.connect(self.start_replace)
            action_bar.addWidget(self.btn_replace)

            root_layout.addWidget(action_container)

        # ---------------- Event & State Handlers ----------------

        def _on_dest_mode_toggled(self, checked: bool):
            self.cb_backup.setEnabled(checked)
            self.edt_out_dir.setEnabled(not checked)
            self.btn_browse_out.setEnabled(not checked)

        def _browse_docs(self):
            filter_str = (
                "Word 文档 (*.docx *.docm);;所有文件 (*.*)"
                if CAPABILITIES.is_macos
                else "Word 文档 (*.docx *.docm *.doc);;所有文件 (*.*)"
            )
            files, _ = QFileDialog.getOpenFileNames(self, "选择待处理的 Word 文档", "", filter_str)
            if files:
                valid_files = []
                for f in files:
                    if CAPABILITIES.is_macos and f.lower().endswith(".doc"):
                        QMessageBox.warning(
                            self, "格式不支持", f"macOS 首版暂不支持旧版 .doc 格式：{os.path.basename(f)}"
                        )
                        continue
                    valid_files.append(f)
                if valid_files:
                    added = self.mapping_model.add_documents(valid_files)
                    if added > 0:
                        self.scan_variables()

        def _remove_selected_docs(self):
            selected = self.tbl_view.selectionModel().selectedRows()
            if selected:
                rows = [idx.row() for idx in selected]
                self.mapping_model.remove_indices(rows)

        def _clear_docs(self):
            if self.mapping_model.rowCount() == 0:
                return
            reply = QMessageBox.question(
                self, "确认清空", f"确定要清空已选的 {self.mapping_model.rowCount()} 个文档吗？"
            )
            if reply == QMessageBox.Yes:
                self.mapping_model.clear()

        def _update_sheet_selector(self):
            path = self.edt_excel_path.text().strip()
            if getattr(self, "_last_excel_path", None) == path:
                return
            self._last_excel_path = path

            sheet_names = get_table_sheet_names(path)
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

        def _on_excel_path_changed(self):
            self._update_sheet_selector()

        def _on_sheet_changed(self, index: int = 0):
            path = self.edt_excel_path.text().strip()
            if path and os.path.isfile(path) and self.mapping_model.rowCount() > 0:
                selected_sheet = self.combo_sheet.currentText().strip() if self.combo_sheet.count() > 1 else None
                try:
                    self.excel_data = load_table_data(path, sheet_name=selected_sheet)
                    self._apply_excel_import(match_by_header=True)
                except Exception as exc:
                    QMessageBox.critical(self, "读取表格失败", f"无法解析工作表数据：{exc}")

        def _browse_excel(self):
            filter_str = "表格数据文件 (*.xlsx *.csv *.xlsm);;所有文件 (*.*)"
            path, _ = QFileDialog.getOpenFileName(self, "选择数据源表格", "", filter_str)
            if path:
                self.edt_excel_path.setText(os.path.abspath(path))
                self._fill_from_excel()

        def _fill_from_excel(self):
            path = self.edt_excel_path.text().strip()
            if not path or not os.path.isfile(path):
                QMessageBox.warning(self, "提示", "请先选择存在的 Excel 或 CSV 表格文件。")
                return

            self._update_sheet_selector()
            selected_sheet = self.combo_sheet.currentText().strip() if self.combo_sheet.count() > 1 else None
            try:
                self.excel_data = load_table_data(path, sheet_name=selected_sheet)
                self._apply_excel_import(match_by_header=True)
            except Exception as exc:
                QMessageBox.critical(self, "读取表格失败", f"无法解析数据表格：{exc}")

        def _apply_excel_import(self, match_by_header: bool = True):
            if not self.excel_data:
                path = self.edt_excel_path.text().strip()
                if path and os.path.isfile(path):
                    try:
                        selected_sheet = self.combo_sheet.currentText().strip() if self.combo_sheet.count() > 1 else None
                        self.excel_data = load_table_data(path, sheet_name=selected_sheet)
                    except Exception:
                        pass

            if not self.excel_data:
                QMessageBox.warning(self, "提示", "请先选择并加载有效的 Excel 或 CSV 数据源表格。")
                return

            filled = self.mapping_model.import_table_data(self.excel_data, match_by_header=match_by_header)
            mode_name = "按列名智能对齐" if match_by_header else "按列顺序"
            sheet_info = f"（工作表：{self.combo_sheet.currentText()}）" if self.combo_sheet.count() > 1 else ""
            self.statusMessage.emit(f"已成功从表格导入 {filled} 行数据{sheet_info}（{mode_name}）", "success")

        def _clear_table_data(self):
            if self.mapping_model.rowCount() == 0:
                return
            reply = QMessageBox.question(self, "确认清空填报数据", "确定要清空下方表格中所有已录入的替换值吗？文档列表将保留。")
            if reply == QMessageBox.Yes:
                self.mapping_model.clear_table_data()
                self.statusMessage.emit("已清空填报数据", "normal")

        def _on_table_cell_clicked(self, index):
            if index.isValid() and 2 <= index.column() < self.mapping_model._status_col():
                self.tbl_view.edit(index)

        def _move_row_up(self):
            selected = self.tbl_view.selectionModel().selectedRows()
            if selected:
                row = selected[0].row()
                if self.mapping_model.move_row_up(row):
                    self.tbl_view.selectRow(row - 1)

        def _move_row_down(self):
            selected = self.tbl_view.selectionModel().selectedRows()
            if selected:
                row = selected[0].row()
                if self.mapping_model.move_row_down(row):
                    self.tbl_view.selectRow(row + 1)

        def _browse_out_dir(self):
            folder = QFileDialog.getExistingDirectory(self, "选择生成文档输出目录")
            if folder:
                self.edt_out_dir.setText(os.path.abspath(folder))

        def _on_docs_changed(self):
            count = self.mapping_model.rowCount()
            self.lbl_doc_count.setText(f"已选 {count} 个文档")
            self.statusMetrics.emit(f"已加载 {count} 个文档" if count > 0 else "")
            self._update_table_summary()

        def _on_table_data_changed(self):
            self._update_table_summary()

        def _update_table_summary(self):
            total_docs = self.mapping_model.rowCount()
            vars_list = self.mapping_model.get_variables()
            if total_docs == 0:
                self.lbl_table_summary.setText("请添加 Word 文档并点击“自动扫描文档变量”以发现各文档中的变量。")
                self.lbl_table_summary.setStyleSheet(f"color: {THEME.tokens.text_secondary}; font-size: 12px;")
                return

            ready_count, _ = self.mapping_model.get_ready_metrics()
            vars_str = "、".join(vars_list[:8]) if vars_list else "（未扫描到变量）"
            if len(vars_list) > 8:
                vars_str += f" 等共 {len(vars_list)} 个"

            summary = (
                f"已加载 {total_docs} 个文档，检测到变量：{vars_str}；"
                f"数据就绪：{ready_count} / {total_docs} 份。"
            )
            color = THEME.tokens.success if ready_count == total_docs and total_docs > 0 else THEME.tokens.text_secondary
            self.lbl_table_summary.setText(summary)
            self.lbl_table_summary.setStyleSheet(f"color: {color}; font-size: 12px;")

        # ---------------- Asynchronous Scanning & Execution ----------------

        def scan_variables(self):
            items = self.mapping_model.get_items()
            if not items:
                return

            file_paths = [item.file_path for item in items]
            self._scan_generation_id += 1
            current_scan_gen = self._scan_generation_id

            self._set_ui_busy(True)
            self.statusMessage.emit("正在扫描各文档变量...", "running")

            def _scan_worker_fn(paths: list[str], use_com: bool, progress_cb=None, cancel_token=None):
                res = MultiDocService.scan_documents(
                    file_paths=paths,
                    use_com=use_com,
                    progress_cb=progress_cb,
                    cancel_token=cancel_token,
                )
                return (res, current_scan_gen)

            worker = TaskWorker(
                _scan_worker_fn,
                paths=file_paths,
                use_com=self.rb_full.isChecked(),
            )
            self._active_worker = worker

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_scan_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_ui_busy(False))

            self._thread_pool.start(worker)

        def _on_scan_finished(self, result_tuple: tuple[ServiceResult, int]):
            self.statusHideProgress.emit()
            res, gen_id = result_tuple
            if gen_id != self._scan_generation_id:
                return

            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("扫描已取消", "warning")
                    return
                self.statusMessage.emit("扫描失败", "error")
                QMessageBox.critical(self, "扫描失败", res.error)
                return

            doc_vars_map, all_vars = res.data
            self.mapping_model.set_variables(all_vars, doc_vars_map)
            self.statusMessage.emit(f"扫描完成：检测到 {len(all_vars)} 项变量", "success")

            # Adjust column widths
            if self.mapping_model.columnCount() > 0:
                h_header = self.tbl_view.horizontalHeader()
                h_header.setSectionResizeMode(0, QHeaderView.Fixed)
                h_header.resizeSection(0, 48)
                h_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
                status_col = self.mapping_model._status_col()
                for c in range(2, status_col):
                    h_header.setSectionResizeMode(c, QHeaderView.Stretch)
                h_header.setSectionResizeMode(status_col, QHeaderView.Fixed)
                h_header.resizeSection(status_col, 110)

        def preview_changes(self):
            items = self.mapping_model.get_items()
            if not items:
                QMessageBox.warning(self, "提示", "请先添加至少一个 Word 文档。")
                return

            all_vars = self.mapping_model.get_variables()
            in_place = self.rb_inplace.isChecked()
            out_dir = self.edt_out_dir.text().strip() if not in_place else None

            res = MultiDocService.generate_preview(
                items=items,
                all_variables=all_vars,
                output_folder=out_dir,
                max_items=10,
            )

            if not res.success:
                QMessageBox.critical(self, "预览失败", res.error)
                return

            snippets = res.data or []
            mode_desc = "直接修改原文件（生成 .backup 备份）" if in_place else f"导出至：{out_dir}"
            lines = [
                "多文档差异化替换预览（前 10 份文档）",
                "=" * 65,
                f"处理模式：{mode_desc}",
                f"文档总数：{len(items)}",
                f"变量总项：{len(all_vars)} （{'、'.join(all_vars) if all_vars else '无'}）",
                "=" * 65,
                "",
            ]

            for snip in snippets:
                lines.append(f"【{snip['index']}】文档：{snip['filename']}  [{snip['status']}]")
                lines.append(f"    目标路径：{snip['destination']}")
                if snip["fields"]:
                    for f in snip["fields"]:
                        tag = f"{{{{{f['field']}}}}} => "
                        if f["is_filled"]:
                            tag += f"【{f['value']}】"
                        elif f["is_detected"]:
                            tag += "⚠️ （待填入）"
                        else:
                            tag += "— （文档不含此变量）"
                        lines.append(f"      {tag}")
                else:
                    lines.append("      （未扫描或无变量）")
                lines.append("")

            dialog = ResultDialog(
                title="多文档替换预览",
                summary_text=f"已生成前 {len(snippets)} 份文档的替换预览信息。",
                details_text="\n".join(lines),
                is_success=True,
                output_folder=out_dir if not in_place else None,
                parent=self,
            )
            dialog.exec()

        def start_replace(self):
            items = self.mapping_model.get_items()
            in_place = self.rb_inplace.isChecked()
            out_dir = self.edt_out_dir.text().strip() if not in_place else None

            err = MultiDocService.validate_inputs(items, in_place=in_place, output_folder=out_dir)
            if err:
                QMessageBox.warning(self, "提示", err)
                return

            ready_count, total_docs = self.mapping_model.get_ready_metrics()
            if ready_count < total_docs:
                reply = QMessageBox.question(
                    self,
                    "部分文档数据未就绪",
                    f"当前有 {total_docs - ready_count} 份文档中存在尚未填入替换值的变量。\n\n"
                    "未填写的变量将保持原样或替换为空，是否继续执行批量替换？",
                )
                if reply != QMessageBox.Yes:
                    return

            mode_name = "完整模式" if self.rb_full.isChecked() else "快速模式"
            dest_hint = "直接修改原文件（自动创建 .backup 备份）" if in_place else f"导出至新目录：{out_dir}"

            reply = QMessageBox.question(
                self,
                "确认开始批量替换",
                f"即将使用【{mode_name}】处理 {len(items)} 个文档。\n\n"
                f"保存方式：{dest_hint}\n\n"
                "确定要开始吗？",
            )
            if reply != QMessageBox.Yes:
                return

            self._set_ui_busy(True)
            self.statusMessage.emit("正在批量替换多文档...", "running")

            worker = TaskWorker(
                MultiDocService.execute_batch_replace,
                items=items,
                in_place=in_place,
                output_folder=out_dir,
                create_backup=self.cb_backup.isChecked(),
                use_com=self.rb_full.isChecked(),
            )
            self._active_worker = worker

            worker.signals.progress.connect(self.statusProgress.emit)
            worker.signals.result.connect(self._on_replace_finished)
            worker.signals.error.connect(self._on_task_error)
            worker.signals.finished.connect(lambda: self._set_ui_busy(False))

            self._thread_pool.start(worker)

        def _on_replace_finished(self, res: ServiceResult[MultiDocBatchResult]):
            self.statusHideProgress.emit()
            if not res.success:
                if res.state == TaskState.CANCELLED:
                    self.statusMessage.emit("替换任务已取消", "warning")
                    return
                self.statusMessage.emit("替换失败", "error")
                QMessageBox.critical(self, "替换失败", res.error)
                return

            data = res.data
            in_place = self.rb_inplace.isChecked()
            out_dir = self.edt_out_dir.text().strip() if not in_place else None

            lines = [
                "多文档差异化替换执行报告",
                "=" * 60,
                f"文档总数：{data.total_docs}",
                f"成功处理：{data.success_docs} 份",
                f"失败处理：{data.failed_docs} 份",
                f"备份文件：{len(data.backup_files)} 个",
                "=" * 60,
                "",
            ]

            for d in data.details:
                st = "✓ 成功" if d["success"] else f"✕ 失败 ({d['error']})"
                lines.append(f"📄 {d['filename']}：{st}（替换 {d['replacements']} 处）")
                if d.get("destination"):
                    lines.append(f"   保存路径：{d['destination']}")
                if d.get("backup_path"):
                    lines.append(f"   备份文件：{d['backup_path']}")
                lines.append("")

            if data.errors:
                lines.append("❌ 错误列表：")
                for e in data.errors:
                    lines.append(f"• {e}")
                self.statusMessage.emit("多文档替换完成（存在部分错误）", "warning")
            else:
                self.statusMessage.emit("多文档替换完成", "success")

            dialog = ResultDialog(
                title="批量替换完成",
                summary_text=f"多文档替换已完成！成功 {data.success_docs} 份，失败 {data.failed_docs} 份。",
                details_text="\n".join(lines),
                is_success=not bool(data.errors),
                output_folder=out_dir,
                parent=self,
            )
            dialog.exec()

        def _on_task_error(self, err_msg: str):
            self.statusHideProgress.emit()
            self.statusMessage.emit("任务异常", "error")
            QMessageBox.critical(self, "执行异常", err_msg)

        def cancel_active_task(self):
            if self._active_worker:
                self._active_worker.cancel()
                self.statusMessage.emit("正在请求取消任务...", "warning")

        def _set_ui_busy(self, is_busy: bool):
            self.btn_add_docs.setEnabled(not is_busy)
            self.btn_remove_docs.setEnabled(not is_busy)
            self.btn_clear_docs.setEnabled(not is_busy)
            self.btn_scan_vars.setEnabled(not is_busy)
            self.btn_fill_excel.setEnabled(not is_busy)
            self.btn_preview.setEnabled(not is_busy)
            self.btn_replace.setEnabled(not is_busy)

else:

    class MultiDocPage:  # type: ignore
        pass
