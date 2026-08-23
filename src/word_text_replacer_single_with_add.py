"""Bulk Text Replacement for Word — Modern Windows 11 Fluent GUI Application.

Integrates single/multi-document search & replace and multi-field Word template
batch generation (mail merge) into a unified, accessible, and elegant interface.
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from platform_capabilities import CAPABILITIES
from replacer_core import (
    BatchProcessResult,
    count_occurrences,
    get_document_text,
    perform_com_preview,
    perform_com_replace,
    perform_standard_preview,
    perform_standard_replace,
    preprocess_text_with_nbsp,
    scan_hyperlinks,
    strip_invisible_chars,
)
from template_merge import (
    ExcelData,
    MergeResult,
    build_field_mapping,
    build_output_filename,
    extract_template_fields,
    extract_template_fields_com,
    generate_batch,
    load_excel_data,
    mapped_row_values,
)
from ui.fluent_widgets import (
    FluentCard,
    FluentStatusBar,
    GhostButton,
    PrimaryButton,
    ResultViewerDialog,
    ScrollableCardContainer,
    SecondaryButton,
    SegmentedNav,
)
from ui.theme import THEME, ThemeColors


class WordTextReplacerApp:
    """Unified Windows 11 Fluent GUI application for Word Text Replacement and Template Merge."""

    def __init__(self, initial_file: str | None = None):
        self.root = tk.Tk()
        self.root.title("Word 批量处理工具 — Bulk Text Replacement for Word")
        self.root.geometry("920x800")
        self.root.minsize(760, 620)
        self.root.configure(bg=THEME.colors.window_bg)

        # State — Text Replace
        self.file_paths: list[str] = []
        self._text_cache: dict[str, str] = {}
        self._live_count_after_id = None
        self.replace_mode_var = tk.StringVar(value="fast")  # 'fast' or 'full'
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.whole_word_var = tk.BooleanVar(value=False)
        self.regex_var = tk.BooleanVar(value=False)
        self.create_backup_var = tk.BooleanVar(value=True)

        # State — Template Merge
        self.template_path_var = tk.StringVar()
        self.excel_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.filename_rule_var = tk.StringVar(value="{{甲方名称}}-{{乙方名称}}-合同.docx")
        self.template_use_com_var = tk.BooleanVar(value=False)
        self.template_replace_empty_var = tk.BooleanVar(value=True)
        self.template_fields: list[str] = []
        self.template_excel_data: ExcelData | None = None
        self.template_mapping: dict[str, str] = {}
        self._mapping_combobox = None

        if initial_file and os.path.exists(initial_file):
            self.file_paths.append(os.path.abspath(initial_file))

        self._build_ui()
        self._setup_keybindings()

        if self.file_paths:
            self._refresh_text_cache()
            self._update_file_list()

    def _build_ui(self):
        # 1. Top Bar: App Title + Segmented Navigation + Action tools (Theme, Help)
        self.top_bar = tk.Frame(self.root, bg=THEME.colors.window_bg, padx=20, pady=12)
        self.top_bar.pack(fill=tk.X)

        # Left: App Icon & Title
        title_box = tk.Frame(self.top_bar, bg=THEME.colors.window_bg)
        title_box.pack(side=tk.LEFT, padx=(0, 24))

        self.app_title = tk.Label(
            title_box,
            text="Word 批量工具",
            font=THEME.font(13, "bold"),
            fg=THEME.colors.text_primary,
            bg=THEME.colors.window_bg,
        )
        self.app_title.pack(side=tk.LEFT)

        # Middle: Segmented Navigation
        self.nav = SegmentedNav(
            self.top_bar,
            tabs=[("replace", "文本查找替换"), ("merge", "模板批量生成")],
            on_tab_change=self._on_tab_change,
        )
        self.nav.pack(side=tk.LEFT)

        # Right: Tools (Theme Switch, Help)
        tools_box = tk.Frame(self.top_bar, bg=THEME.colors.window_bg)
        tools_box.pack(side=tk.RIGHT)

        self.theme_btn = GhostButton(
            tools_box,
            text="🌙 深色" if not THEME.is_dark else "☀️ 浅色",
            command=self._toggle_theme,
        )
        self.theme_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.help_btn = GhostButton(
            tools_box,
            text="❓ 帮助",
            command=self._show_help_dialog,
        )
        self.help_btn.pack(side=tk.LEFT)

        # 2. Main Content Area: Stacked Views inside Scrollable Container
        self.content_frame = tk.Frame(self.root, bg=THEME.colors.window_bg)
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self.view_replace = self._build_replace_view(self.content_frame)
        self.view_merge = self._build_merge_view(self.content_frame)

        # Show Replace view initially
        self.view_replace.pack(fill=tk.BOTH, expand=True)

        # 3. Bottom Status Bar (Shared across all views)
        self.status_bar = FluentStatusBar(self.root)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _on_tab_change(self, tab_id: str):
        if tab_id == "replace":
            self.view_merge.pack_forget()
            self.view_replace.pack(fill=tk.BOTH, expand=True)
            self._update_replace_status()
        elif tab_id == "merge":
            self.view_replace.pack_forget()
            self.view_merge.pack(fill=tk.BOTH, expand=True)
            self._update_merge_status()

    # =========================================================================
    # TAB 1: TEXT REPLACEMENT VIEW
    # =========================================================================

    def _build_replace_view(self, parent: tk.Widget) -> tk.Widget:
        container = ScrollableCardContainer(parent)
        frame = container.scrollable_frame
        frame.configure(padx=20, pady=8)

        # Card 1: File Management
        self.card_files = FluentCard(
            frame,
            title="已选 Word 文档",
            subtitle="支持批量添加 .docx / .doc / .docm 文件",
        )
        self.card_files.pack(fill=tk.X, pady=(0, 12))

        # Listbox with scrollbar
        list_container = tk.Frame(self.card_files, bg=THEME.colors.card_bg)
        list_container.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        self.file_listbox = tk.Listbox(
            list_container,
            height=4,
            font=THEME.font(9),
            selectmode=tk.EXTENDED,
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
        )
        list_scroll = ttk.Scrollbar(list_container, orient="vertical", command=self.file_listbox.yview)
        self.file_listbox.config(yscrollcommand=list_scroll.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.bind("<Delete>", lambda _e: self._remove_selected_files())

        # File operation buttons
        file_btn_row = tk.Frame(self.card_files, bg=THEME.colors.card_bg)
        file_btn_row.pack(fill=tk.X)

        SecondaryButton(file_btn_row, text="+ 添加文件", command=self._add_files).pack(side=tk.LEFT, padx=(0, 6))
        SecondaryButton(file_btn_row, text="移除所选", command=self._remove_selected_files).pack(side=tk.LEFT, padx=(0, 6))
        SecondaryButton(file_btn_row, text="清空列表", command=self._clear_all_files).pack(side=tk.LEFT, padx=(0, 6))
        SecondaryButton(file_btn_row, text="🔗 检查超链接", command=self._check_hyperlinks).pack(side=tk.LEFT, padx=(0, 6))

        self.file_count_badge = tk.Label(
            file_btn_row,
            text="0 个文件",
            font=THEME.font(9),
            fg=THEME.colors.text_secondary,
            bg=THEME.colors.card_bg,
        )
        self.file_count_badge.pack(side=tk.RIGHT)

        # Card 2: Find and Replace Inputs
        self.card_inputs = FluentCard(frame, title="查找与替换")
        self.card_inputs.pack(fill=tk.X, pady=(0, 12))

        # Search Header
        search_hdr = tk.Frame(self.card_inputs, bg=THEME.colors.card_bg)
        search_hdr.pack(fill=tk.X, pady=(0, 4))
        tk.Label(
            search_hdr, text="查找内容：", font=THEME.font(10, "bold"), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg
        ).pack(side=tk.LEFT)

        GhostButton(search_hdr, text="📋 粘贴", command=self._paste_to_search).pack(side=tk.RIGHT)
        GhostButton(search_hdr, text="🔍 NBSP 检查", command=self._debug_nbsp).pack(side=tk.RIGHT, padx=(0, 6))

        self.search_text = tk.Text(
            self.card_inputs,
            height=3,
            wrap=tk.WORD,
            font=THEME.font(10),
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
            padx=6,
            pady=4,
            undo=True,
        )
        self.search_text.pack(fill=tk.X, pady=(0, 8))
        self.search_text.bind("<KeyRelease>", lambda _e: self._schedule_live_count())

        # Replace Header
        replace_hdr = tk.Frame(self.card_inputs, bg=THEME.colors.card_bg)
        replace_hdr.pack(fill=tk.X, pady=(0, 4))
        tk.Label(
            replace_hdr, text="替换为：", font=THEME.font(10, "bold"), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg
        ).pack(side=tk.LEFT)

        GhostButton(replace_hdr, text="📋 粘贴", command=self._paste_to_replace).pack(side=tk.RIGHT)

        self.replace_text = tk.Text(
            self.card_inputs,
            height=3,
            wrap=tk.WORD,
            font=THEME.font(10),
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
            padx=6,
            pady=4,
            undo=True,
        )
        self.replace_text.pack(fill=tk.X, pady=(0, 6))

        # Live match metrics
        self.live_match_lbl = tk.Label(
            self.card_inputs,
            text="",
            font=THEME.font(9, "italic"),
            fg=THEME.colors.text_secondary,
            bg=THEME.colors.card_bg,
        )
        self.live_match_lbl.pack(anchor="w")

        # Card 3: Processing Options
        self.card_options = FluentCard(frame, title="处理选项")
        self.card_options.pack(fill=tk.X, pady=(0, 12))

        # Mode Selection Row
        mode_row = tk.Frame(self.card_options, bg=THEME.colors.card_bg)
        mode_row.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            mode_row, text="处理方式：", font=THEME.font(10, "bold"), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg
        ).pack(side=tk.LEFT)

        self.rb_fast = tk.Radiobutton(
            mode_row,
            text="快速模式（推荐 — 速度快，适合正文与表格）",
            variable=self.replace_mode_var,
            value="fast",
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.card_bg,
            command=self._on_replace_mode_changed,
        )
        self.rb_fast.pack(side=tk.LEFT, padx=(8, 12))

        com_state = tk.NORMAL if CAPABILITIES.has_word_com else tk.DISABLED
        com_hint = "" if CAPABILITIES.has_word_com else "（需 Windows 与 Microsoft Word）"
        self.rb_full = tk.Radiobutton(
            mode_row,
            text=f"完整模式{com_hint}（保留超链接/形状/页眉页脚）",
            variable=self.replace_mode_var,
            value="full",
            state=com_state,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary if CAPABILITIES.has_word_com else THEME.colors.text_placeholder,
            activebackground=THEME.colors.card_bg,
            command=self._on_replace_mode_changed,
        )
        self.rb_full.pack(side=tk.LEFT)

        # Checkboxes row 1
        opts_row1 = tk.Frame(self.card_options, bg=THEME.colors.card_bg)
        opts_row1.pack(fill=tk.X, pady=(0, 4))

        self.cb_backup = tk.Checkbutton(
            opts_row1,
            text="创建备份文件 (.backup)",
            variable=self.create_backup_var,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.card_bg,
        )
        self.cb_backup.pack(side=tk.LEFT, padx=(0, 16))

        self.cb_case = tk.Checkbutton(
            opts_row1,
            text="区分大小写",
            variable=self.case_sensitive_var,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.card_bg,
            command=self._schedule_live_count,
        )
        self.cb_case.pack(side=tk.LEFT, padx=(0, 16))

        self.cb_word = tk.Checkbutton(
            opts_row1,
            text="全字匹配",
            variable=self.whole_word_var,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.card_bg,
            command=self._schedule_live_count,
        )
        self.cb_word.pack(side=tk.LEFT, padx=(0, 16))

        self.cb_regex = tk.Checkbutton(
            opts_row1,
            text="正则表达式（仅限快速模式）",
            variable=self.regex_var,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.card_bg,
            command=self._schedule_live_count,
        )
        self.cb_regex.pack(side=tk.LEFT)

        # Action execution bar (Sticky at card bottom)
        action_bar = tk.Frame(frame, bg=THEME.colors.window_bg)
        action_bar.pack(fill=tk.X, pady=(4, 16))

        SecondaryButton(action_bar, text="预览更改", command=self._preview_replace).pack(side=tk.LEFT)
        self.btn_start_replace = PrimaryButton(action_bar, text="开始替换", command=self._start_replace)
        self.btn_start_replace.pack(side=tk.RIGHT)

        return container

    def _on_replace_mode_changed(self):
        if self.replace_mode_var.get() == "full":
            self.regex_var.set(False)
            self.cb_regex.config(state=tk.DISABLED)
        else:
            self.cb_regex.config(state=tk.NORMAL)
        self._schedule_live_count()

    # =========================================================================
    # TAB 2: TEMPLATE MERGE VIEW (3-STEP WORKFLOW)
    # =========================================================================

    def _build_merge_view(self, parent: tk.Widget) -> tk.Widget:
        container = ScrollableCardContainer(parent)
        frame = container.scrollable_frame
        frame.configure(padx=20, pady=8)

        # Step 1 Card: Data Source Selection
        self.card_step1 = FluentCard(
            frame,
            title="步骤 1：选择数据源与规则",
            subtitle="指定 Word 模板、Excel 数据源与输出文件名",
        )
        self.card_step1.pack(fill=tk.X, pady=(0, 12))

        # Grid rows for path inputs
        s1_grid = tk.Frame(self.card_step1, bg=THEME.colors.card_bg)
        s1_grid.pack(fill=tk.X, pady=(4, 0))

        # Row 0: Word Template
        tk.Label(s1_grid, text="Word 模板：", font=THEME.font(9), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg).grid(
            row=0, column=0, sticky="w", pady=4
        )
        self.ent_template = tk.Entry(
            s1_grid,
            textvariable=self.template_path_var,
            font=THEME.font(9),
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
        )
        self.ent_template.grid(row=0, column=1, sticky="ew", padx=8, pady=4)
        SecondaryButton(s1_grid, text="浏览…", command=self._browse_template).grid(row=0, column=2, sticky="ew", pady=4)

        # Row 1: Excel Data
        tk.Label(s1_grid, text="Excel 数据：", font=THEME.font(9), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg).grid(
            row=1, column=0, sticky="w", pady=4
        )
        self.ent_excel = tk.Entry(
            s1_grid,
            textvariable=self.excel_path_var,
            font=THEME.font(9),
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
        )
        self.ent_excel.grid(row=1, column=1, sticky="ew", padx=8, pady=4)
        SecondaryButton(s1_grid, text="浏览…", command=self._browse_excel).grid(row=1, column=2, sticky="ew", pady=4)

        # Row 2: Output Folder
        tk.Label(s1_grid, text="输出文件夹：", font=THEME.font(9), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg).grid(
            row=2, column=0, sticky="w", pady=4
        )
        self.ent_output = tk.Entry(
            s1_grid,
            textvariable=self.output_dir_var,
            font=THEME.font(9),
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
        )
        self.ent_output.grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        SecondaryButton(s1_grid, text="浏览…", command=self._browse_output_dir).grid(row=2, column=2, sticky="ew", pady=4)

        # Row 3: Filename rule
        tk.Label(s1_grid, text="命名规则：", font=THEME.font(9), fg=THEME.colors.text_primary, bg=THEME.colors.card_bg).grid(
            row=3, column=0, sticky="w", pady=4
        )
        self.ent_fn_rule = tk.Entry(
            s1_grid,
            textvariable=self.filename_rule_var,
            font=THEME.font(9),
            bg=THEME.colors.entry_bg,
            fg=THEME.colors.text_primary,
            highlightbackground=THEME.colors.entry_border,
            highlightthickness=1,
            bd=0,
        )
        self.ent_fn_rule.grid(row=3, column=1, sticky="ew", padx=8, pady=4)
        SecondaryButton(s1_grid, text="🔍 扫描并匹配", command=self._scan_and_match_template).grid(
            row=3, column=2, sticky="ew", pady=4
        )

        s1_grid.columnconfigure(1, weight=1)

        # Step 2 Card: Field Mapping
        self.card_step2 = FluentCard(
            frame,
            title="步骤 2：字段映射",
            subtitle="双击 Excel 列或选择对应字段完成匹配",
        )
        self.card_step2.pack(fill=tk.X, pady=(0, 12))

        tree_container = tk.Frame(self.card_step2, bg=THEME.colors.card_bg)
        tree_container.pack(fill=tk.BOTH, expand=True, pady=(4, 6))

        self.tree_mapping = ttk.Treeview(
            tree_container,
            columns=("variable", "column", "status"),
            show="headings",
            height=6,
        )
        self.tree_mapping.heading("variable", text="Word 模板变量")
        self.tree_mapping.heading("column", text="对应 Excel 列（双击修改）")
        self.tree_mapping.heading("status", text="匹配状态")
        self.tree_mapping.column("variable", width=280)
        self.tree_mapping.column("column", width=260)
        self.tree_mapping.column("status", width=100, anchor="center")
        self.tree_mapping.tag_configure("missing", foreground=THEME.colors.error)
        self.tree_mapping.tag_configure("matched", foreground=THEME.colors.success)

        tree_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree_mapping.yview)
        self.tree_mapping.configure(yscrollcommand=tree_scroll.set)
        self.tree_mapping.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_mapping.bind("<Double-1>", self._edit_mapping_cell)

        self.mapping_summary_lbl = tk.Label(
            self.card_step2,
            text="请先选择 Word 模板与 Excel 数据并点击“扫描并匹配”。",
            font=THEME.font(9),
            fg=THEME.colors.text_secondary,
            bg=THEME.colors.card_bg,
        )
        self.mapping_summary_lbl.pack(anchor="w")

        # Step 3 Card: Options, Preview and Batch Generation
        self.card_step3 = FluentCard(
            frame,
            title="步骤 3：预览并批量生成",
            subtitle="查看前 5 行预览日志并执行批量生成",
        )
        self.card_step3.pack(fill=tk.X, pady=(0, 12))

        opts_row = tk.Frame(self.card_step3, bg=THEME.colors.card_bg)
        opts_row.pack(fill=tk.X, pady=(0, 8))

        com_state = tk.NORMAL if CAPABILITIES.has_word_com else tk.DISABLED
        tk.Checkbutton(
            opts_row,
            text="启用 Word COM 完整模式（支持 .doc、形状与特殊域）",
            variable=self.template_use_com_var,
            state=com_state,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary if CAPABILITIES.has_word_com else THEME.colors.text_placeholder,
            activebackground=THEME.colors.card_bg,
        ).pack(side=tk.LEFT, padx=(0, 16))

        tk.Checkbutton(
            opts_row,
            text="空字段替换为空文本",
            variable=self.template_replace_empty_var,
            font=THEME.font(9),
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.card_bg,
        ).pack(side=tk.LEFT)

        # Merge Action Bar
        merge_act_row = tk.Frame(self.card_step3, bg=THEME.colors.card_bg)
        merge_act_row.pack(fill=tk.X, pady=(0, 8))

        SecondaryButton(merge_act_row, text="预览前 5 行", command=self._preview_template_merge).pack(side=tk.LEFT)
        self.btn_start_merge = PrimaryButton(merge_act_row, text="开始批量生成", command=self._start_template_merge)
        self.btn_start_merge.pack(side=tk.RIGHT)

        # Log Text Box
        log_container = tk.Frame(self.card_step3, bg=THEME.colors.card_bg)
        log_container.pack(fill=tk.BOTH, expand=True)

        self.merge_log = tk.Text(
            log_container,
            height=6,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg=THEME.colors.log_bg,
            fg=THEME.colors.log_fg,
            highlightbackground=THEME.colors.border,
            highlightthickness=1,
            bd=0,
            padx=6,
            pady=6,
        )
        log_scroll = ttk.Scrollbar(log_container, orient="vertical", command=self.merge_log.yview)
        self.merge_log.configure(yscrollcommand=log_scroll.set)
        self.merge_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.merge_log.config(state=tk.DISABLED)

        return container

    # =========================================================================
    # THEME TOGGLING & STYLING
    # =========================================================================

    def _toggle_theme(self):
        new_mode = THEME.toggle()
        self.theme_btn.config(text="🌙 深色" if new_mode == "light" else "☀️ 浅色")
        colors = THEME.colors

        self.root.configure(bg=colors.window_bg)
        self.top_bar.configure(bg=colors.window_bg)
        self.app_title.configure(bg=colors.window_bg, fg=colors.text_primary)
        self.content_frame.configure(bg=colors.window_bg)
        self.nav.update_theme(colors)
        self.status_bar.update_theme(colors)

        # Update Text Replace Cards
        self.card_files.update_theme(colors)
        self.file_listbox.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.file_count_badge.configure(bg=colors.card_bg, fg=colors.text_secondary)
        self.card_inputs.update_theme(colors)
        self.search_text.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.replace_text.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.live_match_lbl.configure(bg=colors.card_bg, fg=colors.text_secondary)
        self.card_options.update_theme(colors)

        # Update Template Merge Cards
        self.card_step1.update_theme(colors)
        self.ent_template.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.ent_excel.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.ent_output.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.ent_fn_rule.configure(bg=colors.entry_bg, fg=colors.text_primary, highlightbackground=colors.entry_border)
        self.card_step2.update_theme(colors)
        self.mapping_summary_lbl.configure(bg=colors.card_bg, fg=colors.text_secondary)
        self.card_step3.update_theme(colors)
        self.merge_log.configure(bg=colors.log_bg, fg=colors.log_fg, highlightbackground=colors.border)

    # =========================================================================
    # TEXT REPLACEMENT LOGIC
    # =========================================================================

    def _get_processed_search(self) -> str:
        raw = self.search_text.get("1.0", tk.END).rstrip("\n")
        return preprocess_text_with_nbsp(raw)

    def _get_processed_replace(self) -> str:
        raw = self.replace_text.get("1.0", tk.END).rstrip("\n")
        return preprocess_text_with_nbsp(raw)

    def _paste_to_search(self):
        try:
            clip = self.root.clipboard_get()
            self.search_text.delete("1.0", tk.END)
            self.search_text.insert("1.0", clip)
            self._schedule_live_count()
        except tk.TclError:
            pass

    def _paste_to_replace(self):
        try:
            clip = self.root.clipboard_get()
            self.replace_text.delete("1.0", tk.END)
            self.replace_text.insert("1.0", clip)
        except tk.TclError:
            pass

    def _debug_nbsp(self):
        text = self.search_text.get("1.0", tk.END).rstrip("\n")
        nbsp_count = text.count("\u00a0")
        literal_count = text.count("[NBSP]") + text.count("&nbsp;")
        messagebox.showinfo(
            "NBSP 空格检测",
            f"当前查找框内容字符数：{len(text)}\n"
            f"不间断空格 (U+00A0)：{nbsp_count} 个\n"
            f"NBSP 占位符：{literal_count} 个\n\n"
            "提示：程序在查找和替换时会自动规范化处理不间断空格。",
            parent=self.root,
        )

    def _add_files(self):
        types = [("Word 文档", "*.docx *.docm *.doc"), ("所有文件", "*.*")]
        selected = filedialog.askopenfilenames(title="选择 Word 文档", filetypes=types)
        if selected:
            added = 0
            for path in selected:
                abs_path = os.path.abspath(path)
                if abs_path not in self.file_paths:
                    self.file_paths.append(abs_path)
                    added += 1
            if added > 0:
                self._refresh_text_cache()
                self._update_file_list()
                self._schedule_live_count()

    def _remove_selected_files(self):
        selected_indices = list(self.file_listbox.curselection())
        if not selected_indices:
            return
        for index in reversed(selected_indices):
            del self.file_paths[index]
        self._refresh_text_cache()
        self._update_file_list()
        self._schedule_live_count()

    def _clear_all_files(self):
        if not self.file_paths:
            return
        if messagebox.askyesno("确认清空", f"确定要清空列表中的 {len(self.file_paths)} 个文件吗？", parent=self.root):
            self.file_paths.clear()
            self._text_cache.clear()
            self._update_file_list()
            self._schedule_live_count()

    def _update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for path in self.file_paths:
            name = os.path.basename(path)
            try:
                size_kb = os.path.getsize(path) / 1024
                size_str = f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.1f} MB"
            except OSError:
                size_str = "未知大小"
            self.file_listbox.insert(tk.END, f"{name}  ({size_str}) — {path}")

        count = len(self.file_paths)
        self.file_count_badge.config(text=f"已选 {count} 个文件")
        self._update_replace_status()

    def _refresh_text_cache(self):
        self._text_cache.clear()
        for path in self.file_paths:
            if path.lower().endswith((".docx", ".docm")):
                try:
                    from docx import Document
                    doc = Document(path)
                    self._text_cache[path] = get_document_text(doc)
                except Exception:
                    self._text_cache[path] = ""

    def _schedule_live_count(self):
        if self._live_count_after_id:
            self.root.after_cancel(self._live_count_after_id)
        self._live_count_after_id = self.root.after(200, self._do_live_count)

    def _do_live_count(self):
        search_for = self._get_processed_search()
        if not search_for or not self.file_paths:
            self.live_match_lbl.config(text="")
            self.status_bar.set_metrics("")
            return

        case_sensitive = self.case_sensitive_var.get()
        use_regex = self.regex_var.get() and self.replace_mode_var.get() == "fast"
        whole_word = self.whole_word_var.get()

        total = 0
        matching_files = 0
        for path in self.file_paths:
            text = self._text_cache.get(path, "")
            if text:
                occ = count_occurrences(text, search_for, case_sensitive, use_regex, whole_word)
                if occ > 0:
                    total += occ
                    matching_files += 1

        msg = f"实时统计：在 {matching_files} 个文件中找到 {total} 处匹配"
        self.live_match_lbl.config(text=msg)
        self.status_bar.set_metrics(f"{len(self.file_paths)} 个文件 | {total} 处匹配")

    def _check_hyperlinks(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先添加至少一个 Word 文档。", parent=self.root)
            return
        results = scan_hyperlinks(self.file_paths)
        total_links = sum(r["count"] for r in results)
        files_with_links = sum(1 for r in results if r["count"] > 0)

        lines = [
            "🔗 文档超链接扫描报告",
            "=" * 50,
            f"扫描文件总数：{len(self.file_paths)}",
            f"含超链接文件：{files_with_links}",
            f"超链接总数量：{total_links}",
            "=" * 50,
            "",
        ]
        for r in results:
            lines.append(f"📄 {r['filename']}：{r['count']} 个链接")
            for url in r["urls"][:5]:
                lines.append(f"   • {url}")
            if len(r["urls"]) > 5:
                lines.append(f"   • ... 以及其他 {len(r['urls']) - 5} 个链接")
            lines.append("")

        if total_links > 0:
            lines.append("💡 建议：若需替换超链接显示文字且保留链接地址，请使用【完整模式】。")

        ResultViewerDialog(
            self.root,
            title="超链接检查结果",
            summary_text=f"共扫描 {len(self.file_paths)} 个文档，发现 {total_links} 处超链接。",
            details_text="\n".join(lines),
            is_success=True,
        )

    def _preview_replace(self):
        search_for = self._get_processed_search()
        if not search_for:
            messagebox.showwarning("提示", "请输入要查找的内容。", parent=self.root)
            return
        if not self.file_paths:
            messagebox.showwarning("提示", "请先添加 Word 文档。", parent=self.root)
            return

        mode = self.replace_mode_var.get()
        case_sensitive = self.case_sensitive_var.get()
        whole_word = self.whole_word_var.get()
        use_regex = self.regex_var.get() and mode == "fast"

        self.status_bar.set_status("正在生成预览...", "running")
        self.root.update()

        try:
            if mode == "full":
                result = perform_com_preview(
                    self.file_paths, search_for, case_sensitive, whole_word, self.status_bar.show_progress
                )
            else:
                result = perform_standard_preview(
                    self.file_paths, search_for, case_sensitive, use_regex, whole_word, self.status_bar.show_progress
                )

            self.status_bar.hide_progress()
            self.status_bar.set_status("预览完成", "success")

            lines = [
                f"查找替换预览报告（模式：{'快速模式' if mode == 'fast' else '完整模式'}）",
                "=" * 60,
                f"查找内容：{search_for}",
                f"替换为：{self._get_processed_replace()}",
                f"匹配总数：{result.total_count} 处（分布于 {result.files_with_matches} 个文件）",
                "=" * 60,
                "",
            ]

            for detail in result.details:
                lines.append(f"📁 文件：{detail.filename}")
                lines.append(f"   匹配数量：{detail.total}")
                for d in detail.details:
                    lines.append(f"   {d}")
                if detail.contexts:
                    lines.append("   📝 匹配上下文样例：")
                    for ctx in detail.contexts:
                        lines.append(f"      {ctx}")
                lines.append("")

            summary = f"在 {len(self.file_paths)} 个文件中找到 {result.total_count} 处匹配（涉及 {result.files_with_matches} 个文件）。"
            ResultViewerDialog(
                self.root,
                title="查找替换预览结果",
                summary_text=summary,
                details_text="\n".join(lines),
                is_success=True,
            )
        except Exception as exc:
            self.status_bar.hide_progress()
            self.status_bar.set_status("预览失败", "error")
            messagebox.showerror("预览失败", str(exc), parent=self.root)

    def _start_replace(self):
        search_for = self._get_processed_search()
        replace_with = self._get_processed_replace()
        if not search_for:
            messagebox.showwarning("提示", "请输入要查找的内容。", parent=self.root)
            return
        if not self.file_paths:
            messagebox.showwarning("提示", "请先添加 Word 文档。", parent=self.root)
            return

        mode = self.replace_mode_var.get()
        case_sensitive = self.case_sensitive_var.get()
        whole_word = self.whole_word_var.get()
        use_regex = self.regex_var.get() and mode == "fast"
        create_backup = self.create_backup_var.get()

        mode_name = "快速模式" if mode == "fast" else "完整模式"
        backup_hint = "（将创建 .backup 备份文件）" if create_backup else "（未勾选备份）"
        if not messagebox.askyesno(
            "确认执行替换",
            f"即将使用【{mode_name}】在 {len(self.file_paths)} 个文档中执行查找替换 {backup_hint}。\n\n"
            f"查找：{search_for[:40]}\n"
            f"替换：{replace_with[:40]}\n\n"
            "是否继续？",
            parent=self.root,
        ):
            return

        self.btn_start_replace.config(state=tk.DISABLED)
        self.status_bar.set_status("正在执行替换...", "running")
        self.root.update()

        try:
            if mode == "full":
                result = perform_com_replace(
                    self.file_paths,
                    search_for,
                    replace_with,
                    case_sensitive,
                    whole_word,
                    create_backup,
                    self.status_bar.show_progress,
                )
            else:
                result = perform_standard_replace(
                    self.file_paths,
                    search_for,
                    replace_with,
                    case_sensitive,
                    use_regex,
                    whole_word,
                    create_backup,
                    self.status_bar.show_progress,
                )

            self.status_bar.hide_progress()
            self._refresh_text_cache()
            self._schedule_live_count()

            lines = [
                f"查找替换执行报告（{mode_name}）",
                "=" * 60,
                f"成功处理文件：{result.successful_files} / {len(self.file_paths)}",
                f"替换总次数：{result.total_count}",
                f"已创建备份：{len(result.backup_files)} 个文件",
                "=" * 60,
                "",
            ]
            for detail in result.details:
                lines.append(f"📁 文件：{detail.filename}")
                for d in detail.details:
                    lines.append(f"   {d}")
                lines.append("")

            if result.errors:
                lines.append("❌ 错误列表：")
                for err in result.errors:
                    lines.append(f"• {err}")
                self.status_bar.set_status("替换完成（存在错误）", "warning")
            else:
                self.status_bar.set_status("替换完成", "success")

            summary = f"处理完成！成功替换 {result.total_count} 处，涉及 {result.successful_files} 个文件。"
            ResultViewerDialog(
                self.root,
                title="替换完成",
                summary_text=summary,
                details_text="\n".join(lines),
                is_success=not bool(result.errors),
            )
        except Exception as exc:
            self.status_bar.hide_progress()
            self.status_bar.set_status("替换失败", "error")
            messagebox.showerror("替换失败", str(exc), parent=self.root)
        finally:
            self.btn_start_replace.config(state=tk.NORMAL)

    def _update_replace_status(self):
        count = len(self.file_paths)
        self.status_bar.set_status("就绪", "normal")
        self.status_bar.set_metrics(f"已加载 {count} 个文件" if count > 0 else "")

    # =========================================================================
    # TEMPLATE MERGE LOGIC
    # =========================================================================

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="选择 Word 模板",
            filetypes=[("Word 模板", "*.docx *.docm *.doc"), ("所有文件", "*.*")],
        )
        if path:
            self.template_path_var.set(os.path.abspath(path))
            if not self.output_dir_var.get():
                self.output_dir_var.set(str(Path(path).parent / "Generated"))

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 数据表",
            filetypes=[("Excel 表格", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.excel_path_var.set(os.path.abspath(path))

    def _browse_output_dir(self):
        path = filedialog.askdirectory(title="选择生成文档输出目录")
        if path:
            self.output_dir_var.set(os.path.abspath(path))

    def _scan_and_match_template(self):
        t_path = self.template_path_var.get().strip()
        e_path = self.excel_path_var.get().strip()

        if not os.path.isfile(t_path):
            messagebox.showwarning("提示", "请选择存在的 Word 模板文件。", parent=self.root)
            return
        if not os.path.isfile(e_path):
            messagebox.showwarning("提示", "请选择存在的 Excel 数据文件。", parent=self.root)
            return

        self.status_bar.set_status("正在扫描模板与数据表...", "running")
        self.root.update()

        try:
            suffix = Path(t_path).suffix.lower()
            if suffix == ".doc":
                if not CAPABILITIES.has_word_com:
                    raise RuntimeError(".doc 格式模板扫描需要 Windows 与 Microsoft Word。")
                self.template_fields = extract_template_fields_com(t_path)
            else:
                self.template_fields = extract_template_fields(t_path)

            if not self.template_fields:
                raise ValueError("未在 Word 模板中检测到 {{字段名}} 格式的变量。")

            self.template_excel_data = load_excel_data(e_path)
            self.template_mapping = build_field_mapping(self.template_fields, self.template_excel_data.headers)
            self._refresh_mapping_tree()

            missing = sum(1 for v in self.template_mapping.values() if not v)
            summary = (
                f"扫描完成：检测到 {len(self.template_fields)} 个模板变量，"
                f"{len(self.template_excel_data.rows)} 行数据；"
                f"已自动匹配 {len(self.template_fields) - missing} 个，缺失 {missing} 个。"
            )
            self.mapping_summary_lbl.config(
                text=summary,
                fg=THEME.colors.error if missing else THEME.colors.success,
            )
            self.status_bar.set_status("扫描完成", "success")
        except Exception as exc:
            self.status_bar.set_status("扫描失败", "error")
            messagebox.showerror("扫描失败", str(exc), parent=self.root)

    def _refresh_mapping_tree(self):
        self.tree_mapping.delete(*self.tree_mapping.get_children())
        for field in self.template_fields:
            col = self.template_mapping.get(field, "")
            status = "✓ 已匹配" if col else "未匹配"
            tag = "matched" if col else "missing"
            self.tree_mapping.insert(
                "",
                "end",
                iid=field,
                values=(f"{{{{{field}}}}}", col or "（双击选择对应列）", status),
                tags=(tag,),
            )

    def _edit_mapping_cell(self, event):
        if not self.template_excel_data:
            return
        row_id = self.tree_mapping.identify_row(event.y)
        column = self.tree_mapping.identify_column(event.x)
        if not row_id or column != "#2":
            return

        bbox = self.tree_mapping.bbox(row_id, "column")
        if not bbox:
            return

        if self._mapping_combobox:
            self._mapping_combobox.destroy()

        combo = ttk.Combobox(self.tree_mapping, values=[""] + self.template_excel_data.headers, state="readonly")
        combo.set(self.template_mapping.get(row_id, ""))
        combo.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        combo.focus_set()
        self._mapping_combobox = combo

        def save(_event=None):
            self.template_mapping[row_id] = combo.get()
            combo.destroy()
            self._mapping_combobox = None
            self._refresh_mapping_tree()

        combo.bind("<<ComboboxSelected>>", save)
        combo.bind("<FocusOut>", save)

    def _preview_template_merge(self):
        if not self.template_excel_data or not self.template_fields:
            self._scan_and_match_template()
        if not self.template_excel_data or not self.template_fields:
            return

        t_path = self.template_path_var.get().strip()
        out_dir = self.output_dir_var.get().strip() or str(Path.cwd() / "Generated")
        extension = Path(t_path).suffix or ".docx"
        fn_rule = self.filename_rule_var.get().strip() or "{{甲方名称}}-合同.docx"

        reserved: set[str] = set()
        lines = [
            "📑 模板批量生成预览（前 5 行）",
            "=" * 60,
            f"Word 模板：{os.path.basename(t_path)}",
            f"Excel 数据：{os.path.basename(self.excel_path_var.get())}",
            f"输出目录：{out_dir}",
            f"命名规则：{fn_rule}",
            "=" * 60,
            "",
        ]

        for excel_row, row in list(zip(self.template_excel_data.excel_rows, self.template_excel_data.rows))[:5]:
            fn = build_output_filename(
                fn_rule, mapped_row_values(row, self.template_mapping), out_dir, extension, reserved
            )
            lines.append(f"【Excel 第 {excel_row} 行】→ 输出文件: {fn}")
            for field in self.template_fields:
                header = self.template_mapping.get(field, "")
                val = row.get(header, "") if header else "【未匹配】"
                lines.append(f"   {{{{{field}}}}} => {val}")
            lines.append("")

        self._set_merge_log("\n".join(lines))
        self.status_bar.set_status("预览就绪", "normal")

    def _start_template_merge(self):
        if not self.template_excel_data or not self.template_fields:
            self._scan_and_match_template()
        if not self.template_excel_data or not self.template_fields:
            return

        t_path = self.template_path_var.get().strip()
        out_dir = self.output_dir_var.get().strip()
        fn_rule = self.filename_rule_var.get().strip()

        if not out_dir:
            messagebox.showwarning("提示", "请指定输出文件夹路径。", parent=self.root)
            return
        if not fn_rule:
            messagebox.showwarning("提示", "请输入输出文件名规则。", parent=self.root)
            return

        missing = [f"{{{{{f}}}}}" for f, h in self.template_mapping.items() if not h]
        if missing and not messagebox.askyesno(
            "存在未匹配变量",
            "以下模板变量未在 Excel 中找到对应列：\n\n" + "\n".join(missing) + "\n\n是否继续批量生成？",
            parent=self.root,
        ):
            return

        self.btn_start_merge.config(state=tk.DISABLED)
        self.status_bar.set_status("正在批量生成文档...", "running")
        self._set_merge_log("正在批量生成文档...\n")
        self.root.update()

        def progress_cb(current, total, result: MergeResult):
            self.status_bar.show_progress(current, total, result.filename)
            status_text = "成功" if result.success else f"失败: {result.error}"
            self._append_merge_log(f"第 {result.excel_row} 行 | {result.filename} | {status_text}\n")
            self.root.update()

        try:
            results = generate_batch(
                template_path=t_path,
                data=self.template_excel_data,
                mapping=self.template_mapping,
                output_folder=out_dir,
                filename_rule=fn_rule,
                use_com=self.template_use_com_var.get(),
                replace_empty=self.template_replace_empty_var.get(),
                progress=progress_cb,
            )

            self.status_bar.hide_progress()
            success_count = sum(1 for r in results if r.success)
            fail_count = len(results) - success_count
            summary = f"批量生成完成：成功 {success_count} 份，失败 {fail_count} 份；保存至：{out_dir}"
            self.status_bar.set_status("生成完成", "success" if fail_count == 0 else "warning")
            messagebox.showinfo("生成完成", summary, parent=self.root)
        except Exception as exc:
            self.status_bar.hide_progress()
            self.status_bar.set_status("生成失败", "error")
            messagebox.showerror("批量生成失败", str(exc), parent=self.root)
        finally:
            self.btn_start_merge.config(state=tk.NORMAL)

    def _set_merge_log(self, text: str):
        self.merge_log.config(state=tk.NORMAL)
        self.merge_log.delete("1.0", tk.END)
        self.merge_log.insert("1.0", text)
        self.merge_log.config(state=tk.DISABLED)

    def _append_merge_log(self, text: str):
        self.merge_log.config(state=tk.NORMAL)
        self.merge_log.insert(tk.END, text)
        self.merge_log.see(tk.END)
        self.merge_log.config(state=tk.DISABLED)

    def _update_merge_status(self):
        self.status_bar.set_status("就绪", "normal")
        if self.template_excel_data:
            self.status_bar.set_metrics(f"已加载 {len(self.template_excel_data.rows)} 行数据")
        else:
            self.status_bar.set_metrics("")

    # =========================================================================
    # HELP & SHORTCUTS DIALOG
    # =========================================================================

    def _show_help_dialog(self):
        help_text = (
            "Word 批量处理工具 — 使用指南与快捷键\n\n"
            "【功能模式】\n"
            "• 文本查找替换：批量查找替换多个 Word 文档中的内容，支持快速/完整模式、正则与全字匹配。\n"
            "• 模板批量生成：根据 Word 模板与 Excel 数据表，按行批量生成填充后的个性化文档。\n\n"
            "【快捷键】\n"
            "• Ctrl+O / Cmd+O：添加文档 / 浏览模板\n"
            "• Ctrl+R / Cmd+R：执行替换 / 开始生成\n"
            "• Ctrl+P / Cmd+P：预览更改\n"
            "• Delete：在文件列表中移除选中的文档\n"
            "• Tab：在各输入框与按钮间顺畅切换焦点\n\n"
            "【模式差异说明】\n"
            "• 快速模式：基于 python-docx，速度快，适合正文与表格，支持正则表达式。\n"
            "• 完整模式：基于 Microsoft Word COM，完整覆盖页眉页脚、文本框与超链接保护。"
        )
        ResultViewerDialog(
            self.root,
            title="使用帮助与快捷键",
            summary_text="高效、现代的 Word 批量文本处理与模板生成工具。",
            details_text=help_text,
            is_success=True,
        )

    def _setup_keybindings(self):
        modifier = "Command" if CAPABILITIES.is_macos else "Control"
        self.root.bind(f"<{modifier}-o>", lambda _e: self._add_files())
        self.root.bind(f"<{modifier}-O>", lambda _e: self._add_files())
        self.root.bind(f"<{modifier}-r>", lambda _e: self._start_replace())
        self.root.bind(f"<{modifier}-R>", lambda _e: self._start_replace())
        self.root.bind(f"<{modifier}-p>", lambda _e: self._preview_replace())
        self.root.bind(f"<{modifier}-P>", lambda _e: self._preview_replace())
        self.root.bind("<F1>", lambda _e: self._show_help_dialog())

    def run(self):
        self.root.mainloop()


# Backwards compatibility wrapper alias
WordTextReplacerSingle = WordTextReplacerApp


def main():
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    if initial_file and not os.path.exists(initial_file):
        initial_file = None
    if initial_file and not initial_file.lower().endswith((".docx", ".doc", ".docm")):
        initial_file = None

    try:
        app = WordTextReplacerApp(initial_file)
        app.run()
    except Exception as exc:
        messagebox.showerror("启动失败", f"应用程序启动异常：{str(exc)}")


if __name__ == "__main__":
    main()
