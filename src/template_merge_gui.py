"""Tkinter window for the template batch-generation workflow."""

from __future__ import annotations

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from template_merge import (
    build_field_mapping,
    build_output_filename,
    extract_template_fields,
    extract_template_fields_com,
    generate_batch,
    load_excel_data,
    mapped_row_values,
)


class TemplateMergeWindow:
    def __init__(self, parent, has_win32com=False):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("Template Merge / 模板批量生成")
        self.window.geometry("900x720")
        self.window.minsize(760, 620)
        self.has_win32com = has_win32com
        self.data = None
        self.fields = []
        self.mapping = {}
        self._mapping_editor = None

        self.template_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.filename_var = tk.StringVar(value="{{甲方名称}}-{{乙方名称}}-合同.docx")
        self.use_com_var = tk.BooleanVar(value=has_win32com)
        self.replace_empty_var = tk.BooleanVar(value=True)
        self._build()

    def _build(self):
        frame = tk.Frame(self.window, padx=14, pady=14)
        frame.pack(fill=tk.BOTH, expand=True)

        self._path_row(frame, 0, "Word 模板：", self.template_var, self._browse_template)
        self._path_row(frame, 1, "Excel 数据：", self.excel_var, self._browse_excel)
        self._path_row(frame, 2, "输出文件夹：", self.output_var, self._browse_output)

        tk.Label(frame, text="文件名规则：").grid(row=3, column=0, sticky="w", pady=5)
        tk.Entry(frame, textvariable=self.filename_var).grid(row=3, column=1, sticky="ew", padx=6)
        tk.Button(frame, text="扫描并匹配", command=self.scan_and_match, bg="#2196f3", fg="white").grid(
            row=3, column=2, sticky="ew"
        )

        options = tk.Frame(frame)
        options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 8))
        com = tk.Checkbutton(
            options, text="高级 Word COM（文本框/脚注/尾注/超链接）",
            variable=self.use_com_var, state=tk.NORMAL if self.has_win32com else tk.DISABLED,
        )
        com.pack(side=tk.LEFT)
        tk.Checkbutton(options, text="空字段替换为空", variable=self.replace_empty_var).pack(side=tk.LEFT, padx=16)

        mapping_frame = tk.LabelFrame(frame, text="字段映射（双击 Excel 列可修改）", padx=6, pady=6)
        mapping_frame.grid(row=5, column=0, columnspan=3, sticky="nsew")
        self.tree = ttk.Treeview(mapping_frame, columns=("variable", "column", "status"), show="headings", height=10)
        self.tree.heading("variable", text="Word 变量")
        self.tree.heading("column", text="Excel 列")
        self.tree.heading("status", text="状态")
        self.tree.column("variable", width=300)
        self.tree.column("column", width=260)
        self.tree.column("status", width=100, anchor="center")
        self.tree.tag_configure("missing", foreground="#d32f2f")
        scrollbar = ttk.Scrollbar(mapping_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", self._edit_mapping)

        actions = tk.Frame(frame)
        actions.grid(row=6, column=0, columnspan=3, sticky="ew", pady=8)
        tk.Button(actions, text="预览前 5 行", command=self.preview, bg="#e3f2fd").pack(side=tk.LEFT)
        self.generate_button = tk.Button(
            actions, text="批量生成", command=self.generate, bg="#4caf50", fg="white", font=("Arial", 10, "bold")
        )
        self.generate_button.pack(side=tk.LEFT, padx=8)
        tk.Button(actions, text="关闭", command=self.window.destroy).pack(side=tk.RIGHT)

        self.progress_label = tk.Label(frame, text="请选择模板和 Excel。", anchor="w", fg="#555555")
        self.progress_label.grid(row=7, column=0, columnspan=3, sticky="ew")
        self.log = scrolledtext.ScrolledText(frame, height=9, state=tk.DISABLED, font=("Consolas", 9))
        self.log.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(5, 0))

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(5, weight=2)
        frame.rowconfigure(8, weight=1)

    def _path_row(self, parent, row, label, variable, command):
        tk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=5)
        tk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=6)
        tk.Button(parent, text="浏览…", command=command).grid(row=row, column=2, sticky="ew")

    def _browse_template(self):
        value = filedialog.askopenfilename(filetypes=[("Word", "*.docx *.docm *.doc"), ("All files", "*.*")])
        if value:
            self.template_var.set(value)
            if not self.output_var.get():
                self.output_var.set(str(Path(value).parent / "Generated"))

    def _browse_excel(self):
        value = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if value:
            self.excel_var.set(value)

    def _browse_output(self):
        value = filedialog.askdirectory()
        if value:
            self.output_var.set(value)

    def _validate_sources(self):
        template = self.template_var.get().strip()
        excel = self.excel_var.get().strip()
        if not os.path.isfile(template):
            raise ValueError("请选择存在的 Word 模板。")
        if not os.path.isfile(excel):
            raise ValueError("请选择存在的 .xlsx 数据文件。")
        return template, excel

    def scan_and_match(self):
        try:
            template, excel = self._validate_sources()
            if Path(template).suffix.lower() == ".doc":
                if not self.use_com_var.get():
                    raise ValueError(".doc 模板需要勾选高级 Word COM 模式。")
                self.fields = extract_template_fields_com(template)
            else:
                self.fields = extract_template_fields(template)
            if not self.fields:
                raise ValueError("模板中没有检测到 {{字段名}} 变量。")
            self.data = load_excel_data(excel)
            self.mapping = build_field_mapping(self.fields, self.data.headers)
            self._refresh_mapping_tree()
            missing = sum(not value for value in self.mapping.values())
            self.progress_label.config(
                text=f"检测到 {len(self.fields)} 个变量、{len(self.data.rows)} 行数据；缺失映射 {missing} 个。",
                fg="#d32f2f" if missing else "green",
            )
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc), parent=self.window)

    def _refresh_mapping_tree(self):
        self.tree.delete(*self.tree.get_children())
        for field in self.fields:
            header = self.mapping.get(field, "")
            status = "✓" if header else "Missing"
            self.tree.insert("", "end", iid=field, values=(f"{{{{{field}}}}}", header, status), tags=(() if header else ("missing",)))

    def _edit_mapping(self, event):
        if self.data is None:
            return
        row_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        if not row_id or column != "#2":
            return
        bbox = self.tree.bbox(row_id, "column")
        if not bbox:
            return
        if self._mapping_editor:
            self._mapping_editor.destroy()
        editor = ttk.Combobox(self.tree, values=[""] + self.data.headers, state="readonly")
        editor.set(self.mapping.get(row_id, ""))
        editor.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        editor.focus_set()
        self._mapping_editor = editor

        def save(_event=None):
            self.mapping[row_id] = editor.get()
            editor.destroy()
            self._mapping_editor = None
            self._refresh_mapping_tree()

        editor.bind("<<ComboboxSelected>>", save)
        editor.bind("<FocusOut>", save)

    def _ensure_scanned(self):
        if self.data is None:
            self.scan_and_match()
        if self.data is None:
            raise ValueError("请先完成模板和 Excel 扫描。")

    def preview(self):
        try:
            self._ensure_scanned()
            output = self.output_var.get().strip() or str(Path.cwd())
            extension = Path(self.template_var.get()).suffix
            reserved = set()
            lines = []
            for excel_row, row in list(zip(self.data.excel_rows, self.data.rows))[:5]:
                filename = build_output_filename(
                    self.filename_var.get(), mapped_row_values(row, self.mapping), output, extension, reserved
                )
                lines.extend([f"Excel Row {excel_row}", f"Output: {filename}"])
                for field in self.fields:
                    header = self.mapping.get(field, "")
                    value = row.get(header, "") if header else "[未映射]"
                    lines.append(f"{{{{{field}}}}} -> {value}")
                lines.append("")
            self._set_log("\n".join(lines))
        except Exception as exc:
            messagebox.showerror("预览失败", str(exc), parent=self.window)

    def generate(self):
        try:
            self._ensure_scanned()
            template, _excel = self._validate_sources()
            output = self.output_var.get().strip()
            if not output:
                raise ValueError("请选择输出文件夹。")
            if not self.filename_var.get().strip():
                raise ValueError("请输入文件名规则。")
            missing = [f"{{{{{field}}}}}" for field, header in self.mapping.items() if not header]
            if missing and not messagebox.askyesno(
                "存在未映射变量", "以下字段没有数据来源：\n\n" + "\n".join(missing) + "\n\n是否继续？", parent=self.window
            ):
                return

            self.generate_button.config(state=tk.DISABLED)
            self._set_log("")

            def progress(current, total, result):
                self.progress_label.config(text=f"Generating {current} / {total}: {result.filename}", fg="#555555")
                status = "OK" if result.success else f"FAILED: {result.error}"
                self._append_log(f"Excel Row {result.excel_row} | {result.filename} | {status}\n")
                self.window.update()

            results = generate_batch(
                template, self.data, self.mapping, output, self.filename_var.get(),
                use_com=self.use_com_var.get(), replace_empty=self.replace_empty_var.get(), progress=progress,
            )
            success = sum(result.success for result in results)
            failed = len(results) - success
            summary = f"Completed — Success: {success}, Failed: {failed}, Output: {output}"
            self.progress_label.config(text=summary, fg="green" if not failed else "#d32f2f")
            messagebox.showinfo("模板生成完成", summary, parent=self.window)
        except Exception as exc:
            messagebox.showerror("生成失败", str(exc), parent=self.window)
        finally:
            self.generate_button.config(state=tk.NORMAL)

    def _set_log(self, text):
        self.log.config(state=tk.NORMAL)
        self.log.delete("1.0", tk.END)
        self.log.insert("1.0", text)
        self.log.config(state=tk.DISABLED)

    def _append_log(self, text):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, text)
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)
