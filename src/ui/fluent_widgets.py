"""Reusable Windows 11 Fluent Design widgets built on Tkinter / ttk."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Optional

from ui.theme import THEME, ThemeColors


class FluentCard(tk.Frame):
    """A card container with smooth background, 1px border, and comfortable padding."""

    def __init__(
        self,
        parent,
        title: str = "",
        subtitle: str = "",
        padx: int = 14,
        pady: int = 12,
        **kwargs,
    ):
        super().__init__(
            parent,
            bg=THEME.colors.card_bg,
            highlightbackground=THEME.colors.border,
            highlightthickness=1,
            bd=0,
            padx=padx,
            pady=pady,
            **kwargs,
        )
        self.title_label = None
        self.subtitle_label = None

        if title or subtitle:
            header_frame = tk.Frame(self, bg=THEME.colors.card_bg)
            header_frame.pack(fill=tk.X, pady=(0, 8))

            if title:
                self.title_label = tk.Label(
                    header_frame,
                    text=title,
                    font=THEME.font(11, "bold"),
                    fg=THEME.colors.text_primary,
                    bg=THEME.colors.card_bg,
                )
                self.title_label.pack(side=tk.LEFT)

            if subtitle:
                self.subtitle_label = tk.Label(
                    header_frame,
                    text=f"  {subtitle}",
                    font=THEME.font(9, "normal"),
                    fg=THEME.colors.text_secondary,
                    bg=THEME.colors.card_bg,
                )
                self.subtitle_label.pack(side=tk.LEFT, pady=(2, 0))

    def update_theme(self, colors: ThemeColors):
        self.config(bg=colors.card_bg, highlightbackground=colors.border)
        for child in self.winfo_children():
            self._apply_child_theme(child, colors)

    def _apply_child_theme(self, widget, colors: ThemeColors):
        if isinstance(widget, tk.Frame):
            widget.config(bg=colors.card_bg)
            for sub in widget.winfo_children():
                self._apply_child_theme(sub, colors)
        elif isinstance(widget, tk.Label):
            widget.config(bg=colors.card_bg)


class PrimaryButton(tk.Button):
    """The single emphasized primary call-to-action button."""

    def __init__(
        self,
        parent,
        text: str,
        command: Optional[Callable] = None,
        height: int = 1,
        padx: int = 18,
        pady: int = 6,
        **kwargs,
    ):
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=THEME.colors.accent,
            fg=THEME.colors.accent_text,
            activebackground=THEME.colors.accent_pressed,
            activeforeground=THEME.colors.accent_text,
            font=THEME.font(10, "bold"),
            relief=tk.FLAT,
            bd=0,
            cursor="hand2",
            padx=padx,
            pady=pady,
            **kwargs,
        )
        self._normal_bg = THEME.colors.accent
        self._hover_bg = THEME.colors.accent_hover
        self._pressed_bg = THEME.colors.accent_pressed
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event):
        if str(self["state"]) != tk.DISABLED:
            self.config(bg=self._hover_bg)

    def _on_leave(self, _event):
        if str(self["state"]) != tk.DISABLED:
            self.config(bg=self._normal_bg)

    def update_theme(self, colors: ThemeColors):
        self._normal_bg = colors.accent
        self._hover_bg = colors.accent_hover
        self._pressed_bg = colors.accent_pressed
        self.config(
            bg=colors.accent,
            fg=colors.accent_text,
            activebackground=colors.accent_pressed,
            activeforeground=colors.accent_text,
        )


class SecondaryButton(tk.Button):
    """A clean neutral secondary button with 1px border."""

    def __init__(
        self,
        parent,
        text: str,
        command: Optional[Callable] = None,
        padx: int = 12,
        pady: int = 5,
        **kwargs,
    ):
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=THEME.colors.button_bg,
            fg=THEME.colors.text_primary,
            activebackground=THEME.colors.button_pressed,
            activeforeground=THEME.colors.text_primary,
            highlightbackground=THEME.colors.button_border,
            highlightthickness=1,
            font=THEME.font(9, "normal"),
            relief=tk.FLAT,
            bd=0,
            cursor="hand2",
            padx=padx,
            pady=pady,
            **kwargs,
        )
        self._normal_bg = THEME.colors.button_bg
        self._hover_bg = THEME.colors.button_hover
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, _event):
        if str(self["state"]) != tk.DISABLED:
            self.config(bg=self._hover_bg)

    def _on_leave(self, _event):
        if str(self["state"]) != tk.DISABLED:
            self.config(bg=self._normal_bg)

    def update_theme(self, colors: ThemeColors):
        self._normal_bg = colors.button_bg
        self._hover_bg = colors.button_hover
        self.config(
            bg=colors.button_bg,
            fg=colors.text_primary,
            activebackground=colors.button_pressed,
            activeforeground=colors.text_primary,
            highlightbackground=colors.button_border,
        )


class GhostButton(tk.Button):
    """Subtle transparent button for inline tools (e.g. Paste, Search tool)."""

    def __init__(
        self,
        parent,
        text: str,
        command: Optional[Callable] = None,
        padx: int = 6,
        pady: int = 3,
        **kwargs,
    ):
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=THEME.colors.card_bg,
            fg=THEME.colors.text_secondary,
            activebackground=THEME.colors.button_hover,
            activeforeground=THEME.colors.text_primary,
            font=THEME.font(9, "normal"),
            relief=tk.FLAT,
            bd=0,
            cursor="hand2",
            padx=padx,
            pady=pady,
            **kwargs,
        )
        self.bind("<Enter>", lambda _e: self.config(bg=THEME.colors.button_hover, fg=THEME.colors.text_primary))
        self.bind("<Leave>", lambda _e: self.config(bg=THEME.colors.card_bg, fg=THEME.colors.text_secondary))


class SegmentedNav(tk.Frame):
    """Windows 11 top level segmented tab navigation."""

    def __init__(self, parent, tabs: list[tuple[str, str]], on_tab_change: Callable[[str], None], **kwargs):
        super().__init__(parent, bg=THEME.colors.window_bg, **kwargs)
        self.tabs = tabs  # list of (tab_id, display_name)
        self.on_tab_change = on_tab_change
        self.active_tab = tabs[0][0] if tabs else ""
        self._buttons: dict[str, tk.Label] = {}
        self._indicators: dict[str, tk.Frame] = {}

        for tab_id, label_text in tabs:
            tab_container = tk.Frame(self, bg=THEME.colors.window_bg, cursor="hand2")
            tab_container.pack(side=tk.LEFT, padx=(0, 8))

            lbl = tk.Label(
                tab_container,
                text=label_text,
                font=THEME.font(10, "bold" if tab_id == self.active_tab else "normal"),
                fg=THEME.colors.accent if tab_id == self.active_tab else THEME.colors.text_secondary,
                bg=THEME.colors.window_bg,
                padx=12,
                pady=6,
                cursor="hand2",
            )
            lbl.pack(fill=tk.X)

            indicator = tk.Frame(
                tab_container,
                height=3,
                bg=THEME.colors.accent if tab_id == self.active_tab else THEME.colors.window_bg,
            )
            indicator.pack(fill=tk.X)

            self._buttons[tab_id] = lbl
            self._indicators[tab_id] = indicator

            # Bind clicks
            lbl.bind("<Button-1>", lambda _e, tid=tab_id: self.select_tab(tid))
            tab_container.bind("<Button-1>", lambda _e, tid=tab_id: self.select_tab(tid))

            # Bind hovers
            lbl.bind("<Enter>", lambda _e, tid=tab_id: self._on_hover(tid, True))
            lbl.bind("<Leave>", lambda _e, tid=tab_id: self._on_hover(tid, False))

    def _on_hover(self, tab_id: str, is_hover: bool):
        if tab_id != self.active_tab:
            lbl = self._buttons[tab_id]
            lbl.config(fg=THEME.colors.text_primary if is_hover else THEME.colors.text_secondary)

    def select_tab(self, tab_id: str):
        if tab_id == self.active_tab:
            return
        self.active_tab = tab_id
        for tid, lbl in self._buttons.items():
            is_active = tid == tab_id
            lbl.config(
                font=THEME.font(10, "bold" if is_active else "normal"),
                fg=THEME.colors.accent if is_active else THEME.colors.text_secondary,
            )
            self._indicators[tid].config(
                bg=THEME.colors.accent if is_active else THEME.colors.window_bg
            )
        self.on_tab_change(tab_id)

    def update_theme(self, colors: ThemeColors):
        self.config(bg=colors.window_bg)
        for tid, lbl in self._buttons.items():
            is_active = tid == self.active_tab
            lbl.master.config(bg=colors.window_bg)
            lbl.config(
                bg=colors.window_bg,
                fg=colors.accent if is_active else colors.text_secondary,
            )
            self._indicators[tid].config(
                bg=colors.accent if is_active else colors.window_bg
            )


class FluentStatusBar(tk.Frame):
    """Windows 11 bottom status bar with status indicator, metrics, and progress."""

    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            bg=THEME.colors.card_bg,
            highlightbackground=THEME.colors.border,
            highlightthickness=1,
            padx=14,
            pady=7,
            **kwargs,
        )

        # Left: Status indicator dot + text
        self.status_dot = tk.Label(
            self, text="●", font=THEME.font(8), fg=THEME.colors.success, bg=THEME.colors.card_bg
        )
        self.status_dot.pack(side=tk.LEFT, padx=(0, 6))

        self.status_text = tk.Label(
            self,
            text="就绪",
            font=THEME.font(9, "normal"),
            fg=THEME.colors.text_primary,
            bg=THEME.colors.card_bg,
        )
        self.status_text.pack(side=tk.LEFT)

        # Middle: Metrics tag
        self.metrics_label = tk.Label(
            self,
            text="",
            font=THEME.font(9, "normal"),
            fg=THEME.colors.text_secondary,
            bg=THEME.colors.card_bg,
        )
        self.metrics_label.pack(side=tk.LEFT, padx=(20, 0))

        # Right: Progress Bar and percentage
        self.progress_frame = tk.Frame(self, bg=THEME.colors.card_bg)
        self.progress_label = tk.Label(
            self.progress_frame,
            text="",
            font=THEME.font(8),
            fg=THEME.colors.text_secondary,
            bg=THEME.colors.card_bg,
        )
        self.progress_label.pack(side=tk.LEFT, padx=(0, 8))

        self.progress_bar = ttk.Progressbar(self.progress_frame, orient="horizontal", length=160, mode="determinate")
        self.progress_bar.pack(side=tk.RIGHT)

    def set_status(self, text: str, state: str = "normal"):
        """State can be 'normal', 'success', 'running', 'warning', 'error'."""
        self.status_text.config(text=text)
        if state == "running":
            self.status_dot.config(fg=THEME.colors.accent, text="●")
        elif state == "success":
            self.status_dot.config(fg=THEME.colors.success, text="●")
        elif state == "warning":
            self.status_dot.config(fg=THEME.colors.warning, text="▲")
        elif state == "error":
            self.status_dot.config(fg=THEME.colors.error, text="✕")
        else:
            self.status_dot.config(fg=THEME.colors.success, text="●")

    def set_metrics(self, text: str):
        self.metrics_label.config(text=text)

    def show_progress(self, current: int, total: int, filename: str = ""):
        self.progress_frame.pack(side=tk.RIGHT)
        if total > 0:
            pct = int((current / total) * 100)
            self.progress_bar["value"] = pct
            self.progress_label.config(text=f"{current}/{total} ({pct}%) {filename[:25]}")
        else:
            self.progress_bar["value"] = 0
            self.progress_label.config(text="")

    def hide_progress(self):
        self.progress_frame.pack_forget()

    def update_theme(self, colors: ThemeColors):
        self.config(bg=colors.card_bg, highlightbackground=colors.border)
        self.status_dot.config(bg=colors.card_bg)
        self.status_text.config(bg=colors.card_bg, fg=colors.text_primary)
        self.metrics_label.config(bg=colors.card_bg, fg=colors.text_secondary)
        self.progress_frame.config(bg=colors.card_bg)
        self.progress_label.config(bg=colors.card_bg, fg=colors.text_secondary)


class ScrollableCardContainer(tk.Frame):
    """Smooth scrollable container for cards."""

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=THEME.colors.window_bg, **kwargs)

        self.canvas = tk.Canvas(self, bg=THEME.colors.window_bg, bd=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=THEME.colors.window_bg)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

        self.canvas_frame = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.bind_all_mousewheel(self)

    def _on_canvas_configure(self, event):
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def _on_mousewheel(self, event):
        if event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")

    def bind_all_mousewheel(self, widget):
        widget.bind("<MouseWheel>", self._on_mousewheel)
        widget.bind("<Button-4>", self._on_mousewheel)
        widget.bind("<Button-5>", self._on_mousewheel)


class ResultViewerDialog:
    """Modern structured results viewer window."""

    def __init__(self, parent, title: str, summary_text: str, details_text: str, is_success: bool = True):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("680x560")
        self.window.minsize(540, 420)
        self.window.configure(bg=THEME.colors.window_bg)
        self.window.transient(parent)
        self.window.grab_set()

        main_frame = tk.Frame(self.window, bg=THEME.colors.window_bg, padx=16, pady=16)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Summary card
        summary_card = FluentCard(main_frame, title="执行结果概览")
        summary_card.pack(fill=tk.X, pady=(0, 12))

        header_row = tk.Frame(summary_card, bg=THEME.colors.card_bg)
        header_row.pack(fill=tk.X, pady=(4, 0))

        status_icon = "✓" if is_success else "✕"
        status_color = THEME.colors.success if is_success else THEME.colors.error
        icon_lbl = tk.Label(
            header_row,
            text=status_icon,
            font=THEME.font(16, "bold"),
            fg=status_color,
            bg=THEME.colors.card_bg,
        )
        icon_lbl.pack(side=tk.LEFT, padx=(0, 8))

        summary_lbl = tk.Label(
            header_row,
            text=summary_text,
            font=THEME.font(10, "normal"),
            fg=THEME.colors.text_primary,
            bg=THEME.colors.card_bg,
            justify=tk.LEFT,
        )
        summary_lbl.pack(side=tk.LEFT, fill=tk.X)

        # Details card
        details_card = FluentCard(main_frame, title="详细报告")
        details_card.pack(fill=tk.BOTH, expand=True, pady=(0, 12))

        text_container = tk.Frame(details_card, bg=THEME.colors.card_bg)
        text_container.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(
            text_container,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg=THEME.colors.log_bg,
            fg=THEME.colors.log_fg,
            relief=tk.FLAT,
            highlightbackground=THEME.colors.border,
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        scrollbar = ttk.Scrollbar(text_container, orient="vertical", command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget.insert("1.0", details_text)
        text_widget.config(state=tk.DISABLED)

        # Buttons row
        btn_row = tk.Frame(main_frame, bg=THEME.colors.window_bg)
        btn_row.pack(fill=tk.X)

        SecondaryButton(btn_row, text="复制详情", command=lambda: self._copy_to_clipboard(details_text)).pack(
            side=tk.LEFT
        )
        PrimaryButton(btn_row, text="确定", command=self.window.destroy).pack(side=tk.RIGHT)

    def _copy_to_clipboard(self, text: str):
        self.window.clipboard_clear()
        self.window.clipboard_append(text)
        messagebox.showinfo("已复制", "详情内容已复制到剪贴板。", parent=self.window)
