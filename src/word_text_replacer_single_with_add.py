import sys
import os
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
from docx import Document
import re
import subprocess  # For VBScript integration

class WordTextReplacerSingle:
    def __init__(self, initial_file=None):
        self.file_paths = []
        if initial_file and os.path.exists(initial_file):
            self.file_paths.append(initial_file)
        
        self.root = tk.Tk()
        self.setup_gui()
        
    def setup_gui(self):
        self.update_title()
        self.root.geometry("700x710")
        self.root.resizable(True, True)
        
        # Main frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header frame with info icon
        header_frame = tk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Invisible spacer to push info icon to the right
        spacer = tk.Label(header_frame, text="")
        spacer.pack(side=tk.LEFT, expand=True)
        
        # Info icon with tooltip
        self.info_icon = tk.Label(header_frame, text="?", font=("Arial", 12, "bold"), 
                                 fg="#666666", cursor="hand2", 
                                 relief="solid", borderwidth=1, padx=4, pady=2)
        self.info_icon.pack(side=tk.RIGHT)
        
        # Bind hover events for tooltip
        self.info_icon.bind("<Enter>", self.show_shortcuts_tooltip)
        self.info_icon.bind("<Leave>", self.hide_shortcuts_tooltip)
        
        # File management section
        files_frame = tk.LabelFrame(main_frame, text="Selected files", 
                                   font=("Arial", 11, "bold"), padx=10, pady=10)
        files_frame.pack(fill=tk.X, pady=(0, 15))
        
        # File list with scrollbar
        list_frame = tk.Frame(files_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.file_listbox = tk.Listbox(list_frame, height=4, font=("Arial", 9), selectmode=tk.EXTENDED)
        scrollbar = tk.Scrollbar(list_frame, orient="vertical")
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Bind Delete key to remove files when listbox is focused
        self.file_listbox.bind('<Delete>', self.handle_delete_key)
        
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # File management buttons
        file_buttons_frame = tk.Frame(files_frame)
        file_buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        add_files_btn = tk.Button(file_buttons_frame, text="Add more files", 
                                 command=self.add_files,
                                 bg="#2196f3", fg="white", font=("Arial", 9))
        add_files_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        remove_file_btn = tk.Button(file_buttons_frame, text="Remove selected", 
                                   command=self.remove_selected_file,
                                   bg="#ff9800", fg="white", font=("Arial", 9))
        remove_file_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        clear_files_btn = tk.Button(file_buttons_frame, text="Clear all", 
                                   command=self.clear_all_files,
                                   bg="#f44336", fg="white", font=("Arial", 9))
        clear_files_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # NEW: Add hyperlink check button here
        hyperlink_btn = tk.Button(file_buttons_frame, text="🔗 Hyperlink check", 
                                 command=self.check_hyperlinks,
                                 bg="#9c27b0", fg="white", font=("Arial", 9))
        hyperlink_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        # File count label
        self.file_count_label = tk.Label(file_buttons_frame, text="", 
                                        font=("Arial", 9), fg="gray")
        self.file_count_label.pack(side=tk.RIGHT)
        
        # Search text section
        search_header_frame = tk.Frame(main_frame)
        search_header_frame.pack(fill=tk.X, pady=(0, 5))
        
        search_label = tk.Label(search_header_frame, text="Search for:", 
                               font=("Arial", 11, "bold"))
        search_label.pack(side=tk.LEFT)
        
        paste_search_btn = tk.Button(search_header_frame, text="📋 Paste clipboard", 
                                    command=self.paste_to_search,
                                    bg="#e8f5e8", font=("Arial", 9))
        paste_search_btn.pack(side=tk.RIGHT)

        debug_btn = tk.Button(search_header_frame, text="🔍 nbsp check", 
                             command=self.debug_nbsp_conversion,
                             bg="#ffeb3b", font=("Arial", 9))
        debug_btn.pack(side=tk.RIGHT, padx=(10, 5))
        
        self.search_text = scrolledtext.ScrolledText(main_frame, height=6, width=70,
                                                    wrap=tk.WORD, font=("Arial", 10), undo=True)
        self.search_text.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Replace text section
        replace_header_frame = tk.Frame(main_frame)
        replace_header_frame.pack(fill=tk.X, pady=(0, 5))
        
        replace_label = tk.Label(replace_header_frame, text="Replace with:", 
                                font=("Arial", 11, "bold"))
        replace_label.pack(side=tk.LEFT)
        
        paste_replace_btn = tk.Button(replace_header_frame, text="📋 Paste clipboard", 
                                     command=self.paste_to_replace,
                                     bg="#e8f5e8", font=("Arial", 9))
        paste_replace_btn.pack(side=tk.RIGHT)
        
        self.replace_text = scrolledtext.ScrolledText(main_frame, height=6, width=70,
                                                     wrap=tk.WORD, font=("Arial", 10), undo=True)
        self.replace_text.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Options frame
        options_frame = tk.Frame(main_frame)
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        # First row of options
        options_row1 = tk.Frame(options_frame)
        options_row1.pack(fill=tk.X)
        
        self.create_backup_var = tk.BooleanVar(value=False)
        backup_checkbox = tk.Checkbutton(options_row1, text="Create backup files (.backup)", 
                                        variable=self.create_backup_var, font=("Arial", 10))
        backup_checkbox.pack(side=tk.LEFT)
        
        # Second row of options
        options_row2 = tk.Frame(options_frame)
        options_row2.pack(fill=tk.X, pady=(5, 0))
        
        self.case_sensitive_var = tk.BooleanVar(value=False)
        case_checkbox = tk.Checkbutton(options_row2, text="Case sensitive search", 
                                      variable=self.case_sensitive_var, font=("Arial", 10))
        case_checkbox.pack(side=tk.LEFT)
        
        # Buttons frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Preview button
        preview_btn = tk.Button(button_frame, text="Preview", 
                               command=self.preview_changes,
                               bg="#e3f2fd", font=("Arial", 10))
        preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Standard Replace button
        self.replace_btn = tk.Button(button_frame, text="Standard Replace", 
                                    command=self.replace_text_in_documents,
                                    bg="#4caf50", fg="white", font=("Arial", 10, "bold"))
        self.replace_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Advanced Replace button
        advanced_replace_btn = tk.Button(button_frame, text="Advanced Replace", 
                                        command=self.advanced_replace_with_vba,
                                        bg="#ff9800", fg="white", font=("Arial", 10, "bold"))
        advanced_replace_btn.pack(side=tk.LEFT, padx=(0, 10))
                
        # Cancel button
        cancel_btn = tk.Button(button_frame, text="Close", 
                              command=self.root.destroy,
                              bg="#333333", fg="white", font=("Arial", 10))
        cancel_btn.pack(side=tk.RIGHT)
        
        # Progress bar frame (initially hidden)
        self.progress_frame = tk.Frame(main_frame)
        self.progress_label = tk.Label(self.progress_frame, text="", font=("Arial", 9))
        self.progress_label.pack(anchor="w")
        
        # Progress percentage label
        self.progress_percent_label = tk.Label(self.progress_frame, text="", font=("Arial", 9), fg="blue")
        self.progress_percent_label.pack(anchor="w")
        
        # Status label
        self.status_label = tk.Label(main_frame, text="Ready", fg="green", font=("Arial", 9))
        self.status_label.pack(anchor="w")
        
        # Setup keyboard bindings
        self.setup_keyboard_bindings()
        
        # Update the file list display
        self.update_file_list()
        
        # Focus on search field
        self.search_text.focus()

    # NEW: NBSP Processing Methods
    def preprocess_text_with_nbsp(self, text):
        """Convert _nbsp_ placeholders to actual non-breaking spaces"""
        return text.replace('_nbsp_', '\u00A0')

    def get_processed_search_text(self):
        """Get search text with _nbsp_ converted to actual non-breaking spaces"""
        raw_text = self.search_text.get("1.0", tk.END).strip()
        return self.preprocess_text_with_nbsp(raw_text)

    def get_processed_replace_text(self):
        """Get replace text with _nbsp_ converted to actual non-breaking spaces"""
        raw_text = self.replace_text.get("1.0", tk.END).strip()
        return self.preprocess_text_with_nbsp(raw_text)

    def debug_nbsp_conversion(self):
        """Debug method to show _nbsp_ conversion"""
        raw_search = self.search_text.get("1.0", tk.END).strip()
        processed_search = self.get_processed_search_text()
        raw_replace = self.replace_text.get("1.0", tk.END).strip()
        processed_replace = self.get_processed_replace_text()
        
        # Define the NBSP character outside the f-string
        nbsp_char = '\u00A0'
        
        debug_info = "nbsp found:\n\n"
        debug_info += f"Raw search text: '{raw_search}'\n"
        debug_info += f"Processed search text: '{processed_search}'\n"
        debug_info += f"nbsp in search text: {processed_search.count(nbsp_char)}\n\n"
        debug_info += f"Raw replace text: '{raw_replace}'\n"
        debug_info += f"Processed replace text: '{processed_replace}'\n"
        debug_info += f"nbsp in replace text: {processed_replace.count(nbsp_char)}\n\n"
        
        messagebox.showinfo("nbsp check", debug_info)

    def check_hyperlinks(self):
        """Check selected files for hyperlinks using python-docx"""
        if not self.file_paths:
            messagebox.showwarning("Warning", "Please add at least one Word document.")
            return
        
        try:
            self.status_label.config(text="Checking for hyperlinks...", fg="orange")
            self.progress_frame.pack(fill=tk.X, pady=(10, 0))
            self.root.update()
            
            hyperlink_results = []
            total_hyperlinks = 0
            files_with_hyperlinks = 0
            
            for i, file_path in enumerate(self.file_paths):
                # Update progress
                progress_percent = int((i / len(self.file_paths)) * 100)
                self.progress_label.config(text=f"Checking {i+1}/{len(self.file_paths)}: {os.path.basename(file_path)}")
                self.progress_percent_label.config(text=f"Progress: {progress_percent}%")
                self.root.update()
                
                try:
                    doc = Document(file_path)
                    file_hyperlinks = []
                    file_hyperlink_count = 0
                    
                    # Check hyperlinks in paragraphs
                    for paragraph in doc.paragraphs:
                        for run in paragraph.runs:
                            if hasattr(run.element, 'hyperlink') and run.element.hyperlink is not None:
                                hyperlink = run.element.hyperlink
                                # Get hyperlink address
                                if hasattr(hyperlink, 'address') and hyperlink.address:
                                    file_hyperlinks.append({
                                        'text': run.text,
                                        'url': hyperlink.address,
                                        'location': 'paragraph'
                                    })
                                    file_hyperlink_count += 1
                    
                    # Check hyperlinks in tables
                    for table_idx, table in enumerate(doc.tables):
                        for row_idx, row in enumerate(table.rows):
                            for cell_idx, cell in enumerate(row.cells):
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        if hasattr(run.element, 'hyperlink') and run.element.hyperlink is not None:
                                            hyperlink = run.element.hyperlink
                                            if hasattr(hyperlink, 'address') and hyperlink.address:
                                                file_hyperlinks.append({
                                                    'text': run.text,
                                                    'url': hyperlink.address,
                                                    'location': f'table {table_idx+1}, row {row_idx+1}, cell {cell_idx+1}'
                                                })
                                                file_hyperlink_count += 1
                    
                    # Alternative method: Check relationships for hyperlinks
                    if hasattr(doc, 'part') and hasattr(doc.part, 'rels'):
                        for rel in doc.part.rels.values():
                            if rel.reltype.endswith('/hyperlink'):
                                # This is a hyperlink relationship
                                # Try to find the text that uses this relationship
                                file_hyperlink_count += 1
                    
                    hyperlink_results.append({
                        'filename': os.path.basename(file_path),
                        'hyperlink_count': file_hyperlink_count,
                        'hyperlinks': file_hyperlinks[:10],  # Limit to first 10 for display
                        'total_found': len(file_hyperlinks)
                    })
                    
                    if file_hyperlink_count > 0:
                        files_with_hyperlinks += 1
                        total_hyperlinks += file_hyperlink_count
                        
                except Exception as e:
                    hyperlink_results.append({
                        'filename': os.path.basename(file_path),
                        'hyperlink_count': 0,
                        'hyperlinks': [],
                        'error': str(e)
                    })
            
            # Hide progress
            self.progress_frame.pack_forget()
            self.status_label.config(text="Hyperlink check completed!", fg="green")
            
            # Show results
            self.show_hyperlink_results(hyperlink_results, total_hyperlinks, files_with_hyperlinks)
            
        except Exception as e:
            self.progress_frame.pack_forget()
            self.status_label.config(text="Error occurred", fg="red")
            messagebox.showerror("Error", f"Error during hyperlink check: {str(e)}")

    def show_hyperlink_results(self, results, total_hyperlinks, files_with_hyperlinks):
        """Show hyperlink check results in scrollable window"""
        message = "-" * 26 + "\n"
        message += f"🔗 Hyperlink check results\n"
        message += "-" * 26 + "\n\n"
        message += f"📊 Summary:\n"
        message += f"Files analyzed: {len(results)}\n"
        message += f"Total hyperlinks found: {total_hyperlinks}\n"
        message += f"Files with hyperlinks: {files_with_hyperlinks}\n\n"
        message += "=" * 77 + "\n\n"
        
        if total_hyperlinks > 0:
            message += "⚠️  Reminder for hyperlinks:\n"
            message += "• Standard Replace updates display text but removes the actual link!\n"
            message += "• Advanced Replace preserves hyperlinks when editing their display text.\n"
            message += "• If you are not going to replace the display text of a hyperlink, you can still use Standard Replace as it is faster.\n\n"
            message += "=" * 77 + "\n\n"
        
        # Show detailed file results
        for i, result in enumerate(results, 1):
            message += f"📁 File {i}: {result['filename']}\n"
            
            if 'error' in result:
                message += f"   ❌ Error: {result['error']}\n"
            elif result['hyperlink_count'] > 0:
                message += f"   🔗 Hyperlinks found: {result['hyperlink_count']}\n"
                
                # Show sample hyperlinks
                if result['hyperlinks']:
                    message += f"   📋 Sample hyperlinks:\n"
                    for idx, hyperlink in enumerate(result['hyperlinks'][:5], 1):  # Show max 5
                        text_preview = hyperlink['text'][:30] + "..." if len(hyperlink['text']) > 30 else hyperlink['text']
                        url_preview = hyperlink['url'][:40] + "..." if len(hyperlink['url']) > 40 else hyperlink['url']
                        message += f"      {idx}. Text: '{text_preview}'\n"
                        message += f"         URL: {url_preview}\n"
                        message += f"         Location: {hyperlink['location']}\n"
                    
                    if result['total_found'] > 5:
                        message += f"      ... and {result['total_found'] - 5} more hyperlinks\n"
            else:
                message += f"   ✅ No hyperlinks found\n"
            
            message += "\n" + "=" * 77 + "\n\n"
        
        if total_hyperlinks < 1:
            message += "💡 Both Standard and Advanced Replace are safe to use."
        else:
            message += "💡 Only use Standard Replace if you are not going to change display texts of hyperlinks."

        self.show_scrollable_results(message, "Hyperlink check results")

    def handle_delete_key(self, event):
        """Handle Delete key press in file listbox"""
        self.remove_selected_file()
        return "break"  # Prevent any other handling
    
    def show_shortcuts_tooltip(self, event):
        """Show shortcuts tooltip on hover - SMART VISIBILITY MANAGEMENT"""
        # Prevent multiple tooltips - but check if existing one is still valid
        if hasattr(self, 'tooltip'):
            try:
                if self.tooltip.winfo_exists():
                    return
            except tk.TclError:
                # Tooltip was destroyed but attribute still exists
                del self.tooltip

        # Create tooltip window - COMPLETELY INDEPENDENT
        self.tooltip = tk.Toplevel()  # Don't pass self.root as parent!
        self.tooltip.wm_overrideredirect(True)  # Remove window decorations
        self.tooltip.configure(bg="lightyellow", relief="solid", borderwidth=1)
        
        # Make tooltip stay on top but don't make it transient
        self.tooltip.attributes('-topmost', True)
        
        # BIND FOCUS EVENTS TO MAIN WINDOW TO CONTROL TOOLTIP VISIBILITY
        def on_main_focus_in(event):
            """Show tooltip when main window gets focus"""
            if hasattr(self, 'tooltip'):
                try:
                    self.tooltip.deiconify()  # Show tooltip
                except tk.TclError:
                    pass
        
        def on_main_focus_out(event):
            """Hide tooltip when main window loses focus"""
            if hasattr(self, 'tooltip'):
                try:
                    self.tooltip.withdraw()  # Hide tooltip (but don't destroy)
                except tk.TclError:
                    pass
        
        # Bind focus events to main window
        self.root.bind('<FocusIn>', on_main_focus_in, add='+')
        self.root.bind('<FocusOut>', on_main_focus_out, add='+')
        
        # Also bind to window state changes (minimize/restore)
        def on_window_state_change(event):
            """Handle window minimize/restore"""
            if hasattr(self, 'tooltip'):
                try:
                    if self.root.state() == 'iconic':  # Minimized
                        self.tooltip.withdraw()
                    else:  # Normal or zoomed
                        self.tooltip.deiconify()
                except tk.TclError:
                    pass
        
        self.root.bind('<Unmap>', on_window_state_change, add='+')
        self.root.bind('<Map>', on_window_state_change, add='+')
        
        # Create a frame to hold everything
        main_frame = tk.Frame(self.tooltip, bg="lightyellow")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Tooltip content
        shortcuts_text = """    𝗥𝗲𝗽𝗹𝗮𝗰𝗲𝗺𝗲𝗻𝘁 𝗺𝗲𝘁𝗵𝗼𝗱𝘀 
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    Preview:
    • Shows what will be changed

    Standard Replace:
    • Handles standard text & tables (fast)
    • ⚠️ Hyperlinks: Editing the display text of a hyperlink updates the display text, but removes the link.

    Advanced Replace:
    • Handles all document areas (complete coverage) and handles hyperlinks correctly

    𝗞𝗲𝘆𝗯𝗼𝗮𝗿𝗱 𝘀𝗵𝗼𝗿𝘁𝗰𝘂𝘁𝘀
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    Actions:
    • F5 - Preview
    • Ctrl+Enter - Standard Replace
    • Shift+Enter - Advanced Replace
    
    Files:
    • Ctrl+O - Add more files
    • Delete - Remove selected files (when file list is focused)
    
    Navigation:
    • Ctrl+1 - Focus Search box
    • Ctrl+2 - Focus Replace box
    • Tab - Switch between boxes

    Text Editing:
    • Ctrl+Tab - Add a tab (\\t) in search/replace box
    • Ctrl+Z - Undo
    • Ctrl+Y - Redo
    • Shift+Space - Insert _nbsp_ at cursor position

    Exit:
    • Escape/Ctrl+Q - Close application

    𝗠𝗶𝘀𝗰𝗲𝗹𝗹𝗮𝗻𝗲𝗼𝘂𝘀
    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    Non-breaking spaces:
    • Type '_nbsp_' for non-breaking spaces or use Shift+Space
    • 🔍 nbsp check - Counts non-breaking spaces in search/replace boxes

    Hyperlinks:
    • 🔗 Hyperlink check - Checks the selected files for hyperlinks
    
    𝗙𝗼𝗿𝘇𝗮 𝗜𝗻𝘁𝗲𝗿!"""
        
        # Create scrollable text widget instead of label
        text_widget = tk.Text(main_frame, 
                            font=("Arial", 9), 
                            bg="lightyellow", 
                            wrap=tk.WORD,
                            width=80,  # Set reasonable width
                            height=30,  # Set reasonable height
                            relief="flat",
                            borderwidth=0,
                            padx=5,
                            pady=3)
        
        # Create scrollbar
        scrollbar = tk.Scrollbar(main_frame, orient="vertical", command=text_widget.yview)
        text_widget.config(yscrollcommand=scrollbar.set)
        
        # Pack text widget and scrollbar
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Insert text and make read-only
        text_widget.insert("1.0", shortcuts_text)
        text_widget.config(state=tk.DISABLED)
        
        # SIMPLE BUT EFFECTIVE POSITIONING - STAY NEAR MAIN WINDOW
        # Get main window bounds
        main_x = self.root.winfo_rootx()
        main_y = self.root.winfo_rooty()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        
        # Get icon position
        icon_x = self.info_icon.winfo_rootx()
        icon_y = self.info_icon.winfo_rooty()
        
        # Set a reasonable tooltip size first
        self.tooltip.geometry("650x600")  # Fixed size
        self.tooltip.update_idletasks()
        
        tooltip_width = 650
        tooltip_height = 600
        
        # Strategy: Try positions in order of preference, all relative to main window
        positions_to_try = [
            # 1. Left of icon (preferred)
            (icon_x - tooltip_width - 10, icon_y),
            # 2. Right of main window
            (main_x + main_width + 10, main_y),
            # 3. Left of main window
            (main_x - tooltip_width - 10, main_y),
            # 4. Above main window
            (main_x, main_y - tooltip_height - 10),
            # 5. Below main window
            (main_x, main_y + main_height + 10),
            # 6. Centered on main window (last resort)
            (main_x + (main_width - tooltip_width) // 2, main_y + (main_height - tooltip_height) // 2)
        ]
        
        # Find the first position that keeps tooltip reasonably close to main window
        final_x, final_y = positions_to_try[0]  # Default to first option
        
        for test_x, test_y in positions_to_try:
            # Check if this position keeps tooltip close to main window
            distance_from_main = abs(test_x - main_x) + abs(test_y - main_y)
            if distance_from_main < 1000:  # Within reasonable distance
                final_x, final_y = test_x, test_y
                break
        
        # Ensure tooltip doesn't go off-screen (but prioritize staying near main window)
        # Only adjust if absolutely necessary
        if final_x < main_x - 800:  # Too far left
            final_x = main_x - 400
        if final_y < main_y - 400:  # Too far up
            final_y = main_y - 200
        
        # Apply final position
        self.tooltip.geometry(f"650x600+{final_x}+{final_y}")
        
        # Add close button in the top-right corner - ONLY WAY TO CLOSE
        close_btn = tk.Button(main_frame, text="🗙", 
                            command=self.hide_shortcuts_tooltip_manual,
                            font=("Arial", 8, "bold"),
                            bg="lightcoral", fg="white",
                            width=2, height=1,
                            relief="flat")
        close_btn.place(relx=1.0, rely=0.0, anchor="ne", x=-23, y=0)

    def hide_shortcuts_tooltip(self, event):
        """Do nothing - tooltip only closes when X is clicked"""
        pass

    def hide_shortcuts_tooltip_manual(self):
        """Manually hide tooltip when close button is clicked"""
        if hasattr(self, 'tooltip'):
            try:
                self.tooltip.destroy()
            except tk.TclError:
                pass  # Already destroyed
            del self.tooltip



#########END OF PART 1######




#########PART 2######
    def setup_keyboard_bindings(self):
        """Setup custom keyboard bindings"""
        # Existing text box bindings
        self.search_text.bind('<Tab>', self.focus_replace_text)
        self.replace_text.bind('<Tab>', self.focus_search_text)
        self.search_text.bind('<Control-Tab>', self.insert_tab)
        self.replace_text.bind('<Control-Tab>', self.insert_tab)
        self.search_text.bind('<Control-z>', self.undo_search)
        self.search_text.bind('<Control-y>', self.redo_search)
        self.replace_text.bind('<Control-z>', self.undo_replace)
        self.replace_text.bind('<Control-y>', self.redo_replace)
        self.search_text.bind('<Control-Shift-Z>', self.redo_search)
        self.replace_text.bind('<Control-Shift-Z>', self.redo_replace)

        # NEW: NBSP shortcut - Shift+Space inserts _nbsp_ at cursor position
        self.search_text.bind('<Shift-space>', self.insert_nbsp_placeholder)
        self.replace_text.bind('<Shift-space>', self.insert_nbsp_placeholder)
        
        # GLOBAL KEYBOARD SHORTCUTS (bind to root)
        self.root.bind('<Control-Return>', lambda e: self.replace_text_in_documents())
        self.root.bind('<F5>', lambda e: self.preview_changes())
        self.root.bind('<Control-o>', lambda e: self.add_files())
        self.root.bind('<Control-O>', lambda e: self.add_files())
        self.root.bind('<Escape>', lambda e: self.root.destroy())
        self.root.bind('<Control-q>', lambda e: self.root.destroy())
        self.root.bind('<Control-Q>', lambda e: self.root.destroy())
        self.root.bind('<Control-Shift-Return>', lambda e: self.replace_and_close())
        self.root.bind('<Control-Key-1>', lambda e: self.search_text.focus())
        self.root.bind('<Control-Key-2>', lambda e: self.replace_text.focus())
        self.root.bind('<Shift-Return>', lambda e: self.advanced_replace_with_vba())
        
        # BIND SHORTCUTS TO TEXT WIDGETS TOO (to prevent line breaks)
        # Search text shortcuts
        self.search_text.bind('<Control-Return>', self.handle_shortcut_replace)
        self.search_text.bind('<Control-Shift-Return>', self.handle_shortcut_replace_close)
        self.search_text.bind('<F5>', self.handle_shortcut_preview)
        self.search_text.bind('<Control-o>', self.handle_shortcut_add_files)
        self.search_text.bind('<Control-O>', self.handle_shortcut_add_files)
        self.search_text.bind('<Escape>', self.handle_shortcut_close)
        self.search_text.bind('<Control-q>', self.handle_shortcut_close)
        self.search_text.bind('<Control-Q>', self.handle_shortcut_close)
        self.search_text.bind('<Shift-Return>', self.handle_shortcut_advanced_replace)
        
        # Replace text shortcuts
        self.replace_text.bind('<Control-Return>', self.handle_shortcut_replace)
        self.replace_text.bind('<Control-Shift-Return>', self.handle_shortcut_replace_close)
        self.replace_text.bind('<F5>', self.handle_shortcut_preview)
        self.replace_text.bind('<Control-o>', self.handle_shortcut_add_files)
        self.replace_text.bind('<Control-O>', self.handle_shortcut_add_files)
        self.replace_text.bind('<Escape>', self.handle_shortcut_close)
        self.replace_text.bind('<Control-q>', self.handle_shortcut_close)
        self.replace_text.bind('<Control-Q>', self.handle_shortcut_close)
        self.replace_text.bind('<Shift-Return>', self.handle_shortcut_advanced_replace)
        
        # Make sure the window can receive focus for global shortcuts
        self.root.focus_set()
    
    def handle_shortcut_replace(self, event):
        """Handle Ctrl+Enter shortcut from text widgets"""
        self.replace_text_in_documents()
        return "break"  # Prevent line break insertion
    
    def handle_shortcut_replace_close(self, event):
        """Handle Ctrl+Shift+Enter shortcut from text widgets"""
        self.replace_and_close()
        return "break"  # Prevent line break insertion
    
    def handle_shortcut_preview(self, event):
        """Handle F5 shortcut from text widgets"""
        self.preview_changes()
        return "break"  # Prevent any default behavior
    
    def handle_shortcut_add_files(self, event):
        """Handle Ctrl+O shortcut from text widgets"""
        self.add_files()
        return "break"  # Prevent any default behavior
    
    def handle_shortcut_close(self, event):
        """Handle Escape/Ctrl+Q shortcut from text widgets"""
        self.root.destroy()
        return "break"  # Prevent any default behavior

    def handle_shortcut_advanced_replace(self, event):
        """Handle Shift+Enter shortcut from text widgets"""
        self.advanced_replace_with_vba()
        return "break"  # Prevent line break insertion
    
    def replace_and_close(self):
        """Replace and close application"""
        if self.replace_text_in_documents(close_after=True):
            pass  # The close will be handled in the results window
    
    def insert_tab(self, event):
        """Insert a tab character at cursor position"""
        widget = event.widget
        widget.insert(tk.INSERT, '\t')
        return "break"  # Prevent any other handling

    def insert_nbsp_placeholder(self, event):
        """Insert _nbsp_ placeholder at cursor position"""
        widget = event.widget
        cursor_pos = widget.index(tk.INSERT)
        widget.insert(cursor_pos, '_nbsp_')
        return "break"  # Prevent any other handling
    
    def focus_replace_text(self, event):
        """Move focus from search to replace text box"""
        self.replace_text.focus()
        return "break"  # Prevent default Tab behavior
    
    def focus_search_text(self, event):
        """Move focus from replace to search text box"""
        self.search_text.focus()
        return "break"  # Prevent default Tab behavior
    
    def undo_search(self, event):
        """Undo in search text box"""
        try:
            self.search_text.edit_undo()
        except tk.TclError:
            pass  # No more undo operations
        return "break"
    
    def redo_search(self, event):
        """Redo in search text box"""
        try:
            self.search_text.edit_redo()
        except tk.TclError:
            pass  # No more redo operations
        return "break"
    
    def undo_replace(self, event):
        """Undo in replace text box"""
        try:
            self.replace_text.edit_undo()
        except tk.TclError:
            pass  # No more undo operations
        return "break"
    
    def redo_replace(self, event):
        """Redo in replace text box"""
        try:
            self.replace_text.edit_redo()
        except tk.TclError:
            pass  # No more redo operations
        return "break"
    
    def paste_to_search(self):
        """Paste clipboard content to search text box"""
        try:
            clipboard_content = self.root.clipboard_get()
            self.search_text.delete("1.0", tk.END)
            self.search_text.insert("1.0", clipboard_content)
            self.status_label.config(text="Pasted clipboard content to search box.", fg="blue")
        except tk.TclError:
            self.status_label.config(text="Clipboard is empty or contains non-text data.", fg="orange")
    
    def paste_to_replace(self):
        """Paste clipboard content to replace text box"""
        try:
            clipboard_content = self.root.clipboard_get()
            self.replace_text.delete("1.0", tk.END)
            self.replace_text.insert("1.0", clipboard_content)
            self.status_label.config(text="Pasted clipboard content to replace box.", fg="blue")
        except tk.TclError:
            self.status_label.config(text="Clipboard is empty or contains non-text data.", fg="orange")
    
    def update_title(self):
        """Update window title with file count"""
        file_count = len(self.file_paths)
        if file_count == 0:
            self.root.title("Bulk Text Replacement for Word - No files selected.")
        else:
            self.root.title(f"Bulk Text Replacement for Word - {file_count} file{'s' if file_count != 1 else ''}")
    
    def update_file_list(self):
        """Update the file listbox and related UI elements"""
        self.file_listbox.delete(0, tk.END)
        
        for file_path in self.file_paths:
            self.file_listbox.insert(tk.END, os.path.basename(file_path))
        
        # Update file count label
        file_count = len(self.file_paths)
        if file_count == 0:
            self.file_count_label.config(text="No files selected.")
            self.replace_btn.config(state="disabled")
        else:
            self.file_count_label.config(text=f"{file_count} file{'s' if file_count != 1 else ''} selected.")
            self.replace_btn.config(state="normal")
        
        self.update_title()
    
    def add_files(self):
        """Add more Word documents to the list"""
        file_types = [
            ("Word Documents", "*.docx *.doc *.docm"),
            ("Word 2007+ Documents", "*.docx *.docm"),
            ("Word 97-2003 Documents", "*.doc"),
            ("All files", "*.*")
        ]
        
        selected_files = filedialog.askopenfilenames(
            title="Select Word Documents to add",
            filetypes=file_types,
            initialdir=os.path.dirname(self.file_paths[0]) if self.file_paths else None
        )
        
        added_count = 0
        for file_path in selected_files:
            # Normalize the path to avoid duplicates due to different path formats
            normalized_path = os.path.normpath(os.path.abspath(file_path))
            
            # Check if this normalized path is already in our list (also normalized)
            already_exists = False
            for existing_path in self.file_paths:
                if os.path.normpath(os.path.abspath(existing_path)) == normalized_path:
                    already_exists = True
                    break
            
            if not already_exists:
                self.file_paths.append(file_path)
                added_count += 1
        
        if added_count > 0:
            self.update_file_list()
            self.status_label.config(text=f"Added {added_count} file{'s' if added_count != 1 else ''}", fg="blue")
        else:
            if len(selected_files) > 0:
                self.status_label.config(text="Files already in list - no duplicates added.", fg="orange")
            else:
                self.status_label.config(text="No new files added.", fg="orange")
    
    def remove_selected_file(self):
        """Remove the selected file(s) from the list"""
        selections = self.file_listbox.curselection()
        if selections:
            # Get filenames before removing (for status message)
            removed_files = []
            for index in selections:
                removed_files.append(os.path.basename(self.file_paths[index]))
            
            # Remove files in reverse order to maintain correct indices
            for index in reversed(selections):
                del self.file_paths[index]
            
            self.update_file_list()
            
            # Update status message
            if len(removed_files) == 1:
                self.status_label.config(text=f"Removed: {removed_files[0]}", fg="orange")
            else:
                self.status_label.config(text=f"Removed {len(removed_files)} files", fg="orange")
        else:
            messagebox.showwarning("Warning", "Please select one or more files to remove.")
    
    def clear_all_files(self):
        """Clear all files from the list"""
        if self.file_paths:
            if messagebox.askyesno("Confirm", "Remove all files from the list?"):
                file_count = len(self.file_paths)
                self.file_paths.clear()
                self.update_file_list()
                self.status_label.config(text=f"Cleared {file_count} file{'s' if file_count != 1 else ''}", fg="orange")
    
    def get_document_text(self, doc):
        """Extract all text from document including paragraphs and tables"""
        full_text = []
        
        # Get text from paragraphs
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        
        # Get text from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        full_text.append(paragraph.text)
        
        return '\n'.join(full_text)
    
    def count_occurrences(self, text, search_for, case_sensitive=False):
        """Count occurrences with case sensitivity option"""
        if not case_sensitive:
            return text.lower().count(search_for.lower())
        else:
            return text.count(search_for)

    def preview_changes(self):
        """Show preview with choice between Standard (fast) or Advanced (comprehensive) mode"""
        if not self.file_paths:
            messagebox.showwarning("Warning", "Please add at least one Word document.")
            return
            
        search_for = self.get_processed_search_text()
        
        if not search_for:
            messagebox.showwarning("Warning", "Please enter text to search for.")
            return
        
        # NEW: Ask user for preview mode
        preview_choice = messagebox.askyesnocancel(
            "Preview Mode Selection",
            "Do you want to run the fast Standard Preview?\n\n" +
            "YES = Standard Preview (Fast)\n" +
            "   • Main content only (paragraphs + tables)\n\n" +
            "NO = Advanced Preview (Comprehensive)\n" +
            "   • Main content + shapes + headers + footers + footnotes + endnotes + form fields + hyperlinks"
        )
        
        if preview_choice is None:  # User clicked Cancel
            return
        elif preview_choice:  # User clicked Yes - Standard Preview
            self.preview_standard_only()
        else:  # User clicked No - Advanced Preview
            self.preview_comprehensive()

    def preview_standard_only(self):
        """Fast preview - Main content only (python-docx)"""
        search_for = self.get_processed_search_text()
        case_sensitive = self.case_sensitive_var.get()
        
        try:
            preview_details = []
            total_matches = 0
            files_with_matches = 0
            
            # Show progress for standard preview
            self.status_label.config(text="Standard preview in progress...", fg="orange")
            self.progress_frame.pack(fill=tk.X, pady=(10, 0))
            self.root.update()
            
            for i, file_path in enumerate(self.file_paths):
                # Update progress
                progress_percent = int((i / len(self.file_paths)) * 100)
                self.progress_label.config(text=f"Analyzing {i+1}/{len(self.file_paths)}: {os.path.basename(file_path)}")
                self.progress_percent_label.config(text=f"Progress: {progress_percent}%")
                self.root.update()
                
                file_result = {
                    'filename': os.path.basename(file_path),
                    'main_total': 0,
                    'details': []
                }
                
                try:
                    # Analyze main content only (python-docx)
                    doc = Document(file_path)
                    document_text = self.get_document_text(doc)
                    main_occurrences = self.count_occurrences(document_text, search_for, case_sensitive)
                    
                    file_result['main_total'] = main_occurrences
                    if main_occurrences > 0:
                        files_with_matches += 1
                        total_matches += main_occurrences
                        file_result['details'].append(f"  📄 Main content: {main_occurrences} match(es)")
                    else:
                        file_result['details'].append(f"  📄 Main content: No matches")
                    
                    preview_details.append(file_result)
                    
                except Exception as e:
                    file_result['details'] = [f"  ❌ Error: {str(e)}"]
                    preview_details.append(file_result)
            
            # Hide progress
            self.progress_frame.pack_forget()
            self.status_label.config(text="Standard preview completed!", fg="green")
            
            # Show standard preview results
            self.show_standard_preview_results(
                preview_details, search_for, self.get_processed_replace_text(),
                total_matches, files_with_matches, case_sensitive
            )
            
        except Exception as e:
            self.progress_frame.pack_forget()
            self.status_label.config(text="Error occurred", fg="red")
            messagebox.showerror("Error", f"Error during standard preview: {str(e)}")

    def show_standard_preview_results(self, results, search_text, replace_text, total_matches, files_with_matches, case_sensitive):
        """Show standard preview results (main content only)"""
        case_info = " (case sensitive)" if case_sensitive else " (case insensitive)"
        
        message = "-" * 46 + "\n"
        message += f"📄 Standard Preview results{case_info}\n"
        message += "-" * 46 + "\n\n"
        message += f"⚡ Fast Analysis: {total_matches} match(es) found in main content\n\n"
        message += f"📊 Summary:\n"
        message += f"Files analyzed: {len(results)}\n"
        message += f"Main content matches: {total_matches} (in {files_with_matches} files)\n"
        message += f"Search text: '{search_text[:40]}{'...' if len(search_text) > 40 else ''}'\n"
        message += f"Replace text: '{replace_text[:40]}{'...' if len(replace_text) > 40 else ''}'\n"
        message += f"Coverage: Main text and tables only\n"
        message += "=" * 77 + "\n\n"
        
        # Show detailed file results
        for i, result in enumerate(results, 1):
            message += f"📁 FILE {i}: {result['filename']}\n"
            
            if result['main_total'] > 0:
                message += f"   ✅ Matches found: {result['main_total']}\n"
            else:
                message += f"   ✅ No matches found\n"
            
            for detail in result['details']:
                message += f"   {detail}\n"
            message += "\n" + "=" * 77 + "\n\n"
        
        # Recommendations
        if total_matches > 0:
            message += "💡 Replacement recommendations:\n"
            message += "Use 'Standard Replace' for fast replacement of main content\n"
            message += "Use 'Advanced Replace' if you also need shapes/headers/footers\n"
        else:
            message += "💡 Tip: No matches found in main content. Try 'Advanced Preview' to run a comprehensive check."
        
        self.show_scrollable_results(message, "Standard Preview results")

    def preview_comprehensive(self):
        """Comprehensive preview - All areas using VBScript only"""
        search_for = self.get_processed_search_text()
        case_sensitive = self.case_sensitive_var.get()
        
        try:
            preview_details = []
            total_advanced_matches = 0
            files_with_advanced_matches = 0
            
            # Show progress for comprehensive preview
            self.status_label.config(text="Comprehensive analysis in progress...", fg="orange")
            self.progress_frame.pack(fill=tk.X, pady=(10, 0))
            self.root.update()
            
            # Check if VBScript is available for advanced preview
            script_dir = os.path.dirname(os.path.abspath(__file__))
            vbs_path = os.path.join(script_dir, "word_advanced_preview.vbs")
            advanced_available = os.path.exists(vbs_path)
            
            for i, file_path in enumerate(self.file_paths):
                # Update progress
                progress_percent = int((i / len(self.file_paths)) * 100)
                self.progress_label.config(text=f"Analyzing {i+1}/{len(self.file_paths)}: {os.path.basename(file_path)}")
                self.progress_percent_label.config(text=f"Progress: {progress_percent}%")
                self.root.update()
                
                try:
                    file_result = {
                        'filename': os.path.basename(file_path),
                        'total': 0,
                        'details': []
                    }
                    
                    # Use VBScript for everything (matches what Advanced Replace does)
                    if advanced_available:
                        advanced_result = self.preview_advanced_areas(file_path, search_for, case_sensitive)
                        file_result['total'] = advanced_result['total']
                        
                        if advanced_result['total'] > 0:
                            files_with_advanced_matches += 1
                            total_advanced_matches += advanced_result['total']
                            file_result['details'].extend(advanced_result['details'])
                        else:
                            file_result['details'].append(f"  📄 No matches found in any areas")
                    else:
                        file_result['details'].append(f"  📄 VBScript not available")
                    
                    preview_details.append(file_result)
                    
                except Exception as e:
                    file_result['details'] = [f"  ❌ Error: {str(e)}"]
                    preview_details.append(file_result)
            
            # Hide progress
            self.progress_frame.pack_forget()
            self.status_label.config(text="Comprehensive preview completed!", fg="green")
            
            # Show comprehensive preview results
            self.show_comprehensive_preview_results(
                preview_details, search_for, self.get_processed_replace_text(),
                total_advanced_matches, files_with_advanced_matches, 
                case_sensitive, advanced_available
            )
            
        except Exception as e:
            self.progress_frame.pack_forget()
            self.status_label.config(text="Error occurred", fg="red")
            messagebox.showerror("Error", f"Error during comprehensive preview: {str(e)}")
#########END OF PART 2######





#########PART 3A######
    def preview_advanced_areas(self, file_path, search_text, case_sensitive=False):
        """Preview advanced areas using VBScript - READ ONLY"""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        vbs_preview_path = os.path.join(script_dir, "word_advanced_preview.vbs")
        
        # Create a preview-only VBScript if it doesn't exist
        if not os.path.exists(vbs_preview_path):
            self.create_preview_vbscript(vbs_preview_path)
        
        result = {
            'total': 0,
            'details': []
        }
        
        try:
            full_path = os.path.abspath(file_path)
            
            # HIDE CMD WINDOW
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            
            # Use case sensitivity flag: 1 for case sensitive, 0 for case insensitive
            case_flag = "1" if case_sensitive else "0"
            
            vbs_result = subprocess.run([
                'cscript', '//NoLogo', vbs_preview_path, full_path, search_text, case_flag
            ], capture_output=True, text=True, cwd=script_dir,
               startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
            
            if vbs_result.returncode == 0:
                # Parse successful output
                output_lines = vbs_result.stdout.strip().split('\n')
                
                for line in output_lines:
                    if "SHAPES:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📦 Shapes: {count} match(es)")
                            result['total'] += count
                    elif "HEADERS:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📄 Headers: {count} match(es)")
                            result['total'] += count
                    elif "FOOTERS:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📄 Footers: {count} match(es)")
                            result['total'] += count
                    elif "FOOTNOTES:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📝 Footnotes: {count} match(es)")
                            result['total'] += count
                    elif "ENDNOTES:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📝 Endnotes: {count} match(es)")
                            result['total'] += count
                    elif "FORMFIELDS:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📋 Form fields: {count} match(es)")
                            result['total'] += count
                    elif "HYPERLINKS:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  🔗 Hyperlinks: {count} match(es)")
                            result['total'] += count
                    elif "MAINCONTENT:" in line:
                        count = int(line.split(':')[1].split()[0])
                        if count > 0:
                            result['details'].append(f"  📄 Main content: {count} match(es)")
                            result['total'] += count
            
            return result
            
        except Exception as e:
            result['details'] = [f"  ❌ Advanced preview error: {str(e)}"]
            return result

    def create_preview_vbscript(self, vbs_path):
        """Create a preview-only VBScript for advanced area scanning"""
        vbs_content = '''
' Word Advanced Preview Script - READ ONLY
Dim args, filePath, searchText, caseSensitive
Set args = WScript.Arguments

If args.Count < 3 Then
    WScript.Echo "Usage: script.vbs <filepath> <searchtext> <casesensitive>"
    WScript.Quit 1
End If

filePath = args(0)
searchText = args(1)
caseSensitive = CBool(args(2))

Dim wordApp, doc
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
wordApp.DisplayAlerts = False

On Error Resume Next
Set doc = wordApp.Documents.Open(filePath, , True) ' Read-only
If Err.Number <> 0 Then
    WScript.Echo "Error opening document"
    wordApp.Quit
    WScript.Quit 1
End If

Dim shapesCount, headersCount, footersCount
shapesCount = 0
headersCount = 0
footersCount = 0

' Count in shapes
For Each shape In doc.Shapes
    If shape.HasTextFrame Then
        If shape.TextFrame.HasText Then
            Dim shapeText
            shapeText = shape.TextFrame.TextRange.Text
            If caseSensitive Then
                shapesCount = shapesCount + CountOccurrences(shapeText, searchText, True)
            Else
                shapesCount = shapesCount + CountOccurrences(shapeText, searchText, False)
            End If
        End If
    End If
Next

' Count in headers and footers
For Each section In doc.Sections
    ' Headers
    For i = 1 To 3
        If section.Headers(i).Exists Then
            Dim headerText
            headerText = section.Headers(i).Range.Text
            If caseSensitive Then
                headersCount = headersCount + CountOccurrences(headerText, searchText, True)
            Else
                headersCount = headersCount + CountOccurrences(headerText, searchText, False)
            End If
        End If
    Next
    
    ' Footers
    For i = 1 To 3
        If section.Footers(i).Exists Then
            Dim footerText
            footerText = section.Footers(i).Range.Text
            If caseSensitive Then
                footersCount = footersCount + CountOccurrences(footerText, searchText, True)
            Else
                footersCount = footersCount + CountOccurrences(footerText, searchText, False)
            End If
        End If
    Next
Next

WScript.Echo "SHAPES: " & shapesCount & " matches"
WScript.Echo "HEADERS: " & headersCount & " matches"
WScript.Echo "FOOTERS: " & footersCount & " matches"

doc.Close False
wordApp.Quit

Function CountOccurrences(text, searchFor, caseSensitive)
    If caseSensitive Then
        CountOccurrences = (Len(text) - Len(Replace(text, searchFor, ""))) / Len(searchFor)
    Else
        CountOccurrences = (Len(text) - Len(Replace(LCase(text), LCase(searchFor), ""))) / Len(searchFor)
    End If
End Function
'''
        
        try:
            with open(vbs_path, 'w', encoding='utf-8') as f:
                f.write(vbs_content)
        except Exception as e:
            pass  # If we can't create it, advanced preview won't work

    def show_comprehensive_preview_results(self, results, search_text, replace_text, 
                                         total_matches, files_with_matches, 
                                         case_sensitive, advanced_available):
        """Show comprehensive preview results"""
        case_info = " (case sensitive)" if case_sensitive else " (case insensitive)"

        message = "-" * 46 + "\n"    
        message += f"🔍 Advanced Preview results{case_info}\n"
        message += "-" * 46 + "\n\n" 
        message += f"🎯 Complete Analysis: {total_matches} total match(es) found\n\n"
        message += f"📊 Summary:\n"
        message += f"Files analyzed: {len(results)}\n"

        if advanced_available:
            message += f"Total matches: {total_matches} (in {files_with_matches} files)\n"
        else:
            message += f"Analysis not available (VBScript missing)\n"

        
        message += f"Search text: '{search_text[:40]}{'...' if len(search_text) > 40 else ''}'\n"
        message += f"Replace text: '{replace_text[:40]}{'...' if len(replace_text) > 40 else ''}'\n"
        message += "=" * 77 + "\n\n"
        
        # Show detailed file results
        for i, result in enumerate(results, 1):
            message += f"📁 FILE {i}: {result['filename']}\n"
            
            if result['total'] > 0:
                message += f"   ✅ Total matches: {result['total']}\n"
            else:
                message += f"   ✅ No matches found\n"

            
            for detail in result['details']:
                message += f"   {detail}\n"
            message += "\n" + "=" * 77 + "\n\n"
        
        # Recommendations
        if total_matches > 0:
            message += "💡 Replacement recommendations:\n"
            message += "Use 'Advanced Replace' for complete coverage with hyperlink preservation\n"
            message += "Or use 'Standard Replace' for faster processing (main content only)\n"
        else:
            message += "💡 Tip: No matches found. Try different search terms or check spelling."
        
        self.show_scrollable_results(message, "Advanced Preview results")

    def replace_in_paragraph_advanced(self, paragraph, search_text, replace_text, case_sensitive=False):
        """Replace text in a paragraph that may span multiple runs"""
        paragraph_text = paragraph.text
        
        # Check for matches based on case sensitivity
        if case_sensitive:
            if search_text not in paragraph_text:
                return 0
            replacements = paragraph_text.count(search_text)
        else:
            if search_text.lower() not in paragraph_text.lower():
                return 0
            replacements = paragraph_text.lower().count(search_text.lower())
        
        # If the search text is contained within a single run, use simple replacement
        for run in paragraph.runs:
            run_text = run.text
            if case_sensitive:
                if search_text in run_text:
                    run.text = run_text.replace(search_text, replace_text)
                    return replacements
            else:
                if search_text.lower() in run_text.lower():
                    # Case insensitive replacement - need to preserve original case positions
                    pattern = re.escape(search_text)
                    run.text = re.sub(pattern, replace_text, run_text, flags=re.IGNORECASE)
                    return replacements
        
        # If we get here, the text spans multiple runs - rebuild the paragraph
        if case_sensitive:
            if search_text in paragraph_text:
                new_paragraph_text = paragraph_text.replace(search_text, replace_text)
            else:
                return 0
        else:
            if search_text.lower() in paragraph_text.lower():
                pattern = re.escape(search_text)
                new_paragraph_text = re.sub(pattern, replace_text, paragraph_text, flags=re.IGNORECASE)
            else:
                return 0
        
        # Clear all runs and add new text to first run
        for run in paragraph.runs:
            run.text = ""
        
        if paragraph.runs:
            paragraph.runs[0].text = new_paragraph_text
        else:
            paragraph.text = new_paragraph_text
        
        return replacements
#########END OF PART 3A######






#########PART 3B######
    def replace_text_in_documents(self, close_after=False):
        """Perform the actual text replacement across all files - Enhanced results"""
        if not self.file_paths:
            messagebox.showwarning("Warning", "Please add at least one Word document.")
            return False
            
        search_for = self.get_processed_search_text()    # CHANGED: Use processed text
        replace_with = self.get_processed_replace_text() # CHANGED: Use processed text
        
        if not search_for:
            messagebox.showwarning("Warning", "Please enter text to search for.")
            return False
        
        # Confirm action
        file_count = len(self.file_paths)
        case_sensitive = self.case_sensitive_var.get()
        case_info = " (case sensitive)" if case_sensitive else " (case insensitive)"
        action_text = "replace all occurrences and close the application" if close_after else "replace all occurrences"
        if not messagebox.askyesno("Confirm", 
                                  f"This will {action_text} in {file_count} file(s){case_info}.\n\nContinue?"):
            return False
        
        try:
            self.status_label.config(text="Processing files...", fg="orange")
            self.progress_frame.pack(fill=tk.X, pady=(10, 0))
            self.root.update()
            
            total_replacements = 0
            successful_files = 0
            backup_files = []
            all_results = []  # NEW: Store detailed results
            
            for i, file_path in enumerate(self.file_paths):
                # Update progress with percentage
                progress_percent = int((i / file_count) * 100)
                self.progress_label.config(text=f"Processing {i+1}/{file_count}: {os.path.basename(file_path)}")
                self.progress_percent_label.config(text=f"Progress: {progress_percent}%")
                self.root.update()
                
                try:
                    # Create backup if requested
                    if self.create_backup_var.get():
                        backup_path = file_path + ".backup"
                        import shutil
                        shutil.copy2(file_path, backup_path)
                        backup_files.append(backup_path)
                    
                    # Process the document
                    doc = Document(file_path)
                    file_replacements = 0
                    
                    # Replace in paragraphs
                    for paragraph in doc.paragraphs:
                        file_replacements += self.replace_in_paragraph_advanced(paragraph, search_for, replace_with, case_sensitive)
                    
                    # Replace in tables
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for paragraph in cell.paragraphs:
                                    file_replacements += self.replace_in_paragraph_advanced(paragraph, search_for, replace_with, case_sensitive)
                    
                    # Save the document
                    doc.save(file_path)
                    
                    # Store detailed results
                    all_results.append({
                        'filename': os.path.basename(file_path),
                        'total': file_replacements,
                        'details': [f"  📄 Main content replacements: {file_replacements}"] if file_replacements > 0 else ["  ✅ No replacements needed"]
                    })
                    
                    total_replacements += file_replacements
                    successful_files += 1
                    
                except Exception as e:
                    all_results.append({
                        'filename': os.path.basename(file_path),
                        'total': 0,
                        'details': [f"  ❌ Error: {str(e)}"]
                    })
            
            # Final progress update
            self.progress_percent_label.config(text="Progress: 100%")
            self.root.update()
            
            self.progress_frame.pack_forget()
            self.status_label.config(text="Replacement completed!", fg="green")
            
            # Show enhanced results instead of simple messagebox
            self.show_replace_all_results(all_results, search_for, replace_with, total_replacements, 
                                         successful_files, backup_files, close_after)
            
            return True
            
        except Exception as e:
            self.status_label.config(text="Error occurred", fg="red")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            return False

    def show_replace_all_results(self, results, search_text, replace_text, total_replacements, successful_files, backup_files, close_after):
        """Show Replace results in scrollable window"""
        case_sensitive = self.case_sensitive_var.get()
        case_info = " (case sensitive)" if case_sensitive else " (case insensitive)"
        
        message = "-" * 46 + "\n"
        message += f"✅ Standard Replace results{case_info}\n"
        message += "-" * 46 + "\n\n" 
        message += f"🎉 Completed: {total_replacements} replacement(s) in main document content\n\n"
        message += f"📊 Summary:\n"
        message += f"Files processed: {successful_files}/{len(self.file_paths)}\n"
        message += f"Total replacements: {total_replacements}\n"
        message += f"Search text: '{search_text[:40]}{'...' if len(search_text) > 40 else ''}'\n"
        message += f"Replace text: '{replace_text[:40]}{'...' if len(replace_text) > 40 else ''}'\n"
        message += f"Coverage: Main text and tables\n"
        
        if backup_files:
            message += f"Backup files created: {len(backup_files)}\n"
        
        message += "=" * 77 + "\n\n"
        
        # Show detailed file results
        for i, result in enumerate(results, 1):
            message += f"📁 FILE {i}: {result['filename']}\n"
            if result['total'] > 0:
                message += f"   ✅ Replacements made: {result['total']}\n"
            else:
                message += f"   ✅ No replacements needed\n"
            
            for detail in result['details']:
                message += f"   {detail}\n"
            message += "\n" + "=" * 77 + "\n\n"
        
        if close_after:
            message += "🚪 The application will close after you click OK.\n\n"
        
        # Show results and handle close_after
        if close_after:
            # For close_after, show results then close
            result_window = tk.Toplevel(self.root)
            result_window.title("Standard Replace results")
            result_window.geometry("600x500")
            result_window.resizable(True, True)
            
            text_widget = scrolledtext.ScrolledText(result_window, wrap=tk.WORD, 
                                                   font=("Consolas", 9), padx=10, pady=10)
            text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            text_widget.insert("1.0", message)
            text_widget.config(state=tk.DISABLED)
            
            def close_and_exit():
                result_window.destroy()
                self.root.destroy()
            
            close_btn = tk.Button(result_window, text="Close application", 
                                 command=close_and_exit,
                                 bg="#333333", fg="white", font=("Arial", 10, "bold"))
            close_btn.pack(pady=(0, 10))
            
            result_window.transient(self.root)
            result_window.grab_set()
        else:
            # Normal scrollable results
            self.show_scrollable_results(message, "Standard Replace results")

    # NEW: Advanced Replace with VBScript Integration + Hidden CMD Window 🚀
    def advanced_replace_with_vba(self):
        """Advanced replace using VBScript for ALL document areas - WITH HIDDEN CMD"""
        if not self.file_paths:
            messagebox.showwarning("Warning", "Please add at least one Word document.")
            return
            
        search_for = self.get_processed_search_text()    # CHANGED: Use processed text
        replace_with = self.get_processed_replace_text() # CHANGED: Use processed text
        
        if not search_for:
            messagebox.showwarning("Warning", "Please enter text to search for.")
            return
        
        # Confirm action
        file_count = len(self.file_paths)
        case_sensitive = self.case_sensitive_var.get()
        case_info = " (case sensitive)" if case_sensitive else " (case insensitive)"
        if not messagebox.askyesno("Confirm Advanced Replace", 
                                  f"This will replace text in all document areas in {file_count} file(s){case_info}.\n\n" +
                                  f"Areas covered: Main content, shapes, headers, footers, footnotes, endnotes, form fields, hyperlinks\n\n" +
                                  f"Search: '{search_for[:30]}{'...' if len(search_for) > 30 else ''}'\n" +
                                  f"Replace: '{replace_with[:30]}{'...' if len(replace_with) > 30 else ''}'\n\n" +
                                  "Continue?"):
            return
        
        try:
            # Get the directory where the Python script is located
            script_dir = os.path.dirname(os.path.abspath(__file__))
            vbs_path = os.path.join(script_dir, "word_advanced_replace.vbs")
            
            if not os.path.exists(vbs_path):
                messagebox.showerror("Error", f"VBScript not found: {vbs_path}\n\nPlease make sure word_advanced_replace.vbs is in the same folder as this Python script.")
                return
            
            self.status_label.config(text="Processing advanced replacements...", fg="orange")
            self.progress_frame.pack(fill=tk.X, pady=(10, 0))
            self.root.update()
            
            all_results = []
            total_advanced_replacements = 0
            successful_files = 0
            errors = []
            
            for i, file_path in enumerate(self.file_paths):
                # Update progress
                progress_percent = int((i / file_count) * 100)
                self.progress_label.config(text=f"Processing {i+1}/{file_count}: {os.path.basename(file_path)}")
                self.progress_percent_label.config(text=f"Progress: {progress_percent}%")
                self.root.update()
                
                try:
                    # Create backup if requested
                    if self.create_backup_var.get():
                        backup_path = file_path + ".backup"
                        import shutil
                        shutil.copy2(file_path, backup_path)
                    
                    # Call VBScript with HIDDEN CMD WINDOW
                    full_path = os.path.abspath(file_path)
                    
                    # HIDE CMD WINDOW - This is the key improvement!
                    startupinfo = subprocess.STARTUPINFO()
                    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    startupinfo.wShowWindow = subprocess.SW_HIDE

                    # Pass case sensitivity to VBScript
                    case_flag = "1" if case_sensitive else "0"
                    
                    result = subprocess.run([
                        'cscript', '//NoLogo', vbs_path, full_path, search_for, replace_with, case_flag
                    ], capture_output=True, text=True, cwd=script_dir,
                       startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    if result.returncode == 0:
                        # Parse successful output - ENHANCED to include all areas
                        output_lines = result.stdout.strip().split('\n')
                        file_replacements = 0
                        details = []
                        
                        for line in output_lines:
                            if "SHAPES:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📦 Shapes: {count} replacement(s)")
                                file_replacements += count
                            elif "HEADERS:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📄 Headers: {count} replacement(s)")
                                file_replacements += count
                            elif "FOOTERS:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📄 Footers: {count} replacement(s)")
                                file_replacements += count
                            elif "FOOTNOTES:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📝 Footnotes: {count} replacement(s)")
                                file_replacements += count
                            elif "ENDNOTES:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📝 Endnotes: {count} replacement(s)")
                                file_replacements += count
                            elif "FORMFIELDS:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📋 Form fields: {count} replacement(s)")
                                file_replacements += count
                            elif "HYPERLINKS:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  🔗 Hyperlinks: {count} replacement(s)")
                                file_replacements += count
                            elif "MAINCONTENT:" in line:
                                count = int(line.split(':')[1].split()[0])
                                details.append(f"  📄 Main content: {count} replacement(s)")
                                file_replacements += count               
                            elif "RESULT:" in line:
                                # Get total from result line as backup
                                total_from_vbs = int(line.split(':')[1].split()[0])
                                if file_replacements == 0:  # If we didn't parse individual counts
                                    file_replacements = total_from_vbs
                        
                        all_results.append({
                            'filename': os.path.basename(file_path),
                            'total': file_replacements,
                            'details': details if details else ["  ✅ No advanced matches found"]
                        })
                        
                        total_advanced_replacements += file_replacements
                        successful_files += 1
                        
                    else:
                        # Handle errors
                        error_msg = result.stderr.strip() if result.stderr else "Unknown VBScript error"
                        all_results.append({
                            'filename': os.path.basename(file_path),
                            'total': 0,
                            'details': [f"  ❌ Error: {error_msg}"]
                        })
                        errors.append(f"{os.path.basename(file_path)}: {error_msg}")
                        
                except Exception as e:
                    all_results.append({
                        'filename': os.path.basename(file_path),
                        'total': 0,
                        'details': [f"  ❌ Error: {str(e)}"]
                    })
                    errors.append(f"{os.path.basename(file_path)}: {str(e)}")
            
            # Final progress update
            self.progress_percent_label.config(text="Progress: 100%")
            self.root.update()
            
            self.progress_frame.pack_forget()
            self.status_label.config(text="Advanced replacement completed!", fg="green")
            
            # Show results in scrollable window
            self.show_advanced_replace_results(all_results, search_for, replace_with, total_advanced_replacements, successful_files, errors)
            
        except Exception as e:
            self.status_label.config(text="Error occurred", fg="red")
            messagebox.showerror("Error", f"Advanced replace error: {str(e)}")

    def show_advanced_replace_results(self, results, search_text, replace_text, total_replacements, successful_files, errors):
        """Show advanced replace results in scrollable window"""
        case_sensitive = self.case_sensitive_var.get()
        case_info = " (case sensitive)" if case_sensitive else " (case insensitive)"
        
        message = "-" * 46 + "\n"
        message += f"🔧 Advanced Replace results{case_info}\n"
        message += "-" * 46 + "\n\n"
        message += f"✅ Completed: {total_replacements} replacement(s) in ALL document areas\n\n"
        message += f"📊 Summary:\n"
        message += f"Files processed: {successful_files}/{len(self.file_paths)}\n"
        message += f"Total advanced replacements: {total_replacements}\n"
        message += f"Search text: '{search_text[:40]}{'...' if len(search_text) > 40 else ''}'\n"
        message += f"Replace text: '{replace_text[:40]}{'...' if len(replace_text) > 40 else ''}'\n"
        message += f"Coverage: Complete document (main content + all advanced areas)\n"
        message += "=" * 77 + "\n\n"
        
        # Show detailed file results
        for i, result in enumerate(results, 1):
            message += f"📁 FILE {i}: {result['filename']}\n"
            if result['total'] > 0:
                message += f"   ✅ Advanced replacements: {result['total']}\n"
            else:
                message += f"   ✅ No advanced matches found\n"
            
            for detail in result['details']:
                message += f"   {detail}\n"
            message += "\n" + "=" * 77 + "\n\n"
        
        if errors:
            message += f"❌ Errors:\n"
            for error in errors:
                message += f"• {error}\n"
            message += "\n"
        
        self.show_scrollable_results(message, "Advanced Replace results")
    
    def show_scrollable_results(self, message, title):
        """Show results in a scrollable window for long messages"""
        result_window = tk.Toplevel(self.root)
        result_window.title(title)
        result_window.geometry("600x500")
        result_window.resizable(True, True)
        
        # Create scrolled text widget
        text_widget = scrolledtext.ScrolledText(result_window, wrap=tk.WORD, 
                                               font=("Consolas", 9), padx=10, pady=10)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Insert the message
        text_widget.insert("1.0", message)
        text_widget.config(state=tk.DISABLED)  # Make it read-only
        
        # Button frame for two buttons
        button_frame = tk.Frame(result_window)
        button_frame.pack(pady=(0, 10))
        
        # OK button (just closes result window)
        ok_btn = tk.Button(button_frame, text="OK", 
                          command=result_window.destroy,
                          bg="#4caf50", fg="white", font=("Arial", 10))
        ok_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Close application button (closes result window AND main app)
        def close_application():
            result_window.destroy()
            self.root.destroy()
        
        close_app_btn = tk.Button(button_frame, text="Close application", 
                                 command=close_application,
                                 bg="#333333", fg="white", font=("Arial", 10))
        close_app_btn.pack(side=tk.LEFT)
        
        # Center the window
        result_window.transient(self.root)
        result_window.grab_set()
#########END OF PART 3B######




#########PART 4 (FINAL)######
    def run(self):
        self.root.mainloop()

def main():
    # Get the initial file from command line (if any)
    initial_file = sys.argv[1] if len(sys.argv) > 1 else None
    
    # Validate the initial file
    if initial_file and not os.path.exists(initial_file):
        initial_file = None
    
    if initial_file and not initial_file.lower().endswith(('.docx', '.doc', '.docm')):
        initial_file = None
    
    try:
        app = WordTextReplacerSingle(initial_file)
        app.run()
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start application: {str(e)}")

if __name__ == "__main__":
    main()

#########END OF PART 4 (FINAL)######
