# Bulk Text Replacement for Word

🎯 Modern, cross-platform desktop application for batch text replacement and mail-merge document generation across Microsoft Word documents. Built with **PySide6** and **Windows 11 Fluent Design / macOS Native styling**.

---

## 🚀 Quick Start

### Option 1: Run Pre-built Executable
- **Windows**: Download `WordTextReplacer.exe` from GitHub Releases / Actions artifacts. No installation required.
- **macOS**: Download `WordTextReplacer.app` from GitHub Releases / Actions artifacts (currently available via Actions developer build artifacts).

### Option 2: Run from Source

```bash
# Clone the repository
git clone https://github.com/chaos-lawyer/Bulk-Text-Replacement-for-Word.git
cd Bulk-Text-Replacement-for-Word

# Set up Python environment (Python 3.10 - 3.13 recommended)
python3 -m venv .venv

# On macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
python src/app.py

# On Windows:
.venv\Scripts\activate
pip install -r requirements.txt
python src\app.py
```

---

## 🧩 Key Features

### 1. 🔍 Batch Text Replacement (`文本查找替换`)
- **Fast Mode (Recommended)**: Powered by `python-docx`. Blazing fast, preserves run formatting, supports Python regular expressions, and works seamlessly across macOS, Windows, and Linux.
- **Full Mode (Windows Word COM)**: Powered by Microsoft Word COM automation (`pywin32`). Comprehensively processes headers, footers, shapes, text boxes, and preserves existing hyperlinks.
- **Real-Time Match Counter**: Live match count updates asynchronously with generational race protection.
- **Preview with Context**: Inspect replacements and surrounding snippets before modifying any files on disk.
- **Hyperlink Inspector**: Asynchronously scan and inspect all embedded URLs in selected documents with cancellation support.
- **Non-Breaking Spaces (NBSP)**: Automatically cleans soft hyphens, zero-width spaces, and normalizes `\u00a0`.

### 2. 📑 Template Batch Merge (`模板批量生成`)
- **Strict 3-Step Guided Workflow**:
  - **Step 1**: Choose Word template (`.docx` / `.docm`), Excel or CSV data source (`.xlsx` / `.csv` / `.xlsm`), output folder, and dynamic filename rule (e.g. `{{Client}}-{{ContractNo}}.docx` or `{{序号}}-{{Client}}.docx`). Source files are verified on disk before scanning.
  - **Step 2**: Interactive 4-column field mapping table with dropdown selection (`QTableView` + `ComboBoxDelegate`) mapping `{{Variable}}` to table columns, plus editable empty-field default fallbacks. Changing output folder/filename rule preserves existing field mappings.
  - **Step 3**: Instant preview of first 5 rows and asynchronous batch generation with real-time log and platform fixed-width typography.
- **Safe & Non-Destructive**: Never alters template files and automatically avoids duplicate output filename collisions.

### 3. 📑 Multi-Document Differentiated Replacement (`多文档匹配替换`)
- **1-to-1 Document-to-Row Data Alignment**: Bind multiple distinct Word documents with multiple rows of variable data to fill individual variables across documents simultaneously.
- **Three Data Entry Modes**:
  - **Pure Manual Entry**: Type replacement values directly into table cells without needing an Excel file.
  - **Excel / CSV Smart Alignment**: Import table data with automatic column header matching or sequential row filling.
  - **In-Place Manual Override**: Cells remain editable after Excel import to make instant adjustments without modifying the original spreadsheet.
- **Flexible Output & Backup**: Supports modifying files in place (with automatic `.backup` creation) or exporting to a designated output folder.

### 4. 🎨 Modern Cross-Platform UI & Architecture
- **Windows 11 Fluent & macOS Native Design**: 8px card surfaces, fixed bottom action bar for persistent 100% visibility, single primary accent button per view, dynamic system theme synchronization, and light/dark theme toggle.
- **Non-Blocking Background Workers**: All file parsing, replacements, and generations run asynchronously on `QThreadPool` with signature pre-inspection, progress reporting, cancel support, and clean graceful window closing.
- **Accessibility (a11y)**: Built-in accessible names, descriptions, and keyboard arrow navigation across tabs.

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|---|---|
| `Ctrl+O` / `Cmd+O` | Add Word documents / browse template |
| `Ctrl+R` / `Cmd+R` | Execute Replace / Start Batch Merge / Start Multi-Doc Replace |
| `Ctrl+P` / `Cmd+P` | Preview changes |
| `Delete` | Remove selected document(s) from list |
| `Left` / `Right` Arrow | Switch between Navigation Tabs |
| `F1` | Open User Guide & Shortcuts dialog |

---

## 🏗️ Project Architecture

```text
Bulk-Text-Replacement-for-Word/
├── src/
│   ├── app.py                         # Primary application entry point (PySide6)
│   ├── core/                          # Pure Python core engines (Zero GUI dependencies)
│   │   ├── replacer_core.py           # Text search, replacement & regex engine
│   │   ├── template_merge.py          # Word template + Excel/CSV mail merge engine
│   │   ├── multi_doc_replacer.py      # Multi-doc variable scan & differentiated replace engine
│   │   └── models.py                  # Core data classes (ExcelData, MultiDocItem, etc.)
│   ├── platform_adapter/              # OS capability, theme & file manager integration
│   │   ├── capabilities.py            # OS & Word COM capability detection
│   │   ├── appearance.py              # System dark mode & DWM accent color
│   │   └── file_manager.py            # Explorer / Finder file revelation
│   ├── application/                   # Services, task models & async workers
│   │   ├── replace_service.py         # Text replacement orchestration
│   │   ├── merge_service.py           # Template merge orchestration
│   │   ├── multi_doc_service.py       # Multi-doc replacement orchestration
│   │   ├── task_models.py             # TaskProgress, TaskState, CancellationToken
│   │   └── workers.py                 # QThreadPool / QRunnable background workers
│   └── ui/                            # PySide6 UI layer
│       ├── main_window.py             # QMainWindow with top navigation & status bar
│       ├── pages/                     # ReplacePage, MergePage & MultiDocPage
│       ├── models/                    # FileListModel, FieldMappingModel, MultiDocMappingModel
│       ├── delegates/                 # MappingComboDelegate
│       ├── widgets/                   # FluentCard, SegmentedNav, FluentStatusBar
│       └── theme/                     # Token definitions & QSS stylesheets
├── tests/                             # Unit & integration test suite (54 automated tests)

├── WordTextReplacer.spec              # Windows PyInstaller standalone spec
├── WordTextReplacer_mac.spec          # macOS PyInstaller .app bundle spec
└── requirements.txt                   # Dependency definitions
```

---

## 📦 Building Standalone Packages

### Windows (.exe)
```bat
.venv\Scripts\pip install -r requirements.txt
.venv\Scripts\pyinstaller --clean --noconfirm WordTextReplacer.spec
# Output: dist/WordTextReplacer.exe
```

### macOS (.app)
```bash
./build_macos.sh
# Output: dist/WordTextReplacer.app
```

---

## 🧪 Testing

Run the full automated test suite:

```bash
python -m unittest discover -s tests -v
```

---

## 📄 License

Source-Available License (Apache 2.0 + No Selling clause). See [LICENSE](LICENSE) for details.
