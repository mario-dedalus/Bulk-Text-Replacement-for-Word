# macOS PyInstaller specification for generating standalone .app bundle
# Build on macOS with:
#   pyinstaller --clean --noconfirm WordTextReplacer_mac.spec

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = (
    collect_submodules("openpyxl")
    + collect_submodules("docx")
    + [
        "PySide6.QtCore", "PySide6.QtGui", "PySide6.QtWidgets",
        "core", "core.models", "core.replacer_core", "core.template_merge", "core.multi_doc_replacer",
        "platform_adapter", "platform_adapter.capabilities", "platform_adapter.appearance", "platform_adapter.file_manager",
        "application", "application.task_models", "application.replace_service", "application.merge_service", "application.multi_doc_service", "application.workers",
        "ui", "ui.theme", "ui.theme.tokens", "ui.theme.theme_manager",
        "ui.widgets", "ui.widgets.card", "ui.widgets.status_bar", "ui.widgets.segmented_nav",
        "ui.models", "ui.models.file_list_model", "ui.models.field_mapping_model", "ui.models.multi_doc_mapping_model",
        "ui.delegates", "ui.delegates.mapping_combo_delegate",
        "ui.dialogs", "ui.dialogs.result_dialog", "ui.dialogs.help_dialog",
        "ui.pages", "ui.pages.replace_page", "ui.pages.merge_page", "ui.pages.multi_doc_page",
        "ui.main_window", "replacer_core", "template_merge", "platform_capabilities",
    ]
)

datas = [
    ("src/ui/theme/windows.qss", "ui/theme"),
    ("src/ui/theme/macos.qss", "ui/theme"),
]

a = Analysis(
    ["src/app.py"],
    pathex=["src"],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PySide6.QtNetwork", "PySide6.QtQml", "PySide6.QtQuick",
        "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.Qt3DCore",
        "PySide6.QtMultimedia", "PySide6.QtBluetooth", "PySide6.QtSensors",
        "tkinter",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="WordTextReplacer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="WordTextReplacer",
)

app = BUNDLE(
    coll,
    name="WordTextReplacer.app",
    icon=None,
    bundle_identifier="com.bulkreplace.wordtextreplacer",
    info_plist={
        "CFBundleShortVersionString": "2.0.0",
        "CFBundleVersion": "2.0.0",
        "NSHighResolutionCapable": "True",
        "CFBundleDisplayName": "WordTextReplacer",
        "LSApplicationCategoryType": "public.app-category.utilities",
    },
)
