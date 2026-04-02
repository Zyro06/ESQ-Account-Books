# accounting.spec
import os
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# ── Collect qt_material and qtawesome assets ──────────────────────────────
qt_material_datas = collect_data_files('qt_material')
qtawesome_datas   = collect_data_files('qtawesome')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('resources', 'resources'),   # style.qss, user_manual.html, etc.
        *qt_material_datas,
        *qtawesome_datas,
    ],
    hiddenimports=[
        # PySide6
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'PySide6.QtPrintSupport',
        'PySide6.QtSvg',
        'PySide6.QtXml',
        # qt_material / qtawesome
        'qt_material',
        'qtawesome',
        # openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.drawing.image',
        # reportlab (PDF export)
        'reportlab',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.units',
        'reportlab.lib.styles',
        'reportlab.lib.colors',
        'reportlab.lib.enums',
        'reportlab.platypus',
        'reportlab.platypus.tables',
        # stdlib (sometimes missed)
        'calendar',
        'sqlite3',
        'json',
        'collections',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ESQ Accounting System',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # ← no console window
    # icon='resources/icon.ico',  # uncomment when you have an icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ESQ Accounting System',
)
