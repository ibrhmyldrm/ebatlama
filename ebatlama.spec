# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import os

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        *collect_data_files('webview'),
        *collect_data_files('openpyxl'),
        *collect_data_files('reportlab'),
    ],
    hiddenimports=[
        # webview
        'webview',
        'webview.platforms.winforms',
        # tkinter
        'tkinter',
        'tkinter.filedialog',
        'tkinter.simpledialog',
        '_tkinter',
        # flask
        'flask',
        'jinja2',
        'werkzeug',
        'werkzeug.serving',
        'werkzeug.routing',
        # excel
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.drawing.image',
        'openpyxl.worksheet.page',
        # pdf
        'reportlab',
        'reportlab.lib',
        'reportlab.platypus',
        'reportlab.pdfbase',
        'reportlab.pdfbase.ttfonts',
        'reportlab.graphics.shapes',
        'reportlab.graphics',
        # pillow
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        # other
        'clr',
        'pythonnet',
        'pkg_resources',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AtolyemHanem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
