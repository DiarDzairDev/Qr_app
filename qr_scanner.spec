# -*- mode: python ; coding: utf-8 -*-
import sys
import os

a = Analysis(
    ['qr_scanner.py'],
    pathex=[],
    binaries=[],
    datas=[('qrcodescan.ico', '.'), ('logo_diardzair.jpg', '.')],
    hiddenimports=[
        'PIL', 'PIL._tkinter_finder', 'tkinter', 'tkinter.ttk', 
        'pandas', 'openpyxl', 'requests', 'qrcode', 
        'win32print', 'win32ui', 'win32con',
        'tkinter.messagebox', 'tkinter.filedialog',
        'PIL.Image', 'PIL.ImageTk', 'PIL.ImageDraw',
        'dataclasses', 'datetime', 'threading',
        'subprocess', 'shutil', 'logging'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'numpy.distutils', 'test', 'unittest'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Mouvement Stock',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['qrcodescan.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='Mouvement Stock'
)
