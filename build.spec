# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller specification file for Header Hunter v8.0

Build command:
    pyinstaller build.spec

Result:
    - Standalone executable in dist/HeaderHunter/
    - No Python installation required on user's machine
"""

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# Collect CustomTkinter data files (themes, assets, etc.)
customtkinter_datas = collect_data_files('customtkinter')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('header_hunter_config.json', '.'),  # Include default config
        ('icon.ico', '.'),                    # App icon
        ('logo.png', '.'),                    # Logo for GUI
    ] + customtkinter_datas,
    hiddenimports=[
        'customtkinter',
        'pandas',
        'xlsxwriter',
        'openpyxl',
        'numpy',
        'tkinterdnd2',  # Optional, but include if available
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # STRIP PROBLEM BINARIES - these cause UPX decompression errors
    excludedimports=['PIL._avif', 'PIL._webp', 'tkinter.test'],
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
    name='HeaderHunter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,              # DISABLED: Prevents 'return code -3' decompression errors
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI-only - no console window
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Application icon
)

