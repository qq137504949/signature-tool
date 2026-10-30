# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec：单文件 GUI exe，内置 fonts 目录
# 构建: pyinstaller signature_tool.spec --noconfirm --clean

a = Analysis(
    ['signature_tool.py'],
    pathex=[],
    binaries=[],
    datas=[('fonts', 'fonts')],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        'numpy', 'pandas', 'matplotlib', 'scipy', 'pytest',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'IPython', 'notebook',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SignatureTool',
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
    icon='assets/icon.ico',
)
