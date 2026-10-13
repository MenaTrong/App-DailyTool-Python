# -*- mode: python ; coding: utf-8 -*-
"""Đóng DailyTool thành một file DailyTool.exe cho Windows."""

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules("openpyxl") + ["et_xmlfile"]

a = Analysis(
    ["run_dailytool.py"],
    pathex=["."],
    binaries=[],
    datas=[("dailytool/ma_bac_si_seed.json", "dailytool")],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="DailyTool",
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
)
