# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置 - 桌面版
解决打包后找不到 index.html 的问题
"""

import os
import sys
from pathlib import Path

block_cipher = None
SPEC_DIR = Path(SPECPATH)
BASE_DIR = SPEC_DIR

# 打包时需要包含的文件
datas = [
    (str(BASE_DIR / 'index.html'), '.'),
    (str(BASE_DIR / '2026年启用报销与费用填写.xlsx'), '.'),
]

a = Analysis(
    ['server_cloud.py'],
    pathex=[str(BASE_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        # openpyxl - 让PyInstaller自动处理
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.drawing',
        'openpyxl.drawing.image',
        # PIL - 让PyInstaller自动处理
        'PIL',
        'PIL.Image',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
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
    name='报销工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='报销工具',
)