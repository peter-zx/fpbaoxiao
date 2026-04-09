# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置
报销费用填写工具 - 桌面版
"""

import os
import sys
from pathlib import Path

block_cipher = None

# 获取当前目录（spec文件所在目录）
SPEC_DIR = Path(SPECPATH)
BASE_DIR = SPEC_DIR

# 数据文件（运行时需要）
datas = [
    (str(BASE_DIR / 'index.html'), '.'),
    (str(BASE_DIR / 'data.json'), '.'),
    (str(BASE_DIR / 'config.json'), '.'),
    (str(BASE_DIR / '2026年启用报销与费用填写.xlsx'), '.'),
    (str(BASE_DIR / 'images'), 'images'),
    (str(BASE_DIR / 'exports'), 'exports'),
]

a = Analysis(
    ['server_cloud.py'],
    pathex=[str(BASE_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.drawing',
        'openpyxl.drawing.image',
        'openpyxl.utils',
        'PIL',
        'PIL.Image',
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
    name='报销工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 显示控制台窗口，方便查看日志
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可添加 icon.ico
)