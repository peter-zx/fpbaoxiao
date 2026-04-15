# -*- coding: utf-8 -*-
"""
server_cloud.py — 兼容层
所有逻辑已迁移到 app/ 模块；本文件保留向后兼容入口。
用户仍可通过 python server_cloud.py 启动（start.bat 无需修改）。
"""

import sys
from pathlib import Path

# 确保 app/ 包可被导入
_root = Path(__file__).parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from app.excel_export import (
    detect_office_tool,
    get_excel_creator_info,
    create_excel_xlsxwriter,
    create_excel_com,
    create_excel,
)
from app.store import load as _store_load, save as _store_save

# ============================================================================
# 旧版全局变量（兼容 start.bat 中的检测逻辑 / 外部调用者）
# ============================================================================
HAS_XLSXWRITER = True   # 新架构固定走 xlsxwriter，不再检测
HAS_OPENPYXL   = True
HAS_PIL         = True
HAS_WIN32COM    = False  # 由 excel_export.detect_office_tool() 动态检测
OFFICE_TYPE     = None   # 同上

# ============================================================================
# 旧版函数别名（兼容外部调用）
# ============================================================================
def load_data():
    return _store_load(_root / 'data.json')

def save_data(data):
    return _store_save(_root / 'data.json', data)

def create_excel_with_images(data):
    from app.excel_export import create_excel
    return create_excel(data, _root / 'exports')

def create_excel_with_com(data):
    from app.excel_export import create_excel_com
    return create_excel_com(data, _root / 'exports')

def detect_office_tool():
    from app.excel_export import detect_office_tool as _detect
    return _detect()

def get_excel_creator_info():
    from app.excel_export import get_excel_creator_info as _info
    return _info()

# ============================================================================
# 入口：python server_cloud.py
# ============================================================================
if __name__ == '__main__':
    from main import main
    main()
