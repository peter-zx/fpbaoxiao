# -*- coding: utf-8 -*-
"""
main.py — 报销费用填写工具主入口
职责：初始化配置、日志、环境，启动 HTTP 服务器
"""

import os
import sys
import json
import logging
import shutil
import socket
from pathlib import Path

# ============================================================================
# Windows 终端 UTF-8 输出兼容
# ============================================================================
if sys.platform == 'win32':
    os.system('')
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except AttributeError:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'replace')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'replace')

# ============================================================================
# 路径解析（兼容打包后的 exe）
# ============================================================================
def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)
    return Path(__file__).parent.resolve()

def _get_data_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent / 'data'
    return _get_base_dir()

BASE_DIR = _get_base_dir()
DATA_DIR = _get_data_dir()

# ============================================================================
# 配置管理
# ============================================================================
DEFAULT_CONFIG = {
    "port": 8765,
    "port_range_start": 8765,
    "port_range_max": 100,
    "cors_origins": ["*"],
    "log_level": "INFO",
    "max_content_length": 20971520,
    "auto_open_browser": True,
    "single_instance": True,
}


def _migrate_config():
    cfg_file = DATA_DIR / 'config.json'
    if not cfg_file.exists():
        cfg_file.write_text(json.dumps(DEFAULT_CONFIG, ensure_ascii=False, indent=2),
                           encoding='utf-8')
        return

    try:
        raw = json.loads(cfg_file.read_text(encoding='utf-8'))
    except Exception:
        return

    if 'server' in raw or 'logging' in raw:
        return  # 已是新版格式

    migrated = dict(DEFAULT_CONFIG)
    for key in ('port', 'port_range_start', 'port_range_max', 'cors_origins',
                'log_level', 'max_content_length', 'auto_open_browser', 'single_instance'):
        if key in raw:
            migrated[key] = raw[key]

    cfg_file.write_text(json.dumps(migrated, ensure_ascii=False, indent=2),
                       encoding='utf-8')
    print("config.json 已迁移到新版格式")


def _load_config():
    _migrate_config()
    cfg_file = DATA_DIR / 'config.json'
    try:
        raw = json.loads(cfg_file.read_text(encoding='utf-8'))
    except Exception:
        raw = {}
    config = dict(DEFAULT_CONFIG)
    for k, v in raw.items():
        if k in DEFAULT_CONFIG:
            config[k] = v
    return config


# ============================================================================
# 日志初始化
# ============================================================================
def _init_logging(log_file, log_level):
    log_format = '%(asctime)s [%(levelname)s] %(message)s'
    handlers = [
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
    logging.basicConfig(
        level=getattr(logging, log_level, logging.INFO),
        format=log_format,
        handlers=handlers,
    )
    for lib in ('PIL', 'PIL.Image', 'win32com', 'win32com.client'):
        logging.getLogger(lib).setLevel(logging.WARNING)


# ============================================================================
# 静态文件目录
# ============================================================================
def _get_static_dir():
    s1 = BASE_DIR / 'static'
    if s1.exists() and (s1 / 'index.html').exists():
        return s1
    s2 = BASE_DIR / 'index.html'
    if s2.exists():
        return BASE_DIR
    return s1


# ============================================================================
# 横幅
# ============================================================================
def _get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _print_banner(port, data_dir, excel_info):
    H = '=' * 50
    print()
    print(f'  {H}')
    print(f'  |{" " * 48}|')
    print(f'  |    {"报销费用填写工具":<42} |')
    print(f'  |    Expense & Reimbursement Tool              |')
    print(f'  |{" " * 48}|')
    print(f'  {H}')
    print()
    print(f'    Author : aigc创意人竹相左边')
    print(f'    Engine : {excel_info.get("tool", "unknown")} / xlsxwriter')
    print()
    print(f'    {"-" * 48}')
    print(f'    Status : Running')
    print(f'    Local  : http://localhost:{port}')
    print(f'    LAN    : http://{_get_local_ip()}:{port}')
    print(f'    Data   : {data_dir}')
    print(f'    {"-" * 48}')
    print()
    print(f'    Stop   : Ctrl+C')
    print()


def _cleanup_old_images():
    old_img_dir = BASE_DIR / 'images'
    if old_img_dir.exists() and old_img_dir.is_dir():
        try:
            shutil.rmtree(old_img_dir)
            logging.info(f"已清理旧版遗留目录: {old_img_dir}")
        except Exception:
            pass


# ============================================================================
# 主入口
# ============================================================================
def main():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / 'exports').mkdir(exist_ok=True)

    config = _load_config()
    log_file = DATA_DIR / 'server.log'
    _init_logging(log_file, config.get('log_level', 'INFO'))
    logging.info(f"启动中... (env={'frozen' if getattr(sys,'frozen',False) else 'dev'})")

    _cleanup_old_images()

    from app.excel_export import detect_office_tool, get_excel_creator_info
    detect_office_tool()
    excel_info = get_excel_creator_info()
    logging.info(f"Excel 生成器: {excel_info}")

    # 依赖检查（无交互，缺依赖直接退出）
    try:
        import xlsxwriter
    except ImportError:
        print("\n缺少 xlsxwriter，请运行: pip install xlsxwriter")
        sys.exit(1)

    try:
        from PIL import Image
    except ImportError:
        print("\n缺少 Pillow，请运行: pip install Pillow")
        sys.exit(1)

    # 单实例检测
    if config.get('single_instance', True):
        from app.server import check_single_instance
        lock_file = DATA_DIR / 'server.lock'
        check_single_instance(lock_file)

    # 端口检测
    start_port = config.get('port', 8765)
    max_range  = config.get('port_range_max', 100)
    from app.server import find_available_port
    port = find_available_port(start_port, max_range)
    if port is None:
        print("无法找到可用端口")
        sys.exit(1)
    if port != start_port:
        print(f"端口 {start_port} 已被占用，自动使用端口 {port}")

    # 业务层
    from app.store import load, save
    from app.server import Server

    class DataStore:
        def load(self): return load(DATA_DIR / 'data.json')
        def save(self, data): return save(DATA_DIR / 'data.json', data)

    class ExcelFactory:
        def create(self, data, output_dir):
            from app.excel_export import create_excel
            return create_excel(data, output_dir)

    # 启动服务器
    static_dir = _get_static_dir()
    server = Server(
        static_dir=static_dir,
        output_dir=DATA_DIR / 'exports',
        data_store=DataStore(),
        excel_factory=ExcelFactory(),
        port=port,
    )

    _print_banner(port, DATA_DIR, excel_info)

    server.start(
        auto_open_browser=config.get('auto_open_browser', True),
        local_url=f'http://localhost:{port}',
    )


if __name__ == '__main__':
    main()
