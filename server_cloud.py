# -*- coding: utf-8 -*-
"""
报销费用填写工具 - 桌面版 (server_cloud.py)
===============================================
为打包优化版本：
  1. 端口自动检测（8765被占用自动换下一个）
  2. 单实例检测（防止重复启动）
  3. 自动打开浏览器
  4. 启动时显示二维码（手机扫码）
"""

import os
import sys
import json
import base64
import io
import logging
import signal
import socket
import webbrowser
import threading
import time

# Windows 终端 UTF-8 输出兼容
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, quote
from pathlib import Path

# ---- 依赖检查 ----
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("❌ 缺少 openpyxl，请运行: pip install openpyxl Pillow")

try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    print("❌ 缺少 Pillow，请运行: pip install Pillow")


# ==================== 打包兼容处理 ====================
def get_base_dir():
    """获取运行目录（兼容打包后的exe）"""
    if getattr(sys, 'frozen', False):
        # 打包后的exe
        return Path(sys._MEIPASS)
    else:
        # 源代码运行
        return Path(__file__).parent.resolve()


def get_data_dir():
    """获取数据目录（exe运行时在exe旁边创建data文件夹）"""
    if getattr(sys, 'frozen', False):
        # exe同目录下创建data文件夹
        return Path(sys.executable).parent / 'data'
    else:
        # 源代码运行用当前目录
        return Path(__file__).parent.resolve()


# ---- 配置路径 ----
BASE_DIR   = get_base_dir()
DATA_DIR   = get_data_dir()
DATA_FILE  = DATA_DIR / 'data.json'
IMAGE_DIR  = DATA_DIR / 'images'
CONFIG_FILE = DATA_DIR / 'config.json'
LOG_FILE   = DATA_DIR / 'server.log'
OUTPUT_DIR = DATA_DIR / 'exports'

# 确保目录存在
DATA_DIR.mkdir(exist_ok=True)
IMAGE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ---- 默认配置 ----
DEFAULT_CONFIG = {
    "port":      8765,
    "cors_origins": ["*"],
    "log_level": "INFO",
    "max_content_length": 20 * 1024 * 1024,
}

config = dict(DEFAULT_CONFIG)


def load_config():
    """加载 config.json"""
    global config
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            config = {**DEFAULT_CONFIG, **loaded}
        except Exception as e:
            config = dict(DEFAULT_CONFIG)
    else:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)


def get_local_ip():
    """获取本机局域网IP"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def find_available_port(start_port=8765, max_attempts=100):
    """自动寻找可用端口"""
    for port in range(start_port, start_port + max_attempts):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(('', port))
            s.close()
            return port
        except OSError:
            continue
    return None


def check_port_in_use(port):
    """检查端口是否被占用"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(('', port))
        s.close()
        return False
    except OSError:
        return True


def check_single_instance():
    """单实例检测 - 防止重复启动"""
    import tempfile
    lock_file = DATA_DIR / 'server.lock'
    
    if lock_file.exists():
        # 读取之前的PID
        try:
            with open(lock_file, 'r') as f:
                old_pid = int(f.read().strip())
            
            # 检查进程是否还在 - 尝试终止旧进程
            if sys.platform == 'win32':
                import subprocess
                try:
                    subprocess.run(['taskkill', '/F', '/PID', str(old_pid)], 
                                   capture_output=True, timeout=3)
                    logging.info(f"已终止旧进程 {old_pid}")
                    time.sleep(1)  # 等待端口释放
                except:
                    pass
            
            # 清理锁文件
            lock_file.unlink()
        except:
            pass
    
    # 写入当前PID
    with open(lock_file, 'w') as f:
        f.write(str(os.getpid()))
    return None


# ---- 日志初始化 ----
def init_logging():
    logging.basicConfig(
        level=getattr(logging, config.get('log_level', 'INFO')),
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )


# ---- 数据存取 ----
def load_data():
    if DATA_FILE.exists():
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logging.warning(f"加载数据失败: {e}")
    return {'expense': [], 'reimburse': []}


def save_data(data):
    try:
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        logging.error(f"保存数据失败: {e}")
        return False


def parse_base64_image(base64_data):
    """解析base64图片"""
    if not base64_data:
        return None
    try:
        if ',' in base64_data:
            header, data = base64_data.split(',', 1)
        else:
            data = base64_data
        img_data = base64.b64decode(data)
        pil_img = PILImage.open(io.BytesIO(img_data))
        return pil_img
    except Exception as e:
        logging.error(f"解析图片失败: {e}")
        return None


def create_excel_with_images(data):
    """创建带图片的Excel"""
    if not HAS_OPENPYXL or not HAS_PIL:
        raise Exception("缺少必要库")

    wb = Workbook()
    wb.remove(wb.active)

    # 样式
    header_font = Font(bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='667eea', end_color='667eea', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # 费用模板
    ws1 = wb.create_sheet('费用模板')
    headers = ['时间', '产品', '关联项目', '发生原因', '金额', '详情截图', '是否有票', '开票主体']
    
    ws1.merge_cells('A1:H1')
    ws1['A1'] = '报销详情'
    ws1['A1'].font = Font(bold=True, size=14)
    ws1['A1'].alignment = Alignment(horizontal='center')

    for col, h in enumerate(headers, 1):
        cell = ws1.cell(row=2, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    for row_idx, r in enumerate(data.get('expense', []), 3):
        ws1.cell(row=row_idx, column=1, value=r.get('time', '')).border = thin_border
        ws1.cell(row=row_idx, column=2, value=r.get('product', '')).border = thin_border
        ws1.cell(row=row_idx, column=3, value=r.get('related', '')).border = thin_border
        ws1.cell(row=row_idx, column=4, value=r.get('reason', '')).border = thin_border
        ws1.cell(row=row_idx, column=5, value=r.get('amount', 0)).border = thin_border
        ws1.cell(row=row_idx, column=7, value=r.get('hasTicket', '')).border = thin_border
        ws1.cell(row=row_idx, column=8, value=r.get('ticketEntity', '')).border = thin_border
        ws1.row_dimensions[row_idx].height = 80

        if r.get('image'):
            pil_img = parse_base64_image(r['image'])
            if pil_img:
                pil_img.thumbnail((150, 100))
                img_buffer = io.BytesIO()
                fmt = pil_img.format if pil_img.format else 'PNG'
                pil_img.save(img_buffer, format='JPEG' if fmt == 'JPEG' else 'PNG')
                img_buffer.seek(0)
                xl_img = XLImage(img_buffer)
                xl_img.anchor = f'F{row_idx}'
                ws1.add_image(xl_img)

    # 报销模板
    ws2 = wb.create_sheet('报销模板')
    for col, h in enumerate(headers, 1):
        cell = ws2.cell(row=2, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    for row_idx, r in enumerate(data.get('reimburse', []), 3):
        ws2.cell(row=row_idx, column=1, value=r.get('time', '')).border = thin_border
        ws2.cell(row=row_idx, column=2, value=r.get('product', '')).border = thin_border
        ws2.cell(row=row_idx, column=3, value=r.get('related', '')).border = thin_border
        ws2.cell(row=row_idx, column=4, value=r.get('reason', '')).border = thin_border
        ws2.cell(row=row_idx, column=5, value=r.get('amount', 0)).border = thin_border
        ws2.cell(row=row_idx, column=7, value=r.get('hasTicket', '')).border = thin_border
        ws2.cell(row=row_idx, column=8, value=r.get('ticketEntity', '')).border = thin_border
        ws2.row_dimensions[row_idx].height = 80

        if r.get('image'):
            pil_img = parse_base64_image(r['image'])
            if pil_img:
                pil_img.thumbnail((150, 100))
                img_buffer = io.BytesIO()
                fmt = pil_img.format if pil_img.format else 'PNG'
                pil_img.save(img_buffer, format='JPEG' if fmt == 'JPEG' else 'PNG')
                img_buffer.seek(0)
                xl_img = XLImage(img_buffer)
                xl_img.anchor = f'F{row_idx}'
                ws2.add_image(xl_img)

    # 保存
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'报销汇总_{timestamp}.xlsx'
    out_path = OUTPUT_DIR / filename
    wb.save(out_path)
    return out_path, filename


# ==================== HTTP 服务器 ====================
class APIHandler(SimpleHTTPRequestHandler):
    """API处理器"""

    def _add_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self._add_cors_headers()
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def _read_json_body(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        return json.loads(body.decode('utf-8'))

    def do_OPTIONS(self):
        self.send_response(200)
        self._add_cors_headers()
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        # 健康检查
        if path == '/api/health':
            self.send_json({'status': 'ok', 'time': datetime.now().isoformat()})

        # 加载数据
        elif path == '/api/load':
            self.send_json(load_data())

        # 最新导出文件下载
        elif path == '/api/download-latest':
            # 获取最新的导出文件
            export_files = list(OUTPUT_DIR.glob('*.xlsx'))
            if export_files:
                latest = max(export_files, key=lambda p: p.stat().st_mtime)
                self.send_file(str(latest), as_download=True)
            else:
                self.send_json({'success': False, 'error': '没有找到导出的文件'}, status=404)

        # 导出文件下载 (兼容直接访问)
        elif path.startswith('/exports/'):
            # URL解码获取文件名
            from urllib.parse import unquote
            filename = unquote(path.split('/')[-1])
            file_path = OUTPUT_DIR / filename
            if file_path.exists():
                self.send_file(str(file_path), as_download=True)
            else:
                self.send_error(404, 'File not found')

        # 静态文件
        else:
            if path == '/' or path == '':
                path = '/index.html'
            file_path = BASE_DIR / path.lstrip('/')
            
            if file_path.is_file():
                if path.endswith('.html'):
                    ct = 'text/html; charset=utf-8'
                    self.send_file(str(file_path), content_type=ct, as_download=False)
                elif path.endswith('.css'):
                    ct = 'text/css; charset=utf-8'
                    self.send_file(str(file_path), content_type=ct)
                elif path.endswith('.js'):
                    ct = 'application/javascript; charset=utf-8'
                    self.send_file(str(file_path), content_type=ct)
                else:
                    self.send_file(str(file_path))
            else:
                # SPA fallback
                self.send_file(str(BASE_DIR / 'index.html'), content_type='text/html; charset=utf-8', as_download=False)

    def send_file(self, filepath, content_type=None, as_download=False):
        """返回文件"""
        filepath = Path(filepath)
        if not filepath.exists():
            self.send_error(404, 'File not found')
            return
        
        self.send_response(200)
        ct = content_type or 'application/octet-stream'
        self.send_header('Content-Type', ct)
        
        if as_download:
            safe_name = quote(filepath.name)
            self.send_header('Content-Disposition', f'attachment; filename="{safe_name}"; filename*=UTF-8\'\'{safe_name}')
        
        self.send_header('Content-Length', filepath.stat().st_size)
        self.send_header('Cache-Control', 'no-cache')
        self._add_cors_headers()
        self.end_headers()
        
        with open(filepath, 'rb') as f:
            self.wfile.write(f.read())

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        try:
            if path == '/api/save':
                data = self._read_json_body()
                ok = save_data(data)
                self.send_json({'success': ok})

            elif path == '/api/export':
                data = self._read_json_body()
                out_path, filename = create_excel_with_images(data)
                self.send_json({
                    'success': True,
                    'filename': filename,
                    'download_url': f'/exports/{filename}'
                })

            elif path == '/api/load':
                self.send_json(load_data())

            else:
                self.send_error(404)
        except Exception as e:
            logging.error(f"处理请求出错: {path} → {e}", exc_info=True)
            self.send_json({'success': False, 'error': str(e)}, status=500)

    def log_message(self, format, *args):
        if args and 'favicon' in str(args[0]):
            return
        logging.info(args[0] if args else format)


# ==================== 主入口 ====================
def main():
    load_config()
    init_logging()

    if not HAS_OPENPYXL or not HAS_PIL:
        logging.error("缺少必要依赖！")
        print("\n❌ 缺少必要依赖！请先安装：")
        print("   pip install openpyxl Pillow\n")
        input("按回车键退出...")
        sys.exit(1)

    # 单实例检测
    old_pid = check_single_instance()
    if old_pid:
        print(f"\n⚠️  程序已在运行中（PID: {old_pid}）")
        print(f"   请关闭后重试，或前往 http://localhost:{config['port']} 使用")
        input("按回车键退出...")
        sys.exit(0)

    # 端口检测
    start_port = config['port']
    port = find_available_port(start_port)
    if port is None:
        print("❌ 无法找到可用端口")
        sys.exit(1)
    
    if port != start_port:
        print(f"⚠️  端口 {start_port} 已被占用，自动使用端口 {port}")

    host = "0.0.0.0"
    server = HTTPServer((host, port), APIHandler)

    local_ip = get_local_ip()
    local_url = f"http://localhost:{port}"
    lan_url = f"http://{local_ip}:{port}"

    # 打印启动信息
    print("\n" + "=" * 52)
    print("  报销费用填写工具 · 桌面版")
    print("=" * 52)
    print(f"\n  ✅ 服务已启动!")
    print(f"\n  📱 访问方式:")
    print(f"     本机:    {local_url}")
    print(f"     局域网:  {lan_url}")
    print(f"\n  📋 数据目录: {DATA_DIR}")
    print(f"  📝 日志文件: {LOG_FILE}")
    print(f"\n  停止服务: Ctrl+C")
    print("=" * 52 + "\n")

    # 自动打开浏览器
    def open_browser():
        time.sleep(1.5)
        webbrowser.open(local_url)

    threading.Thread(target=open_browser, daemon=True).start()

    # 优雅退出
    def shutdown(signum, frame):
        print("\n正在停止服务...")
        # 清理锁文件
        lock_file = DATA_DIR / 'server.lock'
        if lock_file.exists():
            lock_file.unlink()
        logging.info("服务关闭")
        server.shutdown()
        sys.exit(0)

    signal.signal(signal.SIGTERM, shutdown)
    signal.signal(signal.SIGINT, shutdown)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        shutdown(None, None)


if __name__ == '__main__':
    main()
