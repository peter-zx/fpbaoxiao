# -*- coding: utf-8 -*-
"""
报销费用填写工具 - 云服务器版 (server_cloud.py)
===============================================
改动点（相对于 server.py）：
  1. 绑定 0.0.0.0，对外暴露端口
  2. 增加 CORS 跨域支持（手机浏览器必须）
  3. 中文文件名 Content-Disposition 正确编码
  4. JSON/HTML 响应统一 UTF-8
  5. 可通过 config.json 自定义配置
  6. 启动时显示外网访问地址
  7. 日志输出到文件
  8. SIGTERM / SIGINT 优雅退出
  9. 健康检查接口 /api/health
  10. 关闭自动打开浏览器（服务器无GUI）
"""

import os
import sys
import json
import base64
import io
import logging
import signal
import socket

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

# ---- 配置路径 ----
BASE_DIR   = Path(__file__).parent.resolve()
DATA_FILE  = BASE_DIR / 'data.json'
IMAGE_DIR  = BASE_DIR / 'images'
CONFIG_FILE = BASE_DIR / 'config.json'
LOG_FILE   = BASE_DIR / 'server.log'
OUTPUT_DIR = BASE_DIR / 'exports'

IMAGE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# ---- 默认配置 ----
DEFAULT_CONFIG = {
    "host":      "0.0.0.0",
    "port":      8765,
    "cors_origins": ["*"],          # 生产环境建议改为具体域名
    "log_level": "INFO",
    "max_content_length": 20 * 1024 * 1024,  # 20MB
}

config = dict(DEFAULT_CONFIG)

def load_config():
    """加载 config.json，缺失则创建"""
    global config
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                loaded = json.load(f)
            config = {**DEFAULT_CONFIG, **loaded}
            logging.info(f"配置文件已加载: {CONFIG_FILE}")
        except Exception as e:
            logging.warning(f"配置文件读取失败，使用默认值: {e}")
            config = dict(DEFAULT_CONFIG)
    else:
        # 首次运行，写入默认配置
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        logging.info(f"已创建默认配置文件: {CONFIG_FILE}")

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

def get_external_url():
    """获取外网访问URL（用户需自行替换为公网IP或域名）"""
    ip = get_local_ip()
    port = config['port']
    return f"http://{ip}:{port}", f"http://<你的公网IP>:{port}"

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

# ---- 图片处理 ----
def parse_base64_image(base64_str):
    """解析 base64 图片字符串，返回 (PIL.Image, extension)"""
    if not base64_str or not base64_str.startswith('data:'):
        return None, None
    try:
        header, data = base64_str.split(',', 1)
        ext = 'png'
        if 'jpeg' in header or 'jpg' in header:
            ext = 'jpg'
        elif 'gif' in header:
            ext = 'gif'
        elif 'png' in header:
            ext = 'png'
        img_bytes = base64.b64decode(data)
        pil_img = PILImage.open(io.BytesIO(img_bytes))
        # 统一转为 RGB（JPEG 不支持 RGBA）
        if pil_img.mode in ('RGBA', 'P'):
            pil_img = pil_img.convert('RGB')
        return pil_img, ext
    except Exception as e:
        logging.warning(f"解析图片失败: {e}")
        return None, None

# ---- Excel 生成 ----
def build_styles():
    return {
        'header_font':   Font(bold=True, color='FFFFFF', size=11),
        'header_fill':   PatternFill(start_color='4F46E5', end_color='4F46E5', fill_type='solid'),
        'title_font':    Font(bold=True, size=14, color='1F2937'),
        'total_font':    Font(bold=True, size=11, color='DC2626'),
        'cell_font':     Font(size=10),
        'center_align':  Alignment(horizontal='center', vertical='center', wrap_text=True),
        'left_align':    Alignment(horizontal='left',   vertical='center', wrap_text=True),
        'thin_border':   Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'),  bottom=Side(style='thin')
        ),
    }

def create_excel_with_images(data):
    """生成包含图片的 Excel，返回保存路径"""
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl 未安装")

    wb = Workbook()
    wb.remove(wb.active)
    s = build_styles()

    SHEETS = [
        {
            'title': '费用模板（不计入销售成本）',
            'headers': ['时间', '产品（购买的物品、服务）', '关联项目', '发生原因', '金额（元）', '凭证截图', '是否有票', '开票主体'],
            'records': data.get('expense', []),
            'related_label': '关联项目',
        },
        {
            'title': '报销模板（计入销售订单成本）',
            'headers': ['时间', '产品（购买的物品、服务）', '关联客户', '发生原因', '金额（元）', '凭证截图', '是否有票', '开票主体'],
            'records': data.get('reimburse', []),
            'related_label': '关联客户',
        },
    ]

    col_widths = [12, 32, 20, 24, 14, 20, 10, 22]

    for sheet_def in SHEETS:
        ws = wb.create_sheet(sheet_def['title'])
        records = sheet_def['records']

        # 合并标题行
        ws.merge_cells(f'A1:{get_column_letter(len(sheet_def["headers"]))}1')
        title_cell = ws['A1']
        title_cell.value = f'报销详情 — {datetime.now().strftime("%Y年%m月%d日")}'
        title_cell.font = s['title_font']
        title_cell.alignment = s['center_align']
        ws.row_dimensions[1].height = 30

        # 表头
        for col, header in enumerate(sheet_def['headers'], 1):
            cell = ws.cell(row=2, column=col, value=header)
            cell.font = s['header_font']
            cell.fill = s['header_fill']
            cell.alignment = s['center_align']
            cell.border = s['thin_border']
        ws.row_dimensions[2].height = 22

        # 数据行
        for ri, record in enumerate(records, 3):
            ws.cell(row=ri, column=1, value=record.get('time', '')).border = s['thin_border']
            ws.cell(row=ri, column=2, value=record.get('product', '')).border = s['thin_border']
            ws.cell(row=ri, column=3, value=record.get(sheet_def['related_label'].lower().replace('关联',''), record.get('related', ''))).border = s['thin_border']
            ws.cell(row=ri, column=4, value=record.get('reason', '')).border = s['thin_border']

            amt_cell = ws.cell(row=ri, column=5, value=record.get('amount', 0))
            amt_cell.border = s['thin_border']
            amt_cell.number_format = '#,##0.00'

            ws.cell(row=ri, column=6, value='').border = s['thin_border']
            ws.cell(row=ri, column=7, value=record.get('hasTicket', '')).border = s['thin_border']
            ws.cell(row=ri, column=8, value=record.get('ticketEntity', '')).border = s['thin_border']

            # 行高（给图片留空间）
            ws.row_dimensions[ri].height = 80

            # 嵌入图片
            if record.get('image') and HAS_PIL:
                pil_img, _ = parse_base64_image(record['image'])
                if pil_img:
                    try:
                        # 缩放
                        pil_img.thumbnail((160, 110))
                        buf = io.BytesIO()
                        pil_img.save(buf, format='JPEG', quality=85)
                        buf.seek(0)
                        xl_img = XLImage(buf)
                        xl_img.anchor = f'F{ri}'
                        ws.add_image(xl_img)
                    except Exception as e:
                        logging.warning(f"嵌入图片失败 record.id={record.get('id')}: {e}")

        # 合计行
        if records:
            total_row = len(records) + 3
            ws.cell(row=total_row, column=1, value='合计').font = s['total_font']
            ws.cell(row=total_row, column=1).border = s['thin_border']
            ws.cell(row=total_row, column=5, value=f'=SUM(E3:E{total_row-1})').font = s['total_font']
            ws.cell(row=total_row, column=5).number_format = '#,##0.00'
            ws.cell(row=total_row, column=5).border = s['thin_border']
            for c in [2, 3, 4, 6, 7, 8]:
                ws.cell(row=total_row, column=c).border = s['thin_border']
            ws.row_dimensions[total_row].height = 22

        # 列宽
        for col, width in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width

        # 冻结首行
        ws.freeze_panes = 'A3'

    # 保存
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'报销费用_{timestamp}.xlsx'
    out_path = OUTPUT_DIR / filename
    wb.save(str(out_path))
    logging.info(f"Excel 已生成: {out_path}")
    return str(out_path), filename

# ---- HTTP Handler ----
class APIHandler(SimpleHTTPRequestHandler):

    def __init__(self, *args, **kwargs):
        self.base_dir = str(BASE_DIR)
        super().__init__(*args, directory=self.base_dir, **kwargs)

    # ---- 通用工具 ----

    def send_json(self, data, status=200):
        """返回 JSON 响应（统一 UTF-8）"""
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', 'no-cache')
        self._add_cors_headers()
        self.end_headers()
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.wfile.write(body)

    def send_file(self, filepath, content_type=None, as_download=False):
        """返回文件响应
        - as_download=False : 让浏览器直接渲染（如HTML页面）
        - as_download=True  : 弹出下载框（如Excel）
        """
        filepath = Path(filepath)
        if not filepath.exists():
            self.send_error(404, 'File not found')
            return
        self.send_response(200)
        ct = content_type or 'application/octet-stream'
        self.send_header('Content-Type', ct)
        if as_download:
            # 中文文件名 RFC 5987 编码，触发下载
            safe_name = quote(filepath.name)
            self.send_header('Content-Disposition', f'attachment; filename="{safe_name}"; filename*=UTF-8\'\'{safe_name}')
        self.send_header('Content-Length', filepath.stat().st_size)
        self.send_header('Cache-Control', 'no-cache')
        self._add_cors_headers()
        self.end_headers()
        with open(filepath, 'rb') as f:
            self.wfile.write(f.read())

    def _add_cors_headers(self):
        origins = config.get('cors_origins', ['*'])
        if '*' in origins:
            self.send_header('Access-Control-Allow-Origin', '*')
        else:
            origin = self.headers.get('Origin', '')
            if origin in origins:
                self.send_header('Access-Control-Allow-Origin', origin)
                self.send_header('Vary', 'Origin')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def _read_json_body(self):
        """读取并解析 JSON 请求体"""
        cl = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(cl)
        return json.loads(body.decode('utf-8'))

    # ---- OPTIONS 预检 ----
    def do_OPTIONS(self):
        self.send_response(204)
        self._add_cors_headers()
        self.end_headers()

    # ---- GET ----
    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == '/api/health':
            self.send_json({'status': 'ok', 'time': datetime.now().isoformat()})

        elif path == '/api/load':
            self.send_json(load_data())

        elif path.startswith('/exports/'):
            # 导出文件下载（强制弹出保存窗口）
            filename = path.split('/')[-1]
            file_path = OUTPUT_DIR / filename
            self.send_file(file_path, as_download=True)

        elif path == '/api/config':
            # 返回配置（不包含敏感字段）
            safe_config = {k: v for k, v in config.items() if k not in ('secret', 'password')}
            self.send_json(safe_config)

        else:
            # 静态文件（index.html）
            if path == '/' or path == '':
                path = '/index.html'
            file_path = BASE_DIR / path.lstrip('/')
            if file_path.is_file():
                if path.endswith('.html'):
                    ct = 'text/html; charset=utf-8'
                elif path.endswith('.css'):
                    ct = 'text/css; charset=utf-8'
                    self.send_file(str(file_path), content_type=ct)
                elif path.endswith('.js'):
                    ct = 'application/javascript; charset=utf-8'
                    self.send_file(str(file_path), content_type=ct)
                else:
                    self.send_file(str(file_path))
            else:
                # SPA fallback → index.html（不触发下载，直接在浏览器渲染）
                self.send_file(str(BASE_DIR / 'index.html'), content_type='text/html; charset=utf-8', as_download=False)

    # ---- POST ----
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

    # ---- 日志 ----
    def log_message(self, format, *args):
        # 过滤 favicon 日志噪音
        if args and 'favicon' in str(args[0]):
            return
        logging.info(args[0] if args else format)


# ---- 主入口 ----
def main():
    load_config()
    init_logging()

    if not HAS_OPENPYXL or not HAS_PIL:
        logging.error("缺少必要依赖，请先安装: pip install openpyxl Pillow")
        print("\n❌ 缺少必要依赖！请运行：")
        print("   pip install openpyxl Pillow\n")
        sys.exit(1)

    host = config['host']
    port = config['port']
    server = HTTPServer((host, port), APIHandler)

    local_url, ext_url = get_external_url()
    logging.info("=" * 52)
    logging.info("报销费用填写工具 · 云服务器版")
    logging.info("=" * 52)
    logging.info(f"监听地址: http://{host}:{port}")
    logging.info(f"局域网访问: http://{get_local_ip()}:{port}")
    logging.info(f"外网访问（请替换为你的公网IP）: http://<公网IP>:{port}")
    logging.info(f"数据文件: {DATA_FILE}")
    logging.info(f"导出目录: {OUTPUT_DIR}")
    logging.info(f"日志文件: {LOG_FILE}")
    logging.info("=" * 52)
    print("\n" + "=" * 52)
    print(">> 报销费用填写工具 . 云服务器版")
    print("=" * 52)
    print(f"  [OK] 服务已启动，请访问:")
    print(f"       本机:    http://localhost:{port}")
    print(f"       局域网:  http://{get_local_ip()}:{port}")
    print(f"       外网:    http://<公网IP>:{port}")
    print(f"\n  停止服务: Ctrl+C")
    print("=" * 52 + "\n")

    # 优雅退出
    def shutdown(signum, frame):
        print("\n正在停止服务...")
        logging.info("收到退出信号，服务关闭")
        server.shutdown()
        sys.exit(0)

    signal.signal(signal.SIGTERM, shutdown)
    signal.signal(signal.SIGINT,  shutdown)

    server.serve_forever()


if __name__ == '__main__':
    main()
