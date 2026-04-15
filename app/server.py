# -*- coding: utf-8 -*-
"""
server.py — HTTP 服务器与路由层
职责：接收 HTTP 请求、调用业务层、返回响应，不包含 Excel 或数据存取细节
"""

import json
import logging
import socket
import signal
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, quote

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 端口与单实例管理
# ---------------------------------------------------------------------------

def get_local_ip():
    """获取本机局域网 IP"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def find_available_port(start_port=8765, max_attempts=100):
    """遍历端口范围，返回第一个可用端口"""
    for port in range(start_port, start_port + max_attempts):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(('', port))
            s.close()
            return port
        except OSError:
            continue
    return None


def check_single_instance(lock_file: Path):
    """单实例检测：检查锁文件是否存在并尝试终止旧进程"""
    if lock_file.exists():
        try:
            pid = int(lock_file.read_text(encoding='utf-8').strip())
            if sys.platform == 'win32':
                import subprocess
                try:
                    subprocess.run(['taskkill', '/F', '/PID', str(pid)],
                                   capture_output=True, timeout=3)
                    logger.info(f"已终止旧进程 {pid}")
                    time.sleep(1)
                except Exception:
                    pass
            lock_file.unlink(missing_ok=True)
        except Exception:
            pass

    lock_file.write_text(str(__import__('os').getpid()), encoding='utf-8')


# ---------------------------------------------------------------------------
# API Handler
# ---------------------------------------------------------------------------

class APIHandler(SimpleHTTPRequestHandler):

    def __init__(self, *args, static_dir=None, output_dir=None, data_store=None, excel_factory=None, **kwargs):
        self._static_dir  = static_dir
        self._output_dir = output_dir
        self._data_store = data_store
        self._excel_factory = excel_factory
        super().__init__(*args, **kwargs)

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self._cors()
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def send_file(self, filepath, content_type=None, inline=False):
        """返回文件：inline=True → 浏览器直接展示，False → 下载"""
        fp = Path(filepath)
        if not fp.exists():
            self.send_error(404, 'File not found')
            return
        self.send_response(200)
        ct = content_type or 'application/octet-stream'
        self.send_header('Content-Type', ct)
        if not inline:
            safe = quote(fp.name)
            self.send_header('Content-Disposition',
                             f'attachment; filename="{safe}"; filename*=UTF-8\'\'{safe}')
        self.send_header('Content-Length', fp.stat().st_size)
        self.send_header('Cache-Control', 'no-cache')
        self._cors()
        self.end_headers()
        with open(fp, 'rb') as f:
            self.wfile.write(f.read())

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()

    def log_message(self, format, *args):
        if args and 'favicon' in str(args[0]):
            return
        logger.info(args[0] if args else format)

    # ------------------------------------------------------------------
    # GET 路由
    # ------------------------------------------------------------------
    def do_GET(self):
        parsed = urlparse(self.path)
        path   = parsed.path

        if path == '/api/health':
            from app.excel_export import get_excel_creator_info
            self.send_json({
                'status': 'ok',
                'time': datetime.now().isoformat(),
                'excel_creator': get_excel_creator_info(),
            })

        elif path == '/api/load':
            data = self._data_store.load()
            self.send_json(data)

        elif path == '/api/download-latest':
            files = list(self._output_dir.glob('*.xlsx'))
            if files:
                latest = max(files, key=lambda p: p.stat().st_mtime)
                self.send_file(str(latest), inline=False)
            else:
                self.send_json({'success': False, 'error': '没有找到导出的文件'}, status=404)

        elif path.startswith('/exports/'):
            from urllib.parse import unquote
            filename = unquote(path.split('/')[-1])
            self.send_file(str(self._output_dir / filename), inline=False)

        else:
            # 静态文件
            if path in ('/', '', '/index.html'):
                static_index = self._static_dir / 'index.html'
                if static_index.exists():
                    self.send_file(str(static_index),
                                   content_type='text/html; charset=utf-8',
                                   inline=True)
                else:
                    self.send_error(404, 'index.html not found')
                return

            # 其他静态资源
            rel = path.lstrip('/')
            fp = self._static_dir / rel
            if fp.is_file():
                ct_map = {'.html': 'text/html; charset=utf-8',
                           '.css':  'text/css; charset=utf-8',
                           '.js':   'application/javascript; charset=utf-8'}
                ct = ct_map.get(fp.suffix, 'application/octet-stream')
                self.send_file(str(fp), content_type=ct, inline=True)
            else:
                # SPA fallback → index.html
                idx = self._static_dir / 'index.html'
                if idx.exists():
                    self.send_file(str(idx),
                                   content_type='text/html; charset=utf-8',
                                   inline=True)
                else:
                    self.send_error(404, 'Not found')

    # ------------------------------------------------------------------
    # POST 路由
    # ------------------------------------------------------------------
    def do_POST(self):
        parsed = urlparse(self.path)
        path   = parsed.path

        try:
            cl = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(cl).decode('utf-8'))
        except Exception as e:
            logger.error(f"解析请求体失败: {e}")
            self.send_json({'success': False, 'error': 'Invalid request body'}, status=400)
            return

        try:
            if path == '/api/save':
                ok = self._data_store.save(body)
                self.send_json({'success': ok})

            elif path == '/api/export':
                out_path, filename = self._excel_factory.create(body, self._output_dir)
                from app.excel_export import get_excel_creator_info
                self.send_json({
                    'success': True,
                    'filename': filename,
                    'download_url': f'/exports/{filename}',
                    'creator': get_excel_creator_info(),
                })

            elif path == '/api/load':
                self.send_json(self._data_store.load())

            else:
                self.send_error(404)

        except Exception as e:
            logger.error(f"处理请求 {path} 出错: {e}", exc_info=True)
            self.send_json({'success': False, 'error': str(e)}, status=500)


# ---------------------------------------------------------------------------
# 服务器工厂（供 main.py 使用）
# ---------------------------------------------------------------------------

def make_handler(static_dir, output_dir, data_store, excel_factory):
    """返回配置好的 APIHandler 子类，绕过 __init__ 参数限制"""
    def create_handler(*args, **kwargs):
        return APIHandler(*args,
                          static_dir=static_dir,
                          output_dir=output_dir,
                          data_store=data_store,
                          excel_factory=excel_factory,
                          **kwargs)
    return create_handler


class Server:
    """封装 HTTPServer 的启动/停止"""

    def __init__(self, static_dir, output_dir, data_store, excel_factory, port=8765):
        self._port   = port
        Handler      = make_handler(static_dir, output_dir, data_store, excel_factory)
        self._server = HTTPServer(('0.0.0.0', port), Handler)

    @property
    def port(self):
        return self._port

    def start(self, auto_open_browser=True, local_url=None):
        def _open():
            if auto_open_browser and local_url:
                time.sleep(1.5)
                webbrowser.open(local_url)

        threading.Thread(target=_open, daemon=True).start()
        logger.info(f"服务器已启动，端口 {self._port}")

        def shutdown(signum, frame):
            logger.info("正在关闭服务器...")
            self._server.shutdown()
            sys.exit(0)

        signal.signal(signal.SIGTERM, shutdown)
        signal.signal(signal.SIGINT,  shutdown)

        try:
            self._server.serve_forever()
        except KeyboardInterrupt:
            shutdown(None, None)
