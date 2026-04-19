# -*- coding: utf-8 -*-
"""
server.py - HTTP 服务器与路由层
每个 API 分支末尾必须有 return，否则会继续执行静态文件逻辑
"""

import json, logging, socket, signal, sys, threading, time
import hashlib, secrets, webbrowser
from datetime import datetime, timedelta
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, quote

logger = logging.getLogger(__name__)

_BAOXIAO_SALT = "baoxiao_v1_"


def hash_password(password: str) -> str:
    return hashlib.sha256((_BAOXIAO_SALT + password).encode('utf-8')).hexdigest()


def verify_password(password: str, stored_hash: str) -> bool:
    return hash_password(password) == stored_hash


# 内存 Session
_sessions = {}  # {token: {user_id, username, expire_at}}


def generate_token(user_id, username) -> str:
    token = secrets.token_hex(24)
    _sessions[token] = {
        "user_id": user_id,
        "username": username,
        "expire_at": datetime.now() + timedelta(days=7)
    }
    logger.info(f"Session 生成: {username} (token={token[:8]}...)")
    return token


def verify_token(token: str):
    if not token:
        return None
    s = _sessions.get(token)
    if not s or s["expire_at"] < datetime.now():
        if token in _sessions:
            del _sessions[token]
        return None
    return s


def kill_token(token: str):
    if token in _sessions:
        logger.info(f"Session 销毁: {_sessions[token]['username']}")
        del _sessions[token]


def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def find_available_port(start=8765, max_attempts=100):
    for port in range(start, start + max_attempts):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(('', port))
            s.close()
            return port
        except OSError:
            continue
    return None


def check_single_instance(lock_file: Path):
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


class APIHandler(SimpleHTTPRequestHandler):

    def __init__(self, *args, static_dir=None, output_dir=None, data_store=None,
                 excel_factory=None, **kwargs):
        self._static_dir = static_dir
        self._output_dir = output_dir
        self._data_store = data_store
        self._excel_factory = excel_factory
        super().__init__(*args, **kwargs)

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')

    def _require_auth(self):
        """验证 token，返回 session dict 或 None（已自动发 401）"""
        auth = self.headers.get('Authorization', '')
        token = auth.replace('Bearer ', '').strip()
        session = verify_token(token)
        if not session:
            self.send_json({'success': False, 'error': '未登录或登录已过期，请重新登录'}, status=401)
            return None
        return session

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self._cors()
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def send_file(self, filepath, content_type=None, inline=False):
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

    # ========================================================================
    # GET 路由
    # ========================================================================
    def do_GET(self):
        p = urlparse(self.path).path
        logger.info(f"[do_GET] path_raw={self.path!r}  p={p!r}")

        if p == '/api/health':
            from app.excel_export import get_excel_creator_info
            self.send_json({'status': 'ok', 'time': datetime.now().isoformat(),
                            'excel_creator': get_excel_creator_info()})
            return

        if p == '/api/load':
            s = self._require_auth()
            if s is None:
                return
            from app.store import get_records_for_user
            filtered = get_records_for_user(self._data_store._data_file, s['user_id'])
            self.send_json({'success': True, 'data': filtered,
                            'user': {'id': s['user_id'], 'username': s['username']}})
            return

        if p == '/api/me':
            s = self._require_auth()
            if s is None:
                return
            from app.store import get_user_by_id
            user = get_user_by_id(self._data_store._data_file, s['user_id'])
            public = {k: v for k, v in user.items()} if user else None
            if public and 'password' in public:
                del public['password']
            self.send_json({'success': True, 'user': public, 'username': s.get('username', '')})
            return

        if p == '/api/download-latest':
            files = list(self._output_dir.glob('*.xlsx'))
            if files:
                latest = max(files, key=lambda f: f.stat().st_mtime)
                self.send_file(str(latest), inline=False)
                return
            self.send_json({'success': False, 'error': '没有找到导出的文件'}, status=404)
            return

        if p.startswith('/exports/'):
            from urllib.parse import unquote
            fname = unquote(p.split('/')[-1])
            self.send_file(str(self._output_dir / fname), inline=False)
            return

        # 静态文件
        if p in ('/', '', '/index.html'):
            idx = self._static_dir / 'index.html'
            if idx.exists():
                self.send_file(str(idx), content_type='text/html; charset=utf-8', inline=True)
                return
            self.send_error(404, 'index.html not found')
            return

        # 静态文件（URL 可能含 %E6%8F%90%E9%86%92 等编码，先解码再处理）
        from urllib.parse import unquote
        p_decoded = unquote(p)
        if p_decoded.startswith('/static/'):
            rel = p_decoded[9:]  # strip '/static/' (9 chars)
        else:
            rel = unquote(p.lstrip('/'))
        fp = self._static_dir / rel
        if fp.is_file():
            ct_map = {'.html': 'text/html; charset=utf-8',
                       '.css':  'text/css; charset=utf-8',
                       '.js':   'application/javascript; charset=utf-8',
                       '.png':  'image/png',
                       '.jpg':  'image/jpeg',
                       '.jpeg': 'image/jpeg',
                       '.gif':  'image/gif',
                       '.svg':  'image/svg+xml'}
            self.send_file(str(fp), content_type=ct_map.get(fp.suffix, 'application/octet-stream'), inline=True)
            return

        # SPA fallback
        idx = self._static_dir / 'index.html'
        if idx.exists():
            self.send_file(str(idx), content_type='text/html; charset=utf-8', inline=True)
        else:
            self.send_error(404, 'Not found')

    # ========================================================================
    # POST 路由
    # ========================================================================
    def do_POST(self):
        try:
            cl = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(cl).decode('utf-8'))
        except Exception as e:
            logger.error(f"解析请求体失败: {e}")
            self.send_json({'success': False, 'error': 'Invalid request body'}, status=400)
            return

        p = urlparse(self.path).path

        try:
            # ---------- 公开接口 ----------
            if p == '/api/register':
                username = (body.get('username', '') or '').strip()
                password = (body.get('password', '') or '').strip()
                if not username or not password:
                    self.send_json({'success': False, 'error': '用户名和密码不能为空'}, status=400)
                    return
                if len(username) < 2:
                    self.send_json({'success': False, 'error': '用户名至少2个字符'}, status=400)
                    return
                if len(password) < 4:
                    self.send_json({'success': False, 'error': '密码至少4位'}, status=400)
                    return
                from app.store import register_user
                user, err = register_user(self._data_store._data_file, username, hash_password(password))
                if err:
                    self.send_json({'success': False, 'error': err}, status=400)
                    return
                token = generate_token(user['id'], user['username'])
                public = {k: v for k, v in user.items() if k != 'password'}
                self.send_json({'success': True, 'token': token, 'user': public})
                return

            if p == '/api/login':
                username = (body.get('username', '') or '').strip()
                password = (body.get('password', '') or '').strip()
                if not username or not password:
                    self.send_json({'success': False, 'error': '用户名和密码不能为空'}, status=400)
                    return
                from app.store import get_user_by_username, update_last_login
                user = get_user_by_username(self._data_store._data_file, username)
                if not user or not verify_password(password, user.get('password', '')):
                    self.send_json({'success': False, 'error': '用户名或密码错误'}, status=401)
                    return
                update_last_login(self._data_store._data_file, user['id'])
                token = generate_token(user['id'], user['username'])
                public = {k: v for k, v in user.items() if k != 'password'}
                self.send_json({'success': True, 'token': token, 'user': public})
                return

            # ---------- 需要登录 ----------
            s = self._require_auth()
            if s is None:
                return
            user_id = s['user_id']
            username = s['username']
            df = self._data_store._data_file

            if p == '/api/logout':
                auth = self.headers.get('Authorization', '')
                kill_token(auth.replace('Bearer ', '').strip())
                self.send_json({'success': True, 'message': '已退出登录'})
                return

            if p == '/api/load':
                from app.store import get_records_for_user
                filtered = get_records_for_user(df, user_id)
                self.send_json({'success': True, 'data': filtered,
                                'user': {'id': user_id, 'username': username}})
                return

            if p == '/api/save':
                from app.store import save_user_records
                expense_recs   = body.get('expense', [])
                reimburse_recs = body.get('reimburse', [])
                ok = save_user_records(df, user_id, expense_recs, reimburse_recs)
                self.send_json({'success': ok})
                return

            if p == '/api/add-record':
                from app.store import add_record
                tt = body.get('template_type', 'expense')
                record = dict(body.get('record', {}))
                record['user_id'] = user_id
                _, errors = add_record(df, tt, record, user_id)
                if errors:
                    self.send_json({'success': False, 'errors': errors}, status=400)
                    return
                from app.store import get_records_for_user
                filtered = get_records_for_user(df, user_id)
                self.send_json({'success': True, 'data': filtered})
                return

            if p == '/api/delete-record':
                from app.store import delete_record
                tt = body.get('template_type', 'expense')
                idx = body.get('index', -1)
                _, err = delete_record(df, tt, idx, user_id)
                if err:
                    self.send_json({'success': False, 'error': err}, status=400)
                    return
                from app.store import get_records_for_user
                filtered = get_records_for_user(df, user_id)
                self.send_json({'success': True, 'data': filtered})
                return

            if p == '/api/export':
                from app.store import get_records_for_user
                export_data = get_records_for_user(df, user_id)
                out_path, fname = self._excel_factory.create(export_data, self._output_dir)
                from app.excel_export import get_excel_creator_info
                self.send_json({'success': True, 'filename': fname,
                                'download_url': f'/exports/{fname}',
                                'creator': get_excel_creator_info()})
                return

            if p == '/api/user-stats':
                from app.store import get_user_stats
                stats = get_user_stats(df)
                self.send_json({'success': True, 'users': stats})
                return

            if p == '/api/delete-user':
                target = body.get('user_id')
                if not target:
                    self.send_json({'success': False, 'error': '缺少 user_id'}, status=400)
                    return
                from app.store import delete_user
                ok, err = delete_user(df, target)
                if not ok:
                    self.send_json({'success': False, 'error': err}, status=400)
                    return
                self.send_json({'success': True})
                return

            self.send_error(404)

        except Exception as e:
            logger.error(f"处理请求 {p} 出错: {e}", exc_info=True)
            self.send_json({'success': False, 'error': str(e)}, status=500)


def make_handler(static_dir, output_dir, data_store, excel_factory):
    def create_handler(*args, **kwargs):
        return APIHandler(*args, static_dir=static_dir, output_dir=output_dir,
                          data_store=data_store, excel_factory=excel_factory, **kwargs)
    return create_handler


class Server:
    def __init__(self, static_dir, output_dir, data_store, excel_factory, port=8765):
        self._port = port
        Handler = make_handler(static_dir, output_dir, data_store, excel_factory)
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
