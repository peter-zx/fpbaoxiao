# -*- coding: utf-8 -*-
"""
配置管理 — 环境区分、路径配置化、外置 YAML
"""

import os
import sys
import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# ---- 运行环境检测 ----

def detect_env():
    """检测当前运行环境: development / production / frozen(exe)"""
    if getattr(sys, 'frozen', False):
        return 'frozen'
    env = os.environ.get('BAOXIAO_ENV', 'development')
    return env if env in ('development', 'production') else 'development'


def get_project_root():
    """项目根目录（绝对路径，不写死）
    - frozen: exe 所在目录
    - development: main.py 所在目录
    """
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent.resolve()
    # 往上找包含 main.py 的目录
    current = Path(__file__).resolve().parent
    while current != current.parent:
        if (current / 'main.py').exists():
            return current
        current = current.parent
    # fallback: app 的父目录
    return Path(__file__).resolve().parent.parent


# ---- 默认配置 ----

_DEFAULTS = {
    'server': {
        'host': '0.0.0.0',
        'port': 8765,
        'port_max': 8780,
        'max_content_length': 20 * 1024 * 1024,  # 20MB
    },
    'paths': {
        'data_dir': 'data',
        'export_dir': 'exports',
        'static_dir': 'static',
        'log_file': 'server.log',
        'lock_file': 'server.lock',
        'data_file': 'data.json',
    },
    'excel': {
        'row_height': 200,
        'header_height': 40,
        'total_row_height': 40,
        'image_col_width': 15,
        'font_name': '微软雅黑',
        'header_font_size': 12,
        'data_font_size': 10,
    },
    'logging': {
        'level': 'INFO',
        'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    },
    'cors': {
        'origins': ['*'],
    },
}


def _deep_merge(base, override):
    """深度合并字典，override 覆盖 base"""
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


class Config:
    """全局配置对象"""

    def __init__(self):
        self._data = {}
        self._root = get_project_root()
        self._env = detect_env()

    def load(self):
        """加载配置：默认值 → 外置文件 → 环境变量"""
        self._data = _deep_merge({}, _DEFAULTS)

        # 外置 YAML / JSON 配置
        for suffix in ('', f'.{self._env}'):
            for ext in ('.yaml', '.yml', '.json'):
                cfg_path = self._root / f'config{suffix}{ext}'
                if cfg_path.exists():
                    self._load_file(cfg_path)

        # 环境变量覆盖 (BAOXIAO_SERVER_PORT=9000 → server.port=9000)
        for key, value in os.environ.items():
            if key.startswith('BAOXIAO_'):
                parts = key[8:].lower().split('__')
                d = self._data
                for p in parts[:-1]:
                    d = d.setdefault(p, {})
                d[parts[-1]] = self._cast_env(value)

        logger.info(f"配置已加载 (env={self._env}, root={self._root})")

    def _load_file(self, path):
        """从文件加载配置"""
        try:
            text = path.read_text(encoding='utf-8')
            if path.suffix in ('.yaml', '.yml'):
                try:
                    import yaml
                    data = yaml.safe_load(text) or {}
                except ImportError:
                    logger.warning(f"跳过 {path.name}: 缺少 PyYAML")
                    return
            else:
                data = json.loads(text)
            self._data = _deep_merge(self._data, data)
            logger.debug(f"已加载配置文件: {path}")
        except Exception as e:
            logger.warning(f"加载配置文件 {path} 失败: {e}")

    @staticmethod
    def _cast_env(value):
        """环境变量类型推断"""
        if value.lower() in ('true', 'yes', '1'):
            return True
        if value.lower() in ('false', 'no', '0'):
            return False
        try:
            return int(value)
        except ValueError:
            pass
        try:
            return float(value)
        except ValueError:
            pass
        return value

    # ---- 路径解析（相对于项目根目录）----

    def resolve_path(self, *parts):
        """解析配置中的相对路径为绝对路径"""
        return self._root.joinpath(*parts)

    def get_data_dir(self):
        return self.resolve_path(self._data['paths']['data_dir'])

    def get_export_dir(self):
        return self.resolve_path(self._data['paths']['export_dir'])

    def get_static_dir(self):
        if getattr(sys, 'frozen', False):
            return Path(sys._MEIPASS) / self._data['paths']['static_dir']
        return self.resolve_path(self._data['paths']['static_dir'])

    def get_data_file(self):
        return self.get_data_dir() / self._data['paths']['data_file']

    def get_log_file(self):
        return self.resolve_path(self._data['paths']['log_file'])

    def get_lock_file(self):
        return self.resolve_path(self._data['paths']['lock_file'])

    # ---- 快捷访问 ----

    def get(self, dotted_key, default=None):
        """点号访问: config.get('server.port') → 8765"""
        keys = dotted_key.split('.')
        d = self._data
        for k in keys:
            if isinstance(d, dict) and k in d:
                d = d[k]
            else:
                return default
        return d

    def __getitem__(self, dotted_key):
        v = self.get(dotted_key)
        if v is None:
            raise KeyError(dotted_key)
        return v

    @property
    def env(self):
        return self._env

    @property
    def root(self):
        return self._root

    def as_dict(self):
        return dict(self._data)


# 全局单例
config = Config()
