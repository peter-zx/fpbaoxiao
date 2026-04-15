# -*- coding: utf-8 -*-
"""
图片处理 — base64 解析、临时文件管理
职责：纯图片 I/O，不包含 HTTP 或 Excel 逻辑
"""

import base64
import io
import logging
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

# 依赖检查
try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    logger.error("缺少 Pillow，请运行: pip install Pillow")


def decode_base64(base64_data: str):
    """解析 base64 图片字符串 → PIL Image 对象
    
    Args:
        base64_data: data:image/png;base64,... 或纯 base64 字符串
    
    Returns:
        PILImage 对象，或 None
    """
    if not base64_data or not HAS_PIL:
        return None

    try:
        if ',' in base64_data:
            _, data = base64_data.split(',', 1)
        else:
            data = base64_data

        img_bytes = base64.b64decode(data)
        pil_img = PILImage.open(io.BytesIO(img_bytes))
        return pil_img
    except Exception as e:
        logger.error(f"解析 base64 图片失败: {e}")
        return None


def save_to_file(pil_img, directory: Path, prefix: str = 'img', 
                 index: int = 0, fmt: str = 'PNG') -> Path:
    """将 PIL Image 保存到文件
    
    Args:
        pil_img: PIL Image 对象
        directory: 目标目录
        prefix: 文件名前缀
        index: 序号
        fmt: 图片格式 (PNG/JPEG)
    
    Returns:
        保存后的文件路径
    """
    directory.mkdir(parents=True, exist_ok=True)
    ext = 'png' if fmt.upper() == 'PNG' else 'jpg'
    filename = f'{prefix}_{index}.{ext}'
    path = directory / filename

    try:
        pil_img.save(str(path), format=fmt)
        return path
    except Exception as e:
        logger.error(f"保存图片失败 {path}: {e}")
        raise


class ImageTempDir:
    """图片临时目录管理器（上下文管理器）
    
    用法:
        with ImageTempDir('export_xxx') as tmp:
            path = tmp.save(pil_img, 'expense', 0)
            ... 使用 path ...
        # 退出时自动清理
    """

    def __init__(self, parent_dir: Path, prefix: str = 'img_tmp'):
        self._parent = parent_dir
        self._prefix = prefix
        self._path = None

    def __enter__(self):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self._path = self._parent / f'_img_tmp_{self._prefix}_{timestamp}'
        self._path.mkdir(parents=True, exist_ok=True)
        return self

    def __exit__(self, *exc):
        self.cleanup()
        return False

    @property
    def path(self) -> Path:
        return self._path

    def save(self, pil_img, prefix: str = 'img', index: int = 0) -> Path:
        """保存图片并返回路径"""
        return save_to_file(pil_img, self._path, prefix, index)

    def cleanup(self):
        """清理临时目录"""
        if self._path and self._path.exists():
            try:
                shutil.rmtree(self._path, ignore_errors=True)
            except Exception as e:
                logger.warning(f"清理临时目录失败: {e}")


def prepare_images(records: list, prefix: str, tmp_dir: Path):
    """准备图片：从 base64 解码并保存到临时目录
    
    Args:
        records: 记录列表（每条含 image 字段）
        prefix: 文件名前缀 ('expense' / 'reimburse')
        tmp_dir: 临时目录
    
    Returns:
        dict: {excel_row: (img_path, pil_img)}
               excel_row = 数据行在 Excel 中的 1-based 行号
    """
    img_map = {}
    for idx, record in enumerate(records):
        image_data = record.get('image')
        if not image_data:
            continue

        try:
            pil_img = decode_base64(image_data)
            if pil_img:
                img_path = save_to_file(pil_img, tmp_dir, prefix, idx)
                # Excel 行号: header=1, 数据从 2 开始
                img_map[idx + 2] = (str(img_path), pil_img)
        except Exception as e:
            logger.error(f"准备图片失败 (record {idx}): {e}")

    return img_map
