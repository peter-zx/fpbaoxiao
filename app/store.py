# -*- coding: utf-8 -*-
"""
数据存储层 — JSON 文件读写、原子写入、异常处理
职责：纯数据 I/O，不包含任何 HTTP 或 Excel 逻辑
"""

import json
import logging
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

logger = logging.getLogger(__name__)

# 数据结构骨架
EMPTY_DATA = {'expense': [], 'reimburse': []}


def _validate_record(record):
    """校验单条记录的必要字段，返回 (cleaned_record, errors)"""
    errors = []
    cleaned = dict(record)

    # 必填字段
    if not cleaned.get('time'):
        errors.append('缺少时间')
    if not cleaned.get('product'):
        errors.append('缺少产品')
    if not cleaned.get('reason'):
        errors.append('缺少原因')

    # 金额校验
    try:
        cleaned['amount'] = float(cleaned.get('amount', 0))
        if cleaned['amount'] < 0:
            errors.append('金额不能为负数')
    except (TypeError, ValueError):
        errors.append('金额格式错误')
        cleaned['amount'] = 0.0

    # ID 保证
    if not cleaned.get('id'):
        cleaned['id'] = int(datetime.now().timestamp() * 1000)

    # 勾选默认值
    if '_checked' not in cleaned:
        cleaned['_checked'] = True

    return cleaned, errors


def load(data_file: Path):
    """加载数据，返回完整 data dict"""
    if not data_file.exists():
        logger.info(f"数据文件不存在，返回空结构: {data_file}")
        return dict(EMPTY_DATA)

    try:
        text = data_file.read_text(encoding='utf-8')
        data = json.loads(text)
    except json.JSONDecodeError as e:
        logger.error(f"数据文件 JSON 解析失败: {e}")
        # 备份损坏文件
        backup = data_file.with_suffix('.json.bak')
        shutil.copy2(data_file, backup)
        logger.info(f"已备份损坏文件到: {backup}")
        return dict(EMPTY_DATA)
    except Exception as e:
        logger.error(f"加载数据失败: {e}")
        return dict(EMPTY_DATA)

    # 确保结构完整
    for key in ('expense', 'reimburse'):
        if key not in data:
            data[key] = []
        if not isinstance(data[key], list):
            logger.warning(f"data.{key} 不是列表，已重置")
            data[key] = []

    return data


def save(data_file: Path, data: dict) -> bool:
    """原子写入数据（写临时文件 → 重命名），避免写入中断导致数据丢失"""
    data_dir = data_file.parent
    data_dir.mkdir(parents=True, exist_ok=True)

    try:
        # 原子写入：先写临时文件
        tmp_fd, tmp_path = tempfile.mkstemp(
            dir=str(data_dir),
            prefix='.data_',
            suffix='.tmp'
        )
        try:
            with open(tmp_fd, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            # 写入内容失败，清理临时文件
            Path(tmp_path).unlink(missing_ok=True)
            raise

        # 重命名（原子操作）
        shutil.move(tmp_path, str(data_file))
        logger.debug(f"数据已保存: {data_file}")
        return True

    except Exception as e:
        logger.error(f"保存数据失败: {e}")
        return False


def add_record(data_file: Path, template_type: str, record: dict):
    """添加一条记录，返回 (data, errors)"""
    data = load(data_file)
    cleaned, errors = _validate_record(record)

    if errors:
        return data, errors

    if template_type not in ('expense', 'reimburse'):
        return data, [f'无效的模板类型: {template_type}']

    data[template_type].append(cleaned)
    save(data_file, data)
    return data, []


def delete_record(data_file: Path, template_type: str, index: int):
    """删除指定索引的记录，返回 (data, error_msg)"""
    data = load(data_file)

    if template_type not in ('expense', 'reimburse'):
        return data, f'无效的模板类型: {template_type}'

    records = data[template_type]
    if index < 0 or index >= len(records):
        return data, f'索引越界: {index}'

    records.pop(index)
    save(data_file, data)
    return data, None


def clear_records(data_file: Path, template_type: str = None):
    """清空记录，template_type=None 时清空全部"""
    data = load(data_file)

    if template_type:
        if template_type in data:
            data[template_type] = []
    else:
        data = dict(EMPTY_DATA)

    save(data_file, data)
    return data
