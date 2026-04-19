# -*- coding: utf-8 -*-
"""
store.py — 数据存储层
职责：纯数据 I/O、原子写入、异常处理、用户管理、记录增删改查
"""

import json
import logging
import tempfile
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

# =============================================================================
# 数据结构骨架
# =============================================================================
EMPTY_DATA = {
    "users": [],       # 用户列表
    "expense": [],     # 报销记录
    "reimburse": []    # 借款记录
}

# =============================================================================
# 基础数据 I/O（原子写入）
# =============================================================================

def _load_raw(data_file: Path) -> dict:
    """加载原始数据字典"""
    if not data_file.exists():
        logger.info(f"数据文件不存在，返回空结构: {data_file}")
        return dict(EMPTY_DATA)
    try:
        text = data_file.read_text(encoding='utf-8')
        data = json.loads(text)
    except json.JSONDecodeError as e:
        logger.error(f"数据文件 JSON 解析失败: {e}")
        backup = data_file.with_suffix('.json.bak')
        shutil.copy2(data_file, backup)
        logger.info(f"已备份损坏文件到: {backup}")
        return dict(EMPTY_DATA)
    except Exception as e:
        logger.error(f"加载数据失败: {e}")
        return dict(EMPTY_DATA)

    # 确保结构完整
    for key in ('users', 'expense', 'reimburse'):
        if key not in data:
            data[key] = []
        if not isinstance(data[key], list):
            data[key] = []

    return data


def _save_raw(data_file: Path, data: dict) -> bool:
    """原子写入数据（写临时文件 → 重命名），避免写入中断导致数据丢失"""
    data_dir = data_file.parent
    data_dir.mkdir(parents=True, exist_ok=True)

    try:
        tmp_fd, tmp_path = tempfile.mkstemp(
            dir=str(data_dir),
            prefix='.data_',
            suffix='.tmp'
        )
        try:
            with open(tmp_fd, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            Path(tmp_path).unlink(missing_ok=True)
            raise

        shutil.move(tmp_path, str(data_file))
        logger.debug(f"数据已保存: {data_file}")
        return True

    except Exception as e:
        logger.error(f"保存数据失败: {e}")
        return False


# =============================================================================
# 用户管理
# =============================================================================

def load(data_file: Path) -> dict:
    """加载完整数据（兼容旧接口）"""
    return _load_raw(data_file)


def save(data_file: Path, data: dict) -> bool:
    """保存完整数据（兼容旧接口）"""
    return _save_raw(data_file, data)


def save_user_records(data_file: Path, user_id, expense_records: list, reimburse_records: list) -> bool:
    """
    仅替换指定用户的记录，保留其他用户数据和 users 数组。
    前端 saveToServer() 应调用此接口，避免全量覆盖导致 users 丢失。
    """
    data = _load_raw(data_file)

    # 给每条记录打上 user_id 标记
    for r in expense_records:
        r['user_id'] = user_id
    for r in reimburse_records:
        r['user_id'] = user_id

    # 保留其他用户的记录，替换当前用户的
    data['expense']   = [r for r in data.get('expense', [])   if r.get('user_id') != user_id] + expense_records
    data['reimburse'] = [r for r in data.get('reimburse', []) if r.get('user_id') != user_id] + reimburse_records

    return _save_raw(data_file, data)


def get_user_by_username(data_file: Path, username: str) -> Optional[dict]:
    """根据用户名查找用户，返回用户 dict 或 None"""
    data = _load_raw(data_file)
    for user in data.get('users', []):
        if user.get('username', '').strip() == username.strip():
            return user
    return None


def get_user_by_id(data_file: Path, user_id: str) -> Optional[dict]:
    """根据用户ID查找用户"""
    data = _load_raw(data_file)
    for user in data.get('users', []):
        if user.get('id') == user_id:
            return user
    return None


def register_user(data_file: Path, username: str, password_hash: str) -> Tuple[Optional[dict], str]:
    """
    注册新用户。
    返回: (用户 dict, error_msg)
    """
    username = username.strip()
    if not username:
        return None, "用户名不能为空"
    if len(username) < 2:
        return None, "用户名至少2个字符"
    if len(username) > 20:
        return None, "用户名最多20个字符"

    # 检查是否重名
    data = _load_raw(data_file)
    if get_user_by_username(data_file, username):
        return None, "用户名已存在"

    user_id = int(datetime.now().timestamp() * 1000)
    new_user = {
        "id": user_id,
        "username": username,
        "password": password_hash,
        "created_at": datetime.now().strftime("%Y/%m/%d %H:%M"),
        "last_login": None
    }
    data['users'].append(new_user)
    if not _save_raw(data_file, data):
        return None, "保存用户失败，请重试"
    logger.info(f"新用户注册成功: {username} (id={user_id})")
    return new_user, ""


def update_last_login(data_file: Path, user_id) -> bool:
    """更新用户最后登录时间"""
    data = _load_raw(data_file)
    for user in data.get('users', []):
        if user.get('id') == user_id:
            user['last_login'] = datetime.now().strftime("%Y/%m/%d %H:%M")
            return _save_raw(data_file, data)
    return False


# =============================================================================
# 记录管理（加 user_id 隔离）
# =============================================================================

def _validate_record(record):
    """校验单条记录的必要字段"""
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


def add_record(data_file: Path, template_type: str, record: dict, user_id: str = None) -> Tuple[dict, list]:
    """
    添加一条记录，返回 (data, errors)
    如果传入 user_id，记录会关联该用户
    """
    data = _load_raw(data_file)
    cleaned, errors = _validate_record(record)

    if errors:
        return data, errors

    if template_type not in ('expense', 'reimburse'):
        return data, [f'无效的模板类型: {template_type}']

    # 关联用户
    if user_id:
        cleaned['user_id'] = user_id

    data[template_type].append(cleaned)
    _save_raw(data_file, data)
    return data, []


def delete_record(data_file: Path, template_type: str, index: int, user_id: str = None) -> Tuple[dict, str]:
    """删除指定索引的记录。返回 (data, error_msg)"""
    data = _load_raw(data_file)

    if template_type not in ('expense', 'reimburse'):
        return data, f'无效的模板类型: {template_type}'

    records = data[template_type]
    if index < 0 or index >= len(records):
        return data, f'索引越界: {index}'

    # 如果传了 user_id，验证该记录是否属于该用户（安全隔离）
    if user_id is not None:
        if records[index].get('user_id') != user_id:
            return data, '无权删除此记录'

    records.pop(index)
    _save_raw(data_file, data)
    return data, None


def clear_records(data_file: Path, template_type: str = None, user_id: str = None) -> dict:
    """
    清空记录。
    - template_type=None → 清空全部（管理员用）
    - user_id 有值 → 只清空该用户的记录（普通用户清空自己的）
    """
    data = _load_raw(data_file)

    if user_id is None:
        # 清空全部
        if template_type:
            if template_type in data:
                data[template_type] = []
        else:
            data = dict(EMPTY_DATA)
    else:
        # 只清空当前用户的记录
        for key in ('expense', 'reimburse'):
            if template_type is None or template_type == key:
                data[key] = [r for r in data.get(key, []) if r.get('user_id') != user_id]

    _save_raw(data_file, data)
    return data


# =============================================================================
# 数据过滤（按 user_id）
# =============================================================================

def get_records_for_user(data_file: Path, user_id: str, template_type: str = None) -> dict:
    """
    返回只包含指定用户记录的 data dict。
    - 如果 user_id 为 None，返回所有记录（管理员用）
    - template_type 有值，只过滤该类型
    """
    data = _load_raw(data_file)
    result = {"users": data.get("users", []), "expense": [], "reimburse": []}

    types_to_filter = [template_type] if template_type else ['expense', 'reimburse']

    for key in types_to_filter:
        if user_id is None:
            result[key] = data.get(key, [])
        else:
            result[key] = [r for r in data.get(key, []) if r.get('user_id') == user_id]

    return result


# =============================================================================
# 用户数据统计（给管理页用）
# =============================================================================

def get_user_stats(data_file: Path) -> list:
    """返回所有用户的统计数据列表"""
    data = _load_raw(data_file)
    stats = []

    for user in data.get('users', []):
        uid = user.get('id')
        username = user.get('username', '')
        expense_count = sum(1 for r in data.get('expense', []) if r.get('user_id') == uid)
        reimburse_count = sum(1 for r in data.get('reimburse', []) if r.get('user_id') == uid)
        expense_total = sum(float(r.get('amount', 0)) for r in data.get('expense', []) if r.get('user_id') == uid)
        reimburse_total = sum(float(r.get('amount', 0)) for r in data.get('reimburse', []) if r.get('user_id') == uid)
        stats.append({
            "id": uid,
            "username": username,
            "created_at": user.get('created_at', ''),
            "last_login": user.get('last_login', ''),
            "expense_count": expense_count,
            "reimburse_count": reimburse_count,
            "expense_total": round(expense_total, 2),
            "reimburse_total": round(reimburse_total, 2),
        })

    return stats


def delete_user(data_file: Path, user_id) -> Tuple[bool, str]:
    """删除用户及其所有记录"""
    data = _load_raw(data_file)

    # 删除该用户
    original_len = len(data['users'])
    data['users'] = [u for u in data['users'] if u.get('id') != user_id]
    if len(data['users']) == original_len:
        return False, "用户不存在"

    # 删除该用户的记录
    for key in ('expense', 'reimburse'):
        data[key] = [r for r in data.get(key, []) if r.get('user_id') != user_id]

    if not _save_raw(data_file, data):
        return False, "删除时保存失败"

    logger.info(f"用户 id={user_id} 及其所有记录已删除")
    return True, ""
