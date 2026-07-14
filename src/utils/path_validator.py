"""
路径验证工具模块 - 提供安全的路径验证功能
"""

import os
from pathlib import Path

from .exceptions import PathSecurityError


def validate_safe_path(path: str, base_dir: str | None = None, allow_absolute: bool = True) -> Path:
    """
    验证路径安全性，防止路径遍历攻击

    Args:
        path: 要验证的路径
        base_dir: 基础目录，如果提供则验证路径是否在基础目录内
        allow_absolute: 是否允许绝对路径

    Returns:
        Path: 安全的Path对象

    Raises:
        PathSecurityError: 如果路径不安全
    """
    if not path:
        raise PathSecurityError('路径不能为空')

    try:
        # Normalize user-selected paths before applying policy checks.
        path_obj = Path(path).expanduser()

        # 检查路径是否为绝对路径
        if path_obj.is_absolute() and not allow_absolute:
            raise PathSecurityError(f'不允许使用绝对路径: {path}')

        # 解析路径（包括符号链接）
        resolved_path = path_obj.resolve()

        # 如果提供了基础目录，确保路径在基础目录内
        if base_dir:
            base_path = Path(base_dir).expanduser().resolve()
            if not resolved_path.is_relative_to(base_path):
                raise PathSecurityError(f'路径不在允许的目录内: {resolved_path}')

        # 检查路径是否包含特殊字符（Windows）
        if os.name == 'nt':
            invalid_chars = '<>:"|?*'
            filename = resolved_path.name
            if any(char in filename for char in invalid_chars):
                raise PathSecurityError(f'文件名包含无效字符: {filename}')

        return resolved_path

    except PathSecurityError:
        raise
    except Exception as error:
        raise PathSecurityError(f'路径验证失败: {error}') from error
