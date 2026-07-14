"""
Utils module - 工具模块
导出常用的工具类和异常
"""

from .config_validator import ConfigValidator
from .constants import Patterns
from .exceptions import (
    FileProcessingError,
    ImageProcessingError,
    Md2WordError,
    PandocError,
    PathSecurityError,
    XMLProcessingError,
)
from .path_validator import validate_safe_path
from .xpath_cache import OptimizedXMLProcessor, XPathCache

__all__ = [
    # 常量和模式
    'Patterns',
    # 异常类
    'Md2WordError',
    'FileProcessingError',
    'PandocError',
    'ImageProcessingError',
    'XMLProcessingError',
    'PathSecurityError',
    # XML处理工具
    'XPathCache',
    'OptimizedXMLProcessor',
    # 配置验证
    'ConfigValidator',
    # 路径验证
    'validate_safe_path',
]
