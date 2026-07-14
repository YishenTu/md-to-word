"""
md-to-word package - Markdown到Word公文格式转换工具
符合GB/T 9704-2012《党政机关公文格式》国家标准
"""

__version__ = '2.7.0'
__author__ = 'md-to-word Project'
__description__ = 'Markdown到Word公文格式转换工具'

# 导出主要组件
from .config import DocumentConfig
from .utils import (
    ConfigValidator,
    FileProcessingError,
    ImageProcessingError,
    Md2WordError,
    OptimizedXMLProcessor,
    PandocError,
    PathSecurityError,
    Patterns,
    XMLProcessingError,
    XPathCache,
)

__all__ = [
    # 配置
    'DocumentConfig',
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
]
