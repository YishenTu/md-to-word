"""
自定义异常类模块 - 提供具体的异常类型以改进错误处理
"""


class Md2WordError(Exception):
    """Markdown到Word转换的基础异常类"""

    pass


class FileProcessingError(Md2WordError):
    """文件处理相关错误"""

    pass


class PandocError(Md2WordError):
    """Pandoc处理相关错误"""

    pass


class ImageProcessingError(Md2WordError):
    """图片处理相关错误"""

    pass


class XMLProcessingError(Md2WordError):
    """XML处理相关错误"""

    pass


class PathSecurityError(Md2WordError):
    """路径安全相关错误"""

    pass
