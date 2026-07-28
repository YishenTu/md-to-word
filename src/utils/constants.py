"""
共享常量模块 - 存放各个模块共用的正则表达式模式和常量
避免代码重复，保证一致性
"""

import re


class ControlTokens:
    """Reserved position-only tokens shared by Markdown and Word pipeline stages."""

    PAGE_BREAK = '[PAGEBREAK]'
    SIGNATURE = '[MD_TO_WORD_SIGNATURE]'


# 正则表达式模式 - 预编译提高性能
class Patterns:
    """预编译的正则表达式模式集合"""

    # 标题相关模式
    HEADING_PATTERNS = [
        re.compile(r'^[一二三四五六七八九十]+、'),  # 一、二、三、
        re.compile(r'^（[一二三四五六七八九十]+）'),  # （一）（二）（三）
        re.compile(r'^[0-9]+\.'),  # 1. 2. 3.
        re.compile(r'^[0-9]+、'),  # 1、2、3、
    ]

    # 列表相关模式
    UNORDERED_LIST_PATTERN = re.compile(r'^\s*\*\s+')  # 无序列表模式 "* "
    UNORDERED_LIST_REPLACE_PATTERN = re.compile(r'^(\s*)\*\s+')  # 无序列表替换模式

    # 图片相关模式
    MARKDOWN_IMAGE_PATTERN = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')  # ![alt](url)
    OBSIDIAN_IMAGE_PATTERN = re.compile(r'!\[\[([^\]]+)\]\]')  # ![[filename]]

    # 数学公式模式
    LATEX_INLINE_MATH_PATTERN = re.compile(r'\$[^$]+\$')  # 行内LaTeX数学公式 $...$
    LATEX_BLOCK_MATH_PATTERN = re.compile(r'\$\$[^$]+\$\$')  # 块级LaTeX数学公式 $$...$$

    # 表格相关模式
    TABLE_ROW_PATTERN = re.compile(r'^\s*\|')  # 表格行 "|..."

    # 图表标题模式 - 匹配 图/图片/表/表格/图表 + 可选空格 + 数字 + 可选空格 + 标点(:：.) + 描述
    CAPTION_PATTERN = re.compile(r'^(图片?|表格?|图表)\s*(\d+)\s*[:：.]\s*(.*)$')
    CAPTION_PREFIX_PATTERN = re.compile(r'^(图片?|表格?|图表)\s*(\d+)\s*[:：.]\s*')  # 仅匹配前缀部分

    # 多级编号模式
    MULTI_LEVEL_NUMBER_PATTERN = re.compile(r'^(\d+\.\d+(?:\.\d+)*)\s+(.+)$')  # 2.1.1 内容
    SIMPLE_ORDERED_LIST_WITH_CONTENT = re.compile(r'^(\s*)(\d+)\.\s+(.+)$')  # 1. 内容（捕获内容）

    # 附件说明模式
    ATTACHMENT_HEADER_PATTERN = re.compile(r'^\s*附件[:：]\s*$')
    ATTACHMENT_INLINE_PATTERN = re.compile(r'^\s*附件[:：]\s*\d+\.\s+.+?\s*$')
    ATTACHMENT_ITEM_PATTERN = re.compile(r'^\s*(\d+)\.\s+(.+?)\s*$')

    # 图片语法清理模式
    OBSIDIAN_IMAGE_CLEANUP_PATTERN = re.compile(r'!\[\[[^\]]+\]\]')  # 清理 ![[filename]]
    MARKDOWN_IMAGE_CLEANUP_PATTERN = re.compile(r'!\[[^\]]*\]\([^)]+\)')  # 清理 ![alt](path)

    # 通用文本清理模式
    WHITESPACE_CLEANUP_PATTERN = re.compile(r'\s+')  # 清理多余空格
