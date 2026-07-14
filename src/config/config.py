import os
from pathlib import Path
from typing import TypedDict

from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Mm, Pt


class UnorderedListConfig(TypedDict):
    marker: str
    marker_position_chars: int
    text_gap_chars: int
    nested_step_chars: int


class TableConfig(TypedDict):
    auto_fit: bool
    auto_fit_mode: str
    preferred_width_percent: int
    allow_row_breaks: bool


class PandocConfig(TypedDict):
    math_method: str
    table_style: str
    table_caption: bool
    list_style: str
    extra_args: list[str]
    image_wrap_text: bool
    image_wrap_type: str


class ObsidianConfig(TypedDict):
    vault_name: str | None
    attachments_folder: str
    vault_path: str | None


class ImageConfig(TypedDict):
    search_paths: list[str]
    supported_formats: list[str]
    copy_images: bool
    output_dir: str


class DocumentConfig:
    """公文格式配置类，基于GB/T 9704-2012标准"""

    # A4 paper dimensions (units: EMU-backed Length values)
    PAGE_SIZE = {'width': Mm(210), 'height': Mm(297)}
    PAGE_ORIENTATION = WD_ORIENT.PORTRAIT

    # Word page margins recommended by GB/T 33476.2-2016.
    PAGE_MARGINS = {
        'top': Mm(34.58),
        'bottom': Mm(32.58),
        'left': Mm(28),
        'right': Mm(26),
    }

    # 字体配置
    FONTS = {
        'fangsong': 'FangSong',  # 正文字体
        'xiaobiaosong': 'FZXiaoBiaoSong-B05S',  # 标题字体
        'heiti': 'SimHei',  # 一级标题
        'kaiti': 'STKaiTi',  # 二级标题
        'songti': 'SimSun',  # 页码字体
        'latin': 'Times New Roman',  # 拉丁文字和数字
    }

    # 字号配置 (单位: 磅)
    FONT_SIZES = {
        'title': Pt(22),  # 二号 - 标题
        'body': Pt(16),  # 三号 - 正文
        'table': Pt(12),  # 四号 - 表格
        'page_num': Pt(14),  # 四号 - 页码
        'header': Pt(16),  # 三号 - 发文字号
    }

    # 版心配置
    CHARS_PER_LINE = 28  # 每行字符数
    LINES_PER_PAGE = 22  # 每页行数
    BODY_LINE_PITCH = Mm(10.39)
    TITLE_LINE_PITCH = Mm(12.51)

    # 段落缩进
    FIRST_LINE_INDENT = Pt(32)  # 首行缩进2字符

    # Unordered-list positions are measured in body-character cells.
    UNORDERED_LIST: UnorderedListConfig = {
        'marker': '•',
        'marker_position_chars': 2,
        'text_gap_chars': 1,
        'nested_step_chars': 2,
    }

    # 对齐方式
    ALIGNMENTS = {
        'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
        'justify': WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
        'left': WD_PARAGRAPH_ALIGNMENT.LEFT,
        'right': WD_PARAGRAPH_ALIGNMENT.RIGHT,
    }

    # 表格自动适应配置
    TABLE_CONFIG: TableConfig = {
        'auto_fit': True,  # 启用表格自动适应
        'auto_fit_mode': 'window',  # 适应模式：'window'(窗口)
        'preferred_width_percent': 100,  # 首选宽度百分比
        'allow_row_breaks': True,  # 允许跨页断行
    }

    # Pandoc相关配置
    PANDOC_CONFIG: PandocConfig = {
        # 数学公式处理方式
        'math_method': 'mathml',  # 使用MathML渲染数学公式
        # 表格处理
        'table_style': 'grid',  # 表格样式
        'table_caption': True,  # 是否显示表格标题
        # 列表处理
        'list_style': 'chinese',  # 中文列表样式
        # 其他pandoc参数
        'extra_args': ['--fail-if-warnings', '--preserve-tabs', '--wrap=none'],
        # 图片文字环绕设置
        'image_wrap_text': True,  # 是否启用图片文字环绕
        'image_wrap_type': 'topAndBottom',  # 环绕类型：topAndBottom, square, tight, through, none
    }

    # Obsidian路径配置
    OBSIDIAN_CONFIG: ObsidianConfig = {
        # Obsidian Vault名称（用户可配置）
        'vault_name': os.getenv('OBSIDIAN_VAULT_NAME'),
        # 附件文件夹名称（用户可配置）
        'attachments_folder': os.getenv('OBSIDIAN_ATTACHMENTS_FOLDER', '- Attachments'),
        # 完整Vault路径（如果指定，优先使用此路径）
        'vault_path': os.getenv('OBSIDIAN_VAULT_PATH', None),
    }

    # 图片路径配置
    IMAGE_CONFIG: ImageConfig = {
        # 动态生成的搜索路径列表（通过 _build_search_paths() 方法构建）
        'search_paths': [],  # 将在运行时动态填充
        # 支持的图片格式
        'supported_formats': ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp'],
        # 是否复制图片到输出目录
        'copy_images': True,
        # 图片输出目录（相对于Word文档）
        'output_dir': 'images',
    }

    @classmethod
    def _build_search_paths(cls) -> list[str]:
        """根据Obsidian配置动态构建图片搜索路径"""
        paths: list[str] = []
        config = cls.OBSIDIAN_CONFIG
        vault_path_value = config['vault_path']
        vault_name = config['vault_name']
        attachments_folder = config['attachments_folder']

        # 优先使用用户指定的完整Vault路径
        if vault_path_value:
            vault_path = Path(vault_path_value)
            if vault_path.exists():
                attachments_path = vault_path / attachments_folder
                if attachments_path.exists():
                    paths.append(str(attachments_path))
                paths.append(str(vault_path))

        # 自动检测常见的Obsidian路径
        elif vault_name:
            # 检测路径列表：iCloud、Documents、Desktop
            search_locations = [
                Path.home() / 'Library/Mobile Documents/iCloud~md~obsidian/Documents' / vault_name,
                Path.home() / 'Documents' / vault_name,
                Path.home() / 'Desktop' / vault_name,
            ]

            for vault_path in search_locations:
                if vault_path.exists():
                    attachments_path = vault_path / attachments_folder
                    if attachments_path.exists():
                        paths.append(str(attachments_path))
                    paths.append(str(vault_path))
                    break

        # 添加标准备选路径
        paths.extend(['./images', './assets', './'])

        return paths

    @classmethod
    def get_image_search_paths(cls) -> list[str]:
        """获取图片搜索路径，如果未初始化则动态构建"""
        if not cls.IMAGE_CONFIG['search_paths']:
            cls.IMAGE_CONFIG['search_paths'] = cls._build_search_paths()
        return cls.IMAGE_CONFIG['search_paths']

    def get_content_width_emu(self) -> int:
        """Return the usable page width in English Metric Units (EMU)."""
        return int(self.PAGE_SIZE['width'] - self.PAGE_MARGINS['left'] - self.PAGE_MARGINS['right'])

    def get_content_height_emu(self) -> int:
        """Return the usable page height in English Metric Units (EMU)."""
        return int(self.PAGE_SIZE['height'] - self.PAGE_MARGINS['top'] - self.PAGE_MARGINS['bottom'])
