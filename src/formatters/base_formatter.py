"""
基础格式化器类 - 提供共同的功能和配置
"""

from docx.oxml.shared import OxmlElement, qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from ..config import DocumentConfig
from ..utils.constants import Patterns


class BaseFormatter:
    """格式化器基类，提供共同的功能和配置"""

    def __init__(self, config: DocumentConfig | None = None):
        self.config = config or DocumentConfig()

    def _set_run_fonts(self, run: Run, east_asia_font: str) -> None:
        """Set Latin and East Asian typefaces independently on one run."""
        rPr = run._element.rPr
        if rPr is None:
            rPr = OxmlElement('w:rPr')
            run._element.insert(0, rPr)

        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)

        latin_font = self.config.FONTS['latin']
        rFonts.set(qn('w:ascii'), latin_font)
        rFonts.set(qn('w:hAnsi'), latin_font)
        rFonts.set(qn('w:cs'), latin_font)
        rFonts.set(qn('w:eastAsia'), east_asia_font)
        rFonts.set(qn('w:hint'), 'eastAsia')

    def _has_math_formula(self, paragraph: Paragraph) -> bool:
        """检查段落是否包含数学公式（统一方法，避免重复）"""
        try:
            # 1. 检查段落的XML内容是否包含MathML元素（转换后的数学公式）
            xml_str = paragraph._element.xml
            if 'oMath' in xml_str or 'oMathPara' in xml_str:
                return True

            # 2. 检查段落文本是否包含LaTeX格式的数学公式（原始格式）
            text = paragraph.text
            if text:
                # 检查行内数学公式 $...$
                if Patterns.LATEX_INLINE_MATH_PATTERN.search(text):
                    return True
                # 检查块级数学公式 $$...$$
                if Patterns.LATEX_BLOCK_MATH_PATTERN.search(text):
                    return True

            return False
        except Exception as e:
            # 记录错误但不中断处理
            import logging

            logging.debug(f'检查数学公式时出错: {e}')
            return False

    def _enable_snap_to_grid(self, paragraph: Paragraph) -> None:
        """启用段落的文档网格对齐"""
        self._set_snap_to_grid(paragraph, True)

    def _disable_snap_to_grid(self, paragraph: Paragraph) -> None:
        """Disable document-grid alignment for paragraphs with a custom pitch."""
        self._set_snap_to_grid(paragraph, False)

    @staticmethod
    def _set_snap_to_grid(paragraph: Paragraph, enabled: bool) -> None:
        """Set the OOXML document-grid alignment flag on one paragraph."""
        from docx.oxml.ns import qn as qn_func
        from docx.oxml.shared import OxmlElement

        pPr = paragraph._element.get_or_add_pPr()

        # 检查是否已有snapToGrid元素
        snapToGrid = pPr.find(qn_func('w:snapToGrid'))
        if snapToGrid is None:
            # 创建新的snapToGrid元素
            snapToGrid = OxmlElement('w:snapToGrid')
            pPr.append(snapToGrid)

        snapToGrid.set(qn_func('w:val'), 'true' if enabled else 'false')
