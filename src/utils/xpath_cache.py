"""XPath helpers for mutable WordprocessingML element trees."""

from typing import Any

from docx.oxml.shared import OxmlElement


class XPathCache:
    """Centralized XPath helpers for mutable WordprocessingML trees."""

    @staticmethod
    def find_first(element: Any, xpath: str) -> Any | None:
        """Return the first match without caching mutable XML state.

        Args:
            element: XML元素
            xpath: XPath查询字符串

        Returns:
            第一个匹配的元素或None
        """
        results = element.xpath(xpath)
        return results[0] if results else None

    @staticmethod
    def find_or_create(parent: Any, xpath: str, tag_name: str, namespace: str | None = None) -> Any:
        """
        查找元素，如果不存在则创建

        Args:
            parent: 父元素
            xpath: XPath查询字符串
            tag_name: 标签名称
            namespace: XML命名空间

        Returns:
            找到的或新创建的元素
        """
        element = XPathCache.find_first(parent, xpath)
        if element is None:
            from docx.oxml.ns import nsdecls
            from docx.oxml.parser import parse_xml

            if namespace:
                element = parse_xml(f'<{tag_name} {nsdecls(namespace)}/>')
            else:
                element = OxmlElement(tag_name)
            parent.append(element)
        return element

    @staticmethod
    def batch_query(element: Any, queries: dict) -> dict:
        """
        批量执行XPath查询，减少重复遍历

        Args:
            element: XML元素
            queries: 查询字典 {key: xpath}

        Returns:
            结果字典 {key: result}
        """
        results = {}
        for key, xpath in queries.items():
            results[key] = XPathCache.find_first(element, xpath)
        return results


class OptimizedXMLProcessor:
    """优化的XML处理器，减少DOM遍历次数"""

    def __init__(self):
        self.cache = XPathCache()

    def process_table_properties(self, tbl):
        """优化的表格属性处理 - 单次遍历获取所有需要的元素"""
        tblPr = tbl.tblPr

        # 批量查询所有需要的元素
        queries = {'tblW': './/w:tblW', 'tblLayout': './/w:tblLayout', 'tblPrEx': './/w:tblPrEx'}

        elements = self.cache.batch_query(tblPr, queries)
        return elements

    def process_row_properties(self, row):
        """优化的行属性处理"""
        if not hasattr(row, '_tr'):
            return None

        # 先获取trPr
        trPr = self.cache.find_first(row._tr, './w:trPr')
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row._tr.insert(0, trPr)

        # 再查询trHeight
        trHeight = self.cache.find_first(trPr, './/w:trHeight')
        return {'trPr': trPr, 'trHeight': trHeight}

    def process_cell_properties(self, cell):
        """优化的单元格属性处理"""
        tc = cell._tc

        # 获取或创建tcPr
        tcPr = self.cache.find_first(tc, './/w:tcPr')
        if tcPr is None:
            from docx.oxml.ns import nsdecls
            from docx.oxml.parser import parse_xml

            tcPr = parse_xml(f'<w:tcPr {nsdecls("w")}></w:tcPr>')
            tc.insert(0, tcPr)

        # 获取vAlign
        vAlign = self.cache.find_first(tcPr, './/w:vAlign')

        return {'tcPr': tcPr, 'vAlign': vAlign}

    def find_drawings_in_paragraphs(self, paragraphs):
        """优化的图片查找 - 单次遍历所有段落"""
        drawings_map = {}

        for i, paragraph in enumerate(paragraphs):
            if hasattr(paragraph, '_element'):
                drawings = paragraph._element.xpath('.//w:drawing')
                if drawings:
                    drawings_map[i] = drawings

        return drawings_map

    def process_image_properties(self, drawing_element):
        """优化的图片属性处理 - 批量获取所有相关元素"""
        queries = {
            'docPr': './/wp:docPr',
            'cNvPr': './/pic:cNvPr',
            'blip': './/a:blip',
            'inline': './/wp:inline',
            'extent': './/wp:extent',
            'graphic': './/a:graphic',
        }

        # 批量查询所有元素
        elements = {}
        for key, xpath in queries.items():
            results = drawing_element.xpath(xpath)
            elements[key] = results if results else []

        return elements
