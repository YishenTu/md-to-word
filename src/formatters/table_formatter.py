"""
表格格式化器 - 负责表格格式化和样式设置
"""

from docx.document import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.shared import qn
from docx.shared import Pt

from ..utils.xpath_cache import OptimizedXMLProcessor
from .base_formatter import BaseFormatter


class TableFormatter(BaseFormatter):
    """表格格式化器 - 负责表格格式化和样式设置"""

    def __init__(self, config=None):
        super().__init__(config)
        self.xml_processor = OptimizedXMLProcessor()

    def format_tables(self, doc: Document):
        """格式化表格，包含完整的自动适应功能"""
        for table in doc.tables:
            # 启用表格自动适应
            if self.config.TABLE_CONFIG['auto_fit']:
                table.autofit = True

                # 设置表格对齐方式为居中
                table.alignment = WD_TABLE_ALIGNMENT.CENTER

                # 通过XML设置表格自动适应窗口
                tbl = table._tbl
                tblPr = tbl.tblPr

                # 使用优化的批量查询获取所有表格属性
                elements = self.xml_processor.process_table_properties(tbl)

                # 设置表格宽度为100%
                if self.config.TABLE_CONFIG['auto_fit_mode'] == 'window':
                    tblW = elements.get('tblW')
                    if tblW is None:
                        tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
                        tblPr.append(tblW)
                    else:
                        tblW.set(qn('w:w'), str(self.config.TABLE_CONFIG['preferred_width_percent'] * 50))
                        tblW.set(qn('w:type'), 'pct')

                # 设置表格布局为自动
                tblLayout = elements.get('tblLayout')
                if tblLayout is None:
                    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="autofit"/>')
                    tblPr.append(tblLayout)
                else:
                    tblLayout.set(qn('w:type'), 'autofit')

                # 设置表格允许跨页断行
                if self.config.TABLE_CONFIG['allow_row_breaks']:
                    tblPrEx = elements.get('tblPrEx')
                    if tblPrEx is None:
                        tblPrEx = parse_xml(f'<w:tblPrEx {nsdecls("w")}><w:tblLayout w:type="autofit"/></w:tblPrEx>')
                        tbl.append(tblPrEx)

            # 应用三线表样式
            self._apply_three_line_table_style(table)

            # 格式化表格内容
            for row_index, row in enumerate(table.rows):
                # 使用优化的行属性处理
                row_props = self.xml_processor.process_row_properties(row)
                if row_props:
                    trPr = row_props['trPr']
                    if trPr is None:
                        trPr = parse_xml(f'<w:trPr {nsdecls("w")}></w:trPr>')
                        row._tr.insert(0, trPr)

                    # 设置行高规则为自动
                    trHeight = row_props['trHeight']
                    if trHeight is None:
                        trHeight = parse_xml(f'<w:trHeight {nsdecls("w")} w:hRule="auto"/>')
                        trPr.append(trHeight)
                    else:
                        trHeight.set(qn('w:hRule'), 'auto')

                for cell in row.cells:
                    # 使用优化的单元格属性处理
                    cell_props = self.xml_processor.process_cell_properties(cell)
                    tcPr = cell_props['tcPr']
                    vAlign = cell_props['vAlign']

                    if vAlign is None:
                        vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>')
                        tcPr.append(vAlign)
                    else:
                        vAlign.set(qn('w:val'), 'center')

                    # 应用三线表单元格边框
                    self._apply_three_line_cell_borders(cell, row_index, len(table.rows))

                    # 格式化单元格内的段落
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = self.config.FONT_SIZES['table']  # 使用4号字体
                            run.bold = False
                            run.italic = False
                            self._set_run_fonts(run, self.config.FONTS['fangsong'])

                        # 设置单元格段落格式
                        paragraph.alignment = self.config.ALIGNMENTS['center']  # 表格内容居中
                        paragraph_format = paragraph.paragraph_format
                        paragraph_format.space_before = Pt(0)
                        paragraph_format.space_after = Pt(0)

    def format_caption(self, paragraph) -> None:
        """Format one table caption with the shared table typography."""
        for run in paragraph.runs:
            run.font.size = self.config.FONT_SIZES['table']
            run.bold = False
            self._set_run_fonts(run, self.config.FONTS['fangsong'])

        paragraph.alignment = self.config.ALIGNMENTS['center']
        paragraph_format = paragraph.paragraph_format
        paragraph_format.space_after = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph_format.first_line_indent = Pt(0)

    def _apply_three_line_table_style(self, table):
        """应用三线表样式 - 清除默认边框"""
        tbl = table._tbl
        tblPr = tbl.tblPr

        # 移除表格默认边框
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is not None:
            tblPr.remove(tblBorders)

        # 设置表格无边框样式
        no_border_xml = f"""<w:tblBorders {nsdecls('w')}>
            <w:top w:val="none" w:sz="0"/>
            <w:left w:val="none" w:sz="0"/>
            <w:bottom w:val="none" w:sz="0"/>
            <w:right w:val="none" w:sz="0"/>
            <w:insideH w:val="none" w:sz="0"/>
            <w:insideV w:val="none" w:sz="0"/>
        </w:tblBorders>"""

        new_borders = parse_xml(no_border_xml)
        tblPr.append(new_borders)

    def _apply_three_line_cell_borders(self, cell, row_index, total_rows):
        """为单元格应用三线表边框样式"""
        tc = cell._tc
        tcPr = tc.tcPr

        if tcPr is None:
            tcPr = parse_xml(f'<w:tcPr {nsdecls("w")}></w:tcPr>')
            tc.insert(0, tcPr)

        # 移除现有边框设置
        existing_borders = tcPr.find(qn('w:tcBorders'))
        if existing_borders is not None:
            tcPr.remove(existing_borders)

        # 根据行位置设置边框
        if row_index == 0:
            # 第一行（表头）：顶部1.5磅黑色边框 + 底部0.75磅边框
            borders_xml = f"""<w:tcBorders {nsdecls('w')}>
                <w:top w:val="single" w:sz="21" w:color="000000"/>
                <w:bottom w:val="single" w:sz="9" w:color="000000"/>
                <w:left w:val="none" w:sz="0"/>
                <w:right w:val="none" w:sz="0"/>
            </w:tcBorders>"""
        elif row_index == total_rows - 1:
            # 最后一行：底部1.5磅黑色边框
            borders_xml = f"""<w:tcBorders {nsdecls('w')}>
                <w:top w:val="none" w:sz="0"/>
                <w:bottom w:val="single" w:sz="21" w:color="000000"/>
                <w:left w:val="none" w:sz="0"/>
                <w:right w:val="none" w:sz="0"/>
            </w:tcBorders>"""
        else:
            # 中间行：无边框
            borders_xml = f"""<w:tcBorders {nsdecls('w')}>
                <w:top w:val="none" w:sz="0"/>
                <w:bottom w:val="none" w:sz="0"/>
                <w:left w:val="none" w:sz="0"/>
                <w:right w:val="none" w:sz="0"/>
            </w:tcBorders>"""

        new_borders = parse_xml(borders_xml)
        tcPr.append(new_borders)
