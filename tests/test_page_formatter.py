import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn
from docx.shared import Inches, Mm

from src.config import DocumentConfig
from src.formatters import PageFormatter


class TestPageFormatter(unittest.TestCase):
    def test_setup_page_format_serializes_a4_portrait_for_every_section(self):
        document = Document()
        document.add_section(WD_SECTION.NEW_PAGE)

        for section in document.sections:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Inches(11)
            section.page_height = Inches(8.5)

        PageFormatter(DocumentConfig()).setup_page_format(document)

        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / 'a4.docx'
            document.save(output_path)
            saved_document = Document(output_path)

        self.assertEqual(2, len(saved_document.sections))
        for section in saved_document.sections:
            self.assertEqual(WD_ORIENT.PORTRAIT, section.orientation)
            self.assertAlmostEqual(210, section.page_width.mm, delta=0.01)
            self.assertAlmostEqual(297, section.page_height.mm, delta=0.01)

    def test_content_width_is_derived_from_page_size_and_margins(self):
        config = DocumentConfig()
        self.assertEqual(Mm(156), config.get_content_width_emu())

    def test_content_width_honors_instance_configuration(self):
        config = DocumentConfig()
        config.PAGE_SIZE = {
            'width': Mm(250),
            'height': Mm(297)
        }
        config.PAGE_MARGINS = {
            'top': Mm(37),
            'bottom': Mm(35),
            'left': Mm(20),
            'right': Mm(20)
        }

        self.assertEqual(Mm(210), config.get_content_width_emu())

    def test_document_grid_uses_configured_content_width(self):
        class CustomPageConfig(DocumentConfig):
            PAGE_SIZE = {
                'width': Mm(180),
                'height': Mm(297)
            }
            PAGE_MARGINS = {
                'top': Mm(37),
                'bottom': Mm(35),
                'left': Mm(20),
                'right': Mm(20)
            }
            CHARS_PER_LINE = 20

        config = CustomPageConfig()
        document = Document()
        PageFormatter(config).setup_page_format(document)

        document_grid = document.sections[0]._sectPr.find(qn('w:docGrid'))
        expected_char_space = round(Mm(140).twips / config.CHARS_PER_LINE)
        self.assertEqual(str(expected_char_space), document_grid.get(qn('w:charSpace')))


if __name__ == '__main__':
    unittest.main()
