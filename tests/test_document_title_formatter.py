import unittest

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn

from src.config import DocumentConfig
from src.formatters import DocumentTitleFormatter


class TestDocumentTitleFormatter(unittest.TestCase):
    def test_title_uses_its_own_exact_line_pitch(self):
        config = DocumentConfig()
        document = Document()

        DocumentTitleFormatter(config).add_document_title(document, 'Document Title')

        title_paragraph, blank_paragraph = document.paragraphs[:2]
        title_format = title_paragraph.paragraph_format
        self.assertEqual(WD_LINE_SPACING.EXACTLY, title_format.line_spacing_rule)
        self.assertAlmostEqual(
            config.TITLE_LINE_PITCH.pt,
            title_format.line_spacing.pt,
            delta=0.05,
        )
        title_snap = title_paragraph._element.pPr.find(qn('w:snapToGrid'))
        blank_snap = blank_paragraph._element.pPr.find(qn('w:snapToGrid'))
        self.assertEqual('false', title_snap.get(qn('w:val')))
        self.assertEqual('true', blank_snap.get(qn('w:val')))


if __name__ == '__main__':
    unittest.main()
