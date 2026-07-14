import unittest

from docx import Document
from docx.oxml.shared import OxmlElement

from src.utils.xpath_cache import OptimizedXMLProcessor, XPathCache


class TestXPathHelpers(unittest.TestCase):
    def test_queries_observe_elements_added_after_an_earlier_lookup(self):
        paragraph = OxmlElement('w:p')
        self.assertIsNone(XPathCache.find_first(paragraph, './w:pPr'))
        properties = OxmlElement('w:pPr')
        paragraph.append(properties)

        self.assertIs(properties, XPathCache.find_first(paragraph, './w:pPr'))

    def test_row_processing_creates_missing_row_properties(self):
        document = Document()
        row = document.add_table(rows=1, cols=1).rows[0]
        row_properties = row._tr.find('./{*}trPr')
        if row_properties is not None:
            row._tr.remove(row_properties)

        result = OptimizedXMLProcessor().process_row_properties(row)

        self.assertIsNotNone(result['trPr'])
        self.assertIs(result['trPr'], row._tr[0])


if __name__ == '__main__':
    unittest.main()
