import tempfile
import unittest
from pathlib import Path

from src.core.markdown_preprocessor import MarkdownPreprocessor
from src.utils.constants import ControlTokens


class TestSignatureBlockParser(unittest.TestCase):
    def setUp(self):
        self.preprocessor = MarkdownPreprocessor()

    def _preprocess_file(self, source: str):
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory) / 'notice.md'
            input_path.write_text(source, encoding='utf-8')
            return self.preprocessor.preprocess_file(str(input_path))

    def test_extracts_and_normalizes_a_strict_final_signature_block(self):
        source = '---\ntags:\n  - example\n---\n\n# Notice\n\nBody.\n\nExample Authority\n2026-07-14\n'

        result = self._preprocess_file(source)

        self.assertEqual('Example Authority', result['signatory'])
        self.assertEqual(
            '2026\u5e747\u670814\u65e5',
            result['document_date'],
        )
        self.assertEqual(f'Body.\n\n{ControlTokens.SIGNATURE}', result['content'])

    def test_allows_blank_lines_between_signatory_and_date(self):
        source = 'Body.\n\nExample Authority\n\n2026-07-14'

        result = self._preprocess_file(source)

        self.assertEqual('Example Authority', result['signatory'])
        self.assertEqual('2026\u5e747\u670814\u65e5', result['document_date'])
        self.assertEqual(f'Body.\n\n{ControlTokens.SIGNATURE}', result['content'])

    def test_extracts_signature_before_a_terminal_attachment_declaration(self):
        source = (
            'Body.\n\n'
            'Example Authority\n'
            '2026-07-14\n\n'
            '附件：\n\n'
            '1. Implementation plan\n'
            '2. Acceptance checklist\n\n'
            '---\n'
            '#work\n'
        )

        result = self._preprocess_file(source)

        self.assertEqual('Example Authority', result['signatory'])
        self.assertEqual('2026\u5e747\u670814\u65e5', result['document_date'])
        self.assertEqual(
            (
                f'Body.\n\n{ControlTokens.SIGNATURE}\n\n'
                f'{ControlTokens.ATTACHMENT_FIRST_ITEM}附件：1. Implementation plan\n\n'
                f'{ControlTokens.ATTACHMENT_ITEM}2. Acceptance checklist'
            ),
            result['content'],
        )

    def test_trailing_note_metadata_is_removed_before_final_signature_detection(self):
        source = 'Body.\n\nExample Authority\n2026-07-14\n\n---\n#work\n'

        result = self._preprocess_file(source)

        self.assertEqual('Example Authority', result['signatory'])
        self.assertEqual('2026\u5e747\u670814\u65e5', result['document_date'])
        self.assertNotIn('#work', result['content'])
        self.assertTrue(result['content'].endswith(ControlTokens.SIGNATURE))

    def test_extracts_signature_before_a_compact_terminal_attachment_declaration(self):
        source = 'Body.\n\nExample Authority\n2026-07-14\n\n附件：1. Implementation plan'

        result = self._preprocess_file(source)

        self.assertEqual('Example Authority', result['signatory'])
        self.assertEqual(
            (
                f'Body.\n\n{ControlTokens.SIGNATURE}\n\n'
                f'{ControlTokens.ATTACHMENT_FIRST_ITEM}附件：1. Implementation plan'
            ),
            result['content'],
        )

    def test_frontmatter_date_remains_ignored_metadata(self):
        result = self._preprocess_file('---\nDate: 2026-07-14\n---\n\nBody.')

        self.assertNotIn('signatory', result)
        self.assertNotIn('document_date', result)

    def test_requires_a_blank_line_before_the_signature_block(self):
        source = 'Body.\nExample Authority\n2026-07-14'

        result = self._preprocess_file(source)

        self.assertNotIn('signatory', result)
        self.assertIn('Example Authority', result['content'])
        self.assertIn('2026-07-14', result['content'])

    def test_rejects_sentence_text_and_invalid_dates(self):
        sources = (
            'Body.\n\nThis is a conclusion.\n2026-07-14',
            'Body.\n\nExample Authority\n2026-02-30',
        )

        for source in sources:
            with self.subTest(source=source):
                result = self._preprocess_file(source)
                self.assertNotIn('signatory', result)

    def test_does_not_scan_for_a_signature_before_arbitrary_trailing_content(self):
        source = 'Body.\n\nExample Authority\n2026-07-14\n\nAdditional body content.'

        result = self._preprocess_file(source)

        self.assertNotIn('signatory', result)
        self.assertIn('Example Authority\n2026-07-14', result['content'])


if __name__ == '__main__':
    unittest.main()
