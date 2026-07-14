import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

from md_to_word import convert_document, resolve_output_path
from src.utils.exceptions import ImageProcessingError


class TestCliConversion(unittest.TestCase):
    def test_legacy_doc_extension_is_normalized_to_docx(self):
        input_path = Path('/tmp/report.md')

        output_path = resolve_output_path(input_path, '/tmp/report.doc')

        self.assertEqual(Path('/tmp').resolve() / 'report.docx', output_path)

    def test_failed_postprocessing_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / 'input.md'
            output_path = root / 'output.docx'
            input_path.write_text('正文', encoding='utf-8')
            original_output = b'previous valid document'
            output_path.write_bytes(original_output)

            def create_raw_docx(_content, staging_path, **_kwargs):
                Document().save(staging_path)
                return staging_path

            with patch(
                'md_to_word.PandocProcessor.check_pandoc_available',
                return_value=True,
            ):
                with patch(
                    'md_to_word.PandocProcessor.convert_markdown_to_docx',
                    side_effect=create_raw_docx,
                ):
                    with patch(
                        'md_to_word.WordPostprocessor.apply_formatting',
                        side_effect=ImageProcessingError('missing image'),
                    ):
                        with self.assertRaises(ImageProcessingError):
                            convert_document(input_path, output_path)

            self.assertEqual(original_output, output_path.read_bytes())
            self.assertEqual(
                {'input.md', 'output.docx'},
                {path.name for path in root.iterdir()},
            )

    @unittest.skipUnless(shutil.which('pandoc'), 'Pandoc is required for CLI integration tests')
    def test_successful_conversion_atomically_publishes_a_valid_docx(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / 'input.md'
            output_path = root / 'output.docx'
            input_path.write_text('正文。', encoding='utf-8')

            result = convert_document(input_path, output_path)

            document = Document(result)
            self.assertEqual('input', document.paragraphs[0].text)
            self.assertEqual(
                {'input.md', 'output.docx'},
                {path.name for path in root.iterdir()},
            )


if __name__ == '__main__':
    unittest.main()
