import tempfile
import unittest
from pathlib import Path

from src.core.markdown_preprocessor import MarkdownPreprocessor
from src.utils.constants import ControlTokens


class TestMarkdownPreprocessor(unittest.TestCase):
    def setUp(self):
        self.preprocessor = MarkdownPreprocessor()

    def test_preserves_markdown_soft_breaks_and_blockquotes(self):
        source = 'The first line\nends here\n\n> 第一行\n> 第二行'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual(source, result)

    def test_preserves_fenced_code_blocks_verbatim(self):
        source = '```markdown\n**bold marker**\n* list-like code\n---\n```'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual(source, result)

    def test_repositioned_caption_remains_separate_from_its_image_paragraph(self):
        source = '图1: Workflow\n\n![Workflow](assets/workflow.png)'

        result = self.preprocessor.preprocess_content(source)

        self.assertIn(
            '![Workflow](assets/workflow.png)\n\n图1: Workflow',
            result,
        )

    def test_obsidian_embeds_are_entity_escaped_for_pandoc(self):
        source = '![[assets/review&approval.png]]'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual('!&#91;&#91;assets/review&amp;approval.png&#93;&#93;', result)

    def test_attachment_mentions_and_declarations_remain_body_content(self):
        source = '正文提到附件另发。\n\n附件：1. 实施方案\n附件: 2. 数据表'

        result = self.preprocessor.preprocess_content(source)

        self.assertIn('正文提到附件另发。', result)
        self.assertIn(f'{ControlTokens.ATTACHMENT_FIRST_ITEM}附件：1. 实施方案', result)
        self.assertIn(f'{ControlTokens.ATTACHMENT_FIRST_ITEM}附件：2. 数据表', result)

    def test_authored_attachment_hard_breaks_are_normalized_to_item_paragraphs(self):
        source = '附件：1. 实施方案  \n　　　2. 数据明细表  \n　　　3. 验收清单'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual(
            (
                f'{ControlTokens.ATTACHMENT_FIRST_ITEM}附件：1. 实施方案\n\n'
                f'{ControlTokens.ATTACHMENT_ITEM}2. 数据明细表\n\n'
                f'{ControlTokens.ATTACHMENT_ITEM}3. 验收清单'
            ),
            result,
        )

    def test_natural_attachment_list_is_normalized_for_word_alignment(self):
        source = '附件：\n\n1. 实施方案\n2. 数据明细表\n3. 验收清单'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual(
            (
                f'{ControlTokens.ATTACHMENT_FIRST_ITEM}附件：1. 实施方案\n\n'
                f'{ControlTokens.ATTACHMENT_ITEM}2. 数据明细表\n\n'
                f'{ControlTokens.ATTACHMENT_ITEM}3. 验收清单'
            ),
            result,
        )

    def test_single_h1_becomes_file_metadata_title(self):
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory) / 'filename-title.md'
            input_path.write_text('# Markdown Title #\n\nBody.', encoding='utf-8')

            result = self.preprocessor.preprocess_file(str(input_path))

        self.assertEqual('Markdown Title', result['title'])
        self.assertEqual('Body.', result['content'])

    def test_filename_remains_title_when_no_single_h1_exists(self):
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory) / 'filename-title.md'
            input_path.write_text('Body.', encoding='utf-8')

            result = self.preprocessor.preprocess_file(str(input_path))

        self.assertEqual('filename-title', result['title'])

    def test_h1_inside_fenced_code_does_not_override_filename_title(self):
        with tempfile.TemporaryDirectory() as directory:
            input_path = Path(directory) / 'filename-title.md'
            input_path.write_text('```markdown\n# Not a document title\n```', encoding='utf-8')

            result = self.preprocessor.preprocess_file(str(input_path))

        self.assertEqual('filename-title', result['title'])

    def test_ordered_list_markers_are_escaped_without_inline_code(self):
        source = '1. 第一项\n2. 第二项'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual('1\\. 第一项\n\n2\\. 第二项', result)
        self.assertNotIn('`', result)

    def test_unordered_list_nesting_is_preserved_for_pandoc(self):
        source = '- Parent item\n  - Nested item\n- Final item'

        result = self.preprocessor.preprocess_content(source)

        self.assertEqual(source, result)
        self.assertNotIn('[NESTED]', result)


if __name__ == '__main__':
    unittest.main()
