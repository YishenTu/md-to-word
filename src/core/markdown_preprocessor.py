import html
import os
import re
from pathlib import Path
from typing import Any

from ..parsers import SignatureBlockParser
from ..utils.constants import ControlTokens, Patterns
from ..utils.exceptions import FileProcessingError, PathSecurityError


class MarkdownPreprocessor:
    """Markdown预处理器，用于清理和过滤Markdown内容后交给pandoc处理"""

    # Caption处理相关常量
    CAPTION_SEARCH_BEFORE = 10  # 向前查找行数
    CAPTION_SEARCH_AFTER = 20  # 向后查找行数
    CAPTION_MAX_EMPTY_LINES = 2  # 最大允许空行数
    FENCE_START_PATTERN = re.compile(r'^( {0,3})(`{3,}|~{3,})(.*)$')
    H1_PATTERN = re.compile(r'^ {0,3}#[ \t]+(.+?)[ \t]*$')
    FENCE_PLACEHOLDER_PREFIX = '\ue000MD_TO_WORD_FENCE_'
    FENCE_PLACEHOLDER_SUFFIX = '\ue001'
    ATTACHMENT_HEADER_PATTERN = Patterns.ATTACHMENT_HEADER_PATTERN
    ATTACHMENT_INLINE_PATTERN = Patterns.ATTACHMENT_INLINE_PATTERN
    ATTACHMENT_ITEM_PATTERN = Patterns.ATTACHMENT_ITEM_PATTERN

    def __init__(self):
        self.signature_block_parser = SignatureBlockParser()

    def preprocess_file(self, file_path: str) -> dict[str, Any]:
        """预处理Markdown文件，返回处理结果"""
        # 验证路径安全性
        try:
            safe_path = Path(file_path).expanduser().resolve()
        except Exception as error:
            raise PathSecurityError(f'无效的文件路径: {file_path}') from error

        try:
            with open(safe_path, encoding='utf-8') as f:
                content = f.read()
        except FileNotFoundError as error:
            raise FileProcessingError(f'文件不存在: {safe_path}') from error
        except PermissionError as error:
            raise FileProcessingError(f'没有权限读取文件: {safe_path}') from error
        except Exception as error:
            raise FileProcessingError(f'读取文件时出错: {error}') from error

        # 获取文件名作为标题（去掉扩展名）
        filename = os.path.basename(safe_path)
        title_from_filename = os.path.splitext(filename)[0]

        lines = self._filter_document_boundaries(content.split('\n'))
        markdown_title = self._single_h1_title(lines)
        body_content, document_fields = self.signature_block_parser.parse('\n'.join(lines))

        # 预处理内容
        processed_content = self._preprocess_lines(body_content.split('\n'))

        result = {
            'title': markdown_title or title_from_filename,
            'content': processed_content,
        }
        result.update(document_fields)
        return result

    def preprocess_content(self, content: str, file_path: str = '') -> str:
        """预处理Markdown内容"""
        lines = self._filter_document_boundaries(content.split('\n'))
        return self._preprocess_lines(lines)

    def _preprocess_lines(self, lines: list[str]) -> str:
        """Apply structural Markdown transforms after document boundaries are filtered."""
        lines, fenced_blocks = self._protect_fenced_code_blocks(lines)

        # Apply structural transforms only outside fenced code blocks. Bold is
        # normalized in the Word formatting phase so inline code remains intact.
        lines = self._normalize_attachment_sections(lines)
        lines = self._reposition_captions(lines)
        lines = self._fix_unordered_list_asterisks(lines)
        lines = self._skip_first_level_headers(lines)
        lines = self._convert_ordered_lists_to_text(lines)
        lines = self._convert_hr_to_pagebreak(lines)
        lines = self._escape_obsidian_embeds_for_pandoc(lines)
        lines = self._restore_fenced_code_blocks(lines, fenced_blocks)

        # 重新组合内容
        processed_content = '\n'.join(lines)

        return processed_content.strip()

    def _filter_document_boundaries(self, lines: list[str]) -> list[str]:
        """Remove leading frontmatter and trailing note metadata before document parsing."""
        return self._filter_ending_metadata(self._filter_yaml_frontmatter(lines))

    def _single_h1_title(self, lines: list[str]) -> str | None:
        """Return the visible text of the sole level-one heading outside fenced code."""
        protected_lines, _ = self._protect_fenced_code_blocks(lines)
        titles = []
        for line in protected_lines:
            match = self.H1_PATTERN.fullmatch(line)
            if match is None:
                continue
            title = re.sub(r'[ \t]+#+[ \t]*$', '', match.group(1)).strip()
            if title:
                titles.append(title)
        return titles[0] if len(titles) == 1 else None

    @staticmethod
    def _escape_obsidian_embeds_for_pandoc(lines: list[str]) -> list[str]:
        """Keep Obsidian image syntax as literal text across Pandoc versions."""

        def replace_embed(match: re.Match[str]) -> str:
            escaped_path = html.escape(match.group(1), quote=False)
            return f'!&#91;&#91;{escaped_path}&#93;&#93;'

        return [Patterns.OBSIDIAN_IMAGE_PATTERN.sub(replace_embed, line) for line in lines]

    def _protect_fenced_code_blocks(self, lines: list[str]) -> tuple[list[str], dict[str, list[str]]]:
        """Replace fenced code blocks with opaque placeholders during preprocessing."""
        protected_lines: list[str] = []
        blocks: dict[str, list[str]] = {}
        index = 0
        line_index = 0

        while line_index < len(lines):
            opening = self.FENCE_START_PATTERN.match(lines[line_index])
            if opening is None:
                protected_lines.append(lines[line_index])
                line_index += 1
                continue

            fence = opening.group(2)
            fence_character = fence[0]
            closing_pattern = re.compile(rf'^ {{0,3}}{re.escape(fence_character)}{{{len(fence)},}}[ \t]*$')
            block = [lines[line_index]]
            line_index += 1

            while line_index < len(lines):
                block.append(lines[line_index])
                is_closing_line = closing_pattern.match(lines[line_index]) is not None
                line_index += 1
                if is_closing_line:
                    break

            placeholder = f'{self.FENCE_PLACEHOLDER_PREFIX}{index}{self.FENCE_PLACEHOLDER_SUFFIX}'
            blocks[placeholder] = block
            protected_lines.append(placeholder)
            index += 1

        return protected_lines, blocks

    @staticmethod
    def _restore_fenced_code_blocks(lines: list[str], blocks: dict[str, list[str]]) -> list[str]:
        """Restore fenced code blocks previously replaced by opaque placeholders."""
        restored_lines: list[str] = []
        for line in lines:
            restored_lines.extend(blocks.get(line, [line]))
        return restored_lines

    def _normalize_attachment_sections(self, lines: list[str]) -> list[str]:
        """Mark attachment items as separate paragraphs for deterministic Word layout."""
        normalized_lines: list[str] = []
        line_index = 0

        while line_index < len(lines):
            header = lines[line_index]
            if self.ATTACHMENT_HEADER_PATTERN.match(header) is not None:
                item_index = line_index + 1
                while item_index < len(lines) and not lines[item_index].strip():
                    item_index += 1

                items, next_index = self._collect_attachment_items(lines, item_index)
                if not items:
                    normalized_lines.append(header)
                    line_index += 1
                    continue

                self._append_normalized_attachment_items(normalized_lines, items)
                line_index = next_index
                continue

            inline_match = self.ATTACHMENT_INLINE_PATTERN.match(header)
            if inline_match is None:
                normalized_lines.append(lines[line_index])
                line_index += 1
                continue

            first_item_match = re.match(r'^\s*附件[:：]\s*(\d+)\.\s+(.+?)\s*$', header)
            if first_item_match is None:
                normalized_lines.append(header)
                line_index += 1
                continue

            items = [(first_item_match.group(1), first_item_match.group(2).rstrip())]
            continuation_items, next_index = self._collect_attachment_items(lines, line_index + 1)
            items.extend(continuation_items)
            self._append_normalized_attachment_items(normalized_lines, items)
            line_index = next_index

        return normalized_lines

    def _collect_attachment_items(self, lines: list[str], start_index: int) -> tuple[list[tuple[str, str]], int]:
        """Collect one contiguous sequence of numbered attachment items."""
        items: list[tuple[str, str]] = []
        item_index = start_index
        while item_index < len(lines):
            item_match = self.ATTACHMENT_ITEM_PATTERN.match(lines[item_index])
            if item_match is None:
                break
            items.append((item_match.group(1), item_match.group(2).rstrip()))
            item_index += 1
        return items, item_index

    @staticmethod
    def _append_normalized_attachment_items(
        normalized_lines: list[str],
        items: list[tuple[str, str]],
    ) -> None:
        """Emit attachment items as marked paragraphs consumed by the Word formatter."""
        for item_position, (number, item_content) in enumerate(items):
            if item_position > 0:
                normalized_lines.append('')
            if item_position == 0:
                marker = ControlTokens.ATTACHMENT_FIRST_ITEM
                visible_prefix = f'附件：{number}.'
            else:
                marker = ControlTokens.ATTACHMENT_ITEM
                visible_prefix = f'{number}.'
            normalized_lines.append(f'{marker}{visible_prefix} {item_content}')

    def _filter_yaml_frontmatter(self, lines: list[str]) -> list[str]:
        """过滤YAML front matter"""
        end_index = self._frontmatter_end_index(lines)
        return lines[end_index + 1 :] if end_index is not None else lines

    @staticmethod
    def _frontmatter_end_index(lines: list[str]):
        """Return the closing delimiter index for leading YAML frontmatter."""
        if not lines or lines[0].strip() != '---':
            return None
        for index in range(1, len(lines)):
            if lines[index].strip() == '---':
                return index
        return None

    def _filter_ending_metadata(self, lines: list[str]) -> list[str]:
        """过滤结尾的Date和标签"""
        # 从后往前查找最后一个实质性内容的位置
        last_content_index = len(lines) - 1

        for i in range(len(lines) - 1, -1, -1):
            line_stripped = lines[i].strip()

            # 跳过空行、Date行、单词标签（如#work）、---分隔符
            if (
                line_stripped == ''
                or line_stripped.startswith('Date:')
                or line_stripped == '---'
                or (line_stripped.startswith('#') and ' ' not in line_stripped and not line_stripped.startswith('##'))
            ):
                continue
            else:
                last_content_index = i
                break

        # 返回到最后实质内容位置的所有行
        return lines[: last_content_index + 1]

    def _skip_first_level_headers(self, lines: list[str]) -> list[str]:
        """动态检测和调整标题层级
        - 如果检测到多个一级标题（#），将所有标题层级下移
        - 如果只有一个一级标题，则跳过该标题（其文本用作文档标题）
        """
        # 统计一级标题数量
        h1_count = sum(self.H1_PATTERN.fullmatch(line) is not None for line in lines)

        # 如果有多个一级标题，调整所有标题层级
        if h1_count > 1:
            return self._adjust_header_levels(lines)
        else:
            # 原有逻辑：跳过单个一级标题
            processed_lines = []
            for line in lines:
                if self.H1_PATTERN.fullmatch(line) is not None:
                    continue
                else:
                    processed_lines.append(line)
            return processed_lines

    def _adjust_header_levels(self, lines: list[str]) -> list[str]:
        """将标题层级下移一级，但只处理到三级标题
        # -> ##（Heading 2）
        ## -> ###（Heading 3）
        ### -> 作为正文处理（移除标题标记）
        #### 及更深 -> 作为正文处理（移除标题标记）
        """
        processed_lines = []

        for line in lines:
            stripped_line = line.strip()

            # 检查是否为标题行
            if stripped_line.startswith('#'):
                # 找到第一个空格的位置（标题级别和内容的分隔）
                space_index = stripped_line.find(' ')
                if space_index > 0:
                    # 获取标题级别（#的数量）
                    header_level = stripped_line[:space_index]
                    if header_level and all(c == '#' for c in header_level):
                        level_count = len(header_level)
                        # 获取原始行的缩进
                        indent = line[: len(line) - len(line.lstrip())]

                        if level_count <= 2:
                            # 一级和二级标题：下移一级
                            new_line = indent + '#' + stripped_line
                            processed_lines.append(new_line)
                        else:
                            # 三级及更深的标题：作为正文处理，移除标题标记
                            content = stripped_line[space_index + 1 :]  # 获取标题内容
                            # 检查是否是多级编号格式，如果是，需要保护
                            multi_match = Patterns.MULTI_LEVEL_NUMBER_PATTERN.match(content)
                            if multi_match:
                                # 是多级编号，使用反引号包裹
                                numbering = multi_match.group(1)
                                text = multi_match.group(2)
                                content = f'`{numbering}` {text}'
                            new_line = indent + content
                            processed_lines.append(new_line)
                        continue

            # 非标题行或无法识别的格式，保持原样
            processed_lines.append(line)

        return processed_lines

    def _fix_unordered_list_asterisks(self, lines: list[str]) -> list[str]:
        """修复无序列表的星号，避免被误识别为斜体"""
        processed_lines = []

        for line in lines:
            # 检测无序列表项 (例: "* 项目内容")
            if Patterns.UNORDERED_LIST_PATTERN.match(line):
                # 将星号替换为短横线
                processed_line = Patterns.UNORDERED_LIST_REPLACE_PATTERN.sub(r'\1- ', line)
                processed_lines.append(processed_line)
            else:
                processed_lines.append(line)

        return processed_lines

    def _reposition_captions(self, lines: list[str]) -> list[str]:
        """重新定位图表标题，确保标题始终在图表后面"""
        processed_lines = []
        i = 0

        while i < len(lines):
            line = lines[i].strip()

            # 检查是否为图表标题
            caption_match = Patterns.CAPTION_PATTERN.match(line)
            if not caption_match:
                # 不是标题，直接添加
                processed_lines.append(lines[i])
                i += 1
                continue

            caption_type = caption_match.group(1)  # 图/图片/表/表格/图表

            # 检查标题是否已经在正确位置
            if self._is_caption_after_element(lines, i, caption_type):
                self._append_standalone_caption(processed_lines, lines[i])
                i += 1
                continue

            # 标题不在正确位置，查找对应的图表元素
            element_info = self._find_element_for_caption(lines, i, caption_type)

            if element_info['found']:
                # 需要移动标题到元素后面
                caption_line = lines[i]
                i += 1

                # 添加从标题后到元素（含元素）的所有行
                while i <= element_info['index']:
                    processed_lines.append(lines[i])
                    i += 1

                # 在元素后添加标题
                self._append_standalone_caption(processed_lines, caption_line)
            else:
                # Keep unmatched captions in place, but as standalone paragraphs.
                self._append_standalone_caption(processed_lines, lines[i])
                i += 1

        return processed_lines

    @staticmethod
    def _append_standalone_caption(processed_lines: list[str], caption_line: str) -> None:
        """Keep a caption in its own Markdown paragraph without creating a Word blank line."""
        if processed_lines and processed_lines[-1].strip():
            processed_lines.append('')
        processed_lines.append(caption_line)
        processed_lines.append('')

    def _is_caption_after_element(self, lines: list[str], caption_index: int, caption_type: str) -> bool:
        """检查caption是否已在正确位置（紧跟在对应图表后面）"""
        empty_lines = 0

        # 向前查找，最多查找CAPTION_SEARCH_BEFORE行
        for j in range(caption_index - 1, max(caption_index - self.CAPTION_SEARCH_BEFORE, -1), -1):
            prev_line = lines[j].strip()

            if not prev_line:  # 空行
                empty_lines += 1
                continue

            # 检查是否为匹配的元素
            if self._is_matching_element(prev_line, caption_type):
                # 如果是表格，需要确认是表格的最后一行
                if caption_type in ['表', '表格'] and j + 1 < len(lines):
                    next_line = lines[j + 1].strip()
                    if next_line and Patterns.TABLE_ROW_PATTERN.match(next_line):
                        return False  # 不是表格最后一行

                # 检查空行数是否在允许范围内
                return empty_lines <= self.CAPTION_MAX_EMPTY_LINES
            else:
                # 遇到其他内容，停止查找
                break

        return False

    def _find_element_for_caption(self, lines: list[str], caption_index: int, caption_type: str) -> dict[str, Any]:
        """向后查找caption对应的图表元素"""
        # 从下一行开始，最多查找CAPTION_SEARCH_AFTER行
        for j in range(caption_index + 1, min(caption_index + self.CAPTION_SEARCH_AFTER + 1, len(lines))):
            check_line = lines[j].strip()

            # 如果遇到另一个标题，停止查找
            if Patterns.CAPTION_PATTERN.match(check_line):
                break

            # 检查是否为匹配的元素
            if self._is_matching_element(check_line, caption_type):
                element_index = j

                # 如果是表格，找到表格的结束位置
                if caption_type in ['表', '表格']:
                    element_index = self._find_table_end(lines, j)

                return {'found': True, 'index': element_index}

        return {'found': False, 'index': -1}

    def _is_matching_element(self, line: str, caption_type: str) -> bool:
        """判断是否为匹配的图表元素"""
        if caption_type in ['图', '图片', '图表']:
            return bool(Patterns.MARKDOWN_IMAGE_PATTERN.match(line) or Patterns.OBSIDIAN_IMAGE_PATTERN.match(line))
        if caption_type in ['表', '表格']:
            return Patterns.TABLE_ROW_PATTERN.match(line) is not None
        return False

    def _find_table_end(self, lines: list[str], table_start: int) -> int:
        """找到表格的结束位置"""
        end_index = table_start

        # 确认是表格：检查是否有连续的表格行
        if table_start + 1 < len(lines) and Patterns.TABLE_ROW_PATTERN.match(lines[table_start + 1].strip()):
            # 继续向后查找直到表格结束
            while end_index + 1 < len(lines) and Patterns.TABLE_ROW_PATTERN.match(lines[end_index + 1].strip()):
                end_index += 1

        return end_index

    def _convert_ordered_lists_to_text(self, lines: list[str]) -> list[str]:
        """Convert ordered-list markers into visible body-paragraph text.

        Rules:
        1. Convert every ``1. Content`` item into a body paragraph.
        2. Convert dotted markers such as ``2.1.1 Content`` the same way.
        3. Escape the first period so Pandoc does not create native numbering.
        4. Leave unordered-list indentation intact for Pandoc's native levels.
        """
        processed_lines: list[str] = []
        in_attachment_block = False

        for line in lines:
            stripped_line = line.lstrip()
            if stripped_line.startswith(('附件：', '附件:')):
                in_attachment_block = True
                processed_lines.append(line)
                continue

            if in_attachment_block and re.match(r'^\s+\d+\.\s+', line):
                # Preserve hard breaks and authored attachment alignment.
                processed_lines.append(line)
                continue

            in_attachment_block = False

            # Match dotted multi-level markers before simple ordered markers.
            multi_match = Patterns.MULTI_LEVEL_NUMBER_PATTERN.match(line)
            if multi_match:
                numbering = multi_match.group(1).replace('.', r'\.', 1)
                content = multi_match.group(2)
                new_line = f'{numbering} {content}'
                if processed_lines and processed_lines[-1].strip() != '':
                    processed_lines.append('')
                processed_lines.append(new_line)
                continue

            # Match a number followed by a period and a space.
            simple_match = Patterns.SIMPLE_ORDERED_LIST_WITH_CONTENT.match(line)
            if simple_match:
                indent = simple_match.group(1)
                number = simple_match.group(2)
                content = simple_match.group(3)
                new_line = f'{indent}{number}\\. {content}'
                if processed_lines and processed_lines[-1].strip() != '':
                    processed_lines.append('')
                processed_lines.append(new_line)
                continue

            processed_lines.append(line)

        return processed_lines

    def _convert_hr_to_pagebreak(self, lines: list[str]) -> list[str]:
        """将水平线 --- 转换为分页符标记（排除front/end matter）

        注意：YAML front matter已在之前被过滤，这里只处理内容中的---
        """
        processed_lines: list[str] = []

        for line in lines:
            stripped = line.strip()
            # 检测独立的 --- 行（水平线）
            if stripped == '---' or stripped == '***' or stripped == '___':
                # 确保前后有空行，使其成为独立段落
                if processed_lines and processed_lines[-1].strip() != '':
                    processed_lines.append('')
                processed_lines.append(ControlTokens.PAGE_BREAK)
                processed_lines.append('')  # 后面也加空行
            else:
                processed_lines.append(line)

        return processed_lines
