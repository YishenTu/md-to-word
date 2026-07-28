import re
from datetime import date

from ..utils.constants import ControlTokens, Patterns


class SignatureBlockParser:
    """Extract an unambiguous signatory and date block from the terminal document region."""

    ISO_DATE_PATTERN = re.compile(r'^(\d{4})-(\d{2})-(\d{2})$')
    CHINESE_DATE_PATTERN = re.compile(r'^(\d{4})\u5e74(\d{1,2})\u6708(\d{1,2})\u65e5$')
    MARKDOWN_PREFIX_PATTERN = re.compile(r'^(?:#{1,6}\s|[-+*]\s|\d+[.)]\s|>|\||!\[|!\[\[|`{3,}|~{3,})')
    SENTENCE_PUNCTUATION_PATTERN = re.compile(r'[.!?;:\u3002\uff01\uff1f\uff1b\uff1a]')
    MAX_SIGNATORY_LENGTH = 50

    def parse(self, content: str) -> tuple[str, dict[str, str]]:
        """Replace a strict signature block with its position anchor and return its values."""
        lines = content.split('\n')
        last_index = self._last_nonempty_index(lines)
        if last_index is None or last_index == 0:
            return content, {}

        date_indices = [last_index]
        attachment_start = self._terminal_attachment_start(lines, last_index)
        if attachment_start is not None:
            preceding_index = self._previous_nonempty_index(lines, attachment_start - 1)
            if preceding_index is not None:
                date_indices.append(preceding_index)

        for date_index in date_indices:
            parsed = self._parse_signature_at(lines, date_index)
            if parsed is not None:
                return parsed

        return content, {}

    def _parse_signature_at(self, lines: list[str], date_index: int) -> tuple[str, dict[str, str]] | None:
        normalized_date = self._normalize_document_date(lines[date_index].strip())
        if normalized_date is None or date_index == 0:
            return None

        signatory_index = self._previous_nonempty_index(lines, date_index - 1)
        if signatory_index is None:
            return None

        signatory = lines[signatory_index].strip()
        if not self._is_valid_signatory(signatory):
            return None

        if signatory_index > 0 and lines[signatory_index - 1].strip():
            return None

        body_lines = lines.copy()
        body_lines[signatory_index : date_index + 1] = [ControlTokens.SIGNATURE]
        body = '\n'.join(body_lines).rstrip()
        return body, {
            'signatory': signatory,
            'document_date': normalized_date,
        }

    def _terminal_attachment_start(self, lines: list[str], last_index: int) -> int | None:
        for header_index in range(last_index, -1, -1):
            if self._is_terminal_attachment(lines, header_index, last_index):
                return header_index
        return None

    @staticmethod
    def _is_terminal_attachment(lines: list[str], header_index: int, last_index: int) -> bool:
        header = lines[header_index]
        if Patterns.ATTACHMENT_HEADER_PATTERN.fullmatch(header):
            item_index = header_index + 1
            while item_index <= last_index and not lines[item_index].strip():
                item_index += 1
            if item_index > last_index:
                return False
        elif Patterns.ATTACHMENT_INLINE_PATTERN.fullmatch(header):
            item_index = header_index + 1
        else:
            return False

        return all(
            Patterns.ATTACHMENT_ITEM_PATTERN.fullmatch(lines[index]) is not None
            for index in range(item_index, last_index + 1)
        )

    @staticmethod
    def _last_nonempty_index(lines: list[str]) -> int | None:
        for index in range(len(lines) - 1, -1, -1):
            if lines[index].strip():
                return index
        return None

    @staticmethod
    def _previous_nonempty_index(lines: list[str], start_index: int) -> int | None:
        for index in range(start_index, -1, -1):
            if lines[index].strip():
                return index
        return None

    def _is_valid_signatory(self, value: str) -> bool:
        if not value or len(value) > self.MAX_SIGNATORY_LENGTH:
            return False
        if self.MARKDOWN_PREFIX_PATTERN.match(value):
            return False
        return self.SENTENCE_PUNCTUATION_PATTERN.search(value) is None

    def _normalize_document_date(self, value: str) -> str | None:
        match = self.ISO_DATE_PATTERN.fullmatch(value)
        if match is None:
            match = self.CHINESE_DATE_PATTERN.fullmatch(value)
        if match is None:
            return None

        year, month, day = (int(component) for component in match.groups())
        try:
            normalized = date(year, month, day)
        except ValueError:
            return None

        return f'{normalized.year}\u5e74{normalized.month}\u6708{normalized.day}\u65e5'
