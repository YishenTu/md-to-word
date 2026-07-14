import re
from datetime import date


class SignatureBlockParser:
    """Extract an unambiguous signatory and date block from the document tail."""

    ISO_DATE_PATTERN = re.compile(r'^(\d{4})-(\d{2})-(\d{2})$')
    CHINESE_DATE_PATTERN = re.compile(r'^(\d{4})\u5e74(\d{1,2})\u6708(\d{1,2})\u65e5$')
    MARKDOWN_PREFIX_PATTERN = re.compile(r'^(?:#{1,6}\s|[-+*]\s|\d+[.)]\s|>|\||!\[|!\[\[|`{3,}|~{3,})')
    SENTENCE_PUNCTUATION_PATTERN = re.compile(r'[.!?;:\u3002\uff01\uff1f\uff1b\uff1a]')
    MAX_SIGNATORY_LENGTH = 50

    def parse(self, content: str) -> tuple[str, dict[str, str]]:
        """Remove and return a strict final signatory/date block when present."""
        lines = content.split('\n')
        last_index = self._last_nonempty_index(lines)
        if last_index is None or last_index == 0:
            return content, {}

        normalized_date = self._normalize_document_date(lines[last_index].strip())
        if normalized_date is None:
            return content, {}

        signatory_index = last_index - 1
        signatory = lines[signatory_index].strip()
        if not self._is_valid_signatory(signatory):
            return content, {}

        if signatory_index > 0 and lines[signatory_index - 1].strip():
            return content, {}

        body = '\n'.join(lines[:signatory_index]).rstrip()
        return body, {
            'signatory': signatory,
            'document_date': normalized_date,
        }

    @staticmethod
    def _last_nonempty_index(lines: list[str]):
        for index in range(len(lines) - 1, -1, -1):
            if lines[index].strip():
                return index
        return None

    def _is_valid_signatory(self, value: str) -> bool:
        if not value or len(value) > self.MAX_SIGNATORY_LENGTH:
            return False
        if self.MARKDOWN_PREFIX_PATTERN.match(value):
            return False
        return self.SENTENCE_PUNCTUATION_PATTERN.search(value) is None

    def _normalize_document_date(self, value: str):
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
