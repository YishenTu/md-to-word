import re
from collections.abc import Callable
from decimal import Decimal

from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, RGBColor
from docx.text.paragraph import Paragraph

from ..utils.exceptions import FileProcessingError
from .base_formatter import BaseFormatter


class SignatureFormatter(BaseFormatter):
    """Render a single-authority signatory and date for stamped documents."""

    DATE_PATTERN = re.compile(r'^\d{4}\u5e74\d{1,2}\u6708\d{1,2}\u65e5$')
    DATE_SIDE_BLANK_CHARS = Decimal('4')
    BASE_DATE_LENGTH = 9
    BASE_DATE_WIDTH_CHARS = Decimal('7.25')
    EXTRA_DATE_CHAR_WIDTH = Decimal('0.5')
    SPACING_LINES = 2

    def add_signature(self, doc: Document, signatory: str, document_date: str) -> None:
        """Append adjacent signatory and date paragraphs."""
        self._render_signature(doc.add_paragraph, signatory, document_date)

    def replace_signature_anchor(
        self,
        anchor: Paragraph,
        signatory: str,
        document_date: str,
    ) -> None:
        """Replace one position-only anchor with adjacent signatory and date paragraphs."""
        self._render_signature(anchor.insert_paragraph_before, signatory, document_date)
        anchor._element.getparent().remove(anchor._element)

    def _render_signature(
        self,
        paragraph_factory: Callable[[], Paragraph],
        signatory: str,
        document_date: str,
    ) -> None:
        left_indent_chars = self._calculate_left_indent_chars(document_date)

        for _ in range(self.SPACING_LINES):
            self._format_blank_grid_line(paragraph_factory())

        signatory_paragraph = paragraph_factory()
        self._add_formatted_run(signatory_paragraph, signatory)
        self._format_signature_paragraph(
            signatory_paragraph,
            left_indent_chars,
            keep_with_next=True,
        )

        date_paragraph = paragraph_factory()
        self._add_formatted_run(date_paragraph, document_date)
        self._format_signature_paragraph(
            date_paragraph,
            left_indent_chars,
            keep_with_next=False,
        )
        self._enable_number_spacing(date_paragraph)

    def _format_blank_grid_line(self, paragraph: Paragraph) -> None:
        paragraph.alignment = self.config.ALIGNMENTS['left']
        paragraph_format = paragraph.paragraph_format
        paragraph_format.left_indent = Pt(0)
        paragraph_format.right_indent = Pt(0)
        paragraph_format.first_line_indent = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph_format.keep_together = True
        paragraph_format.keep_with_next = True
        self._enable_snap_to_grid(paragraph)

    def _calculate_left_indent_chars(self, document_date: str) -> Decimal:
        """Return the shared left indent that leaves four date characters per side."""
        if self.DATE_PATTERN.fullmatch(document_date) is None:
            raise FileProcessingError('Document date must use normalized Chinese date syntax')

        date_length = len(document_date)
        if date_length < 9 or date_length > 11:
            raise FileProcessingError('Normalized document date must contain 9 to 11 characters')

        date_width_chars = (
            self.BASE_DATE_WIDTH_CHARS + Decimal(date_length - self.BASE_DATE_LENGTH) * self.EXTRA_DATE_CHAR_WIDTH
        )
        centered_area_padding = self.DATE_SIDE_BLANK_CHARS * 2
        left_indent_chars = Decimal(self.config.CHARS_PER_LINE) - centered_area_padding - date_width_chars
        if left_indent_chars < 0:
            raise FileProcessingError('Signature layout exceeds the document grid')
        return left_indent_chars

    def _add_formatted_run(self, paragraph, text: str) -> None:
        run = paragraph.add_run(text)
        run.font.size = self.config.FONT_SIZES['body']
        run.font.bold = False
        run.font.italic = False
        run.font.color.rgb = RGBColor(0, 0, 0)
        self._set_run_fonts(run, self.config.FONTS['fangsong'])

    def _format_signature_paragraph(
        self,
        paragraph,
        left_indent_chars: Decimal,
        keep_with_next: bool,
    ) -> None:
        paragraph.alignment = self.config.ALIGNMENTS['center']
        paragraph_format = paragraph.paragraph_format
        paragraph_format.left_indent = self._physical_indent(left_indent_chars)
        paragraph_format.right_indent = Pt(0)
        paragraph_format.first_line_indent = Pt(0)
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph_format.keep_together = True
        paragraph_format.keep_with_next = keep_with_next

        indent = paragraph._element.get_or_add_pPr().get_or_add_ind()
        indent.set(qn('w:leftChars'), str(int(left_indent_chars * 100)))
        self._enable_snap_to_grid(paragraph)

    def _physical_indent(self, left_indent_chars: Decimal) -> Emu:
        char_pitch_emu = Decimal(self.config.get_content_width_emu()) / Decimal(self.config.CHARS_PER_LINE)
        return Emu(int((char_pitch_emu * left_indent_chars).to_integral_value()))

    @staticmethod
    def _enable_number_spacing(paragraph) -> None:
        p_pr = paragraph._element.get_or_add_pPr()
        auto_space = p_pr.find(qn('w:autoSpaceDN'))
        if auto_space is None:
            auto_space = OxmlElement('w:autoSpaceDN')
            p_pr.append(auto_space)
        auto_space.set(qn('w:val'), 'true')
