"""Format Markdown unordered lists as deterministic native Word lists."""

from docx.document import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

from .base_formatter import BaseFormatter


class ListFormatter(BaseFormatter):
    """Own the marker, indentation, and text styling of unordered lists."""

    def format_lists(self, doc: Document) -> None:
        """Apply one controlled bullet definition to every unordered list level."""
        bullet_num_ids = self._format_numbering_definitions(doc)

        for paragraph in doc.paragraphs:
            style = paragraph.style
            if style is not None and style.name.startswith('Heading'):
                continue

            level = self._get_bullet_level(paragraph, bullet_num_ids)
            if level is not None:
                self._format_list_item(paragraph, level)

    def _get_bullet_level(self, paragraph, bullet_num_ids: set[str]) -> int | None:
        """Return a paragraph's zero-based bullet level, or ``None`` when it is not a bullet."""
        p_pr = paragraph._element.find(qn('w:pPr'))
        num_pr = p_pr.find(qn('w:numPr')) if p_pr is not None else None
        if num_pr is None:
            return None

        num_id = num_pr.find(qn('w:numId'))
        if num_id is None or num_id.get(qn('w:val')) not in bullet_num_ids:
            return None

        level = num_pr.find(qn('w:ilvl'))
        if level is None:
            return 0

        try:
            return max(0, int(level.get(qn('w:val'), '0')))
        except ValueError:
            return 0

    def _format_list_item(self, paragraph, level: int) -> None:
        """Format one list item while retaining its native numbering properties."""
        text_position_pt, hanging_pt = self._list_positions_pt(level)
        paragraph_format = paragraph.paragraph_format
        paragraph_format.left_indent = Pt(text_position_pt)
        paragraph_format.first_line_indent = Pt(-hanging_pt)
        paragraph_format.space_before = Pt(0)
        paragraph_format.space_after = Pt(0)
        paragraph.alignment = self.config.ALIGNMENTS['justify']

        p_pr = paragraph._element.get_or_add_pPr()
        self._set_numbering_tab(p_pr, self._points_to_twips(text_position_pt))
        self._enable_snap_to_grid(paragraph)

        for run in paragraph.runs:
            run.font.size = self.config.FONT_SIZES['body']
            run.bold = False
            run.italic = False
            run.font.color.rgb = RGBColor(0, 0, 0)
            self._set_run_fonts(run, self.config.FONTS['fangsong'])

    def _format_numbering_definitions(self, doc: Document) -> set[str]:
        """Format bullet levels and return the concrete numbering IDs that use them."""
        if not hasattr(doc, 'part') or not hasattr(doc.part, 'numbering_part'):
            return set()

        numbering_part = doc.part.numbering_part
        if numbering_part is None:
            return set()

        numbering = numbering_part._element
        bullet_abstract_ids = set()

        for abstract_num in numbering.findall(qn('w:abstractNum')):
            contains_bullet = False
            for level in abstract_num.findall(qn('w:lvl')):
                num_format = level.find(qn('w:numFmt'))
                if num_format is None or num_format.get(qn('w:val')) != 'bullet':
                    continue

                contains_bullet = True
                self._format_bullet_level(level)

            if contains_bullet:
                abstract_id = abstract_num.get(qn('w:abstractNumId'))
                if abstract_id is not None:
                    bullet_abstract_ids.add(abstract_id)

        bullet_num_ids = set()
        for num in numbering.findall(qn('w:num')):
            abstract_num_id = num.find(qn('w:abstractNumId'))
            if abstract_num_id is None:
                continue
            if abstract_num_id.get(qn('w:val')) in bullet_abstract_ids:
                num_id = num.get(qn('w:numId'))
                if num_id is not None:
                    bullet_num_ids.add(num_id)

        return bullet_num_ids

    def _format_bullet_level(self, level) -> None:
        """Make one numbering level use the configured marker and character-grid positions."""
        level_index = self._level_index(level)

        level_text = level.find(qn('w:lvlText'))
        if level_text is None:
            level_text = OxmlElement('w:lvlText')
            level.append(level_text)
        level_text.set(qn('w:val'), self.config.UNORDERED_LIST['marker'])

        suffix = level.find(qn('w:suff'))
        if suffix is None:
            suffix = OxmlElement('w:suff')
            level.insert(list(level).index(level_text), suffix)
        suffix.set(qn('w:val'), 'tab')

        justification = level.find(qn('w:lvlJc'))
        if justification is None:
            justification = OxmlElement('w:lvlJc')
            level.append(justification)
        justification.set(qn('w:val'), 'left')

        text_position_pt, hanging_pt = self._list_positions_pt(level_index)
        p_pr = self._get_or_add_level_p_pr(level)
        self._set_indent(
            p_pr,
            self._points_to_twips(text_position_pt),
            self._points_to_twips(hanging_pt),
        )
        self._set_numbering_tab(p_pr, self._points_to_twips(text_position_pt))
        self._set_numbering_fonts(level)

    @staticmethod
    def _level_index(level) -> int:
        try:
            return max(0, int(level.get(qn('w:ilvl'), '0')))
        except ValueError:
            return 0

    def _list_positions_pt(self, level: int) -> tuple[float, float]:
        """Return the item-text position and hanging width in points."""
        settings = self.config.UNORDERED_LIST
        marker_position_chars = settings['marker_position_chars'] + level * settings['nested_step_chars']
        hanging_chars = settings['text_gap_chars']
        character_width_pt = self.config.FONT_SIZES['body'].pt
        text_position_chars = marker_position_chars + hanging_chars
        return (
            text_position_chars * character_width_pt,
            hanging_chars * character_width_pt,
        )

    @staticmethod
    def _points_to_twips(points: float) -> int:
        return int(round(points * 20))

    @staticmethod
    def _get_or_add_level_p_pr(level):
        p_pr = level.find(qn('w:pPr'))
        if p_pr is not None:
            return p_pr

        p_pr = OxmlElement('w:pPr')
        r_pr = level.find(qn('w:rPr'))
        if r_pr is None:
            level.append(p_pr)
        else:
            level.insert(list(level).index(r_pr), p_pr)
        return p_pr

    @staticmethod
    def _set_indent(p_pr, text_position_twips: int, hanging_twips: int) -> None:
        ind = p_pr.find(qn('w:ind'))
        if ind is None:
            ind = OxmlElement('w:ind')
            p_pr.append(ind)

        for attribute in (
            'w:start',
            'w:startChars',
            'w:leftChars',
            'w:firstLine',
            'w:firstLineChars',
            'w:hangingChars',
        ):
            ind.attrib.pop(qn(attribute), None)

        ind.set(qn('w:left'), str(text_position_twips))
        ind.set(qn('w:hanging'), str(hanging_twips))

    @staticmethod
    def _set_numbering_tab(p_pr, text_position_twips: int) -> None:
        tabs = p_pr.find(qn('w:tabs'))
        ind = p_pr.find(qn('w:ind'))
        if tabs is None:
            tabs = OxmlElement('w:tabs')
            if ind is None:
                p_pr.append(tabs)
            else:
                p_pr.insert(list(p_pr).index(ind), tabs)
        else:
            for tab in list(tabs):
                tabs.remove(tab)

        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), 'num')
        tab.set(qn('w:pos'), str(text_position_twips))
        tabs.append(tab)

    def _set_numbering_fonts(self, level) -> None:
        """Set explicit marker fonts and size for Word and WPS compatibility."""
        r_pr = level.find(qn('w:rPr'))
        if r_pr is None:
            r_pr = OxmlElement('w:rPr')
            level.append(r_pr)

        r_fonts = r_pr.find(qn('w:rFonts'))
        if r_fonts is None:
            r_fonts = OxmlElement('w:rFonts')
            r_pr.append(r_fonts)

        r_fonts.set(qn('w:ascii'), self.config.FONTS['latin'])
        r_fonts.set(qn('w:hAnsi'), self.config.FONTS['latin'])
        r_fonts.set(qn('w:cs'), self.config.FONTS['latin'])
        r_fonts.set(qn('w:eastAsia'), self.config.FONTS['fangsong'])

        size_half_points = str(int(round(self.config.FONT_SIZES['body'].pt * 2)))
        for tag in ('w:sz', 'w:szCs'):
            size = r_pr.find(qn(tag))
            if size is None:
                size = OxmlElement(tag)
                r_pr.append(size)
            size.set(qn('w:val'), size_half_points)
