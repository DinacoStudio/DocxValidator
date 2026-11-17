import base64
import io
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from docx_rules import FormulasRule, ListsFormatRule, Severity, TablesFormatRule


W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
PNG_1PX = base64.b64decode(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==')


def messages(found):
    return [violation.message for violation in found]


def add_table_borders(table):
    borders = OxmlElement('w:tblBorders')
    for name in TablesFormatRule.REQUIRED_BORDERS:
        edge = OxmlElement(f'w:{name}')
        edge.set(qn('w:val'), 'single')
        edge.set(qn('w:sz'), '4')
        borders.append(edge)
    table._tbl.tblPr.append(borders)


def repeat_first_row(table):
    tr_pr = table.rows[0]._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement('w:tblHeader'))


def formatted_table(doc, label='Таблица 1', title='Название таблицы'):
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(14)
    normal.paragraph_format.line_spacing = 1.0
    if label:
        paragraph = doc.add_paragraph(label)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if title:
        paragraph = doc.add_paragraph(title)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.first_line_indent = Cm(1.25)
    table = doc.add_table(rows=2, cols=2)
    add_table_borders(table)
    repeat_first_row(table)
    for cell, text in zip(table.rows[0].cells, ('Показатель', 'Значение')):
        cell.text = text
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    table.rows[1].cells[0].text = '1'
    table.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table.rows[1].cells[1].text = 'Данные'
    table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    return table


class TablesDocumentRulesTests(unittest.TestCase):
    def test_natural_reference_counts(self):
        doc = Document()
        formatted_table(doc)
        doc.add_paragraph('Результаты приведены в таблице 1 и подтверждают вывод.')
        self.assertEqual([], TablesFormatRule().check(doc))

    def test_missing_borders_and_repeating_header_are_reported(self):
        doc = Document()
        table = formatted_table(doc)
        table.add_row()
        for edge in table._tbl.tblPr.find(qn('w:tblBorders')):
            edge.set(qn('w:val'), 'none')
        table.rows[0]._tr.get_or_add_trPr().remove(
            table.rows[0]._tr.get_or_add_trPr().find(qn('w:tblHeader')))
        doc.add_paragraph('Данные приведены в таблице 1.')
        found = TablesFormatRule().check(doc)
        self.assertTrue(any('границ' in message for message in messages(found)))
        self.assertTrue(any('w:tblHeader' in message for message in messages(found)))

    def test_continuation_must_refer_to_previously_recognized_number(self):
        doc = Document()
        formatted_table(doc)
        doc.add_paragraph('Продолжение таблицы 1').alignment = WD_ALIGN_PARAGRAPH.RIGHT
        formatted_table(doc, label=None, title=None)
        doc.add_paragraph('См. данные в таблице 1.')
        self.assertFalse(any('Неоднозначная надпись' in message
                             for message in messages(TablesFormatRule().check(doc))))

        ambiguous = Document()
        ambiguous.add_paragraph('Продолжение таблицы 7').alignment = WD_ALIGN_PARAGRAPH.RIGHT
        formatted_table(ambiguous, label=None, title=None)
        found = TablesFormatRule().check(ambiguous)
        self.assertTrue(any(v.severity == Severity.WARNING and 'Неоднозначная надпись' in v.message
                            for v in found))
        self.assertFalse(any("Перед таблицей не найдена надпись" in v.message for v in found))


class ListsDocumentRulesTests(unittest.TestCase):
    @staticmethod
    def add_numbering(doc, abstract_id, num_id, marker='–'):
        numbering = doc.part.numbering_part.element
        numbering.insert(0, parse_xml(
            f'<w:abstractNum {W_NS} w:abstractNumId="{abstract_id}">'
            f'<w:multiLevelType w:val="singleLevel"/>'
            f'<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/>'
            f'<w:lvlText w:val="{marker}"/></w:lvl></w:abstractNum>'))
        numbering.append(parse_xml(
            f'<w:num {W_NS} w:numId="{num_id}"><w:abstractNumId w:val="{abstract_id}"/></w:num>'))

    @staticmethod
    def style_numbering(style, num_id):
        num_pr = parse_xml(
            f'<w:numPr {W_NS}><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>')
        style.element.get_or_add_pPr().append(num_pr)

    def add_styled_items(self, doc, marker='–', intro='Перечислим элементы:'):
        self.add_numbering(doc, 401, 501, marker)
        parent = doc.styles.add_style('AutoListParent', WD_STYLE_TYPE.PARAGRAPH)
        self.style_numbering(parent, 501)
        child = doc.styles.add_style('AutoListChild', WD_STYLE_TYPE.PARAGRAPH)
        child.base_style = parent
        doc.add_paragraph(intro)
        for text in ('первый пункт;', 'последний пункт.'):
            paragraph = doc.add_paragraph(text, child)
            paragraph.paragraph_format.first_line_indent = Cm(1.75)
        return doc

    def test_numpr_is_inherited_through_style_chain(self):
        doc = self.add_styled_items(Document())
        self.assertEqual([], ListsFormatRule().check(doc))

    def test_colon_before_automatic_list(self):
        doc = self.add_styled_items(Document(), intro='Перечислим элементы.')
        self.assertTrue(any('перед автоматическим списком' in message
                            for message in messages(ListsFormatRule().check(doc))))

    def test_only_en_dash_is_valid_automatic_marker(self):
        for marker in ('-', '—', '•'):
            with self.subTest(marker=marker):
                found = ListsFormatRule().check(self.add_styled_items(Document(), marker=marker))
                self.assertTrue(any(v.severity == Severity.ERROR and 'Маркер списка' in v.message
                                    for v in found))


class FormulaImageContextTests(unittest.TestCase):
    def test_image_between_formula_intro_and_where_is_cautious_warning(self):
        doc = Document()
        doc.add_paragraph('Расчет выполняется по формуле:')
        doc.add_paragraph().add_run().add_picture(io.BytesIO(PNG_1PX))
        doc.add_paragraph('где x — исходное значение.')
        found = FormulasRule().check(doc)
        self.assertTrue(any(v.severity == Severity.WARNING and 'как изображение' in v.message
                            for v in found))

    def test_ordinary_figure_does_not_trigger_formula_warning(self):
        doc = Document()
        doc.add_paragraph('Архитектура системы показана на рисунке:')
        doc.add_paragraph().add_run().add_picture(io.BytesIO(PNG_1PX))
        doc.add_paragraph('Рисунок 1 – Архитектура системы')
        self.assertFalse(any('как изображение' in message
                             for message in messages(FormulasRule().check(doc))))


if __name__ == '__main__':
    unittest.main()
