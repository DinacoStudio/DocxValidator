"""Регрессии проверки эффективного форматирования ТЗ по ГОСТ 34."""
import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from docx_rules.tz34_text_format import Tz34TextFormatRule


class Tz34FormatRegressions(unittest.TestCase):
    def setUp(self):
        self.doc = Document()
        normal = self.doc.styles['Normal']
        normal.font.name = 'Times New Roman'
        normal.font.size = Pt(14)
        normal.paragraph_format.line_spacing = 1.5
        normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal.paragraph_format.first_line_indent = Cm(1.25)
        self.body = self.doc.add_paragraph('Обычный текст раздела.')

    def messages(self):
        return [v.message for v in Tz34TextFormatRule().check(self.doc)]

    def test_inherited_paragraph_style_and_direct_override(self):
        base = self.doc.styles.add_style('Test base', WD_STYLE_TYPE.PARAGRAPH)
        base.base_style = self.doc.styles['Normal']
        base.font.name = 'Arial'
        base.font.size = Pt(12)
        base.paragraph_format.line_spacing = 2.0
        base.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        base.paragraph_format.first_line_indent = Cm(0.5)
        child = self.doc.styles.add_style('Test child', WD_STYLE_TYPE.PARAGRAPH)
        child.base_style = base
        self.body.style = child
        messages = ' '.join(self.messages())
        for text in ('Arial', '12.0 pt', 'интервал 2.0', 'выравнивание', 'Красная строка'):
            self.assertIn(text.lower(), messages.lower())
        run = self.body.runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        self.body.paragraph_format.line_spacing = 1.5
        self.body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.body.paragraph_format.first_line_indent = Cm(1.25)
        self.assertEqual([], self.messages())

    def test_character_style_overrides_paragraph_font(self):
        style = self.doc.styles.add_style('Test character', WD_STYLE_TYPE.CHARACTER)
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        self.body.runs[0].style = style
        messages = ' '.join(self.messages())
        self.assertIn('Arial', messages)
        self.assertIn('10.0 pt', messages)

    def test_doc_defaults_are_used(self):
        normal = self.doc.styles['Normal']
        normal.font.name = None
        normal.font.size = None
        normal.paragraph_format.line_spacing = None
        defaults = self.doc.styles.element.find(qn('w:docDefaults'))
        rpr = defaults.find(qn('w:rPrDefault')).find(qn('w:rPr'))
        # Значения задаются через стандартные объекты python-docx.
        from docx.text.font import Font
        from docx.oxml import OxmlElement
        from copy import deepcopy
        wrapper = OxmlElement('w:r')
        wrapper.append(deepcopy(rpr))
        font = Font(wrapper)
        font.name = 'Arial'
        font.size = Pt(12)
        defaults.find(qn('w:rPrDefault')).replace(rpr, wrapper.rPr)
        # В стандартном шаблоне python-docx интервал docDefaults равен 1,15.
        messages = ' '.join(self.messages())
        self.assertIn('Arial', messages)
        self.assertIn('12.0 pt', messages)
        self.assertIn('Междустрочный интервал', messages)

    def test_multiple_and_fixed_spacing(self):
        fmt = self.body.paragraph_format
        for value in (1.0, 2.0):
            with self.subTest(value=value):
                fmt.line_spacing = value
                self.assertTrue(any('Междустрочный интервал' in m for m in self.messages()))
        for value in (1.5, 1.55):
            with self.subTest(value=value):
                fmt.line_spacing = value
                self.assertEqual([], self.messages())
        for rule, word in ((WD_LINE_SPACING.EXACTLY, 'точно'),
                           (WD_LINE_SPACING.AT_LEAST, 'минимум')):
            with self.subTest(rule=rule):
                fmt.line_spacing = Pt(18)
                fmt.line_spacing_rule = rule
                messages = ' '.join(self.messages())
                self.assertIn(f'{word} 18 пт', messages)
                self.assertIn('полуторный', messages)
                self.assertNotIn('228600', messages)  # внутренние единицы EMU

    def test_fixed_spacing_inherited_from_style(self):
        fmt = self.doc.styles['Normal'].paragraph_format
        fmt.line_spacing = Pt(20)
        fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        self.assertIn('минимум 20 пт', ' '.join(self.messages()))

    def test_unspecified_properties_are_not_reported(self):
        normal = self.doc.styles['Normal']
        normal.font.name = normal.font.size = None
        normal.paragraph_format.line_spacing = None
        normal.paragraph_format.alignment = None
        normal.paragraph_format.first_line_indent = None
        defaults = self.doc.styles.element.find(qn('w:docDefaults'))
        defaults.getparent().remove(defaults)
        self.assertEqual([], self.messages())

    def test_bold_and_excluded_paragraphs(self):
        self.body.runs[0].bold = True
        self.body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self.body.paragraph_format.first_line_indent = Cm(0)
        self.assertEqual([], self.messages())
        self.body.paragraph_format.line_spacing = 1.0
        self.assertIn('Междустрочный интервал', ' '.join(self.messages()))
        doc = Document()
        title = doc.add_paragraph('Титульный лист')
        title.runs[0].font.name = 'Arial'
        doc.add_paragraph('СОДЕРЖАНИЕ')
        toc = doc.add_paragraph('Раздел .... 1')
        toc.runs[0].font.name = 'Arial'
        intro = doc.add_paragraph('ВВЕДЕНИЕ')
        intro.paragraph_format.line_spacing = 1.5
        for text in ('Таблица 1', 'Продолжение таблицы 1'):
            p = doc.add_paragraph(text)
            p.runs[0].font.name = 'Arial'
        self.assertFalse(any('Arial' in v.message for v in Tz34TextFormatRule().check(doc)))


if __name__ == '__main__':
    unittest.main()
