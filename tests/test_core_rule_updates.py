import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Twips

from docx_rules import (
    BreaksAndPaginationRule,
    PageMarginsRule,
    Severity,
    TableOfContentsRule,
    TextGridRule,
    TitlePageRule,
)


def set_grid(section, pitch='360', kind='lines'):
    for old in section._sectPr.findall(qn('w:docGrid')):
        section._sectPr.remove(old)
    grid = OxmlElement('w:docGrid')
    grid.set(qn('w:type'), kind)
    if pitch is not None:
        grid.set(qn('w:linePitch'), pitch)
    section._sectPr.append(grid)


def add_toc_field(doc, instruction='TOC \\o "1-3"', cache='Введение\t3'):
    paragraph = doc.add_paragraph()
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), instruction)
    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.text = cache
    run.append(text)
    field.append(run)
    paragraph._p.append(field)
    return paragraph


def title_document(work_type='КУРСОВАЯ РАБОТА', ministry='Министерство просвещения Российской Федерации'):
    doc = Document()
    values = [
        ministry,
        'Государственное образовательное учреждение',
        'Отделение информационных технологий',
        work_type,
        'ТЕМА РАБОТЫ: РАЗРАБОТКА СИСТЕМЫ',
        'Москва 2026',
    ]
    for value in values:
        paragraph = doc.add_paragraph(value)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if value.startswith('ТЕМА'):
            paragraph.runs[0].bold = True
    return doc


class TextGridRuleUpdatesTests(unittest.TestCase):
    def build(self, usable_delta=0, pitch='360'):
        doc = Document()
        section = doc.sections[0]
        section.page_height = Twips(14400 + usable_delta)
        section.top_margin = Twips(1440)
        section.bottom_margin = Twips(1440)
        set_grid(section, pitch)
        return doc

    def test_integer_steps_and_two_twip_tolerance(self):
        self.assertEqual([], TextGridRule().check(self.build()))
        self.assertEqual([], TextGridRule().check(self.build(usable_delta=2)))

    def test_non_integer_step_is_error(self):
        found = TextGridRule().check(self.build(usable_delta=3))
        self.assertTrue(any(v.severity == Severity.ERROR and 'целого числа шагов' in v.message
                            for v in found))


class PageMarginsRuleUpdatesTests(unittest.TestCase):
    def test_tolerance_and_section_location(self):
        doc = Document()
        first = doc.sections[0]
        first.top_margin = Cm(2)
        first.bottom_margin = Cm(2)
        first.left_margin = Cm(3)
        first.right_margin = Cm(1.5)
        doc.add_section()
        second = doc.sections[1]
        second.top_margin = Twips(round(Cm(2).twips) + 2)
        second.bottom_margin = Cm(2)
        second.left_margin = Cm(3)
        second.right_margin = Cm(1.5)
        self.assertEqual([], PageMarginsRule().check(doc))
        second.right_margin = Twips(round(Cm(1.5).twips) + 7)
        found = PageMarginsRule().check(doc)
        self.assertEqual(['Секция 2'], [v.location for v in found])


class TableOfContentsRuleUpdatesTests(unittest.TestCase):
    def test_real_field_with_cache_and_heading_settings(self):
        doc = Document()
        heading = doc.add_paragraph('СОДЕРЖАНИЕ')
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        heading.paragraph_format.first_line_indent = Cm(1.25)
        heading.runs[0].bold = True
        add_toc_field(doc)
        self.assertEqual([], TableOfContentsRule().check(doc))

    def test_manual_toc_is_error(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('Введение\t3')
        found = TableOfContentsRule().check(doc)
        self.assertTrue(any(v.severity == Severity.ERROR and 'ручное оглавление' in v.message
                            for v in found))

    def test_empty_cache_and_missing_heading_settings_are_errors(self):
        doc = Document()
        add_toc_field(doc, instruction='TOC', cache='')
        messages = [v.message for v in TableOfContentsRule().check(doc)]
        self.assertTrue(any('источник заголовков' in message for message in messages))
        self.assertTrue(any('Кэш поля TOC пуст' in message for message in messages))


class BreaksAndPaginationRuleTests(unittest.TestCase):
    def test_reports_suspicious_breaks_and_pagination_as_warnings(self):
        doc = Document()
        doc.add_paragraph('Текст').add_run().add_break()
        doc.add_paragraph('')
        doc.add_paragraph('')
        heading = doc.add_paragraph('ВВЕДЕНИЕ')
        heading.paragraph_format.keep_with_next = False
        body = doc.add_paragraph('Основной текст')
        body.paragraph_format.widow_control = False
        page = doc.add_paragraph('Следующая страница')
        page.add_run().add_break(WD_BREAK.PAGE)
        found = BreaksAndPaginationRule().check(doc)
        self.assertTrue(found)
        self.assertTrue(all(v.severity == Severity.WARNING for v in found))
        joined = ' '.join(v.message for v in found)
        for token in ['soft break', 'пустых абзаца', 'keep_with_next', 'widow_control', 'разрыв страницы']:
            self.assertIn(token, joined)

    def test_single_empty_paragraph_and_defaults_are_not_reported(self):
        doc = Document()
        doc.add_paragraph('Первый абзац')
        doc.add_paragraph('')
        doc.add_paragraph('Второй абзац')
        self.assertEqual([], BreaksAndPaginationRule().check(doc))


class TitlePageRuleUpdatesTests(unittest.TestCase):
    def test_supported_work_types_and_neutral_topic_messages(self):
        for work_type in ['ДИПЛОМНЫЙ ПРОЕКТ', 'КУРСОВАЯ РАБОТА', 'ТЕХНИЧЕСКОЕ ЗАДАНИЕ']:
            with self.subTest(work_type=work_type):
                found = TitlePageRule().check(title_document(work_type))
                self.assertFalse(any('Не указан вид работы' in v.message for v in found))
                self.assertFalse(any('Министерство просвещения' in v.message for v in found))
                self.assertFalse(any('Тема дипломного проекта' in v.message for v in found))

    def test_obsolete_ministry_is_error_and_required_ministry_is_missing(self):
        doc = title_document(ministry='Министерство образования и науки Российской Федерации')
        found = TitlePageRule().check(doc)
        self.assertTrue(any(v.severity == Severity.ERROR and 'устаревшее наименование' in v.message
                            for v in found))
        self.assertTrue(any(v.severity == Severity.ERROR and 'обязательное наименование' in v.message
                            for v in found))


if __name__ == '__main__':
    unittest.main()
