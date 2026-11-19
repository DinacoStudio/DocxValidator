# -*- coding: utf-8 -*-
"""Профиль tz-34-602: оформление ТЗ на автоматизированную систему по образцу ПР2."""
import unittest

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from docx_rules import (BreaksAndPaginationRule, PageMarginsRule, TablesFormatRule,
                        Tz34StructureRule, Tz34TextFormatRule, Tz34TitlePageRule,
                        WhiteFillRule)
from docx_rules.classification import classify_paragraphs
from profiles import get_rules


class Tz34ProfileTests(unittest.TestCase):
    def test_profile_composition(self):
        rules = get_rules('tz-34-602')
        names = {type(r).__name__ for r in rules}
        self.assertIn('Tz34TitlePageRule', names)
        self.assertIn('Tz34StructureRule', names)
        self.assertIn('Tz34TextFormatRule', names)
        # Оформление образца не содержит сетки, нумерации страниц и МДК-титула
        self.assertNotIn('TextGridRule', names)
        self.assertNotIn('PageNumberingRule', names)
        self.assertNotIn('TitlePageRule', names)
        self.assertNotIn('SectionFormatRule', names)
        self.assertNotIn('MainTextFormatRule', names)
        self.assertNotIn('TableOfContentsRule', names)
        # Поля: правое 1 см (в образце — 1.0)
        margins = [r for r in rules if isinstance(r, PageMarginsRule)]
        self.assertEqual([margins[0].left, margins[0].right], [3.0, 1.0])
        # Повторяемая шапка таблиц не требуется (в образце таблицы короткие)
        tables_rule = [r for r in rules if isinstance(r, TablesFormatRule)][0]
        self.assertFalse(tables_rule.require_header_repeat)

    def test_unknown_profile_still_rejected(self):
        with self.assertRaises(ValueError):
            get_rules('tz-999')


class Tz34TitlePageRuleTests(unittest.TestCase):
    def build_title(self):
        doc = Document()
        p = doc.add_paragraph('УТВЕРЖДЕНО')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.bold = True
        name = doc.add_paragraph('АВТОМАТИЗИРОВАННАЯ СИСТЕМА ОРГАНИЗАЦИИ СВАДЕБНЫХ ТОРЖЕСТВ')
        name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        name.runs[0].font.bold = True
        doc.add_paragraph('Листов 20')
        doc.add_paragraph('2026')
        doc.add_paragraph('')
        doc.add_paragraph('Литера')
        doc.add_paragraph('Настоящее техническое задание определяет назначение системы.')
        return doc

    def test_valid_title_passes(self):
        self.assertEqual([], Tz34TitlePageRule().check(self.build_title()))

    def test_missing_requisites_are_reported(self):
        doc = Document()
        doc.add_paragraph('ТЕХНИЧЕСКОЕ ЗАДАНИЕ')
        doc.add_paragraph('Литера')
        found = Tz34TitlePageRule().check(doc)
        joined = ' '.join(v.message for v in found)
        for token in ['УТВЕРЖДЕНО', 'Листов', 'год', 'наименование']:
            self.assertIn(token, joined)


class Tz34StructureRuleTests(unittest.TestCase):
    SECTIONS = [
        'Общие сведения', 'Назначение и цели создания системы',
        'Характеристика объекта автоматизации', 'Требования к системе',
        'Состав и содержание работ по созданию системы',
        'Порядок контроля и приемки системы', 'Требования к документированию',
        'Порядок внесения изменений',
    ]

    def test_correct_structure_passes(self):
        doc = Document()
        for i, name in enumerate(self.SECTIONS, 1):
            doc.add_paragraph(f'{i}. {name}')
            doc.add_paragraph('Текст раздела.')
        self.assertEqual([], Tz34StructureRule().check(doc))

    def test_missing_and_reordered_sections_are_reported(self):
        doc = Document()
        doc.add_paragraph('1. Общие сведения')
        doc.add_paragraph('Текст.')
        doc.add_paragraph('2. Требования к системе')
        doc.add_paragraph('Текст.')
        doc.add_paragraph('3. Назначение и цели создания системы')
        doc.add_paragraph('Текст.')
        found = Tz34StructureRule().check(doc)
        joined = ' '.join(v.message for v in found)
        self.assertIn('Отсутствует обязательный раздел', joined)
        self.assertIn('Нарушен порядок разделов', joined)


class Tz34TextFormatRuleTests(unittest.TestCase):
    def test_sample_formatting_passes(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')  # маркер конца титульного листа
        body = doc.add_paragraph('Обычный текст раздела с отступом.')
        body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        body.paragraph_format.line_spacing = 1.5
        body.paragraph_format.first_line_indent = Cm(1.25)
        for run in body.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
        heading = doc.add_paragraph('1. Требования к системе')
        for run in heading.runs:
            run.font.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
        self.assertEqual([], Tz34TextFormatRule().check(doc))

    def test_wrong_font_size_and_spacing_are_reported(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')  # сбрасывает режим оглавления
        body = doc.add_paragraph('Текст набран не по образцу.')
        body.paragraph_format.line_spacing = 1.0
        body.runs[0].font.size = Pt(12)
        found = Tz34TextFormatRule().check(doc)
        joined = ' '.join(v.message for v in found)
        self.assertIn('12.0 pt', joined)
        self.assertIn('Междустрочный интервал 1.0', joined)

    def test_table_labels_are_skipped(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('ВВЕДЕНИЕ')  # сбрасывает режим оглавления
        doc.add_paragraph('Таблица 1')
        doc.add_paragraph('Продолжение таблицы 1')
        self.assertEqual([], Tz34TextFormatRule().check(doc))


class ClassificationForTzTitleTests(unittest.TestCase):
    def test_title_word_does_not_start_appendix(self):
        doc = Document()
        doc.add_paragraph('УТВЕРЖДЕНО')
        doc.add_paragraph('Приложение «Организация свадебных торжеств»')
        doc.add_paragraph('Литера')
        doc.add_paragraph('Текст ТЗ.')
        classified = classify_paragraphs(doc)
        self.assertFalse(classified[1].is_appendix)
        # Титул заканчивается на «Литера»: первые два абзаца — титульные
        self.assertTrue(classified[0].is_title_page)
        self.assertTrue(classified[1].is_title_page)
        self.assertFalse(classified[2].is_title_page)

    def test_breaks_rule_ignores_title_gaps(self):
        doc = Document()
        doc.add_paragraph('УТВЕРЖДЕНО')
        for _ in range(10):
            doc.add_paragraph('')
        doc.add_paragraph('Литера')
        doc.add_paragraph('Текст.')
        self.assertEqual([], [v for v in BreaksAndPaginationRule().check(doc)
                              if 'пустых абзаца' in v.message])


class TitleTableSkipTests(unittest.TestCase):
    def test_approval_sheet_table_is_not_checked(self):
        # Лист согласования на титуле («Подп. и дата») — часть титула, не данные
        doc = Document()
        approval = doc.add_table(rows=2, cols=2)
        approval.cell(0, 0).text = 'Подп. и дата'
        approval.cell(1, 0).text = 'Иванов И.И.'
        for _ in range(5):
            doc.add_paragraph('')
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('Таблица 1')
        data = doc.add_table(rows=4, cols=2)
        data.cell(0, 0).text = 'Показатель'
        for row in data.rows:
            for cell in row.cells:
                cell.paragraphs[0].add_run('значение')
        found = TablesFormatRule().check(doc)
        self.assertFalse(any('Таблица №1' in v.location for v in found),
                         'Таблица титульного листа не должна проверяться')


if __name__ == '__main__':
    unittest.main()
