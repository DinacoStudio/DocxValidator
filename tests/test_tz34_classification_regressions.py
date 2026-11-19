import unittest

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx_rules.base import Severity
from docx_rules.classification import classify_paragraphs
from docx_rules.tz34_title import Tz34TitlePageRule
from docx_rules.tz34_structure import Tz34StructureRule


class TzClassificationRegressions(unittest.TestCase):
    def test_numbered_boundary_requires_title_evidence(self):
        doc = Document()
        for text in ('УТВЕРЖДЕНО', '09.02.07. Код специальности', 'Листов 20',
                     '2026', '1. Общие сведения', 'Текст', 'Приложение А'):
            doc.add_paragraph(text)
        paragraphs = classify_paragraphs(doc)
        self.assertTrue(all(p.is_title_page for p in paragraphs[:4]))
        self.assertFalse(paragraphs[4].is_title_page)
        self.assertTrue(paragraphs[-1].is_appendix)
        self.assertEqual([], classify_paragraphs(Document()))

        no_title = Document()
        no_title.add_paragraph('1. Общие сведения')
        no_title.add_paragraph('Текст')
        self.assertFalse(any(p.is_title_page for p in classify_paragraphs(no_title)))

    def test_unknown_boundary_warns_and_title_name_is_not_appendix(self):
        doc = Document()
        doc.add_paragraph('УТВЕРЖДЕНО')
        doc.add_paragraph('Приложение «Название системы»')
        doc.add_paragraph('2026')
        self.assertFalse(any(p.is_appendix for p in classify_paragraphs(doc)))
        violations = Tz34TitlePageRule().check(doc)
        self.assertEqual(len(violations), 1)
        self.assertEqual(violations[0].severity, Severity.WARNING)
        self.assertIn('границу', violations[0].message)

    def test_appendix_before_later_word_heading(self):
        doc = Document()
        for text in ('Основной текст', 'Приложение А', 'Описание', 'ЗАКЛЮЧЕНИЕ'):
            doc.add_paragraph(text)
        paragraphs = classify_paragraphs(doc)
        self.assertFalse(paragraphs[0].is_title_page)
        self.assertTrue(paragraphs[1].is_appendix)

    def test_inherited_title_style_and_direct_override(self):
        doc = Document()
        style = doc.styles.add_style('TitleInheritedRegression', WD_STYLE_TYPE.PARAGRAPH)
        style.font.bold = True
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph('УТВЕРЖДЕНО')
        name = doc.add_paragraph('Полное наименование автоматизированной системы', style=style)
        doc.add_paragraph('Листов 20')
        doc.add_paragraph('2026')
        doc.add_paragraph('Литера')
        self.assertEqual([], Tz34TitlePageRule().check(doc))
        name.runs[0].bold = False
        self.assertTrue(any('наименование' in v.message for v in Tz34TitlePageRule().check(doc)))

    def test_manual_toc_without_page_numbers_does_not_satisfy_structure(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        for number, (label, _) in enumerate(Tz34StructureRule.REQUIRED_SECTIONS, 1):
            doc.add_paragraph(f'{number}. {label}')
        self.assertTrue(all(p.is_toc for p in classify_paragraphs(doc)))
        errors = Tz34StructureRule().check(doc)
        self.assertEqual(sum('Отсутствует обязательный раздел' in v.message for v in errors), 8)

    def test_unnumbered_body_heading_matches_paginated_toc_entry(self):
        doc = Document()
        for text in ('СОДЕРЖАНИЕ', '1. Общие сведения\t3',
                     '2. Требования к системе\t4', 'Общие сведения', 'Текст раздела'):
            doc.add_paragraph(text)
        ps = classify_paragraphs(doc)
        self.assertTrue(ps[2].is_toc)
        self.assertFalse(ps[3].is_toc)
        self.assertFalse(ps[4].is_toc)
        self.assertFalse(any('«Общие сведения»' in v.message for v in Tz34StructureRule().check(doc)))

    def test_repeated_manual_entry_marks_body_boundary(self):
        doc = Document()
        for text in ('СОДЕРЖАНИЕ', '1. Общие сведения', '2. Требования к системе',
                     '1. Общие сведения', 'Текст раздела'):
            doc.add_paragraph(text)
        ps = classify_paragraphs(doc)
        self.assertTrue(ps[1].is_toc)
        self.assertTrue(ps[2].is_toc)
        self.assertFalse(ps[3].is_toc)

    def test_styled_unnumbered_body_heading(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('1. Общие сведения\t3')
        doc.add_paragraph('Общие положения', style='Heading 1')
        self.assertFalse(classify_paragraphs(doc)[2].is_toc)

    def test_page_break_alone_does_not_end_manual_toc(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('1. Общие сведения')
        entry = doc.add_paragraph('2. Требования к системе')
        entry.paragraph_format.page_break_before = True
        self.assertTrue(classify_paragraphs(doc)[2].is_toc)

    def test_same_name_at_different_levels_remains_in_toc(self):
        doc = Document()
        for text in ('СОДЕРЖАНИЕ', '1. Общие сведения', '1.1 Общие сведения',
                     '1. Общие сведения', 'Основной текст'):
            doc.add_paragraph(text)
        ps = classify_paragraphs(doc)
        self.assertTrue(ps[1].is_toc)
        self.assertTrue(ps[2].is_toc)
        self.assertFalse(ps[3].is_toc)

    def test_unnumbered_first_section_after_title_before_appendix(self):
        doc = Document()
        for text in ('УТВЕРЖДЕНО', 'Название системы', 'Листов 20', '2026',
                     'Общие сведения', 'Текст раздела', 'Приложение А', 'ЗАКЛЮЧЕНИЕ'):
            doc.add_paragraph(text)
        ps = classify_paragraphs(doc)
        self.assertTrue(ps[3].is_title_page)
        self.assertFalse(ps[4].is_title_page)
        self.assertFalse(ps[5].is_title_page)
        self.assertTrue(ps[6].is_appendix)
        self.assertFalse(any('«Общие сведения»' in v.message for v in Tz34StructureRule().check(doc)))

    def test_late_appendix_does_not_discard_title(self):
        doc = Document()
        for text in ('УТВЕРЖДЕНО', 'Название системы', 'Листов 20', '2026',
                     'Текст документа', 'Приложение А', 'Описание', 'ЗАКЛЮЧЕНИЕ'):
            doc.add_paragraph(text)
        ps = classify_paragraphs(doc)
        self.assertTrue(ps[0].is_title_page)
        self.assertFalse(ps[5].is_title_page)
        self.assertTrue(ps[5].is_appendix)

    def test_missing_approval_still_checks_identifiable_title(self):
        doc = Document()
        p = doc.add_paragraph('Полное наименование автоматизированной системы')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        for text in ('Листов 20', '2026', '1. Общие сведения', 'Текст'):
            doc.add_paragraph(text)
        ps = classify_paragraphs(doc)
        self.assertTrue(ps[0].is_title_page)
        self.assertFalse(ps[3].is_title_page)
        violations = Tz34TitlePageRule().check(doc)
        self.assertTrue(any(v.severity == Severity.ERROR and 'УТВЕРЖДЕНО' in v.message
                            for v in violations))
        self.assertFalse(any('границу' in v.message for v in violations))

    def test_toc_ends_at_first_level_heading_not_entry(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        doc.add_paragraph('1. Общие сведения\t3')
        doc.add_paragraph('1. Общие сведения')
        doc.add_paragraph('Основной текст')
        paragraphs = classify_paragraphs(doc)
        self.assertTrue(paragraphs[1].is_toc)
        self.assertFalse(paragraphs[2].is_toc)
        self.assertFalse(paragraphs[3].is_toc)


if __name__ == '__main__':
    unittest.main()
