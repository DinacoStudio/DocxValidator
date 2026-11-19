# -*- coding: utf-8 -*-
"""Регрессии проверки структуры ТЗ вне титула и содержания."""
import unittest

from docx import Document

from docx_rules.tz34_structure import Tz34StructureRule


SECTIONS = [label for label, _ in Tz34StructureRule.REQUIRED_SECTIONS]


class Tz34StructureRegressions(unittest.TestCase):
    def test_title_and_toc_do_not_supply_missing_sections(self):
        doc = Document()
        doc.add_paragraph(SECTIONS[0])  # До маркера содержания — титул.
        doc.add_paragraph('СОДЕРЖАНИЕ')
        for name in SECTIONS[1:]:
            doc.add_paragraph(name)  # Строки без номеров страниц тоже в содержании.
        doc.add_paragraph('ВВЕДЕНИЕ')
        violations = Tz34StructureRule().check(doc)
        missing = [v.message for v in violations if 'Отсутствует обязательный раздел' in v.message]
        self.assertEqual(len(SECTIONS), len(missing))
        for name in SECTIONS:
            self.assertTrue(any(f'«{name}»' in message for message in missing), name)

    def test_order_uses_body_not_toc_and_preserves_paragraph_locations(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        for name in SECTIONS:
            doc.add_paragraph(name)
        doc.add_paragraph('ВВЕДЕНИЕ')
        body = SECTIONS.copy()
        body[0], body[1] = body[1], body[0]
        for name in body:
            doc.add_paragraph(name)
        violations = Tz34StructureRule().check(doc)
        self.assertFalse(any('Отсутствует обязательный раздел' in v.message for v in violations))
        order = [v for v in violations if 'Нарушен порядок разделов' in v.message]
        self.assertEqual(1, len(order))
        self.assertEqual('Абзац 11', order[0].location)
        self.assertIn(f'«{SECTIONS[1]}»', order[0].message)

    def test_toc_order_does_not_affect_correct_body(self):
        doc = Document()
        doc.add_paragraph('СОДЕРЖАНИЕ')
        for name in reversed(SECTIONS):
            doc.add_paragraph(name)
        doc.add_paragraph('ВВЕДЕНИЕ')
        for index, name in enumerate(SECTIONS, 1):
            doc.add_paragraph(f'{index}. {name}')
        self.assertEqual([], Tz34StructureRule().check(doc))


if __name__ == '__main__':
    unittest.main()
