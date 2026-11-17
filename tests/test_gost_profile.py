# -*- coding: utf-8 -*-
"""Профильные проверки технического задания по ГОСТ 19.201-78."""
import unittest

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_rules import Severity, WhiteFillRule
from gost_rules import (
    AnnotationRule,
    Gost19201PageNumberingRule,
    Gost19201StructureRule,
    TechnicalSpecificationRequisitesRule,
    TechnicalSpecificationTablesRule,
)
from profiles import get_rules


SECTIONS = (
    "1 ВВЕДЕНИЕ",
    "2 ОСНОВАНИЯ ДЛЯ РАЗРАБОТКИ",
    "3 НАЗНАЧЕНИЕ РАЗРАБОТКИ",
    "4 ТРЕБОВАНИЯ К ПРОГРАММЕ ИЛИ ПРОГРАММНОМУ ИЗДЕЛИЮ",
    "5 ТРЕБОВАНИЯ К ПРОГРАММНОЙ ДОКУМЕНТАЦИИ",
    "6 ТЕХНИКО-ЭКОНОМИЧЕСКИЕ ПОКАЗАТЕЛИ",
    "7 СТАДИИ И ЭТАПЫ РАЗРАБОТКИ",
    "8 ПОРЯДОК КОНТРОЛЯ И ПРИЁМКИ",
)


def add_page_field(paragraph):
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    field.append(OxmlElement("w:r"))
    paragraph._p.append(field)


def format_table(table):
    borders = OxmlElement("w:tblBorders")
    for name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = OxmlElement("w:" + name)
        edge.set(qn("w:val"), "single")
        borders.append(edge)
    table._tbl.tblPr.append(borders)
    header = OxmlElement("w:tblHeader")
    header.set(qn("w:val"), "true")
    table.rows[0]._tr.get_or_add_trPr().append(header)


class AnnotationTests(unittest.TestCase):
    def test_complete_annotation(self):
        doc = Document()
        title = doc.add_paragraph("ТЕХНИЧЕСКОЕ ЗАДАНИЕ")
        title.add_run().add_break(WD_BREAK.PAGE)
        doc.add_paragraph("АННОТАЦИЯ")
        doc.add_paragraph(
            "Назначение программы и область применения описаны для пользователей. "
            "Основные функции обеспечивают обработку данных.")
        doc.add_paragraph("ВВЕДЕНИЕ")
        self.assertEqual([], AnnotationRule().check(doc))

    def test_missing_content_and_wrong_page(self):
        doc = Document()
        doc.add_paragraph("Титул")
        doc.add_paragraph("АННОТАЦИЯ")
        doc.add_paragraph("Краткий текст.")
        doc.add_paragraph("ВВЕДЕНИЕ")
        found = AnnotationRule().check(doc)
        self.assertTrue(any(v.severity == Severity.ERROR and "второй" in v.message for v in found))
        self.assertEqual(4, sum(v.severity == Severity.WARNING for v in found))


class StructureTests(unittest.TestCase):
    def test_required_sections_in_order(self):
        doc = Document()
        for heading in SECTIONS:
            doc.add_paragraph(heading)
        doc.add_paragraph("ПРИЛОЖЕНИЕ А")
        self.assertEqual([], Gost19201StructureRule().check(doc))

    def test_missing_and_out_of_order(self):
        doc = Document()
        for heading in reversed(SECTIONS[:-1]):
            doc.add_paragraph(heading)
        found = Gost19201StructureRule().check(doc)
        self.assertTrue(any("Порядок контроля" in v.message for v in found))
        self.assertTrue(any("Нарушен порядок" in v.message for v in found))


class RequisitesTests(unittest.TestCase):
    def test_requisites_and_dates(self):
        doc = Document()
        doc.add_paragraph("Наименование программы: Валидатор документов")
        doc.add_paragraph("Заказчик: ООО Заказчик")
        doc.add_paragraph("Разработчик: ООО Исполнитель")
        doc.add_paragraph("Основание для разработки: договор № 12-А от 01.02.2026")
        doc.add_paragraph("Начало разработки: 02.02.2026")
        doc.add_paragraph("Окончание разработки: 30.06.2026")
        self.assertEqual([], TechnicalSpecificationRequisitesRule().check(doc))

    def test_reverse_dates(self):
        doc = Document()
        doc.add_paragraph("Начало разработки: 30.06.2026")
        doc.add_paragraph("Окончание разработки: 01.02.2026")
        found = TechnicalSpecificationRequisitesRule().check(doc)
        self.assertTrue(any("не может предшествовать" in v.message for v in found))


class TablesTests(unittest.TestCase):
    def test_two_formatted_referenced_tables(self):
        doc = Document()
        doc.add_paragraph("Характеристики приведены в таблице 1, сроки — в таблице 2.")
        for number in (1, 2):
            doc.add_paragraph(f"Таблица {number} – Данные")
            table = doc.add_table(rows=2, cols=2)
            format_table(table)
        self.assertEqual([], TechnicalSpecificationTablesRule().check(doc))

    def test_not_enough_tables(self):
        doc = Document()
        doc.add_paragraph("Таблица 1 – Данные")
        table = doc.add_table(rows=1, cols=1)
        format_table(table)
        found = TechnicalSpecificationTablesRule().check(doc)
        self.assertTrue(any("не менее двух" in v.message for v in found))


class NumberingAndProfilesTests(unittest.TestCase):
    def test_gost_numbering_starts_with_three(self):
        doc = Document()
        doc.add_paragraph("Титул")
        doc.add_section(WD_SECTION.NEW_PAGE)
        doc.add_paragraph("АННОТАЦИЯ")
        doc.add_section(WD_SECTION.NEW_PAGE)
        doc.add_paragraph("ВВЕДЕНИЕ")
        main = doc.sections[2]
        main.header.is_linked_to_previous = False
        add_page_field(main.header.paragraphs[0])
        pg_num = OxmlElement("w:pgNumType")
        pg_num.set(qn("w:start"), "3")
        main._sectPr.append(pg_num)
        self.assertEqual([], Gost19201PageNumberingRule().check(doc))

    def test_profiles_are_independent_and_gost_extends_standard(self):
        standard = get_rules("standard")
        gost = get_rules("gost-19-201")
        self.assertTrue(any(isinstance(rule, WhiteFillRule) for rule in standard))
        self.assertTrue(any(type(rule).__name__ == "BreaksAndPaginationRule"
                            for rule in standard))
        self.assertFalse(any(isinstance(rule, AnnotationRule) for rule in standard))
        self.assertTrue(any(isinstance(rule, AnnotationRule) for rule in gost))
        self.assertEqual([type(rule) for rule in standard],
                         [type(rule) for rule in gost[:len(standard)]])
        self.assertIsNot(get_rules("standard")[0], get_rules("standard")[0])

    def test_unknown_profile(self):
        with self.assertRaises(ValueError):
            get_rules("unknown")


if __name__ == "__main__":
    unittest.main()
