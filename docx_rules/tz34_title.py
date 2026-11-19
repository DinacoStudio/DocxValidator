"""Титульный лист технического задания по ГОСТ 19.104-78 (реквизиты листа согласования).

Образец оформления: «УТВЕРЖДЕНО», полное наименование системы (полужирное,
по центру), «Листов N», год, затем после титула — «Литера» на первой странице
основного текста. Учебная шапка («Министерство просвещения …») в этом титуле
не используется.
"""

import re
from typing import List

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.classification import classify_paragraphs
from docx_rules.helpers import is_word_heading


class Tz34TitlePageRule(BaseRule):
    """Проверяет наличие обязательных реквизитов титула ТЗ по ГОСТ 19.104."""

    APPROVED_REGEX = re.compile(r'^УТВЕРЖДЕНО\b', re.IGNORECASE)
    SHEETS_REGEX = re.compile(r'^ЛИСТОВ\s+\d+', re.IGNORECASE)
    YEAR_REGEX = re.compile(r'^\d{4}$')
    NAME_MIN_LENGTH = 20

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "Tz34TitlePageRule"

        paragraphs = classify_paragraphs(doc)
        first_body = next((i for i, p in enumerate(paragraphs) if not p.is_title_page), None)
        if not first_body:
            # Титул не распознан (нет «Содержания», «Литеры», «Аннотации» и т.п.) —
            # правило не может отличить титул от текста, не сообщаем ничего.
            return violations

        title_paras = paragraphs[:first_body]
        title_texts = [p.text.strip() for p in title_paras if p.text.strip()]

        if not any(self.APPROVED_REGEX.match(t) for t in title_texts):
            violations.append(Violation(
                rule_name, Severity.ERROR, "Титульный лист",
                "На титульном листе ТЗ отсутствует гриф «УТВЕРЖДЕНО»."))

        if not any(self.SHEETS_REGEX.match(t) for t in title_texts):
            violations.append(Violation(
                rule_name, Severity.ERROR, "Титульный лист",
                "На титульном листе не указано количество листов («Листов N»)."))

        if not any(self.YEAR_REGEX.match(t) for t in title_texts):
            violations.append(Violation(
                rule_name, Severity.ERROR, "Титульный лист",
                "На титульном листе не указан год разработки."))

        if not any(self._is_bold_centered(p) and len(p.text.strip()) >= self.NAME_MIN_LENGTH
                   for p in title_paras):
            violations.append(Violation(
                rule_name, Severity.ERROR, "Титульный лист",
                "Не найдено наименование системы/документа "
                "(полужирный центральный абзац на титульном листе)."))

        following_texts = [p.text.strip().upper() for p in paragraphs[first_body:first_body + 10]]
        if not any(t.startswith('ЛИТЕРА') for t in following_texts if t):
            violations.append(Violation(
                rule_name, Severity.WARNING, "Начало документа",
                "После титульного листа ожидается «Литера» "
                "(обозначение документа по ГОСТ 19.104)."))

        return violations

    @staticmethod
    def _is_bold_centered(paragraph) -> bool:
        runs = [r for r in paragraph.runs if r.text.strip()]
        if not runs:
            return False
        bold = all(r.font.bold for r in runs)
        centered = paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
        return bold and centered
