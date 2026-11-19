"""Структура технического задания на автоматизированную систему (по ГОСТ 34.602).

Набор и порядок разделов соответствуют образцу оформления ПР2: «Общие
сведения», «Назначение и цели создания системы», «Характеристика объекта
автоматизации», «Требования к системе», «Состав и содержание работ по
созданию системы», «Порядок контроля и приемки системы», «Требования
к документированию», «Порядок внесения изменений».
"""

import re
from typing import List, Optional, Tuple

from docx import Document

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.classification import classify_paragraphs

# Номера разделов/подразделов и завершающие точки-двоеточия срезаются,
# пробелы схлопываются: «7. Требования к документированию» → «ТРЕБОВАНИЯ К ДОКУМЕНТИРОВАНИЮ».
_NUMBER_PREFIX = re.compile(r'^\s*\d+(?:\.\d+)*[.)]?\s*')


def _clean_heading(text: str) -> str:
    text = _NUMBER_PREFIX.sub('', text.strip()).upper().replace('Ё', 'Е')
    return re.sub(r'\s+', ' ', text).rstrip('.:')


class Tz34StructureRule(BaseRule):
    """Проверяет наличие и порядок разделов ТЗ на автоматизированную систему."""

    REQUIRED_SECTIONS: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
        ('Общие сведения', ('ОБЩИЕ СВЕДЕНИЯ',)),
        ('Назначение и цели создания системы', ('НАЗНАЧЕНИЕ И ЦЕЛИ СОЗДАНИЯ СИСТЕМЫ',)),
        ('Характеристика объекта автоматизации', ('ХАРАКТЕРИСТИКА ОБЪЕКТА АВТОМАТИЗАЦИИ',)),
        ('Требования к системе', ('ТРЕБОВАНИЯ К СИСТЕМЕ',)),
        ('Состав и содержание работ по созданию системы', (
            'СОСТАВ И СОДЕРЖАНИЕ РАБОТ ПО СОЗДАНИЮ СИСТЕМЫ',
            'СОСТАВ И СОДЕРЖАНИЕ РАБОТ',
        )),
        ('Порядок контроля и приемки системы', (
            'ПОРЯДОК КОНТРОЛЯ И ПРИЕМКИ СИСТЕМЫ',
            'ПОРЯДОК КОНТРОЛЯ И ПРИЕМКИ',
        )),
        ('Требования к документированию', ('ТРЕБОВАНИЯ К ДОКУМЕНТИРОВАНИЮ',)),
        ('Порядок внесения изменений', ('ПОРЯДОК ВНЕСЕНИЯ ИЗМЕНЕНИЙ',)),
    )

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = self.__class__.__name__

        headings = [(_clean_heading(p.text), index + 1)
                    for index, p in enumerate(classify_paragraphs(doc))
                    if p.text.strip() and not p.is_title_page and not p.is_toc]
        found: List[Tuple[str, Optional[int]]] = []

        for label, variants in self.REQUIRED_SECTIONS:
            occurrences = [number for heading, number in headings if heading in variants]
            if not occurrences:
                violations.append(Violation(
                    rule_name, Severity.ERROR, 'Структура документа',
                    f'Отсутствует обязательный раздел ТЗ «{label}».'))
                found.append((label, None))
            else:
                found.append((label, occurrences[0]))

        present = [(label, position) for label, position in found if position is not None]
        for (previous_label, previous), (label, position) in zip(present, present[1:]):
            if position < previous:
                violations.append(Violation(
                    rule_name, Severity.ERROR, f'Абзац {position}',
                    f'Нарушен порядок разделов ТЗ: «{label}» расположен перед «{previous_label}».'))

        return violations
