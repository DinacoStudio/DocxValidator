"""Проверка активной сетки текста (w:docGrid) в секциях документа."""

import re

from docx import Document
from docx.oxml.ns import qn
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation


class TextGridRule(BaseRule):
    """Проверяет активную сетку и целое число её шагов в рабочей высоте страницы."""
    def __init__(self, twip_tolerance: int = 2):
        self.twip_tolerance = twip_tolerance

    @staticmethod
    def _twips(length) -> Optional[int]:
        return round(length.twips) if length is not None else None

    def check(self, doc: Document) -> List[Violation]:
        violations = []
        for index, section in enumerate(doc.sections, 1):
            grids = section._sectPr.findall(qn('w:docGrid'))
            errors = []
            if len(grids) > 1:
                errors.append('В секции задано несколько противоречивых w:docGrid.')
            if not grids:
                errors.append('Отсутствует сетка текста (w:docGrid).')

            page = self._twips(section.page_height)
            top = self._twips(section.top_margin)
            bottom = self._twips(section.bottom_margin)
            usable_height = None if None in (page, top, bottom) else page - top - bottom
            if usable_height is None or usable_height <= 0:
                errors.append('Невозможно определить положительную рабочую высоту страницы (page_height - top - bottom).')

            for grid in grids:
                kind = grid.get(qn('w:type'), 'none')
                if kind not in ('lines', 'linesAndChars'):
                    errors.append(f'Ожидается активная сетка строк (type lines или linesAndChars), получено: {kind}.')
                pitch_text = grid.get(qn('w:linePitch'), '')
                pitch = int(pitch_text) if re.fullmatch(r'[+]?[0-9]+', pitch_text) and int(pitch_text) > 0 else None
                if pitch is None:
                    errors.append('linePitch сетки строк должен быть положительным целым.')
                if pitch is not None and usable_height is not None and usable_height > 0:
                    steps = max(1, round(usable_height / pitch))
                    mismatch = abs(usable_height - steps * pitch)
                    if mismatch > self.twip_tolerance:
                        errors.append(
                            f'Рабочая высота {usable_height} twips не содержит целого числа шагов '
                            f'linePitch={pitch} twips (ближайшее: {steps}, расхождение {mismatch} twips; '
                            f'допуск {self.twip_tolerance} twips).')
                space = grid.get(qn('w:charSpace'))
                if space is not None and not re.fullmatch(r'[+-]?[0-9]+', space):
                    errors.append('charSpace должен быть целым числом (включая отрицательные и ноль).')
            violations.extend(Violation('TextGridRule', Severity.ERROR, f'Секция {index}', e) for e in errors)
        return violations
