"""Проверка полей страницы."""

import re

from docx import Document
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation


class PageMarginsRule(BaseRule):
    def __init__(self, top=2.0, bottom=2.0, left=3.0, right=1.5,
                 tolerance_cm: float = 0.01):
        self.top = top
        self.bottom = bottom
        self.left = left
        self.right = right
        self.tolerance_cm = tolerance_cm

    @staticmethod
    def _twips(length) -> Optional[int]:
        return round(length.twips) if length is not None else None

    @staticmethod
    def _cm(length) -> Optional[float]:
        return round(length.cm, 3) if length is not None else None

    def check(self, doc: Document) -> List[Violation]:
        violations = []
        expected = {
            'левое': self.left,
            'правое': self.right,
            'верхнее': self.top,
            'нижнее': self.bottom,
        }
        for i, section in enumerate(doc.sections, 1):
            actual_lengths = {
                'левое': section.left_margin,
                'правое': section.right_margin,
                'верхнее': section.top_margin,
                'нижнее': section.bottom_margin,
            }
            wrong = []
            for name, length in actual_lengths.items():
                actual = length.cm if length is not None else None
                if actual is None or abs(actual - expected[name]) > self.tolerance_cm:
                    wrong.append(f'{name}: {self._cm(length)} см')
            if wrong:
                violations.append(Violation(
                    rule_name='PageMarginsRule',
                    severity=Severity.ERROR,
                    location=f'Секция {i}',
                    message=(f'Поля секции заданы неверно ({"; ".join(wrong)}). Ожидается: '
                             f'левое {self.left} см, правое {self.right} см, верхнее {self.top} см, '
                             f'нижнее {self.bottom} см; допуск {self.tolerance_cm} см.')
                ))
        return violations
