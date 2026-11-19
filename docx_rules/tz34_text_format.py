"""Форматирование текста технического задания по образцу ГОСТ-оформления.

Образец: Times New Roman 14 пт, междустрочный интервал 1,5, выравнивание
по ширине, красная строка 1,25 см. Заголовки разделов — полужирные, без
выравнивания по центру (в отличие от оформления МДК 04.01), поэтому
полужирные абзацы проверяются только по гарнитуре и кеглю.
"""

import re
from typing import List

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.classification import classify_paragraphs
from docx_rules.helpers import effective, full_text

# Надписи таблиц («Таблица 1», «Продолжение таблицы 1») — юрисдикция
# TablesFormatRule: в образце они набраны с одинарным интервалом.
_TABLE_LABEL_REGEX = re.compile(r'^(?:Продолжение\s+)?Таблиц[аы]\s+\d+', re.IGNORECASE)

FONT_NAME = "Times New Roman"
FONT_SIZE = 14.0
LINE_SPACING = 1.5
# В образце встречаются значения 1.52–1.56 (шапки таблиц, часть абзацев):
# Word хранит интервал с округлением, поэтому допуск шире, чем в МДК.
SPACING_TOLERANCE = 0.1
INDENT_CM = 1.25
INDENT_TOLERANCE_CM = 0.1


class Tz34TextFormatRule(BaseRule):
    """Проверяет шрифт, кегль, интервал и выравнивание текста ТЗ."""

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = self.__class__.__name__

        paragraphs = classify_paragraphs(doc)
        for idx, paragraph in enumerate(paragraphs, start=1):
            if getattr(paragraph, 'is_title_page', False) or getattr(paragraph, 'is_toc', False):
                continue
            text = full_text(paragraph).strip()
            if not text or _TABLE_LABEL_REGEX.match(text):
                continue

            runs = [r for r in paragraph.runs if r.text.strip()]
            if not runs:
                continue
            is_bold_paragraph = all(effective(paragraph, 'bold', r) for r in runs)
            loc = f"Абзац {idx} ('{text[:30]}...')"

            for run in runs:
                font_name = run.font.name
                if font_name and font_name != FONT_NAME:
                    violations.append(Violation(
                        rule_name, Severity.ERROR, loc,
                        f"Шрифт '{font_name}', а должен быть '{FONT_NAME}'."))
                size = run.font.size.pt if run.font.size else None
                if size and abs(size - FONT_SIZE) > 0.1:
                    violations.append(Violation(
                        rule_name, Severity.ERROR, loc,
                        f"Размер шрифта {size} pt, а должен быть {FONT_SIZE} pt."))

            line_spacing = paragraph.paragraph_format.line_spacing
            if line_spacing is not None and abs(float(line_spacing) - LINE_SPACING) > SPACING_TOLERANCE:
                violations.append(Violation(
                    rule_name, Severity.ERROR, loc,
                    f"Междустрочный интервал {line_spacing}, ожидается {LINE_SPACING}."))

            # Заголовки (полужирные абзацы) оформляются по центру не обязаны:
            # у них проверяется только гарнитура/кегль/интервал.
            if is_bold_paragraph:
                continue

            # Служебные реквизиты (по правому краю) и подписи (по центру)
            # допустимы; выравнивание влево — отступление от образца.
            alignment = paragraph.paragraph_format.alignment
            if alignment == WD_ALIGN_PARAGRAPH.LEFT:
                violations.append(Violation(
                    rule_name, Severity.WARNING, loc,
                    "Выравнивание текста должно быть по ширине (JUSTIFY)."))

            indent = paragraph.paragraph_format.first_line_indent
            if indent is not None and abs(indent.cm - INDENT_CM) > INDENT_TOLERANCE_CM:
                violations.append(Violation(
                    rule_name, Severity.WARNING, loc,
                    f"Красная строка равна {round(indent.cm, 2)} см, а должна быть {INDENT_CM} см."))

        return violations
