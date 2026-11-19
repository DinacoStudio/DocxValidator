"""Осторожные предупреждения о ручной пагинации и небезопасных настройках абзацев."""

import re

from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from typing import List

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.classification import classify_paragraphs
from docx_rules.helpers import full_text, is_toc_line, effective, is_heading

class BreaksAndPaginationRule(BaseRule):
    """Осторожные предупреждения о ручной пагинации и небезопасных настройках абзацев."""

    # Заголовок раздела/параграфа: словесный («ВВЕДЕНИЕ», «ГЛАВА 1 …») или
    # нумерованный («1. Общие сведения», «2.1. Назначение»), короткая строка.
    PARAGRAPH_HEADING_REGEX = re.compile(r'^\d+(?:\.\d+)*[.)]?\s+\S')

    @classmethod
    def _is_heading_after_gap(cls, paragraph) -> bool:
        """True, если за последовательностью пустых абзацев следует заголовок раздела."""
        text = full_text(paragraph).strip()
        if not text or len(text) > 120:
            return False
        return bool(is_heading(text) or cls.PARAGRAPH_HEADING_REGEX.match(text))

    def check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = 'BreaksAndPaginationRule'
        paragraphs = classify_paragraphs(doc)
        # Титульный лист существует, если есть хотя бы один абзац после него:
        # пустые серии на титуле — норма свободной вёрстки (например, «Листов 20»,
        # год и подписи в ТЗ по ГОСТ 19.104), их не сообщаем.
        title_exists = any(not p.is_title_page for p in paragraphs)

        empty_start = None
        for idx, paragraph in enumerate(paragraphs):
            text = full_text(paragraph).strip()
            has_structural_break = bool(paragraph._p.xpath('.//w:br[@w:type="page"] | ./w:pPr/w:sectPr'))
            is_plain_empty = not text and not has_structural_break and not paragraph._p.xpath('.//w:drawing | .//w:pict')
            if is_plain_empty and empty_start is None:
                empty_start = idx
            if (not is_plain_empty or idx == len(paragraphs) - 1) and empty_start is not None:
                end = idx if is_plain_empty else idx - 1
                count = end - empty_start + 1
                # По требованиям МДК (слайд 204) перед началом очередного параграфа
                # ставится два Enter'а: ровно 2 пустых абзаца, за которыми следует
                # заголовок раздела — норма, а не нарушение (иначе правило
                # противоречит SectionFormatRule, требующему ровно 2 Enter'а).
                follows_heading = (not is_plain_empty and self._is_heading_after_gap(paragraphs[idx]))
                on_title_page = title_exists and paragraphs[empty_start].is_title_page
                if count >= 2 and not (count == 2 and follows_heading) and not on_title_page:
                    violations.append(Violation(
                        rule_name, Severity.WARNING,
                        f'Абзацы {empty_start + 1}–{end + 1}',
                        f'Обнаружено {count} пустых абзаца подряд. Не используйте повторные Enter для отступов или перехода на страницу.'))
                empty_start = None

            location = f'Абзац {idx + 1}' + (f" ('{paragraph.text.strip()[:30]}...')" if paragraph.text.strip() else '')
            soft_breaks = paragraph._p.xpath('.//w:br[not(@w:type) or @w:type="textWrapping"] | .//w:cr')
            if soft_breaks:
                violations.append(Violation(
                    rule_name, Severity.WARNING, location,
                    f'Найден ручной перенос строки (soft break), {len(soft_breaks)} шт. Проверьте, что он нужен, а не заменяет новый абзац.'))

            page_breaks = paragraph._p.xpath('.//w:br[@w:type="page"]')
            if page_breaks:
                violations.append(Violation(
                    rule_name, Severity.WARNING, location,
                    f'Найден ручной разрыв страницы, {len(page_breaks)} шт. Для устойчивой вёрстки предпочтительно свойство «с новой страницы».'))

            sect_pr = paragraph._p.find('./' + qn('w:pPr') + '/' + qn('w:sectPr'))
            if sect_pr is not None:
                type_node = sect_pr.find(qn('w:type'))
                section_type = type_node.get(qn('w:val')) if type_node is not None else 'nextPage'
                if not paragraph.text.strip() or section_type == 'continuous':
                    violations.append(Violation(
                        rule_name, Severity.WARNING, location,
                        f'Найден разрыв раздела ({section_type}). Проверьте, что отдельный раздел необходим и не создан пустым Enter.'))

            stripped = paragraph.text.strip()
            if stripped and is_heading(stripped) and effective(paragraph, 'keep_with_next') is False:
                violations.append(Violation(
                    rule_name, Severity.WARNING, location,
                    'У заголовка явно отключено keep_with_next: заголовок может остаться внизу страницы без текста.'))
            if stripped and not getattr(paragraph, 'is_toc', False) and effective(paragraph, 'widow_control') is False:
                violations.append(Violation(
                    rule_name, Severity.WARNING, location,
                    'Явно отключен запрет висячих строк (widow_control); проверьте пагинацию абзаца.'))

        return violations
