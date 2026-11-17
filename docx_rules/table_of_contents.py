"""Проверка оглавления и его записей."""

import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from typing import List, Optional

from docx_rules.base import BaseRule, Severity, Violation
from docx_rules.helpers import full_text, has_field, is_heading, is_toc_line
from docx_rules.classification import classify_paragraphs


class TableOfContentsRule(BaseRule):
    """Требует автоматическое поле TOC с настройкой заголовков и сохранённым результатом."""
    @staticmethod
    def _toc_fields(doc: Document):
        fields = []
        for simple in doc.element.body.iter(qn('w:fldSimple')):
            instruction = simple.get(qn('w:instr'), '')
            if re.match(r'^\s*TOC\b', instruction, re.I):
                cache = ''.join(node.text or '' for node in simple.iter(qn('w:t')))
                fields.append((instruction, cache))

        stack = []
        for node in doc.element.body.iter():
            if node.tag == qn('w:fldChar'):
                kind = node.get(qn('w:fldCharType'))
                if kind == 'begin':
                    stack.append({'instruction': '', 'cache': '', 'separated': False})
                elif kind == 'separate' and stack:
                    stack[-1]['separated'] = True
                elif kind == 'end' and stack:
                    field = stack.pop()
                    if re.match(r'^\s*TOC\b', field['instruction'], re.I):
                        fields.append((field['instruction'], field['cache']))
            elif node.tag == qn('w:instrText') and stack and not stack[-1]['separated']:
                stack[-1]['instruction'] += node.text or ''
            elif node.tag in (qn('w:t'), qn('w:tab')):
                for field in stack:
                    if field['separated']:
                        field['cache'] += ('\t' if node.tag == qn('w:tab') else node.text or '')
        return fields

    @staticmethod
    def _has_heading_settings(instruction: str) -> bool:
        return bool(re.search(r'\\o\s+"\s*\d+\s*-\s*\d+\s*"', instruction, re.I)
                    or re.search(r'\\t\s+"[^"\r\n]+"', instruction, re.I)
                    or re.search(r'(?:^|\s)\\u(?:\s|$)', instruction, re.I))

    def check(self, doc: Document) -> List[Violation]:
        rule_name = 'TableOfContentsRule'
        fields = self._toc_fields(doc)
        if not fields:
            has_manual_heading = any(
                p.text.strip().upper() in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ') for p in doc.paragraphs)
            return [Violation(
                rule_name, Severity.ERROR,
                'Оглавление' if has_manual_heading else 'Документ',
                ('Обнаружено ручное оглавление. Требуется настоящее автоматическое поле TOC Word.'
                 if has_manual_heading else
                 'Автоматическое оглавление не найдено: требуется настоящее поле TOC Word.'))
            ]

        violations = []
        if len(fields) > 1:
            violations.append(Violation(
                rule_name, Severity.WARNING, 'Оглавление',
                f'Найдено несколько полей TOC ({len(fields)}); проверьте, что оглавление не продублировано.'))
        for index, (instruction, cache) in enumerate(fields, 1):
            location = 'Оглавление' if len(fields) == 1 else f'Оглавление, поле TOC {index}'
            if not self._has_heading_settings(instruction):
                violations.append(Violation(
                    rule_name, Severity.ERROR, location,
                    'В поле TOC не задан источник заголовков (ожидается ключ \\o, \\t или \\u).'))
            if not cache.strip():
                violations.append(Violation(
                    rule_name, Severity.ERROR, location,
                    'Кэш поля TOC пуст: обновите оглавление в Word, чтобы в поле были сохранены пункты.'))
            elif not re.search(r'(?:\t|\.{2,}|…)\s*\d+\b', cache):
                violations.append(Violation(
                    rule_name, Severity.ERROR, location,
                    'В сохранённом результате TOC не найдены номера страниц; обновите оглавление в Word.'))

        # Автоматическое поле не отменяет требования к видимому заголовку оглавления.
        for item in self._legacy_check(doc):
            if item.location.startswith('Заголовок оглавления') or (
                    item.location == 'Документ' and 'Содержание' in item.message):
                violations.append(item)
        return violations

    def _legacy_check(self, doc: Document) -> List[Violation]:
        violations: List[Violation] = []
        rule_name = "TableOfContentsRule"

        toc_header_index: Optional[int] = None

        for idx, p in enumerate(doc.paragraphs):
            clean_text = p.text.strip().upper()
            if clean_text in ["СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"]:
                toc_header_index = idx
                loc = f"Заголовок оглавления (абзац {idx+1})"

                align = p.alignment
                if align is None and p.style and p.style.paragraph_format:
                    align = p.style.paragraph_format.alignment
                if align != WD_ALIGN_PARAGRAPH.CENTER:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Заголовок 'Содержание' должен быть выровнен по центру."
                        )
                    )

                is_bold = any(run.bold for run in p.runs if run.text.strip())
                if not is_bold and p.style and p.style.font:
                    is_bold = p.style.font.bold or False
                if not is_bold:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Название 'Содержание' должно быть выделено полужирным шрифтом."
                        )
                    )

                fmt = p.paragraph_format
                first_indent = fmt.first_line_indent
                if first_indent is None and p.style and p.style.paragraph_format:
                    first_indent = p.style.paragraph_format.first_line_indent
                if first_indent is None or abs(first_indent.cm - 1.25) > 0.05:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message="У заголовка 'Содержание' должна быть установлена красная строка."
                        )
                    )
                break

        if toc_header_index is None:
            violations.append(
                Violation(
                    rule_name=rule_name,
                    severity=Severity.ERROR,
                    location="Документ",
                    message="Раздел 'Содержание' (или 'Оглавление') не найден в документе."
                )
            )
            return violations

        toc_entries = []
        for p_idx in range(toc_header_index + 1, len(doc.paragraphs)):
            p = doc.paragraphs[p_idx]
            text = full_text(p).strip()
            if not text:
                continue

            is_toc_field = has_field(p, 'TOC')

            if is_heading(text) and not is_toc_line(text) and not is_toc_field:
                break

            has_page_break = any(
                'w:br' in run._element.xml and 'w:type="page"' in run._element.xml
                for run in p.runs
            )

            toc_entries.append((p_idx + 1, p))

            if has_page_break:
                break

        # Если содержание построено авто-полем TOC — ручные проверки записей не применяются
        toc_is_field = any(has_field(p, 'TOC') for _, p in toc_entries)

        if not toc_is_field:
            # Общий правый табулятор у записей содержания
            tab_positions = set()
            entries_without_tab = []
            for p_num, p in toc_entries:
                right_tabs = p._p.xpath('./w:pPr/w:tabs/w:tab[@w:val="right"]')
                if not right_tabs:
                    entries_without_tab.append(p_num)
                for tab in right_tabs:
                    pos = tab.get(qn('w:pos'))
                    if pos is not None:
                        tab_positions.add(int(pos))

            for p_num in entries_without_tab:
                p = dict(toc_entries)[p_num]
                text = full_text(p).strip()
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=f"Оглавление, абзац {p_num} ('{text[:30]}...')",
                        message="У пункта оглавления отсутствует правый табулятор (w:tabs, val=right) для позиции номера страницы."
                    )
                )
            if tab_positions and len(toc_entries) > len(entries_without_tab) and len(tab_positions) > 1:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location="Оглавление",
                        message=f"Правый табулятор пунктов оглавления должен быть на одной позиции "
                                f"(найдено разных позиций: {len(tab_positions)})."
                    )
                )

        for p_num, p in toc_entries:
            text = full_text(p).strip()
            loc = f"Оглавление, абзац {p_num} ('{text[:30]}...')"

            if toc_is_field:
                continue  # авто-оглавление: ручные проверки записей пропускаются

            if re.search(r'\.{3,}', text) or re.search(r'(\.\s*){4,}', text):
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Проставление точек не допустимо! Используйте стандартную табуляцию с заполнителем или авто-оглавление Word."
                    )
                )

            fmt = p.paragraph_format
            align = p.alignment
            if align is None and p.style and p.style.paragraph_format:
                align = p.style.paragraph_format.alignment

            if align != WD_ALIGN_PARAGRAPH.JUSTIFY:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message="Пункт оглавления должен быть выровнен по ширине (JUSTIFY)."
                    )
                )

            first_indent = fmt.first_line_indent
            if first_indent is None and p.style and p.style.paragraph_format:
                first_indent = p.style.paragraph_format.first_line_indent

            if first_indent is not None and abs(first_indent.cm) > 0.05:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.ERROR,
                        location=loc,
                        message=f"У пункта оглавления не должно быть красной строки (сейчас {round(first_indent.cm, 2)} см)."
                    )
                )

            num_match = re.match(r'^(?:Глава\s+)?\d+(?:\.\d+)*\.?([ 	]*)(?=[^\d.\s])', text, re.IGNORECASE)
            if num_match:
                spaces = num_match.group(1)
                if spaces != ' ':
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.ERROR,
                            location=loc,
                            message="Между номером главы/параграфа и названием должен быть строго один пробел."
                        )
                    )

        # Сверка заголовков документа с записями содержания
        chapter_entries = set()
        numbered_entries = set()
        for _, p in toc_entries:
            entry_text = full_text(p).strip()
            chapter_match = re.match(r'^Глава\s+(\d+)\b', entry_text, re.IGNORECASE)
            if chapter_match:
                chapter_entries.add(int(chapter_match.group(1)))
                continue
            number_match = re.match(r'^(\d+(?:\.\d+)*)\.?\s', entry_text)
            if number_match:
                numbered_entries.add(number_match.group(1))

        for idx, p in enumerate(classify_paragraphs(doc)):
            if p.is_title_page or p.is_toc:
                continue
            body_text = p.text.strip()
            loc = f"Абзац {idx + 1} ('{body_text[:30]}...')"

            chapter_match = re.match(r'^Глава\s+(\d+)\b', body_text, re.IGNORECASE)
            if chapter_match:
                number = int(chapter_match.group(1))
                covered = number in chapter_entries or any(
                    entry == str(number) or entry.startswith(f'{number}.')
                    for entry in numbered_entries)
                if not covered:
                    violations.append(
                        Violation(
                            rule_name=rule_name,
                            severity=Severity.WARNING,
                            location=loc,
                            message=f"Заголовка 'Глава {number}' нет в содержании."
                        )
                    )
                continue

            number_match = re.match(r'^(\d+(?:\.\d+)+)\.?\s+', body_text)
            if number_match and number_match.group(1) not in numbered_entries:
                violations.append(
                    Violation(
                        rule_name=rule_name,
                        severity=Severity.WARNING,
                        location=loc,
                        message=f"Параграфа {number_match.group(1)} нет в содержании."
                    )
                )

        return violations
