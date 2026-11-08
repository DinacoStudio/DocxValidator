"""
Создание копии документа с замечаниями в виде комментариев Word.

Каждое нарушение из отчёта валидатора привязывается к абзацу документа,
указанному в поле location нарушения. Расположение определяется по номеру
абзаца и текстовому фрагменту из location (в правилах он дублируется для
надёжности: «Абзац N ('текст...')»).
"""
from datetime import datetime
from typing import List, Optional, Tuple
import re

from docx import Document
from docx.text.paragraph import Paragraph
from docx_rules import Violation

AUTHOR = "DocxValidator"
INITIALS = "DV"

PARA_NUM_REGEX = re.compile(r'абза[цц]е?\s+(\d+)', re.IGNORECASE)
SNIPPET_REGEX = re.compile(r"'([^']*)'")
TABLE_NUM_REGEX = re.compile(r'Таблица\s+№\s*(\d+)')
SOURCES_ELEMENT_REGEX = re.compile(r'Список источников, элемент\s+(\d+)')


def _normalize(text: str) -> str:
    return re.sub(r'\s+', ' ', text).strip()


def _parse_paragraph_reference(location: str) -> Tuple[Optional[int], Optional[str]]:
    """
    Достаёт из location номер абзаца и текстовый фрагмент после него.
    Возвращает (номер, фрагмент или None).
    """
    match = PARA_NUM_REGEX.search(location)
    if not match:
        return None, None

    number = int(match.group(1))
    snippet_match = SNIPPET_REGEX.search(location[match.end():])
    snippet = None
    if snippet_match:
        snippet = _normalize(snippet_match.group(1)).rstrip('.')

    return number, snippet


def _find_with_runs(paragraphs: List[Paragraph], start_idx: int) -> Optional[Paragraph]:
    """
    Ищет ближайший абзац с текстом (комментарий можно привязать только к run'ам),
    двигаясь от start_idx вперёд, затем назад.
    """
    for offset in range(len(paragraphs)):
        for idx in (start_idx + offset, start_idx - offset):
            if 0 <= idx < len(paragraphs) and paragraphs[idx].runs:
                return paragraphs[idx]
    return None


def resolve_anchor(doc: Document, location: str) -> Optional[Paragraph]:
    """Определяет абзац документа, к которому привязывается замечание."""
    paragraphs = doc.paragraphs
    if not paragraphs:
        return None

    # 1. Таблица — замечание в первую ячейку
    table_match = TABLE_NUM_REGEX.search(location)
    if table_match:
        table_idx = int(table_match.group(1)) - 1
        if 0 <= table_idx < len(doc.tables):
            cell = doc.tables[table_idx].rows[0].cells[0]
            anchor = _find_with_runs(cell.paragraphs, 0)
            if anchor:
                return anchor

    # 2. Введение
    if location.strip() == "Введение":
        for p in paragraphs:
            if re.match(r'^ВВЕДЕНИЕ', p.text.strip().upper()):
                return _find_with_runs(paragraphs, paragraphs.index(p))

    # 3. Элемент списка источников (N-й абзац после заголовка СПИСОК)
    sources_match = SOURCES_ELEMENT_REGEX.search(location)
    if sources_match:
        header_idx = None
        for idx, p in enumerate(paragraphs):
            if re.match(r'^СПИСОК', p.text.strip().upper()):
                header_idx = idx
                break
        if header_idx is not None:
            element_idx = int(sources_match.group(1))
            non_empty = [
                p for p in paragraphs[header_idx + 1:]
                if p.text.strip() and p.runs
            ]
            if 1 <= element_idx <= len(non_empty):
                return non_empty[element_idx - 1]

    # 4. Нарушения, не привязанные к конкретному абзацу, — в самое начало документа
    if not PARA_NUM_REGEX.search(location):
        return _find_with_runs(paragraphs, 0)

    # 5. Общий случай: «Абзац N» (+ фрагмент текста для проверки)
    number, snippet = _parse_paragraph_reference(location)

    candidates = []
    for candidate in (number - 1, number, number - 2):  # нумерация в правилах 1-базная
        if 0 <= candidate < len(paragraphs):
            candidates.append(paragraphs[candidate])

    if snippet:
        for candidate in candidates:
            if _normalize(candidate.text).startswith(snippet) and candidate.runs:
                return candidate

    if candidates:
        return _find_with_runs(paragraphs, paragraphs.index(candidates[0]))
    return None


def make_notes_docx(src_path: str, violations: List[Violation], dst_path: str) -> Tuple[int, int]:
    """
    Создаёт копию документа src_path, в которой каждое нарушение отмечено
    комментарием Word. Оригинал не изменяется.

    Возвращает (число добавленных замечаний, число непривязанных нарушений).
    """
    doc = Document(src_path)
    placed = 0
    unresolved = 0

    for violation in violations:
        anchor = resolve_anchor(doc, violation.location)
        if anchor is None:
            unresolved += 1
            continue

        lines = [f"[{violation.severity.value}] {violation.rule_name}", violation.location]
        if violation.message:
            lines.append(violation.message)

        comment = doc.add_comment(
            runs=anchor.runs,
            text=lines[0],
            author=AUTHOR,
            initials=INITIALS,
        )
        for line in lines[1:]:
            comment.add_paragraph(line)
        placed += 1

    doc.core_properties.comments = f"Замечания DocxValidator от {datetime.now():%d.%m.%Y %H:%M}"
    doc.save(dst_path)
    return placed, unresolved
