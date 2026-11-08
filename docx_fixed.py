"""
Создание исправленной версии документа по требованиям оформления.

Оригинал не изменяется: документ открывается заново, правится в памяти и
сохраняется в новый файл. Исправляется только то, что определяется reliably
(надёжно) и не разрушает содержимое:

- поля страницы (Л: 3 см, П: 1,5 см, В/Н: 2 см);
- основной текст: по ширине, красная строка 1,25 см, интервал 1,5,
  Times New Roman 14, чёрный цвет, интервалы до/после 0, запрет висячих строк;
- заголовки глав/параграфов: полужирные, по центру, без красной строки;
  крупные разделы начинаются с новой страницы, параграфы — нет;
- убирается белая заливка фона (w:shd fill="FFFFFF") и белое выделение;
- длинное тире (—) заменяется на тире средней длины (–).

Титульный лист, оглавление, списки, формулы, рисунки, подписи и приложения
не трогаются — их форматирование специфично и легко сломать автоматикой.
"""
from typing import Dict
import re

from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

from docx_rules import classify_paragraphs, is_toc_line, SectionFormatRule, ListsFormatRule

# Требуемые значения — те же, что проверяют PageMarginsRule и MainTextFormatRule
MARGINS = {"left": 3.0, "right": 1.5, "top": 2.0, "bottom": 2.0}
FONT_NAME = "Times New Roman"
FONT_SIZE = 14.0
LINE_SPACING = 1.5
FIRST_LINE_INDENT_CM = 1.25

MAJOR_HEADER_REGEX = re.compile(SectionFormatRule.MAJOR_HEADER_REGEX, re.IGNORECASE)
APPENDIX_HEADER_REGEX = re.compile(r'^ПРИЛОЖЕНИЕ', re.IGNORECASE)
SUBHEADING_REGEX = re.compile(r'^\d+(\.\d+)+\.?\s+')
NUMBERED_SECTION_REGEX = re.compile(r'^\d+\.?\s+\S')
CAPTION_REGEX = re.compile(r'^Рис(унок)?\.?\s*(\d+)', re.IGNORECASE)
TABLE_LABEL_REGEX = re.compile(r'^(Таблица\s+\d+|Продолжение таблицы\s+\d+)', re.IGNORECASE)

WHITE_FILLS = {"ffffff", "white"}

_SECTION_RULE = SectionFormatRule()


def _is_white_shd(shd) -> bool:
    if shd is None:
        return False
    if (shd.get(qn('w:val')) or '').lower() == 'nil':
        return False
    fill = (shd.get(qn('w:fill')) or '').strip().lstrip('#').lower()
    return fill in WHITE_FILLS


def _remove_white_shading(element) -> int:
    """Убирает белую заливку и белое выделение внутри элемента. Возвращает число правок."""
    removed = 0
    # Список, а не итератор: удалять элементы прямо во время обхода дерева нельзя
    for shd in list(element.iter(qn('w:shd'))):
        if _is_white_shd(shd):
            shd.getparent().remove(shd)
            removed += 1
    for highlight in list(element.iter(qn('w:highlight'))):
        if (highlight.get(qn('w:val')) or '').lower() in WHITE_FILLS:
            highlight.getparent().remove(highlight)
            removed += 1
    return removed


def _is_list_item(text: str) -> bool:
    text = text.lstrip()
    return bool(
        ListsFormatRule.DOT_ITEM_REGEX.match(text)
        or ListsFormatRule.BRACKET_ITEM_REGEX.match(text)
        or ListsFormatRule.DASH_ITEM_REGEX.match(text)
    )


def _is_body_paragraph(paragraph, paragraphs, idx: int) -> bool:
    """
    Определяет, является ли абзац основным текстом (который можно безопасно править).
    Титульный лист и оглавление отсеиваются до вызова (в цикле create_fixed_docx).
    """
    text = paragraph.text.strip()
    if not text:
        return False

    # Заголовки, подписи, надписи таблиц, пункты списков правятся отдельно
    if MAJOR_HEADER_REGEX.match(text.upper()) or SUBHEADING_REGEX.match(text):
        return False
    if CAPTION_REGEX.match(text) or TABLE_LABEL_REGEX.match(text):
        return False
    if _is_list_item(text):
        return False

    # По центру/по правому краю — подписи и служебные элементы
    if paragraph.alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
        return False

    # Формулы и рисунки не трогаем
    xml = paragraph._element.xml
    if 'oMath' in xml or '<w:drawing' in xml or '<w:pict' in xml:
        return False

    return True


def _fix_body_paragraph(paragraph) -> None:
    fmt = paragraph.paragraph_format
    fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt.line_spacing = LINE_SPACING
    fmt.first_line_indent = Cm(FIRST_LINE_INDENT_CM)
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.widow_control = True

    for run in paragraph.runs:
        run.font.name = FONT_NAME
        run.font.size = Pt(FONT_SIZE)
        # Кегль и гарнитура для сложных сценариев (кириллица, формулы в тексте)
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = rPr.makeelement(qn('w:rFonts'), {})
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:ascii'), FONT_NAME)
        rFonts.set(qn('w:hAnsi'), FONT_NAME)
        rFonts.set(qn('w:cs'), FONT_NAME)
        color = run.font.color
        if color and color.rgb is not None and color.rgb != RGBColor(0, 0, 0):
            run.font.color.rgb = RGBColor(0, 0, 0)


def _is_major_header(paragraph) -> bool:
    text = paragraph.text.strip().upper()
    return bool(MAJOR_HEADER_REGEX.match(text))


def _is_subheading(paragraph) -> bool:
    return bool(SUBHEADING_REGEX.match(paragraph.text.strip()))


def _fix_major_header(paragraph, paragraphs, idx: int) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.line_spacing = LINE_SPACING
    for run in paragraph.runs:
        run.font.bold = True
    # Крупные разделы начинаются с новой страницы
    if not _SECTION_RULE._has_page_break_before(paragraphs, idx):
        paragraph.paragraph_format.page_break_before = True


def _fix_subheading(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.first_line_indent = Cm(0)
    paragraph.paragraph_format.line_spacing = LINE_SPACING
    for run in paragraph.runs:
        run.font.bold = True
    # Параграф не должен принудительно начинаться с новой страницы
    if paragraph.paragraph_format.page_break_before:
        paragraph.paragraph_format.page_break_before = False


def create_fixed_docx(src_path: str, dst_path: str) -> Dict[str, int]:
    """
    Создаёт исправленную копию документа (оригинал не изменяется).
    Возвращает статистику применённых исправлений.
    """
    doc = Document(src_path)
    stats: Dict[str, int] = {
        "поля страницы": 0,
        "основной текст": 0,
        "заголовки": 0,
        "убрана белая заливка": 0,
        "заменено длинное тире": 0,
    }

    # 1. Поля страницы
    for section in doc.sections:
        section.top_margin = Cm(MARGINS["top"])
        section.bottom_margin = Cm(MARGINS["bottom"])
        section.left_margin = Cm(MARGINS["left"])
        section.right_margin = Cm(MARGINS["right"])
        stats["поля страницы"] += 1

    paragraphs = classify_paragraphs(doc)

    # Титульный лист заканчивается на первом крупном заголовке (СОДЕРЖАНИЕ,
    # ВВЕДЕНИЕ, ГЛАВА...) или нумерованном разделе. Эвристика classify_paragraphs
    # помечает весь документ титульным листом, если СОДЕРЖАНИЯ нет вообще, —
    # для правки определяем конец титульного листа самостоятельно.
    title_end_idx = None
    for idx, paragraph in enumerate(paragraphs):
        text = paragraph.text.strip()
        if not text or is_toc_line(text):
            continue
        if MAJOR_HEADER_REGEX.match(text.upper()) or NUMBERED_SECTION_REGEX.match(text):
            title_end_idx = idx
            break

    # Приложения оформляются иначе (12 пт, без нумерации) — основной текст после
    # заголовка ПРИЛОЖЕНИЕ не правится
    appendix_start_idx = None
    for idx, paragraph in enumerate(paragraphs):
        if APPENDIX_HEADER_REGEX.match(paragraph.text.strip().upper()):
            appendix_start_idx = idx
            break

    for idx, paragraph in enumerate(paragraphs):
        in_appendix = appendix_start_idx is not None and idx >= appendix_start_idx
        if title_end_idx is not None:
            in_title_page = idx < title_end_idx
        else:
            in_title_page = getattr(paragraph, 'is_title_page', False)
        in_toc = getattr(paragraph, 'is_toc', False)

        # 2. Основной текст и 3. заголовки
        if not in_appendix:
            if in_title_page or in_toc:
                pass
            elif _is_body_paragraph(paragraph, paragraphs, idx):
                _fix_body_paragraph(paragraph)
                stats["основной текст"] += 1
            elif _is_major_header(paragraph):
                _fix_major_header(paragraph, paragraphs, idx)
                stats["заголовки"] += 1
            elif _is_subheading(paragraph):
                _fix_subheading(paragraph)
                stats["заголовки"] += 1

        # 4. Белая заливка — убирается по всему документу, включая титульный лист:
        # белой подложке не место на белой бумаге нигде
        stats["убрана белая заливка"] += _remove_white_shading(paragraph._element)

        # 5. Длинное тире — по всему документу вне титульного листа
        if not in_title_page:
            for run in paragraph.runs:
                if '—' in run.text:
                    run.text = run.text.replace('—', '–')
                    stats["заменено длинное тире"] += 1

    # Белая заливка в ячейках таблиц (их абзацы в doc.paragraphs не попадают)
    for table in doc.tables:
        stats["убрана белая заливка"] += _remove_white_shading(table._tbl)

    doc.save(dst_path)
    return stats
