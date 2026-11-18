"""Профильные проверки технического задания по ГОСТ 19.201-78."""
from datetime import date
import re
from typing import List, Optional, Sequence, Tuple

from docx import Document
from docx.oxml.ns import qn

from docx_rules import BaseRule, Severity, Violation, has_page_field


_NUMBER_PREFIX = re.compile(r"^\s*\d+(?:\.\d+)*[.)]?\s*")
_PAGE_BREAK = re.compile(r'<w:br\b[^>]*w:type="page"')


def _clean_heading(text: str) -> str:
    text = _NUMBER_PREFIX.sub("", text.strip()).upper().replace("Ё", "Е")
    return re.sub(r"\s+", " ", text).rstrip(".:")


def _violation(rule: str, location: str, message: str,
               severity: Severity = Severity.ERROR) -> Violation:
    return Violation(rule, severity, location, message)


def _has_page_break_before(paragraphs: Sequence, index: int) -> bool:
    paragraph = paragraphs[index]
    if paragraph.paragraph_format.page_break_before:
        return True
    p_pr = paragraph._p.pPr
    if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
        return True
    if index == 0:
        return True
    previous = paragraphs[index - 1]
    if _PAGE_BREAK.search(previous._p.xml):
        return True
    previous_pr = previous._p.pPr
    return previous_pr is not None and previous_pr.find(qn("w:sectPr")) is not None


class AnnotationRule(BaseRule):
    """Проверяет наличие, положение и минимальный состав аннотации ТЗ."""

    REQUIRED_CONTENT = (
        ("назначение", (r"\bНАЗНАЧЕН", r"\bЦЕЛ[ЬИ]")),
        ("область применения", (r"ОБЛАСТ[ЬИ]\s+ПРИМЕНЕН", r"ПРИМЕНЯЕТСЯ")),
        ("пользователи", (r"\bПОЛЬЗОВАТЕЛ", r"\bОПЕРАТОР", r"\bПЕРСОНАЛ")),
        ("основные функции", (r"\bФУНКЦ", r"ВОЗМОЖНОСТ")),
    )
    NEXT_SECTION = re.compile(
        r"^(?:\d+(?:\.\d+)*[.)]?\s*)?(ВВЕДЕНИЕ|СОДЕРЖАНИЕ|ОГЛАВЛЕНИЕ)\s*$", re.I)

    def check(self, doc: Document) -> List[Violation]:
        rule_name = self.__class__.__name__
        paragraphs = list(doc.paragraphs)
        annotation_index = next(
            (index for index, paragraph in enumerate(paragraphs)
             if _clean_heading(paragraph.text) == "АННОТАЦИЯ"), None)
        if annotation_index is None:
            return [_violation(rule_name, "Документ",
                               "Не найден обязательный раздел «Аннотация».")]

        violations: List[Violation] = []
        location = f"Аннотация (абзац {annotation_index + 1})"
        if annotation_index == 0:
            violations.append(_violation(
                rule_name, location,
                "Аннотация должна располагаться после титульного листа."))
        if not _has_page_break_before(paragraphs, annotation_index):
            violations.append(_violation(
                rule_name, location,
                "Аннотация должна начинаться на второй логической странице после разрыва страницы/раздела."))

        content = []
        for paragraph in paragraphs[annotation_index + 1:]:
            text = paragraph.text.strip()
            if text and self.NEXT_SECTION.match(text):
                break
            content.append(text)
        annotation_text = " ".join(content).upper().replace("Ё", "Е")
        if not annotation_text.strip():
            violations.append(_violation(rule_name, location, "Аннотация не должна быть пустой."))
            return violations

        for label, patterns in self.REQUIRED_CONTENT:
            if not any(re.search(pattern, annotation_text) for pattern in patterns):
                violations.append(_violation(
                    rule_name, location,
                    f"В аннотации не раскрыт обязательный элемент: {label}.",
                    Severity.WARNING))
        return violations


class Gost19201StructureRule(BaseRule):
    """Проверяет обязательные разделы ТЗ и их порядок по ГОСТ 19.201-78.

    ГОСТ 19.201-78 п. 1.4 допускает «уточнять содержание разделов, вводить
    новые разделы или объединять отдельные из них», поэтому помимо точных
    названий разделов принимаются равнозначные варианты из практики
    оформления ТЗ (в том числе из выданного примера).
    """

    REQUIRED_SECTIONS: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
        ("Введение", (
            "ВВЕДЕНИЕ",
            "ОБЩИЕ СВЕДЕНИЯ",  # объединяет введение и основания (допустимо по п. 1.4)
        )),
        ("Основания для разработки", (
            "ОСНОВАНИЯ ДЛЯ РАЗРАБОТКИ",
            "ОСНОВАНИЕ ДЛЯ РАЗРАБОТКИ",
            "ОСНОВАНИЕ ДЛЯ ПРОВЕДЕНИЯ РАБОТ",
            "ОБЩИЕ СВЕДЕНИЯ",
        )),
        ("Назначение разработки", (
            "НАЗНАЧЕНИЕ РАЗРАБОТКИ",
            "НАЗНАЧЕНИЕ И ЦЕЛИ СОЗДАНИЯ СИСТЕМЫ",
        )),
        ("Требования к программе или программному изделию", (
            "ТРЕБОВАНИЯ К ПРОГРАММЕ ИЛИ ПРОГРАММНОМУ ИЗДЕЛИЮ",
            "ТРЕБОВАНИЯ К ПРОГРАММЕ",
            "ТРЕБОВАНИЯ К ПРОГРАММНОМУ ИЗДЕЛИЮ",
            "ТРЕБОВАНИЯ К СИСТЕМЕ",
        )),
        ("Требования к программной документации", (
            "ТРЕБОВАНИЯ К ПРОГРАММНОЙ ДОКУМЕНТАЦИИ",
            "ТРЕБОВАНИЯ К ДОКУМЕНТИРОВАНИЮ",
        )),
        ("Технико-экономические показатели", ("ТЕХНИКО-ЭКОНОМИЧЕСКИЕ ПОКАЗАТЕЛИ",)),
        ("Стадии и этапы разработки", (
            "СТАДИИ И ЭТАПЫ РАЗРАБОТКИ",
            "СОСТАВ И СОДЕРЖАНИЕ РАБОТ ПО СОЗДАНИЮ СИСТЕМЫ",
            "СОСТАВ И СОДЕРЖАНИЕ РАБОТ",
        )),
        ("Порядок контроля и приёмки", (
            "ПОРЯДОК КОНТРОЛЯ И ПРИЕМКИ",
            "ПОРЯДОК КОНТРОЛЯ И ПРИЕМКИ СИСТЕМЫ",
        )),
    )

    # «Общие сведения» объединяет два раздела: считаем его одним вхождением
    # для обоих, поэтому дубликатное предупреждение для пары не выводится.
    MERGED_ALIASES = {"ОБЩИЕ СВЕДЕНИЯ"}

    def check(self, doc: Document) -> List[Violation]:
        rule_name = self.__class__.__name__
        headings = [(_clean_heading(p.text), index + 1)
                    for index, p in enumerate(doc.paragraphs) if p.text.strip()]
        found: List[Tuple[str, Optional[int]]] = []
        violations: List[Violation] = []
        reported_duplicates = set()

        for label, variants in self.REQUIRED_SECTIONS:
            occurrences = [number for heading, number in headings if heading in variants]
            if not occurrences:
                violations.append(_violation(
                    rule_name, "Структура документа",
                    f"Отсутствует обязательный раздел «{label}» "
                    f"(допустимы объединённые формулировки по п. 1.4 ГОСТ 19.201-78)."))
                found.append((label, None))
            else:
                found.append((label, occurrences[0]))
                distinct = set()
                for heading, number in headings:
                    if heading in variants and heading not in self.MERGED_ALIASES:
                        distinct.add(heading)
                if len(distinct) > 1 and label not in reported_duplicates:
                    reported_duplicates.add(label)
                    violations.append(_violation(
                        rule_name, f"Абзацы {', '.join(map(str, occurrences))}",
                        f"Заголовок раздела «{label}» встречается более одного раза.",
                        Severity.WARNING))

        present = [(label, position) for label, position in found if position is not None]
        for (previous_label, previous), (label, position) in zip(present, present[1:]):
            if position < previous:
                violations.append(_violation(
                    rule_name, f"Абзац {position}",
                    f"Нарушен порядок разделов: «{label}» расположен перед «{previous_label}»."))

        appendix_positions = [number for heading, number in headings
                              if re.match(r"^ПРИЛОЖЕНИЕ(?:\s+[А-ЯA-Z0-9]+)?$", heading)]
        last_required = max((position for _, position in present), default=0)
        if appendix_positions and min(appendix_positions) < last_required:
            violations.append(_violation(
                rule_name, f"Абзац {min(appendix_positions)}",
                "Приложения должны располагаться после всех обязательных разделов технического задания."))
        return violations


class TechnicalSpecificationRequisitesRule(BaseRule):
    """Проверяет идентифицирующие реквизиты ТЗ и согласованность сроков."""

    REQUIRED = (
        ("название программы", (
            r"НАИМЕНОВАНИЕ\s+(?:ПРОГРАММЫ|ПРОГРАММНОГО\s+ИЗДЕЛИЯ)",
            r"НАИМЕНОВАНИЕ\s+(?:ПОЛНОЕ\s+)?СИСТЕМЫ",
            r"ТЕХНИЧЕСКОЕ\s+ЗАДАНИЕ\s+НА\s+(?:СОЗДАНИЕ|РАЗРАБОТКУ)",
        )),
        ("разработчик", (r"\bРАЗРАБОТЧИК\b", r"ОРГАНИЗАЦИЯ[- ]РАЗРАБОТЧИК")),
        ("заказчик", (r"\bЗАКАЗЧИК\b", r"ОРГАНИЗАЦИЯ[- ]ЗАКАЗЧИК")),
        ("основание разработки", (
            r"ОСНОВАНИ[ЕЯ]\s+ДЛЯ\s+РАЗРАБОТКИ",
            r"ОСНОВАНИ[ЕЯ]\s+ДЛЯ\s+ПРОВЕДЕНИЯ\s+РАБОТ",
        )),
        ("номер документа или договора", (
            r"(?:ДОГОВОР|КОНТРАКТ|ДОКУМЕНТ|ТЗ)\s*(?:ОТ\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*)?[№N]\s*[\wА-ЯЁ/-]+",
            r"[№N]\s*[\wА-ЯЁ/-]+\s+ОТ\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}",
        )),
    )
    DATE_RE = re.compile(r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{4})(?!\d)")

    @classmethod
    def _dates(cls, text: str) -> List[date]:
        result = []
        for day, month, year in cls.DATE_RE.findall(text):
            try:
                result.append(date(int(year), int(month), int(day)))
            except ValueError:
                pass
        return result

    def check(self, doc: Document) -> List[Violation]:
        rule_name = self.__class__.__name__
        text = "\n".join(p.text for p in doc.paragraphs)
        upper = text.upper().replace("Ё", "Е")
        violations: List[Violation] = []

        for label, patterns in self.REQUIRED:
            if not any(re.search(pattern, upper, re.MULTILINE) for pattern in patterns):
                violations.append(_violation(
                    rule_name, "Реквизиты технического задания",
                    f"Не найден обязательный реквизит: {label}."))

        dates = self._dates(text)
        if len(dates) < 2:
            violations.append(_violation(
                rule_name, "Реквизиты технического задания",
                "Должны быть указаны как минимум две корректные календарные даты.",
                Severity.WARNING))

        start_match = re.search(
            r"(?:НАЧАЛ[ОА]|СРОК\s+НАЧАЛА)[^\n\d]{0,40}(\d{1,2}[./-]\d{1,2}[./-]\d{4})", upper)
        end_match = re.search(
            r"(?:ОКОНЧАНИ[ЕЯ]|СРОК\s+ОКОНЧАНИЯ)[^\n\d]{0,40}(\d{1,2}[./-]\d{1,2}[./-]\d{4})", upper)
        if start_match and end_match:
            start = self._dates(start_match.group(1))
            end = self._dates(end_match.group(1))
            if start and end and end[0] < start[0]:
                violations.append(_violation(
                    rule_name, "Сроки разработки",
                    "Дата окончания разработки не может предшествовать дате начала."))
        return violations


class TechnicalSpecificationTablesRule(BaseRule):
    """Требует две таблицы; оформление и ссылки проверяет TablesFormatRule."""

    def check(self, doc: Document) -> List[Violation]:
        count = len(doc.tables)
        if count >= 2:
            return []
        return [_violation(
            self.__class__.__name__,
            "Таблицы технического задания",
            f"Техническое задание должно содержать не менее двух оформленных таблиц; найдено: {count}.",
        )]


class Gost19201PageNumberingRule(BaseRule):
    """Проверяет профильную нумерацию: титул и аннотация без PAGE, основная часть с 3."""

    @staticmethod
    def _section_has_page(section) -> bool:
        containers = (section.header, section.footer, section.first_page_header,
                      section.first_page_footer, section.even_page_header,
                      section.even_page_footer)
        return any(has_page_field(container) for container in containers)

    def check(self, doc: Document) -> List[Violation]:
        rule_name = self.__class__.__name__
        violations: List[Violation] = []
        if len(doc.sections) < 3:
            return [_violation(
                rule_name, "Разделы документа",
                "Титульный лист, аннотация и основная часть должны быть разделены как минимум на три секции Word.")]

        for index in (0, 1):
            if self._section_has_page(doc.sections[index]):
                violations.append(_violation(
                    rule_name, f"Секция {index + 1}",
                    "Титульный лист и аннотация не должны содержать поле PAGE."))

        main = doc.sections[2]
        if not has_page_field(main.header):
            violations.append(_violation(
                rule_name, "Секция 3 (верхний колонтитул)",
                "Основная часть должна содержать поле PAGE в верхнем колонтитуле."))
        pg_num = main._sectPr.find(qn("w:pgNumType"))
        start = pg_num.get(qn("w:start")) if pg_num is not None else None
        if start != "3":
            violations.append(_violation(
                rule_name, "Секция 3",
                "Нумерация основной части должна явно начинаться с 3 (w:pgNumType w:start=\"3\")."))
        return violations
