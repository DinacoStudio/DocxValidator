from argparse import ArgumentParser
from pathlib import Path
import sys
from typing import List, Optional, Sequence

from colorama import Fore
from docx import Document
from tqdm import tqdm

from colors import print_color
from docx_fixed import create_fixed_docx
from docx_notes import make_notes_docx
from docx_rules import BaseRule, Violation
import profiles


class DocxValidator:
    def __init__(self, rules: List[BaseRule]) -> None:
        self.rules = rules

    def validate(self, file_path: str) -> List[Violation]:
        doc = Document(file_path)
        violations = []
        for rule in tqdm(self.rules, desc="Проверка документа на соответствия правилам..."):
            violations.extend(rule.check(doc))
        return violations


def default_rules() -> List[BaseRule]:
    """Обратная совместимость: стандартный профиль правил."""
    return profiles.get_rules("standard")


def build_parser() -> ArgumentParser:
    parser = ArgumentParser(
        description="Проверяет Word документ на правила оформления документации"
    )
    parser.add_argument("input_file", help="Путь до Word документа")
    parser.add_argument(
        "--profile",
        choices=profiles.PROFILE_NAMES,
        default=profiles.PROFILE_STANDARD,
        help="Профиль проверки (по умолчанию: standard)",
    )
    parser.add_argument(
        "--make-notes",
        action="store_true",
        help="Создать рядом копию с замечаниями Word (оригинал не изменяется)",
    )
    parser.add_argument(
        "--create-fixed",
        action="store_true",
        help="Создать рядом исправленную копию (оригинал не изменяется)",
    )
    return parser


def _print_report(report: List[Violation]) -> None:
    print_color(f"Найдено нарушений: {len(report)}\n", Fore.RED)
    for item in report:
        color = Fore.RED if item.severity.value == "ERR" else Fore.YELLOW
        print_color(
            f"[{item.severity.value}] {item.rule_name} | {item.location}", color
        )
        print(f"    -> {item.message}")


def run(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    docx_path = Path(args.input_file)

    print("".rjust(70, "-"))
    print()
    print("DocxValidator".rjust(40, " "))
    print()
    print("".rjust(70, "-"))

    if not docx_path.is_file():
        print_color(f"Файл не найден: {docx_path.absolute()}", Fore.RED)
        return 2

    validator = DocxValidator(profiles.get_rules(args.profile))
    try:
        report = validator.validate(str(docx_path.absolute()))
    except Exception as error:
        print_color(
            f"Не удалось открыть или разобрать документ '{docx_path}': "
            f"{type(error).__name__}: {error}",
            Fore.RED,
        )
        print_color(
            "Убедитесь, что это корректный файл Word (.docx), и он не поврежден.",
            Fore.RED,
        )
        return 2

    if report:
        _print_report(report)
    else:
        print_color("Все в норме!", Fore.GREEN)

    if args.make_notes:
        if report:
            notes_path = docx_path.with_name(docx_path.stem + "_замечания.docx")
            placed, unresolved = make_notes_docx(
                str(docx_path.absolute()), report, str(notes_path)
            )
            suffix = (
                f", не привязано к абзацам: {unresolved}" if unresolved else ""
            )
            print_color(
                f"\nСоздана копия с замечаниями: {notes_path} "
                f"(комментариев: {placed}{suffix})",
                Fore.GREEN,
            )
        else:
            print_color(
                "Замечаний нет — копия с комментариями не требуется.", Fore.CYAN
            )

    if args.create_fixed:
        fixed_path = docx_path.with_name(docx_path.stem + "_исправленный.docx")
        stats = create_fixed_docx(
            str(docx_path.absolute()), str(fixed_path)
        )
        print_color(f"\nСоздана исправленная версия: {fixed_path}", Fore.GREEN)
        for name, count in stats.items():
            print(f"    исправлено «{name}»: {count}")
        try:
            fixed_report = validator.validate(str(fixed_path))
        except Exception as error:
            print_color(
                f"Не удалось повторно проверить '{fixed_path}': "
                f"{type(error).__name__}: {error}",
                Fore.RED,
            )
            return 2
        print_color(
            f"Нарушений было: {len(report)}, в исправленной версии осталось: "
            f"{len(fixed_report)}",
            Fore.CYAN,
        )

    if report:
        print_color(
            "Внимание! Программа может ошибаться в поиске ошибок в формате "
            "документа. Ошибочно найденные нарушения можно игнорировать.",
            Fore.CYAN,
        )
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(run())
