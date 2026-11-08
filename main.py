from docx_rules import *
from docx_notes import make_notes_docx
from argparse import ArgumentParser, Namespace
from colors import print_color
from pathlib import Path
from colorama import Fore

class DocxValidator:
    def __init__(self, rules: List[BaseRule]) -> None:
        self.rules = rules
    def validate(self, file_path: str) -> List[Violation]:
        doc = Document(file_path)
        violations = []
        for rule in tqdm(self.rules, desc="Проверка документа на соответствия правилам..."):
            violations.extend(rule.check(doc))
        return violations

parser = ArgumentParser(description="Проверяет Word документ на правила, указанные на написание документации или диплома")
parser.add_argument("input_file", help="Путь до Word документа")
parser.add_argument("--make-notes", action="store_true",
                    help="Создать рядом копию документа с замечаниями в виде комментариев Word (оригинал не изменяется)")
parser.print_help()

if __name__ == "__main__":
    args = parser.parse_args()

    print("".rjust(70, "-"))
    print()
    print("DocxValidator".rjust(40, " "))
    print()
    print("".rjust(70, "-"))
    validator = DocxValidator(rules=[
        PageMarginsRule(),
        MainTextFormatRule(),
        PageNumberingRule(),
        TitlePageRule(),
        SectionFormatRule(),
        SectionPageFillRule(),
        ReferencesListRule(),
        IntroductionContentRule(),
        ParagraphConclusionRule(),
        SourceReferencesRule(),
        AppendixFormatRule(),
        FiguresRule(),
        TablesFormatRule(),
        FormulasRule(),
        ListsFormatRule(),
        DashesRule(),
        WhiteFillRule()
    ])
    docx_path = Path(args.input_file)
    report = validator.validate(str(docx_path.absolute()))
    if report == []:
        print_color("Все в норме!", Fore.GREEN)
        if args.make_notes:
            print_color("Замечаний нет — копия с комментариями не требуется.", Fore.CYAN)
        __import__("sys").exit(0)

    print_color(f"Найдено нарушений: {len(report)}\n", Fore.RED)
    for item in report:
        print_color(f"[{item.severity.value}] {item.rule_name} | {item.location}", Fore.RED if item.severity.value == "ERR" else Fore.YELLOW)
        print(f"    -> {item.message}")

    if args.make_notes:
        notes_path = docx_path.with_name(docx_path.stem + "_замечания.docx")
        placed, unresolved = make_notes_docx(str(docx_path.absolute()), report, str(notes_path))
        print_color(f"\nСоздана копия с замечаниями: {notes_path} (комментариев: {placed}" +
                    (f", не привязано к абзацам: {unresolved}" if unresolved else "") + ")", Fore.GREEN)

    print_color("Внимание! Программа может ошибаться в поиске ошибок в формате документа! Если вы считаете, что нарушения ошибочны, то можете спокойно игнорировать", Fore.CYAN)
    __import__("sys").exit(1)

    print_color(f"Найдено нарушений: {len(report)}\n", Fore.RED)
    for item in report:
        print_color(f"[{item.severity.value}] {item.rule_name} | {item.location}", Fore.RED if item.severity.value == "ERR" else Fore.YELLOW)
        print(f"    -> {item.message}")
    print_color("Внимание! Программа может ошибаться в поиске ошибок в формате документа! Если вы считаете, что нарушения ошибочны, то можете спокойно игнорировать", Fore.CYAN)
    __import__("sys").exit(1)
