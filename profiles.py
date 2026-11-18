"""Наборы правил валидатора для стандартного и ГОСТ-профилей."""
from typing import Callable, Dict, List

from docx_rules import (
    AppendixFormatRule,
    BaseRule,
    BreaksAndPaginationRule,
    DashesRule,
    FiguresRule,
    FormulasRule,
    IntroductionContentRule,
    ListsFormatRule,
    MainTextFormatRule,
    PageMarginsRule,
    PageNumberingRule,
    ParagraphConclusionRule,
    ReferencesListRule,
    SectionFormatRule,
    SectionPageFillRule,
    SourceReferencesRule,
    TableOfContentsRule,
    TablesFormatRule,
    TextGridRule,
    TitlePageRule,
    WhiteFillRule,
)
from gost import (
    AnnotationRule,
    Gost19201PageNumberingRule,
    Gost19201StructureRule,
    TechnicalSpecificationRequisitesRule,
    TechnicalSpecificationTablesRule,
)


PROFILE_STANDARD = "standard"
PROFILE_GOST_19_201 = "gost-19-201"
PROFILE_NAMES = (PROFILE_STANDARD, PROFILE_GOST_19_201)


def standard_rules() -> List[BaseRule]:
    """Возвращает новый набор общих правил при каждом вызове."""
    return [
        PageMarginsRule(),
        TextGridRule(),
        TableOfContentsRule(),
        MainTextFormatRule(),
        PageNumberingRule(),
        TitlePageRule(),
        SectionFormatRule(),
        BreaksAndPaginationRule(),
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
        WhiteFillRule(),
    ]


def gost_19_201_rules() -> List[BaseRule]:
    """Стандартный набор с профильными правилами ГОСТ 19.201-78.

    Правила курсовой работы, не предусмотренные структурой ТЗ
    (введение с объектом/предметом, список источников, ссылки [N],
    выводы «Таким образом» в конце параграфов), в этом профиле не
    запускаются: по ГОСТ 19.201-78 они обязательными разделами не являются.
    """
    coursework_only = (
        IntroductionContentRule,
        ReferencesListRule,
        SourceReferencesRule,
        ParagraphConclusionRule,
    )
    return [rule for rule in standard_rules()
            if type(rule) not in coursework_only] + [
        AnnotationRule(),
        Gost19201StructureRule(),
        TechnicalSpecificationRequisitesRule(),
        TechnicalSpecificationTablesRule(),
        Gost19201PageNumberingRule(),
    ]


_PROFILE_FACTORIES: Dict[str, Callable[[], List[BaseRule]]] = {
    PROFILE_STANDARD: standard_rules,
    PROFILE_GOST_19_201: gost_19_201_rules,
}


def get_rules(profile: str = PROFILE_STANDARD) -> List[BaseRule]:
    """Возвращает правила профиля или сообщает о неизвестном имени."""
    try:
        factory = _PROFILE_FACTORIES[profile]
    except KeyError as error:
        choices = ", ".join(PROFILE_NAMES)
        raise ValueError(f"Неизвестный профиль '{profile}'. Доступны: {choices}.") from error
    return factory()


def rules_for_profile(profile: str = PROFILE_STANDARD) -> List[BaseRule]:
    """Обратимо читаемый псевдоним для интеграции CLI."""
    return get_rules(profile)
