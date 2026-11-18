"""
Профильные правила ГОСТ 19.201-78 (техническое задание на программное изделие).

Подключаются только в профиле `gost-19-201` (см. profiles.py); стандартный
профиль валидатора их не использует.
"""

from gost.gost_rules import (
    AnnotationRule,
    Gost19201PageNumberingRule,
    Gost19201StructureRule,
    TechnicalSpecificationRequisitesRule,
    TechnicalSpecificationTablesRule,
)

__all__ = [
    'AnnotationRule',
    'Gost19201PageNumberingRule',
    'Gost19201StructureRule',
    'TechnicalSpecificationRequisitesRule',
    'TechnicalSpecificationTablesRule',
]
