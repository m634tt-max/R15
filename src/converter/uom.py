# -*- coding: utf-8 -*-
"""Поиск и извлечение единиц измерения (UoM) для свойств."""

from typing import Dict, List, Optional, Tuple

import pandas as pd

from .config import ConverterConfig
from .formatters import fmt_value, norm_str


def detect_adjacent_uom_columns(raw_headers: List[object]) -> Dict[str, str]:
    """
    Для режима 'adjacent' (equipment): находит для каждого свойства соседнюю справа
    пустую колонку без заголовка и возвращает карту {имя_свойства: имя_колонки_uom_в_df}.
    """
    uom_col_map: Dict[str, str] = {}
    for i, h in enumerate(raw_headers):
        name = norm_str(h)
        if not name or name.startswith("__uom_"):
            continue
        if i + 1 < len(raw_headers):
            nxt = norm_str(raw_headers[i + 1])
            if not nxt or nxt.startswith("__uom_"):
                uom_col_map[name] = f"__uom_{i + 1}"
    return uom_col_map


def resolve_uom(
    prop: str,
    row: pd.Series,
    config: ConverterConfig,
    adjacent_uom_map: Optional[Dict[str, str]] = None,
) -> str:
    """Возвращает единицу измерения (UoM) для конкретного свойства в строке прибора."""
    # 1. Фиксированная единица измерения (например, design lifetime -> year)
    if prop in config.fixed_uom:
        return config.fixed_uom[prop]

    # 2. Соседняя колонка справа без заголовка (режим equipment)
    if config.uom_mode == "adjacent" and adjacent_uom_map:
        uom_col = adjacent_uom_map.get(prop)
        if uom_col and uom_col in row.index:
            return fmt_value(prop, row[uom_col])
        return ""

    # 3. Общая колонка единиц измерения для целевых свойств (режим tags)
    if (
        config.uom_mode == "column"
        and config.uom_column
        and prop in config.uom_target_properties
        and config.uom_column in row.index
    ):
        return fmt_value(prop, row[config.uom_column])

    return ""
