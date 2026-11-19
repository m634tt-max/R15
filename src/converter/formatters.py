# -*- coding: utf-8 -*-
"""Форматирование значений свойств, дат и формул оборудования."""

import re
from datetime import datetime
from typing import Any, Sequence

import pandas as pd

FORMULA_RE = re.compile(r'^="?E-"&F(\d+)"?$')
NA_VALUE: str = "N/A"


def norm_str(value: Any) -> str:
    """Нормализует имя колонки или строковое значение (None/NaN -> '')."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def resolve_equipment_number(value: Any, tag: str) -> str:
    """Вычисляет номер оборудования из формулы вида ="E-"&F... или возвращает значение."""
    if isinstance(value, str):
        if FORMULA_RE.match(value.strip()):
            return f"E-{tag}"
    return fmt_value("equipment number", value)


def fmt_value(prop: str, value: Any, date_properties: Sequence[str] = ()) -> str:
    """Форматирует значение свойства для вывода в длинную таблицу."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    if prop in date_properties:
        if isinstance(value, datetime):
            return value.strftime("%y%m%d")
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        text = str(value).strip()
        # Попытка распарсить строковую дату ISO/Excel, если pandas прочитал её как текст
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(text, fmt).strftime("%y%m%d")
            except ValueError:
                continue
        return text

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()
