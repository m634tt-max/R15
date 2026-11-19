#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_tags.py — конвертация выгрузки приборов в формат образца.

Вход  (широкий формат):  один прибор = одна строка, свойства в колонках.
Выход (длинный формат):   одно свойство тега = одна строка, как в образце.

Запуск:
    python convert_tags.py --input source_data.xlsx --output result.xlsx
Опционально:
    --sample sample.xlsx   — взять порядок свойств и шапку колонок из образца
"""

from __future__ import annotations

import argparse
import sys
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ════════ Константы ════════

DEFAULT_INPUT: str = "source_data.xlsx"
DEFAULT_OUTPUT: str = "result.xlsx"
DEFAULT_SAMPLE: str = "sample.xlsx"

# Порядок свойств по умолчанию (как в образце); точные имена берутся
# из файла-образца, если он передан через --sample
PROPERTY_ORDER: List[str] = [
    "output signal range",
    "type of control system",
    "designed by company name",
    "po code",
    "tag class name",
    "tag description eng",
    "tag number",
    "tag status",
    "wbs code",
    "communication protocol",
    "gas analysed",
    "explosion protection signal type",
    "set point high",
    "set point low",
    "parent tag number",
    "building code",
    "commissioning system code",
    "design lifetime",
]

# Свойства-уставки: ед. изм. берётся из колонки «ед. изм 2» исходника
UOM_COLUMN: str = "ед. изм 2"
UOM_FOR_LIFETIME: str = "year"        # ед. изм. для design lifetime
NA_VALUE: str = "N/A"                 # подстановка для commissioning system code

OUTPUT_COLUMNS: List[str] = [
    "Action", "Wbs Code", "Tag Number", "Property Name",
    "Property Value", "UoM", "Comment",
]

# ──── Константы оформления Excel (hex БЕЗ "#") ────
XL_HEADER_FILL: str = "D9E1F2"   # заливка шапки
XL_EVEN_FILL: str = "F2F7FB"     # чётные строки
XL_ODD_FILL: str = "FFFFFF"      # нечётные строки
XL_HEADER_FONT: str = "1F4E79"   # цвет шрифта шапки
MAX_COL_WIDTH: int = 50          # макс. автоширина колонки


# ════════ Модель строки результата ════════

@dataclass
class TagRow:
    """Одна строка выходного файла (одно свойство одного тега)."""
    __slots__ = ("action", "wbs", "tag", "prop", "value", "uom", "comment")

    action: str
    wbs: str
    tag: str
    prop: str
    value: str
    uom: str
    comment: str


# ════════ Чтение и подготовка данных ════════

def norm(s: object) -> str:
    """Нормализация имени колонки: строка + обрезка пробелов."""
    return str(s).strip() if pd.notna(s) else ""


def read_source(path: Path) -> Tuple[bool, Optional[pd.DataFrame], str]:
    """Чтение исходной выгрузки. Возвращает (success, dataframe, message)."""
    try:
        df: pd.DataFrame = pd.read_excel(path, dtype=str)
    except FileNotFoundError:
        return False, None, f"Файл не найден: {path}"
    except Exception:
        return False, None, f"Не удалось прочитать файл:\n{traceback.format_exc()}"
    if df.empty:
        return False, None, "В исходном файле нет данных"
    return True, df, ""


def find_column(df: pd.DataFrame, target: str) -> Optional[str]:
    """Поиск колонки по нормализованному имени. None — если не найдена."""
    for col in df.columns:
        if norm(col) == target:
            return str(col)
    return None


def prepare_data(df: pd.DataFrame) -> Tuple[bool, Optional[pd.DataFrame], str]:
    """
    Подготовка данных: forward fill пустых ячеек общих свойств
    и нормализация имён колонок.
    """
    try:
        df = df.dropna(how="all").ffill()
    except Exception:
        return False, None, f"Ошибка обработки данных:\n{traceback.format_exc()}"
    return True, df, ""


def extract_order_from_sample(path: Path) -> Tuple[List[str], List[str]]:
    """
    Чтение порядка свойств и шапки колонок из файла-образца.
    При ошибке возвращаются значения по умолчанию.
    """
    try:
        sample: pd.DataFrame = pd.read_excel(path, dtype=str)
        props = [norm(x) for x in sample["Property Name"].tolist() if norm(x)]
        seen: List[str] = []
        for p in props:
            if p not in seen:
                seen.append(p)
        if seen:
            return seen, OUTPUT_COLUMNS
    except Exception:
        pass
    return PROPERTY_ORDER, OUTPUT_COLUMNS


def filter_properties(properties: List[str],
                      df: pd.DataFrame) -> List[str]:
    """
    Оставляем только те свойства, для которых есть колонки в исходнике.
    Порядок сохраняется как в образце.
    """
    source_cols = {norm(c) for c in df.columns}
    return [p for p in properties if p in source_cols]


def build_rows(df: pd.DataFrame,
               properties: List[str]) -> Tuple[bool, Optional[List[TagRow]], str]:
    """
    Формирование строк длинного формата: 18 свойств на каждый прибор.
    """
    uom_col = find_column(df, UOM_COLUMN)
    rows: List[TagRow] = []
    try:
        for _, device in df.iterrows():
            # карта: нормализованное имя свойства -> значение ячейки
            values: Dict[str, str] = {
                norm(col): ("" if pd.isna(device[col]) else str(device[col]).strip())
                for col in df.columns
            }
            tag: str = values.get("tag number", "")
            wbs: str = values.get("wbs code", "")
            if not tag:
                continue  # строка без тега — пропуск
            for prop in properties:
                if prop == "commissioning system code" and not values.get(prop):
                    value = NA_VALUE
                else:
                    value = values.get(prop, "")
                if prop in ("set point high", "set point low"):
                    uom = values.get(UOM_COLUMN, "") if uom_col else ""
                elif prop == "design lifetime":
                    uom = UOM_FOR_LIFETIME
                else:
                    uom = ""
                rows.append(TagRow("", wbs, tag, prop, value, uom, ""))
    except Exception:
        return False, None, f"Ошибка формирования строк:\n{traceback.format_exc()}"
    return True, rows, ""


# ════════ Запись результата ════════

def style_sheet(ws: object, headers: List[str]) -> None:
    """Оформление листа: шапка, чередование строк, границы, ширины, фриз, фильтр."""
    thin: object = Side(style="thin", color="B0B0B0")
    border: object = Border(left=thin, right=thin, top=thin, bottom=thin)
    # Шапка
    for c, name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c)
        cell.value = name
        cell.font = Font(bold=True, color=XL_HEADER_FONT)
        cell.fill = PatternFill("solid", fgColor=XL_HEADER_FILL)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    # Данные: чередование строк
    for r in range(2, ws.max_row + 1):
        fill = XL_EVEN_FILL if r % 2 == 0 else XL_ODD_FILL
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    # Ширина колонок (авто, с ограничением)
    for c, name in enumerate(headers, start=1):
        width = len(str(name)) + 2
        for r in range(2, min(ws.max_row, 200) + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None:
                width = max(width, min(len(str(v)) + 2, MAX_COL_WIDTH))
        ws.column_dimensions[get_column_letter(c)].width = min(width, MAX_COL_WIDTH)
    # Шапка + высота строк + фриз + фильтр
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def write_excel(rows: List[TagRow],
                out_path: Path,
                headers: List[str]) -> Tuple[bool, str]:
    """Запись результата в Excel с оформлением."""
    try:
        wb: Workbook = Workbook()
        ws = wb.active
        ws.title = "Data 1"
        ws.append(headers)
        for row in rows:
            ws.append([row.action, row.wbs, row.tag, row.prop,
                       row.value, row.uom, row.comment])
        style_sheet(ws, headers)
        wb.save(out_path)
        return True, f"Создан файл: {out_path} (строк данных: {len(rows)})"
    except Exception:
        return False, f"Ошибка записи файла:\n{traceback.format_exc()}"


# ════════ Запуск ════════

def parse_args() -> argparse.Namespace:
    """Разбор аргументов командной строки."""
    parser = argparse.ArgumentParser(
        description="Конвертация выгрузки приборов в формат образца")
    parser.add_argument("--input", default=DEFAULT_INPUT,
                        help=f"входной файл (по умолчанию: {DEFAULT_INPUT})")
    parser.add_argument("--output", default=DEFAULT_OUTPUT,
                        help=f"выходной файл (по умолчанию: {DEFAULT_OUTPUT})")
    parser.add_argument("--sample", default=DEFAULT_SAMPLE,
                        help="файл-образец для порядка свойств "
                             f"(по умолчанию: {DEFAULT_SAMPLE})")
    return parser.parse_args()


def main() -> int:
    """Точка входа. Возвращает код возврата: 0 — успех, 1 — ошибка."""
    args = parse_args()

    in_path: Path = Path(args.input)
    out_path: Path = Path(args.output)
    sample_path: Path = Path(args.sample)

    ok, df, msg = read_source(in_path)
    if not ok:
        print(f"[ОШИБКА] {msg}")
        return 1
    assert df is not None

    ok, df, msg = prepare_data(df)
    if not ok:
        print(f"[ОШИБКА] {msg}")
        return 1
    assert df is not None

    if not find_column(df, "tag number"):
        print("[ОШИБКА] В исходном файле нет колонки 'tag number'")
        return 1

    if sample_path.exists():
        properties, headers = extract_order_from_sample(sample_path)
        print(f"[И] Порядок свойств и шапка взяты из образца: {sample_path.name}")
    else:
        properties, headers = PROPERTY_ORDER, OUTPUT_COLUMNS
        print(f"[ПРЕДУПР.] Образец {sample_path.name} не найден — "
              f"используется порядок свойств по умолчанию")

    properties = filter_properties(properties, df)
    print(f"[И] Свойств к выгрузке: {len(properties)}")

    ok, rows, msg = build_rows(df, properties)
    if not ok:
        print(f"[ОШИБКА] {msg}")
        return 1
    assert rows is not None

    ok, msg = write_excel(rows, out_path, headers)
    if not ok:
        print(f"[ОШИБКА] {msg}")
        return 1
    print(f"[ОК] {msg}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
