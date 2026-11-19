# -*- coding: utf-8 -*-
"""
convert_equipment.py — конвертация данных оборудования в формат образца.

Читает лист «Лист2» файла исходных данных (широкий формат: один прибор =
одна строка, свойства в колонках, единицы измерения — в соседних колонках
без заголовка) и создаёт Excel в «длинном» формате образца (одно свойство
оборудования = одна строка).

Запуск:
    python convert_equipment.py --input source_data2.xlsx --output result2.xlsx
"""

# ══════════════════════════════════════════════════════════════════════
# Импорты (сгруппированы: стандартная библиотека / сторонние)
# ══════════════════════════════════════════════════════════════════════
import argparse
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ──── Константы ────
DEFAULT_INPUT: str = "source_data2.xlsx"
DEFAULT_SAMPLE: str = "sample2.xlsx"
DEFAULT_OUTPUT: str = "result2.xlsx"
SOURCE_SHEET: str = "Лист2"          # лист с данными оборудования
HEADER_ROWS: int = 2                  # строка 0 — русский заголовок, строка 1 — английский
MAX_COL_WIDTH: int = 50               # ограничение автоширины колонок
DATE_PROPERTIES: Tuple[str, ...] = (
    "manufacturing date", "warranty end date", "purchase date",
)
# Порядок и оформление выходной таблицы (как в образце)
OUT_COLUMNS: Tuple[str, ...] = (
    "Action", "Wbs Code", "Tag Number", "Property Name",
    "Property Value", "UoM", "Comment",
)
HDR_FILL: PatternFill = PatternFill("solid", fgColor="D9E2F3")
HDR_FONT: Font = Font(bold=True)
THIN: Side = Side(style="thin", color="B0B0B0")
BORDER: Border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BAND_FILL: PatternFill = PatternFill("solid", fgColor="F2F6FC")


# ══════════════════════════════════════════════════════════════════════
# Чтение исходных данных
# ══════════════════════════════════════════════════════════════════════
def read_source(path: Path) -> Tuple[
    Dict[str, Tuple[int, Optional[int]]], List[List[Any]], List[List[Any]]
]:
    """Читает лист исходных данных в двух режимах и возвращает карту колонок.

    Файлы Excel могут содержать как формулы (нужен data_only=True),
    так и обычные строки, которые в режиме data_only иногда приходят как None.
    Поэтому лист читается дважды: вычисленные значения + сырые для резерва.

    Возвращает:
        (карта «имя свойства → (индекс колонки значения, индекс колонки UoM)»,
         строки с вычисленными значениями, сырые строки для резерва)
    """
    if SOURCE_SHEET not in openpyxl.load_workbook(path, read_only=True).sheetnames:
        raise ValueError(f"Лист {SOURCE_SHEET!r} не найден в файле {path.name}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[SOURCE_SHEET]
    rows: List[List[Any]] = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    wb_raw = openpyxl.load_workbook(path, read_only=True)
    ws_raw = wb_raw[SOURCE_SHEET]
    rows_raw: List[List[Any]] = [list(r) for r in ws_raw.iter_rows(values_only=True)]
    wb_raw.close()
    if len(rows) <= HEADER_ROWS:
        raise ValueError("Лист данных пуст: нет строк приборов")

    eng_hdr = rows[1]  # английские заголовки свойств
    cols: Dict[str, Tuple[int, Optional[int]]] = {}
    for i, h in enumerate(eng_hdr):
        name = (h or "").strip() if isinstance(h, str) else (h or "")
        if not name:
            continue  # колонка без заголовка — это UoM предыдущего свойства
        uom_idx: Optional[int] = None
        if i + 1 < len(eng_hdr):
            nxt = eng_hdr[i + 1]
            # Пустая соседняя ячейка (None или пустая строка) — это колонка UoM
            if nxt is None or (isinstance(nxt, str) and not nxt.strip()):
                uom_idx = i + 1
        cols[name] = (i, uom_idx)
    if not cols:
        raise ValueError("Не найдено ни одного заголовка свойств в строке 2")
    return cols, rows[HEADER_ROWS:], rows_raw[HEADER_ROWS:]


def read_sample_order(path: Path) -> List[str]:
    """Возвращает порядок свойств из файла-образца (без повторов)."""
    df = pd.read_excel(path, sheet_name=0, keep_default_na=False)
    if "Property Name" not in df.columns:
        raise ValueError("В образце нет колонки 'Property Name'")
    order: List[str] = []
    for v in df["Property Name"]:
        name = str(v).strip()
        if name and name not in order:
            order.append(name)
    return order


# ══════════════════════════════════════════════════════════════════════
# Преобразование значений
# ══════════════════════════════════════════════════════════════════════
FORMULA_RE = re.compile(r'^="?E-"&F(\d+)"?$')


def resolve_equipment_number(value: Any, tag: str) -> str:
    """Возвращает номер оборудования: вычисляет формулу или берёт значение как есть."""
    if isinstance(value, str):
        m = FORMULA_RE.match(value.strip())
        if m:
            return f"E-{tag}"
    return fmt_value("equipment number", value)


def fmt_value(prop: str, value: Any) -> str:
    """Форматирует значение свойства для вывода в таблицу."""
    if value is None:
        return ""
    if prop in DATE_PROPERTIES:
        if isinstance(value, datetime):
            return value.strftime("%y%m%d")  # формат YYMMDD, как в образце
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def forward_fill(values: List[Any]) -> List[Any]:
    """Заполняет пустые значения последним непустым (общие свойства класса)."""
    filled: List[Any] = []
    last: Any = None
    for v in values:
        if v is None or (isinstance(v, str) and not v.strip()):
            filled.append(last)
        else:
            filled.append(v)
            last = v
    return filled


# ══════════════════════════════════════════════════════════════════════
# Сборка выходной таблицы
# ══════════════════════════════════════════════════════════════════════
def merge_cell(computed: Any, raw: Any) -> Any:
    """Выбирает значение ячейки: вычисленное, иначе сырая величина."""
    if computed is None or (isinstance(computed, str) and not computed.strip()):
        return raw
    return computed


def build_rows(
    cols: Dict[str, Tuple[int, Optional[int]]],
    data: List[List[Any]],
    order: List[str],
    data_raw: Optional[List[List[Any]]] = None,
) -> Tuple[List[Dict[str, str]], List[str]]:
    """Формирует строки «длинного» формата и список предупреждений."""
    warnings: List[str] = []
    # Пропускаем свойства, которых нет в исходнике
    used = [p for p in order if p in cols]
    missing = [p for p in order if p not in cols]
    for name in missing:
        warnings.append(f"Свойство {name!r} из образца не найдено в исходнике — пропущено")

    # Карта значений и UoM по каждому свойству (с forward fill)
    value_map: Dict[str, List[str]] = {}
    uom_map: Dict[str, List[str]] = {}
    for prop, (v_idx, u_idx) in cols.items():
        values = forward_fill([
            merge_cell(row[v_idx] if v_idx < len(row) else None,
                       (data_raw[i][v_idx] if data_raw and v_idx < len(data_raw[i]) else None))
            for i, row in enumerate(data)
        ])
        value_map[prop] = [fmt_value(prop, v) for v in values]
        if u_idx is not None:
            uom = forward_fill([
                merge_cell(row[u_idx] if u_idx < len(row) else None,
                           (data_raw[i][u_idx] if data_raw and u_idx < len(data_raw[i]) else None))
                for i, row in enumerate(data)
            ])
            uom_map[prop] = [fmt_value(prop, u) for u in uom]
        else:
            uom_map[prop] = [""] * len(data)

    out: List[Dict[str, str]] = []
    for i, row in enumerate(data):
        tag = value_map.get("tag number", [""] * len(data))[i]
        wbs = value_map.get("wbs code", [""] * len(data))[i]
        equip_no = resolve_equipment_number(
            value_map.get("equipment number", [None] * len(data))[i], tag
        )
        for prop in used:
            out.append({
                "Action": "",
                "Wbs Code": wbs,
                "Tag Number": equip_no,
                "Property Name": prop,
                "Property Value": value_map[prop][i],
                "UoM": uom_map[prop][i],
                "Comment": "",
            })
    return out, warnings


# ══════════════════════════════════════════════════════════════════════
# Оформление и сохранение
# ══════════════════════════════════════════════════════════════════════
def write_output(rows: List[Dict[str, str]], path: Path) -> None:
    """Записывает результат в Excel с оформлением по правилам проекта."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data 2"
    ws.append(list(OUT_COLUMNS))
    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for i, rec in enumerate(rows, start=2):
        ws.append([rec[c] for c in OUT_COLUMNS])
        for cell in ws[i]:
            cell.border = BORDER
            if i % 2 == 0:
                cell.fill = BAND_FILL
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{len(rows) + 1}"
    for j, col in enumerate(OUT_COLUMNS, start=1):
        width = len(col)
        for rec in rows:
            width = max(width, len(rec[col]))
        ws.column_dimensions[get_column_letter(j)].width = min(width + 2, MAX_COL_WIDTH)
    wb.save(path)


# ══════════════════════════════════════════════════════════════════════
# Точка входа
# ══════════════════════════════════════════════════════════════════════
def main() -> int:
    """Разбор аргументов и запуск конвертации. Возвращает код выхода."""
    parser = argparse.ArgumentParser(
        description="Конвертация исходных данных оборудования в формат образца",
    )
    parser.add_argument("--input", default=DEFAULT_INPUT, help="файл исходных данных")
    parser.add_argument("--sample", default=DEFAULT_SAMPLE, help="файл-образец")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help="файл результата")
    args = parser.parse_args()

    try:
        cols, data, data_raw = read_source(Path(args.input))
        order = read_sample_order(Path(args.sample))
        rows, warnings = build_rows(cols, data, order, data_raw)
        for w in warnings:
            print(f"ПРЕДУПРЕЖДЕНИЕ: {w}")
        write_output(rows, Path(args.output))
        print(f"Готово: {len(rows)} строк ({len(data)} приборов x {len(cols)} свойств) -> {args.output}")
        return 0
    except Exception:
        print("ОШИБКА:\n" + traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
