# -*- coding: utf-8 -*-
"""Основная логика чтения, нормализации, протяжки и конвертации wide -> long."""

import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .config import ConverterConfig, OUTPUT_COLUMNS
from .formatters import NA_VALUE, fmt_value, norm_str, resolve_equipment_number
from .uom import detect_adjacent_uom_columns, resolve_uom

logger = logging.getLogger(__name__)

# Стили оформления Excel
HDR_FILL = PatternFill("solid", fgColor="D9E2F3")
HDR_FONT = Font(bold=True, color="1F4E79")
EVEN_FILL = PatternFill("solid", fgColor="F2F6FC")
ODD_FILL = PatternFill("solid", fgColor="FFFFFF")
THIN_SIDE = Side(style="thin", color="B0B0B0")
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
MAX_COL_WIDTH: int = 50


def _merge_cell_value(computed: Any, raw: Any) -> Any:
    """Возвращает вычисленное значение ячейки, а если оно пустое — сырое (формулу/текст)."""
    if computed is None or (isinstance(computed, str) and not computed.strip()):
        return raw
    return computed


def read_source_dataframe(path: Path, config: ConverterConfig) -> pd.DataFrame:
    """
    Читает исходный лист Excel в DataFrame, сохраняя вычисленные значения
    с резервом на сырые формулы и присваивая технические имена безымянным UoM-колонкам.
    """
    if not path.exists():
        raise FileNotFoundError(f"Исходный файл не найден: {path}")

    wb_val = openpyxl.load_workbook(path, read_only=True, data_only=True)
    wb_raw = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        sheet_names = wb_val.sheetnames
        if isinstance(config.source_sheet, int):
            if config.source_sheet >= len(sheet_names):
                raise ValueError(f"Индекс листа {config.source_sheet} вне диапазона в {path.name}")
            sheet_name = sheet_names[config.source_sheet]
        else:
            if config.source_sheet not in sheet_names:
                raise ValueError(f"Лист {config.source_sheet!r} не найден в файле {path.name}")
            sheet_name = config.source_sheet

        rows_val = [list(r) for r in wb_val[sheet_name].iter_rows(values_only=True)]
        rows_raw = [list(r) for r in wb_raw[sheet_name].iter_rows(values_only=True)]
    finally:
        wb_val.close()
        wb_raw.close()

    if len(rows_val) <= config.header_row + 1:
        raise ValueError(f"Лист {sheet_name!r} пуст: отсутствуют строки данных")

    raw_header = rows_val[config.header_row]
    columns: List[str] = []
    seen_names: Dict[str, int] = {}

    for idx, h in enumerate(raw_header):
        name = norm_str(h)
        if not name:
            # Безымянные колонки справа от свойства хранят UoM (в режиме equipment)
            columns.append(f"__uom_{idx}")
        else:
            if name in seen_names:
                seen_names[name] += 1
                columns.append(f"{name}__{seen_names[name]}")
            else:
                seen_names[name] = 0
                columns.append(name)

    data_start = config.header_row + 1 + config.skip_rows_after_header
    merged_rows: List[List[Any]] = []
    max_cols = len(columns)

    for r_idx in range(data_start, len(rows_val)):
        v_row = rows_val[r_idx]
        r_row = rows_raw[r_idx] if r_idx < len(rows_raw) else []
        row_cells: List[Any] = []
        for c_idx in range(max_cols):
            val_c = v_row[c_idx] if c_idx < len(v_row) else None
            raw_c = r_row[c_idx] if c_idx < len(r_row) else None
            row_cells.append(_merge_cell_value(val_c, raw_c))
        merged_rows.append(row_cells)

    df = pd.DataFrame(merged_rows, columns=columns)
    return df


def read_sample_order(sample_path: Optional[Path], config: ConverterConfig) -> List[str]:
    """Читает порядок свойств из файла-образца или возвращает порядок по умолчанию."""
    if sample_path and sample_path.exists():
        sample_df = pd.read_excel(sample_path, sheet_name=0, keep_default_na=False)
        if "Property Name" not in sample_df.columns:
            raise ValueError(f"В образце {sample_path.name} отсутствует колонка 'Property Name'")
        order: List[str] = []
        for val in sample_df["Property Name"]:
            name = norm_str(val)
            if name and name not in order:
                order.append(name)
        if order:
            logger.info("Порядок свойств загружен из образца %s (%d шт.)", sample_path.name, len(order))
            return order

    if config.default_property_order:
        logger.info(
            "Образец не найден или пуст — используется порядок по умолчанию (%d шт.)",
            len(config.default_property_order),
        )
        return list(config.default_property_order)

    raise FileNotFoundError(f"Файл-образец не найден и порядок по умолчанию не задан: {sample_path}")


def _apply_targeted_ffill(df: pd.DataFrame, config: ConverterConfig) -> pd.DataFrame:
    """
    Удаляет полностью пустые строки и выполняет точечный forward_fill
    по общим колонкам свойств и связанным колонкам единиц измерения.
    """
    cleaned = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").copy()
    if cleaned.empty:
        return cleaned

    excluded = set(config.ffill_exclude_columns)
    cols_to_fill = [col for col in cleaned.columns if col not in excluded]
    cleaned[cols_to_fill] = cleaned[cols_to_fill].ffill()
    return cleaned


def convert_wide_to_long(
    df: pd.DataFrame,
    config: ConverterConfig,
    property_order: Optional[List[str]] = None,
) -> Tuple[pd.DataFrame, int]:
    """
    Конвертирует DataFrame из широкого формата (1 прибор = 1 строка)
    в длинный формат образца (1 свойство = 1 строка).

    Возвращает:
        (выходной DataFrame с колонками OUTPUT_COLUMNS, количество прочитанных приборов)
    """
    prepared_df = _apply_targeted_ffill(df, config)
    if prepared_df.empty:
        return pd.DataFrame(columns=list(OUTPUT_COLUMNS)), 0

    order = property_order if property_order is not None else list(config.default_property_order)
    if not order:
        order = [c for c in prepared_df.columns if not c.startswith("__uom_")]

    available_cols = set(prepared_df.columns)
    used_properties = [p for p in order if p in available_cols]
    missing_properties = [p for p in order if p not in available_cols]

    for prop in missing_properties:
        logger.warning("Свойство %r из образца не найдено в исходнике — пропущено", prop)

    adjacent_uom_map = (
        detect_adjacent_uom_columns(list(prepared_df.columns))
        if config.uom_mode == "adjacent"
        else None
    )

    records: List[Dict[str, str]] = []
    devices_count: int = 0

    for _, row in prepared_df.iterrows():
        raw_tag = fmt_value("tag number", row.get("tag number"))
        if config.require_non_empty_tag and not raw_tag:
            continue

        devices_count += 1
        wbs_code = fmt_value("wbs code", row.get("wbs code"))

        # Для equipment в Tag Number записывается equipment number (включая формулу E-...)
        if config.use_equipment_formula:
            raw_equip = row.get(config.tag_source_column)
            out_tag = resolve_equipment_number(raw_equip, raw_tag)
        else:
            out_tag = fmt_value(config.tag_source_column, row.get(config.tag_source_column))

        for prop in used_properties:
            raw_val = row.get(prop)
            formatted_val = fmt_value(prop, raw_val, config.date_properties)

            if prop in config.default_na_properties and not formatted_val:
                formatted_val = NA_VALUE

            uom_val = resolve_uom(prop, row, config, adjacent_uom_map)

            records.append({
                "Action": "",
                "Wbs Code": wbs_code,
                "Tag Number": out_tag,
                "Property Name": prop,
                "Property Value": formatted_val,
                "UoM": uom_val,
                "Comment": "",
            })

    result_df = pd.DataFrame(records, columns=list(OUTPUT_COLUMNS))
    return result_df, devices_count


def write_output(result_df: pd.DataFrame, out_path: Path, config: ConverterConfig) -> None:
    """Сохраняет итоговый DataFrame в Excel с оформлением по стандарту проекта."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = config.output_sheet

    headers = list(OUTPUT_COLUMNS)
    ws.append(headers)

    for cell in ws[1]:
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for row_idx, row_values in enumerate(result_df.itertuples(index=False, name=None), start=2):
        ws.append(list(row_values))
        fill = EVEN_FILL if row_idx % 2 == 0 else ODD_FILL
        for cell in ws[row_idx]:
            cell.fill = fill
            cell.border = CELL_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    for col_idx, col_name in enumerate(headers, start=1):
        width = len(col_name) + 2
        for val in result_df[col_name].head(500):
            width = max(width, min(len(str(val)) + 2, MAX_COL_WIDTH))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(width, MAX_COL_WIDTH)

    ws.freeze_panes = "A2"
    if not result_df.empty:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(result_df) + 1}"

    wb.save(out_path)
