# -*- coding: utf-8 -*-
"""Конфигурации режимов конвертации (TAGS и EQUIPMENT)."""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union


@dataclass(frozen=True)
class ConverterConfig:
    """Параметры конвертации для конкретного типа выгрузки."""

    name: str
    default_input: str
    default_sample: str
    default_output: str
    source_sheet: Union[int, str]
    output_sheet: str
    header_row: int                                # индекс строки (0-based) с английскими заголовками
    skip_rows_after_header: int = 0
    tag_source_column: str = "tag number"          # колонка, из которой берётся идентификатор для Tag Number
    use_equipment_formula: bool = False            # парсить формулу ="E-"&F... для Tag Number
    uom_mode: str = "column"                       # "column" (общая колонка UoM) или "adjacent" (соседняя справа)
    uom_column: Optional[str] = None
    uom_target_properties: Tuple[str, ...] = ()
    fixed_uom: Dict[str, str] = field(default_factory=dict)
    default_na_properties: Tuple[str, ...] = ()    # свойства, для которых пустое значение заменяется на "N/A"
    date_properties: Tuple[str, ...] = ()
    default_property_order: Tuple[str, ...] = ()
    ffill_exclude_columns: Tuple[str, ...] = ()    # колонки, которые не протягиваются (пусто = тянутся все общие)
    require_non_empty_tag: bool = False            # пропускать строки без tag number


OUTPUT_COLUMNS: Tuple[str, ...] = (
    "Action",
    "Wbs Code",
    "Tag Number",
    "Property Name",
    "Property Value",
    "UoM",
    "Comment",
)

TAGS_CONFIG = ConverterConfig(
    name="tags",
    default_input="source_data.xlsx",
    default_sample="sample.xlsx",
    default_output="result.xlsx",
    source_sheet=0,
    output_sheet="Data 1",
    header_row=0,
    tag_source_column="tag number",
    use_equipment_formula=False,
    uom_mode="column",
    uom_column="ед. изм 2",
    uom_target_properties=("set point high", "set point low"),
    fixed_uom={"design lifetime": "year"},
    default_na_properties=("commissioning system code",),
    date_properties=(),
    default_property_order=(
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
    ),
    require_non_empty_tag=True,
)

EQUIPMENT_CONFIG = ConverterConfig(
    name="equipment",
    default_input="source_data2.xlsx",
    default_sample="sample2.xlsx",
    default_output="result2.xlsx",
    source_sheet="Лист2",
    output_sheet="Data 2",
    header_row=1,  # строка 0 — русский заголовок, строка 1 — английский
    tag_source_column="equipment number",
    use_equipment_formula=True,
    uom_mode="adjacent",
    date_properties=("manufacturing date", "warranty end date", "purchase date"),
    require_non_empty_tag=False,
)

CONVERTER_CONFIGS: Dict[str, ConverterConfig] = {
    "tags": TAGS_CONFIG,
    "equipment": EQUIPMENT_CONFIG,
}
