#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
unified_converter.py — единая точка входа для конвертации выгрузок приборов и оборудования.

Примеры запуска:
    python unified_converter.py --type equipment --input source_data2.xlsx --sample sample2.xlsx
    python unified_converter.py --type tags --input source_data.xlsx --sample sample.xlsx
"""

import argparse
import logging
import sys
from pathlib import Path

from src.converter import (
    CONVERTER_CONFIGS,
    convert_wide_to_long,
    read_sample_order,
    read_source_dataframe,
    write_output,
)


def setup_logging() -> None:
    """Настраивает единый формат логирования в консоль."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def parse_args() -> argparse.Namespace:
    """Разбор аргументов командной строки."""
    parser = argparse.ArgumentParser(
        description="Конвертация выгрузок приборов/оборудования из широкого формата в длинный",
    )
    parser.add_argument(
        "--type",
        required=True,
        choices=sorted(CONVERTER_CONFIGS.keys()),
        help="Режим конвертации: tags или equipment",
    )
    parser.add_argument(
        "--input",
        default=None,
        help="Путь к исходному файлу Excel (по умолчанию берётся из конфига режима)",
    )
    parser.add_argument(
        "--sample",
        default=None,
        help="Путь к файлу-образцу для порядка свойств (по умолчанию берётся из конфига режима)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Путь к выходному файлу Excel (по умолчанию берётся из конфига режима)",
    )
    return parser.parse_args()


def main() -> int:
    """Запускает полный цикл конвертации."""
    setup_logging()
    logger = logging.getLogger("unified_converter")
    args = parse_args()

    config = CONVERTER_CONFIGS[args.type]
    input_path = Path(args.input or config.default_input)
    sample_path = Path(args.sample or config.default_sample)
    output_path = Path(args.output or config.default_output)

    logger.info("Запуск конвертации в режиме '%s'", config.name.upper())
    logger.info("Входной файл: %s | Образец: %s | Выход: %s", input_path, sample_path, output_path)

    try:
        source_df = read_source_dataframe(input_path, config)
        property_order = read_sample_order(sample_path, config)
        result_df, devices_count = convert_wide_to_long(source_df, config, property_order)

        write_output(result_df, output_path, config)

        logger.info("Прочитано приборов: %d", devices_count)
        logger.info("Создано строк в итоговой таблице: %d", len(result_df))
        logger.info("Результат успешно сохранён в %s", output_path)
        return 0
    except Exception as exc:
        logger.exception("Ошибка при выполнении конвертации: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
