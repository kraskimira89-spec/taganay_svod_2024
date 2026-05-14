"""
Универсальный pipeline по ОСВ счёта 60 для любого года (`--year`).

Структура (воспроизводимый расчёт):
  1. `load_osv60` — читает .xls, классифицирует строки (статья + контрагент);
  2. `detect_group` — только правило по тексту статьи (для отладки/тестов);
  3. `build_svod` — группировка включённых строк по внутренней группе, суммы по кредиту и на соцтакси;
  4. `save_report` — Excel: «Строки_ОСВ60», «Свод_по_группам», «Не_включено_в_свод», «Группа_затрат_записка».

Примеры:
  python svod_osv60.py --year 2024
  python svod_osv60.py --year 2025 --osv-path "_extract_osv/Osv_schet_60_2025.xls"
  python svod_osv60.py --year 2024 --article-col 7
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from config_osv60 import classify_by_article
from svod_osv60_core import (
    BASE_DIR,
    EXTRACT_DIR,
    build_svod_dataframe,
    build_zapiska_table,
    load_osv60_rows,
    print_osv60_run_summary,
    write_osv60_workbook,
)


def load_osv60(
    path: Path,
    article_col_idx: int | None = None,
    year: int = 2024,
) -> tuple[pd.DataFrame, pd.DataFrame, int | None]:
    """
    Загрузка ОСВ 60: возвращает (включённые строки, исключённые, индекс колонки статьи или None).
    Колонки — как в `load_osv60_rows` (Контрагент, Оборот_Кт, Внутренняя_группа, На_соцтакси и т.д.).
    `year` — для коэффициента косвенных расходов (`config_allocation.get_soc_taxi_share`).
    """
    return load_osv60_rows(path, article_col_idx=article_col_idx, year=year)


def detect_group(article: str | None) -> str | None:
    """По тексту статьи расходов — внутренняя группа; если правило не сработало — None."""
    r = classify_by_article(article)
    return r[0] if r else None


def build_svod(df: pd.DataFrame) -> pd.DataFrame:
    """Свод по полю «Внутренняя_группа»: сумма оборота по кредиту и сумма на соцтакси."""
    return build_svod_dataframe(df)


def save_report(
    raw_df: pd.DataFrame,
    svod_df: pd.DataFrame,
    zapiska_df: pd.DataFrame,
    skipped_df: pd.DataFrame,
    output_path: Path,
) -> None:
    """Сохраняет четыре листа в один .xlsx."""
    write_osv60_workbook(output_path, raw_df, skipped_df, svod_df, zapiska_df)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Свод прочих расходов по счёту 60 для соцтакси (любой год)")
    p.add_argument("--year", type=int, default=2024, help="Год (подпись итога и имя выхода по умолчанию)")
    p.add_argument(
        "--osv-path",
        type=Path,
        default=None,
        help="Путь к .xls ОСВ 60 (по умолчанию _extract_osv/Osv_schet_60_{year}.xls)",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Выходной .xlsx (по умолчанию Osv60_soc_taxi_{year}.xlsx)",
    )
    p.add_argument(
        "--article-col",
        type=int,
        default=None,
        help="0-based индекс колонки «Статья расходов» (колонка H = 7), если не определяется автоматически",
    )
    return p


def main() -> None:
    args = build_parser().parse_args()
    osv = args.osv_path
    if osv is not None and not osv.is_absolute():
        osv = BASE_DIR / osv
    out = args.output
    if out is not None and not out.is_absolute():
        out = BASE_DIR / out
    if osv is None:
        osv = EXTRACT_DIR / f"Osv_schet_60_{args.year}.xls"
    if out is None:
        out = BASE_DIR / f"Osv60_soc_taxi_{args.year}.xlsx"

    # Тот же pipeline, что и в `run_with_paths`, но явными шагами API
    detail, skipped, used_col = load_osv60(osv, args.article_col, year=args.year)
    svod = build_svod(detail)
    zapiska = build_zapiska_table(detail, args.year)
    save_report(detail, svod, zapiska, skipped, out)
    print_osv60_run_summary(osv, out, used_col, args.article_col, detail, skipped, zapiska)


if __name__ == "__main__":
    main()
