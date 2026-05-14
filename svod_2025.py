"""
Свод ФОТ и страховых взносов за 2025 год → `Personal_2025_soc_taxi.xlsx`.

Логика (матрица долей, группы затрат, k из `config_allocation`) — в `svod_personal_core.py`.
Этот модуль только задаёт аргументы CLI и при необходимости подставляет папку взносов
`Nalogi-i-vznosy-2025/`, если она есть в корне проекта.

Исходники перед запуском:
  1. `Svodnyi-po-ZP-za-2025-god.xlsx` в корне (или `--zp-file`).
  2. Файлы «Налоги и взносы» за 2025 — папка `Nalogi-i-vznosy-2025/*.xlsx` или маска
     `Nalogi-i-vznosy*2025*.xlsx` в корне / `--vznosy-glob` / переменная `VZNOSY_GLOB`.
  3. Матрица участия — та же, что для 2024 (`_participation_rules_for_year` в
     `svod_personal_core.py`); изменения по сотрудникам — правки там.

Сверка: после расчёта в консоли выводится ИТОГО и сравнение со строкой счёта 70
«на соцтакси» из `Svod-2024-2025.xlsx` (~3 629 342,14 руб.). Расхождение возможно,
если в 1С другое распределение или изменился состав/доли.

Примеры:
  python svod_2025.py
  python svod_2025.py --vznosy-dir "C:\\taganay_svod_2024\\Nalogi-i-vznosy-2025"
  python svod_2025.py --zp-file "D:\\data\\Svodnyi-po-ZP-za-2025-god.xlsx" --output "Personal_2025_soc_taxi.xlsx"
"""

from __future__ import annotations

import argparse
from pathlib import Path

from svod_personal_core import PersonalPipelineConfig, run_personal_pipeline

BASE = Path(__file__).resolve().parent


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Сбор Personal_2025_soc_taxi.xlsx (ФОТ + взносы, доли соцтакси, k из config_allocation).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Подробности и сверка со счётом 70 — в docstring модуля svod_2025.py и в svod_personal_core.run_personal_pipeline.",
    )
    ap.add_argument("--year", type=int, default=2025, help="Год (по умолчанию 2025)")
    ap.add_argument(
        "--zp-file",
        type=Path,
        default=None,
        help="Сводная ЗП Excel (по умолчанию Svodnyi-po-ZP-za-{year}-god.xlsx в корне проекта)",
    )
    ap.add_argument(
        "--vznosy-dir",
        type=Path,
        default=None,
        help="Папка с файлами «Налоги и взносы» (*.xlsx). Если не задано и нет --vznosy-glob, "
        "используется папка Nalogi-i-vznosy-2025 в корне проекта, если она существует",
    )
    ap.add_argument(
        "--vznosy-glob",
        type=str,
        default=None,
        help="Маска взносов в корне проекта, напр. Nalogi-i-vznosy*2025*.xlsx (или переменная VZNOSY_GLOB)",
    )
    ap.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Куда сохранить Personal_{year}_soc_taxi.xlsx (по умолчанию корень проекта)",
    )
    args = ap.parse_args()
    year = args.year

    vznosy_dir = args.vznosy_dir
    if vznosy_dir is None and args.vznosy_glob is None:
        cand = BASE / "Nalogi-i-vznosy-2025"
        if cand.is_dir():
            vznosy_dir = cand

    cfg = PersonalPipelineConfig(
        personal_file=args.zp_file,
        vznosy_dir=vznosy_dir,
        vznosy_glob=args.vznosy_glob,
        soc_taxi_out=args.output,
    )
    has_any = any(
        (
            cfg.personal_file is not None,
            cfg.vznosy_dir is not None,
            cfg.vznosy_glob is not None,
            cfg.soc_taxi_out is not None,
        )
    )
    run_personal_pipeline(year, cfg if has_any else None)


if __name__ == "__main__":
    main()
