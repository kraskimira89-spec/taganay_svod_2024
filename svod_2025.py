"""Свод ФОТ и взносов за 2025 → Personal_2025_soc_taxi.xlsx. См. `svod_personal_core`."""

from __future__ import annotations

import argparse
from pathlib import Path

from svod_personal_core import PersonalPipelineConfig, run_personal_pipeline


def main() -> None:
    ap = argparse.ArgumentParser(description="Сбор Personal_2025_soc_taxi.xlsx (матрица как в 2024, k из config_allocation)")
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
        help="Папка с файлами «Налоги и взносы» (*.xlsx), например Nalogi-i-vznosy-2025",
    )
    ap.add_argument(
        "--vznosy-glob",
        type=str,
        default=None,
        help="Маска взносов в корне проекта, напр. Nalogi-i-vznosy*2025*.xlsx",
    )
    ap.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Куда сохранить итоговый Personal_{year}_soc_taxi.xlsx (по умолчанию корень проекта)",
    )
    args = ap.parse_args()
    year = args.year
    cfg = PersonalPipelineConfig(
        personal_file=args.zp_file,
        vznosy_dir=args.vznosy_dir,
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
