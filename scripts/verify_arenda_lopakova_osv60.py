# -*- coding: utf-8 -*-
"""
Сверка: сумма «На_соцтакси» по контрагентам с фамилией Лопакова в Osv60_soc_taxi_2024.xlsx
с ожидаемой годовой арендой VW по договору № 01-20 (90 000 руб./мес → 1 080 000 руб./год).

Запуск:
  python scripts/verify_arenda_lopakova_osv60.py

При отсутствии файла или колонок — сообщение и код выхода 1.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
EXPECTED_YEAR_RENT = 1_080_000.0
OSV60_OUT = ROOT / "Osv60_soc_taxi_2024.xlsx"


def main() -> int:
    if not OSV60_OUT.is_file():
        print(f"Нет файла {OSV60_OUT.name}. Сначала: python svod_osv60.py --year 2024")
        return 1
    df = pd.read_excel(OSV60_OUT, sheet_name="Строки_ОСВ60", engine="openpyxl")
    if "Контрагент" not in df.columns or "На_соцтакси" not in df.columns:
        print("В листе нет колонок Контрагент / На_соцтакси")
        return 1
    m = df["Контрагент"].astype(str).str.contains("Лопакова", case=False, na=False)
    s = float(df.loc[m, "На_соцтакси"].sum())
    print(f"Сумма На_соцтакси по контрагентам «Лопакова»: {s:,.2f} руб.")
    print(f"Ожидается по договору аренды VW (12 мес. по 90 тыс.): {EXPECTED_YEAR_RENT:,.2f} руб.")
    diff = s - EXPECTED_YEAR_RENT
    print(f"Разница: {diff:,.2f} руб.")
    if abs(diff) < 1.0:
        print("Совпадение в пределах 1 руб.")
    elif abs(diff) / EXPECTED_YEAR_RENT < 0.02:
        print("Небольшое расхождение (<2%) — проверить доп.услуги/НДС в строках ОСВ.")
    else:
        print("Существенное расхождение — проверить классификацию контрагента в svod_osv60_core.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
