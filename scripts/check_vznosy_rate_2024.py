# -*- coding: utf-8 -*-
"""
Проверка: фактическая доля взносов к базе (начислено по взносам) по каждому файлу Nalogi-i-vznosy*.xlsx
ожидается около льготной ставки 7,2% (протокол рабочей сессии).

Запуск из корня проекта:
  python scripts/check_vznosy_rate_2024.py

Порог отклонения от 0.072 по умолчанию: 0.005 (0,5 п.п.).
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from svod_personal_core import read_tax_sheets  # noqa: E402


def main() -> None:
    expected = 0.072
    tol = 0.005
    files = sorted(ROOT.glob("Nalogi-i-vznosy*.xlsx"))
    if not files:
        print("Нет файлов Nalogi-i-vznosy*.xlsx в корне проекта.")
        return
    rows = []
    for f in files:
        parts = read_tax_sheets(f, 2024)
        if not parts:
            continue
        full = pd.concat(parts, ignore_index=True)
        base = float(full["Начислено_по_взносам"].sum())
        vzn = float(full["Взносы_2024"].sum())
        rate = (vzn / base) if base else 0.0
        rows.append(
            {
                "Файл": f.name,
                "База_начислено_по_взносам": base,
                "Взносы_2024": vzn,
                "Доля_факт": rate,
                "Отклонение_от_7_2pct": rate - expected,
                "OK": abs(rate - expected) <= tol or base == 0,
            }
        )
    if not rows:
        print("В файлах не найдено листов «Налоги и взносы».")
        return
    out = pd.DataFrame(rows)
    print(out.to_string(index=False))
    bad = out[~out["OK"]]
    if len(bad) > 0:
        print("\nВнимание: есть файлы вне допуска ±7,2%. Проверьте вручную помесячно / по сотрудникам.")
    else:
        print("\nВсе проверенные файлы в допуске по совокупной доле.")


if __name__ == "__main__":
    main()
