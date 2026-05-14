"""
Свод по ОСВ счёта 60 за 2024 год → `Osv60_soc_taxi_2024.xlsx`.

Правила классификации контрагентов — в `svod_osv60_core.py`.
"""

from __future__ import annotations

from svod_osv60_core import run_year

if __name__ == "__main__":
    run_year(2024)
