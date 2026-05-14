"""
Свод по ОСВ счёта 60 за 2025 год → `Osv60_soc_taxi_2025.xlsx`.

Те же правила, что и для 2024 (`svod_osv60_core.py`); исходник: `Osv_schet_60_2025.xls`.
"""

from __future__ import annotations

from svod_osv60_core import run_year

if __name__ == "__main__":
    run_year(2025)
