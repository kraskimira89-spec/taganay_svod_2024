"""Свод ФОТ и взносов за 2025 год → Personal_2025_soc_taxi.xlsx (k админ-долей из `config_allocation`)."""

from __future__ import annotations

from svod_personal_core import run_personal_pipeline


def main() -> None:
    run_personal_pipeline(2025)


if __name__ == "__main__":
    main()
