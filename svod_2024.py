"""Свод ФОТ и взносов за 2024 год → Personal_2024_soc_taxi.xlsx. См. `svod_personal_core.run_personal_pipeline`."""

from __future__ import annotations

from svod_personal_core import run_personal_pipeline


def main() -> None:
    run_personal_pipeline(2024)


if __name__ == "__main__":
    main()
