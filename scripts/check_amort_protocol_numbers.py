# -*- coding: utf-8 -*-
"""
Напоминание цифр амортизации (счёт 02) из протокола рабочей сессии и проверка наличия файлов ОСВ.

Ожидаемые суммы (оборот Кт по амортизации автотранспорта — уточнить в вашей ОСВ):
  2024: 458 333,36 руб.
  2025: 499 999,98 руб.

После копирования выгрузок (`scripts/sync_osv_01_02.py`) сверьте визуально в Excel или расширьте скрипт под вашу разметку столбцов.

Запуск: python scripts/check_amort_protocol_numbers.py
"""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PROTOCOL = {2024: 458_333.36, 2025: 499_999.98}


def main() -> None:
    d = ROOT / "_extract_osv"
    for year, amount in PROTOCOL.items():
        p = d / f"Osv_schet_02_{year}.xls"
        st = "есть" if p.is_file() else f"нет — задайте OSV02_{year} и sync_osv_01_02.py"
        print(f"{year}: протокол Кт 02 ~ {amount:,.2f} руб.; файл {p.name}: {st}")


if __name__ == "__main__":
    main()
