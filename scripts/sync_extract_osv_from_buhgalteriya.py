"""
Копирует выгрузки ОСВ и свод из папки бухгалтерии в `_extract_osv` под именами,
ожидаемыми скриптами (svod_osv60, svod_reconcile_osv76, rename_extract_osv_files).

Источник по умолчанию:
  E:\\ЦИП Таганай\\6-БУХГАЛТЕРИЯ\\Оборотно-сальдовые ведомости

Переопределение: переменная окружения BUHGALTERIYA_OSV_DIR

Запуск из корня репозитория:
  python scripts/sync_extract_osv_from_buhgalteriya.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST = BASE / "_extract_osv"
DEFAULT_SRC = Path(r"E:\ЦИП Таганай\6-БУХГАЛТЕРИЯ\Оборотно-сальдовые ведомости")


def _find_long_osv60_xlsx(src: Path) -> Path | None:
    for p in src.glob("*60*.xlsx"):
        if "2024" in p.name and "2025" in p.name:
            return p
    return None


def main() -> None:
    src = Path(os.environ.get("BUHGALTERIYA_OSV_DIR", str(DEFAULT_SRC)))
    if not src.is_dir():
        raise SystemExit(f"Нет папки-источника: {src}")
    DEST.mkdir(parents=True, exist_ok=True)

    def cp(name: str, dest_name: str) -> None:
        a = src / name
        b = DEST / dest_name
        if not a.is_file():
            print("Пропуск (нет файла):", name)
            return
        shutil.copy2(a, b)
        print(f"OK: {name!r} -> {dest_name}")

    cp("Osv_schet_70_2025.xls", "Osv_schet_70_2025.xls")
    cp("ОБС верно по счету 70 за 2025 г. (1).xls", "Osv_schet_70_2025_vernaya_1c.xls")
    cp("ОБС верно по счету 60 за 2025 г.xls", "Osv_schet_60_2025.xls")
    cp("ОБС вернопо счету 76 за 2025 г.xls", "Osv_schet_76_2025.xls")

    long60 = _find_long_osv60_xlsx(src)
    if long60 is not None:
        shutil.copy2(long60, DEST / "Osv_schet_60_2024-2025.xlsx")
        print(f"OK: {long60.name!r} -> Osv_schet_60_2024-2025.xlsx")
    else:
        print("Пропуск: не найден xlsx ОСВ 60 (2024–2025)")

    cp("Свод 2024г .xlsx", "Svod_finansovye_pokazateli_2024.xlsx")


if __name__ == "__main__":
    main()
