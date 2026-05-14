"""
Копирует «верную» ОСВ по счёту 70 за 2024 в проект как _extract_osv/Osv_schet_70_2024.xls.

Путь: переменная окружения OSV70_VERNAYA_2024 или DEFAULT_SOURCE ниже.

Запуск:
  python scripts/sync_osv70_vernaya_2024.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST = BASE / "_extract_osv" / "Osv_schet_70_2024.xls"

DEFAULT_SOURCE = Path(
    r"E:\ЦИП Таганай\6-БУХГАЛТЕРИЯ\Оборотно-сальдовые ведомости"
    r"\ОБС верно по счету 70 за 2024 г.xls"
)


def main() -> None:
    src = Path(os.environ.get("OSV70_VERNAYA_2024", str(DEFAULT_SOURCE)))
    if not src.is_file():
        raise SystemExit(
            f"Нет файла-источника: {src}. Укажите OSV70_VERNAYA_2024 или положите выгрузку в путь по умолчанию."
        )
    DEST.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, DEST)
    print("Скопировано:", src, "->", DEST)


if __name__ == "__main__":
    main()
