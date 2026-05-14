"""
Копирует «верную» ОСВ по счёту 60 за 2024 в проект как _extract_osv/Osv_schet_60_2024.xls,
чтобы дальше работали команды без --osv-path.

Путь к исходнику задайте при необходимости переменной окружения OSV60_VERNAYA_2024
или поправьте SOURCE в этом файле.

Запуск:
  python scripts/sync_osv60_vernaya_2024.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST = BASE / "_extract_osv" / "Osv_schet_60_2024.xls"

DEFAULT_SOURCE = Path(
    r"E:\ЦИП Таганай\6-БУХГАЛТЕРИЯ\Оборотно-сальдовые ведомости"
    r"\Оборотно-сальдовая ведомость верная по счету 60 за 2024 г.xls"
)


def main() -> None:
    src = Path(os.environ.get("OSV60_VERNAYA_2024", str(DEFAULT_SOURCE)))
    if not src.is_file():
        raise SystemExit(f"Нет файла-источника: {src}")
    DEST.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, DEST)
    print("Скопировано:", src, "->", DEST)


if __name__ == "__main__":
    main()
