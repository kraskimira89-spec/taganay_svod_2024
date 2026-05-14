"""
Копирует выгрузки ОСВ по счетам 01 и 02 (основные средства и амортизация) в _extract_osv/.

Имена в проекте: Osv_schet_01_{год}.xls, Osv_schet_02_{год}.xls

Пути к исходникам задаются переменными окружения (полный путь к каждому файлу).
Если переменная не задана или файла нет — строка пропускается.

  OSV01_2024, OSV02_2024, OSV01_2025, OSV02_2025

Пример (PowerShell):
  $env:OSV02_2024="E:\\...\\ОСВ по счету 02 за 2024.xls"
  python scripts/sync_osv_01_02.py

Запуск:
  python scripts/sync_osv_01_02.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST_DIR = BASE / "_extract_osv"

MAPPING: list[tuple[str, str]] = [
    ("OSV01_2024", "Osv_schet_01_2024.xls"),
    ("OSV02_2024", "Osv_schet_02_2024.xls"),
    ("OSV01_2025", "Osv_schet_01_2025.xls"),
    ("OSV02_2025", "Osv_schet_02_2025.xls"),
]


def main() -> None:
    DEST_DIR.mkdir(parents=True, exist_ok=True)
    copied = 0
    for env_key, dest_name in MAPPING:
        raw = os.environ.get(env_key)
        if not raw:
            print(f"Пропуск {dest_name}: не задана переменная {env_key}")
            continue
        src = Path(raw)
        if not src.is_file():
            print(f"Пропуск {dest_name}: нет файла {src}")
            continue
        dest = DEST_DIR / dest_name
        shutil.copy2(src, dest)
        print("OK:", dest_name, "<-", src)
        copied += 1
    if copied == 0:
        raise SystemExit(
            "Не скопировано ни одного файла. Задайте OSV01_2024, OSV02_2024 и т.д. — см. docstring."
        )


if __name__ == "__main__":
    main()
