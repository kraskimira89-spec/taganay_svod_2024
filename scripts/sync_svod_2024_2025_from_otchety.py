"""
Копирует сводный Excel «Свод 2024-2025 .xlsx» из папки отчётов
(`…\\Отчеты 2024 год\\ОСВ и свод\\`) в корень проекта под стабильным именем
без пробела перед расширением.

Переопределить папку-источник: переменная окружения SVOD_2024_2025_SOURCE_DIR
(ожидается один файл *2024*2025*.xlsx внутри).

Запуск:
  python scripts/sync_svod_2024_2025_from_otchety.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST = BASE / "Svod_2024_2025_iz_arhiva_OSV_i_svod.xlsx"

DEFAULT_DIR = Path(
    r"E:\ЦИП Таганай\Суд\ДСЗН 2025 по тарифу соцтакси\Отчеты 2024 год\ОСВ и свод"
)


def _pick_source() -> Path:
    d = Path(os.environ.get("SVOD_2024_2025_SOURCE_DIR", str(DEFAULT_DIR)))
    if not d.is_dir():
        raise SystemExit(f"Нет папки-источника: {d}")
    matches = sorted(d.glob("*2024*2025*.xlsx"))
    if len(matches) != 1:
        raise SystemExit(
            f"Ожидается ровно один файл *2024*2025*.xlsx в {d}, найдено: {len(matches)}"
        )
    return matches[0]


def main() -> None:
    src = _pick_source()
    if not src.is_file():
        raise SystemExit(f"Нет файла: {src}")
    shutil.copy2(src, DEST)
    print("Скопировано:", src, "->", DEST)


if __name__ == "__main__":
    main()
