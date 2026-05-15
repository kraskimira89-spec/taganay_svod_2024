# -*- coding: utf-8 -*-
"""Ведомости PDF по разнарядкам 2024/2025 (формат как у поставщиков) → data/delo_registers/."""

from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from delo_github_policy import (  # noqa: E402
    copy_register_to_case,
    ensure_registers_dir,
    update_register_opis_row,
    REGISTER_FILES,
)
from sozdat_vedomost_pdf import build_vedomost  # noqa: E402

CONFIG = [
    (
        2024,
        "12.1",
        BASE / "docs" / "Разнорядки 2024",
        "ВЕДОМОСТЬ ДОКУМЕНТОВ РАЗНАРЯДОК 2024",
    ),
    (
        2025,
        "12.2",
        BASE / "docs" / "Разнорядки 2025",
        "ВЕДОМОСТЬ ДОКУМЕНТОВ РАЗНАРЯДОК 2025",
    ),
]


def main() -> None:
    ensure_registers_dir()
    for year, app_id, folder, title in CONFIG:
        if not folder.is_dir():
            print(f"Пропуск {year}: нет папки {folder}")
            continue
        out = ensure_registers_dir() / REGISTER_FILES[app_id]
        print(f"\n=== {year} ===")
        build_vedomost(
            folder,
            out,
            title,
            group_label="Месяц / направление",
            mirror_dir=folder,
        )
        copy_register_to_case(out, "12_Разнарядки")
        update_register_opis_row(
            app_id,
            f"Сводная ведомость PDF разнарядок соцтакси за {year} год",
            2,
        )
    print("\nГотово. Исходные PDF — Яндекс.Диск (12.3 / 12.4).")


if __name__ == "__main__":
    main()
