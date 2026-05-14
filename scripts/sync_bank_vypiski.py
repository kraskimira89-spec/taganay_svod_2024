# -*- coding: utf-8 -*-
"""
Копирует банковские выписки (Сбер, ВТБ) в проект: docs/bank_vypiski/{sber,vtb}/

Переопределение путей — переменные окружения (опционально):
  BANK_VYPISKI_SBER_2024, BANK_VYPISKI_SBER_2025,
  BANK_VYPISKI_VTB_DIR — папка с выписками ВТБ (копируются все перечисленные имена).

Запуск:
  python scripts/sync_bank_vypiski.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST_ROOT = BASE / "docs" / "bank_vypiski"

# --- Сбер (PDF по годам) ---
SBER_2024_DEFAULT = Path(
    r"E:\ЦИП Таганай\5-Сбербанк\СберБизнес. Выписка 7 счетов за 2024.01.01-2024.12.31"
    r"\СберБизнес. Выписка за 2024.01.01-2024.12.31.pdf"
)
SBER_2025_DEFAULT = Path(
    r"E:\ЦИП Таганай\5-Сбербанк\СберБизнес. Выписка 7 счетов за 2025.01.01-2025.12.31"
    r"\СберБизнес. Выписка за 2025.01.01-2025.12.31.pdf"
)

# --- ВТБ (папка по умолчанию + имена файлов) ---
VTB_DIR_DEFAULT = Path(r"E:\ЦИП Таганай\5-ВТБ\Выписки счета ВТБ")
VTB_FILES = [
    "VTB_BankStatement_some_accounts_01.01.2024-31.12.2024_281732.pdf",
    "VTB_BankStatement_some_accounts_01.01.2024-31.12.2024_281731.xlsx",
    "VTB_BankStatement_some_accounts_01.01.2025-31.12.2025_279187.pdf",
    "VTB_BankStatement_some_accounts_01.01.2025-31.12.2025_279189.xlsx",
    "VTB_BankStatement_some_accounts_01.01.2024-31.12.2024_281269.xlsx",
    "VTB_BankStatement_some_accounts_01.01.2024-31.12.2024_281270.pdf",
]


def main() -> None:
    sber_2024 = Path(os.environ.get("BANK_VYPISKI_SBER_2024", str(SBER_2024_DEFAULT)))
    sber_2025 = Path(os.environ.get("BANK_VYPISKI_SBER_2025", str(SBER_2025_DEFAULT)))
    vtb_dir = Path(os.environ.get("BANK_VYPISKI_VTB_DIR", str(VTB_DIR_DEFAULT)))

    dest_sber = DEST_ROOT / "sber"
    dest_vtb = DEST_ROOT / "vtb"
    dest_sber.mkdir(parents=True, exist_ok=True)
    dest_vtb.mkdir(parents=True, exist_ok=True)

    pairs: list[tuple[Path, Path]] = []
    for src in (sber_2024, sber_2025):
        if src.is_file():
            pairs.append((src, dest_sber / src.name))
        else:
            print("Пропуск (нет файла):", src)

    for name in VTB_FILES:
        src = vtb_dir / name
        if src.is_file():
            pairs.append((src, dest_vtb / name))
        else:
            print("Пропуск (нет файла):", src)

    if not pairs:
        raise SystemExit("Нет ни одного файла для копирования. Проверьте пути на диске E:.")

    for src, dst in pairs:
        shutil.copy2(src, dst)
        print("OK:", dst.relative_to(BASE))

    print("Готово. Каталог:", DEST_ROOT.relative_to(BASE))


if __name__ == "__main__":
    main()
