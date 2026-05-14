"""
Переименование файлов в _extract_osv с «кракозябрами» в имени в читаемые UTF-8 имена.

.xls: по заголовку внутри файла — «Оборотно-сальдовая ведомость по счету NN за YYYY г.»
  -> Osv_schet_NN_YYYY.xls

.xlsx: если по содержимому похоже на свод доходов/расходов за год (не ОСВ)
  -> Svod_finansovye_pokazateli_YYYY.xlsx

Запуск из корня проекта:
  python scripts/rename_extract_osv_files.py
"""

from __future__ import annotations

import re
import shutil
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent
EXTRACT = BASE / "_extract_osv"


def _read_osv_account_year(path: Path) -> tuple[str, str] | None:
    try:
        df = pd.read_excel(path, header=None, nrows=8, engine="xlrd")
    except Exception:
        return None
    for i in range(len(df)):
        cell = df.iloc[i, 0]
        if not isinstance(cell, str):
            continue
        m = re.search(r"сч[её]ту\s+(\d+)", cell, flags=re.IGNORECASE)
        if not m:
            continue
        acc = m.group(1)
        y = re.search(r"(\d{4})", cell)
        if not y:
            continue
        return acc, y.group(1)
    return None


def _is_finance_summary_xlsx(path: Path) -> int | None:
    """Возвращает год или None."""
    try:
        df = pd.read_excel(path, header=None, nrows=20, engine="openpyxl")
    except Exception:
        return None
    text = " ".join(
        str(x) for x in df.values.ravel() if isinstance(x, str) or x is not None
    )
    if re.search(r"сч\s*60|счёт\s*60|счету\s*60", text, re.I):
        pass  # might still be summary table with сч60 row
    if "ВСЕГО" in text or "в т.ч" in text.lower() or "Поступило" in text:
        m = re.search(r"(20\d{2})\s*год", text)
        if m:
            return int(m.group(1))
        m2 = re.search(r"(20\d{2})", str(df.iloc[1, 0]) if len(df) > 1 else "")
        if m2:
            return int(m2.group(1))
    return None


def main() -> None:
    if not EXTRACT.is_dir():
        raise SystemExit(f"Нет папки: {EXTRACT}")

    for path in sorted(EXTRACT.iterdir()):
        if re.fullmatch(r"Osv_schet_\d+_\d+\.xls", path.name, flags=re.I):
            print("Уже нормальное имя:", path.name)
            continue
        if re.fullmatch(r"Svod_finansovye_pokazateli_\d{4}\.xlsx", path.name, flags=re.I):
            print("Уже нормальное имя:", path.name)
            continue
        if path.suffix.lower() == ".xls":
            meta = _read_osv_account_year(path)
            if not meta:
                print("Пропуск (не распознана ОСВ):", path.name)
                continue
            acc, year = meta
            new_name = f"Osv_schet_{acc}_{year}.xls"
            new_path = path.with_name(new_name)
            if new_path.resolve() == path.resolve():
                continue
            if new_path.exists():
                raise SystemExit(f"Целевой файл уже есть: {new_path}")
            path.rename(new_path)
            print(f"{path.name!r} -> {new_name}")

        elif path.suffix.lower() == ".xlsx":
            year = _is_finance_summary_xlsx(path)
            if year:
                new_name = f"Svod_finansovye_pokazateli_{year}.xlsx"
            else:
                new_name = f"Import_{path.stem[:40]}.xlsx"
            new_path = path.with_name(new_name)
            if new_path.resolve() == path.resolve():
                continue
            if new_path.exists():
                raise SystemExit(f"Целевой файл уже есть: {new_path}")
            path.rename(new_path)
            print(f"{path.name!r} -> {new_name}")
        else:
            print("Не трогаю:", path.name)


if __name__ == "__main__":
    main()
