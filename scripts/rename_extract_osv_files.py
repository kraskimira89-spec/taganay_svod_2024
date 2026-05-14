"""
Переименование файлов в _extract_osv с «кракозябрами» в имени в читаемые UTF-8 имена.

.xls: по заголовку внутри файла — «Оборотно-сальдовая ведомость по счету NN за YYYY г.»
  -> Osv_schet_NN_YYYY.xls

.xlsx / .xls (не ОСВ): если по содержимому похоже на свод доходов/расходов за год
  -> Svod_finansovye_pokazateli_YYYY.xlsx или .xls

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


def _finance_summary_year_from_cells(df: pd.DataFrame, max_rows: int = 30) -> int | None:
    """Общая эвристика: свод по счетам / поступлениям за год (не ОСВ)."""
    parts: list[str] = []
    for i in range(min(max_rows, len(df))):
        for j in range(min(df.shape[1], 12)):
            v = df.iloc[i, j]
            if isinstance(v, str) and v.strip():
                parts.append(v)
            elif v is not None and not (isinstance(v, float) and pd.isna(v)):
                parts.append(str(v))
    text = " ".join(parts)
    if re.search(r"сч\s*60|счёт\s*60|счету\s*60|\s60\s", text, re.I):
        pass
    if "ВСЕГО" in text or "в т.ч" in text.lower() or "Поступило" in text:
        m = re.search(r"(20\d{2})\s*год", text)
        if m:
            return int(m.group(1))
        for i in range(min(8, len(df))):
            for j in range(min(4, df.shape[1])):
                s = str(df.iloc[i, j])
                m2 = re.search(r"(20\d{2})", s)
                if m2:
                    return int(m2.group(1))
    return None


def _is_finance_summary_xlsx(path: Path) -> int | None:
    """Возвращает год или None."""
    try:
        df = pd.read_excel(path, header=None, nrows=30, engine="openpyxl")
    except Exception:
        return None
    return _finance_summary_year_from_cells(df)


def _is_finance_summary_xls(path: Path) -> int | None:
    """Свод финпоказателей в старом .xls (xlrd), не ОСВ."""
    try:
        df = pd.read_excel(path, header=None, nrows=30, engine="xlrd")
    except Exception:
        return None
    return _finance_summary_year_from_cells(df)


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
        if re.fullmatch(r"Svod_finansovye_pokazateli_\d{4}\.xls", path.name, flags=re.I):
            print("Уже нормальное имя:", path.name)
            continue
        if path.suffix.lower() == ".xls":
            meta = _read_osv_account_year(path)
            if meta:
                acc, year = meta
                new_name = f"Osv_schet_{acc}_{year}.xls"
                new_path = path.with_name(new_name)
                if new_path.resolve() == path.resolve():
                    continue
                if new_path.exists():
                    raise SystemExit(f"Целевой файл уже есть: {new_path}")
                path.rename(new_path)
                print(f"{path.name!r} -> {new_name}")
                continue

            fin_year = _is_finance_summary_xls(path)
            if fin_year:
                new_name = f"Svod_finansovye_pokazateli_{fin_year}.xls"
                new_path = path.with_name(new_name)
                if new_path.resolve() == path.resolve():
                    continue
                if new_path.exists():
                    raise SystemExit(f"Целевой файл уже есть: {new_path}")
                path.rename(new_path)
                print(f"{path.name!r} -> {new_name}")
                continue

            print("Пропуск (не распознана ОСВ и не свод показателей):", path.name)
            continue

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
