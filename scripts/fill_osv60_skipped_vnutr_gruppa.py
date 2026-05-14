"""
Заполняет колонку «Внутренняя_группа» на листе «Не_включено_в_свод» в Osv60_soc_taxi_*.xlsx.

Группы согласуются с `docs/dogovory/README.md` и логикой `svod_osv60_core.classify_row`
(статья расходов, если есть, иначе контрагент).

Пример:
  python scripts/fill_osv60_skipped_vnutr_gruppa.py
  python scripts/fill_osv60_skipped_vnutr_gruppa.py --files "Perplexity_paket_Taganay_2024_2025/taganay_perplexity_2024_2025/Osv60_soc_taxi_2024.xlsx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd

from svod_osv60_core import classify_row

SHEET = "Не_включено_в_свод"
COL = "Внутренняя_группа"


def _article_cell(v: object) -> str | None:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s else None


def fill_sheet(df: pd.DataFrame) -> pd.DataFrame:
    if "Контрагент" not in df.columns:
        raise ValueError("На листе нет колонки «Контрагент»")
    out = df.copy()
    if COL not in out.columns:
        # после «Статья_расходов», иначе в конец
        cols = list(out.columns)
        if "Статья_расходов" in cols:
            i = cols.index("Статья_расходов") + 1
            out.insert(i, COL, "")
        else:
            out[COL] = ""

    arts = out["Статья_расходов"] if "Статья_расходов" in out.columns else pd.Series([None] * len(out))

    def group_for_row(i: int) -> str:
        name = str(out.at[i, "Контрагент"]).strip()
        art = _article_cell(arts.iloc[i]) if "Статья_расходов" in out.columns else None
        r = classify_row(name, art)
        return r[0] if r is not None else ""

    for i in range(len(out)):
        out.at[i, COL] = group_for_row(i)
    return out


def process_file(path: Path) -> None:
    import os
    import shutil

    xl = pd.ExcelFile(path)
    sheets = {s: pd.read_excel(path, sheet_name=s) for s in xl.sheet_names}
    if SHEET not in sheets:
        raise ValueError(f"Нет листа «{SHEET}» в {path}")
    sheets[SHEET] = fill_sheet(sheets[SHEET])
    tmp = path.parent / f"{path.stem}.__tmp_{os.getpid()}.xlsx"
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            for name, sdf in sheets.items():
                sdf.to_excel(writer, sheet_name=name, index=False)
    except Exception:
        if tmp.exists():
            tmp.unlink()
        raise
    try:
        os.replace(tmp, path)
    except OSError:
        # файл открыт в Excel — оставляем рядом готовую копию
        alt = path.with_name(f"{path.stem}_vnutr_gruppa{path.suffix}")
        shutil.move(str(tmp), str(alt))
        print(f"WARN: {path} занят — записано: {alt}")
        return
    print(f"OK: {path} — лист «{SHEET}», колонка «{COL}»")


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    default = [
        root
        / "Perplexity_paket_Taganay_2024_2025"
        / "taganay_perplexity_2024_2025"
        / "Osv60_soc_taxi_2024.xlsx",
        root
        / "Perplexity_paket_Taganay_2024_2025"
        / "taganay_perplexity_2024_2025"
        / "Osv60_soc_taxi_2025.xlsx",
    ]
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--files",
        nargs="*",
        type=Path,
        default=None,
        help="Пути к .xlsx (по умолчанию два файла в Perplexity_paket…)",
    )
    args = p.parse_args()
    files = args.files if args.files else default
    for f in files:
        f = f.resolve() if not f.is_absolute() else f
        if not f.is_file():
            print(f"Пропуск (нет файла): {f}")
            continue
        process_file(f)


if __name__ == "__main__":
    main()
