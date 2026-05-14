"""
Сравнение сводов по счёту 60 за два года (лист «Группа_затрат_записка» в Osv60_soc_taxi_*.xlsx).

Пример:
  python compare_osv60.py
  python compare_osv60.py --y1 2024 --y2 2025 --out Svod_60_2024_2025.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent


def load_zapiska(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Группа_затрат_записка", engine="openpyxl")
    if df.shape[1] < 2:
        raise SystemExit(f"Неожиданный формат: {path}")
    c0, c1 = df.columns[0], df.columns[1]
    out = df[[c0, c1]].copy()
    out.columns = ["Группа_затрат", "Сумма"]
    out = out[out["Группа_затрат"].astype(str).str.strip() != ""]
    out = out[~out["Группа_затрат"].astype(str).str.upper().str.startswith("ИТОГО")]
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Сравнение Osv60_soc_taxi по годам")
    ap.add_argument("--y1", type=int, default=2024)
    ap.add_argument("--y2", type=int, default=2025)
    ap.add_argument(
        "--file1",
        type=Path,
        default=None,
        help="Первый xlsx (по умолчанию Osv60_soc_taxi_{y1}.xlsx)",
    )
    ap.add_argument("--file2", type=Path, default=None)
    ap.add_argument("--out", type=Path, default=BASE_DIR / "Svod_60_2024_2025.xlsx")
    args = ap.parse_args()

    f1 = args.file1 or (BASE_DIR / f"Osv60_soc_taxi_{args.y1}.xlsx")
    f2 = args.file2 or (BASE_DIR / f"Osv60_soc_taxi_{args.y2}.xlsx")
    if not f1.is_file() or not f2.is_file():
        raise SystemExit(f"Нужны файлы: {f1} и {f2}")

    a = load_zapiska(f1)
    b = load_zapiska(f2)
    a = a.rename(columns={"Сумма": f"Сумма_{args.y1}"})
    b = b.rename(columns={"Сумма": f"Сумма_{args.y2}"})

    m = pd.merge(a, b, on="Группа_затрат", how="outer").fillna(0.0)
    s1, s2 = f"Сумма_{args.y1}", f"Сумма_{args.y2}"
    m["Изменение_руб"] = m[s2] - m[s1]
    m["Изменение_%"] = m.apply(
        lambda r: (r["Изменение_руб"] / r[s1] * 100.0) if abs(r[s1]) > 1e-6 else float("nan"),
        axis=1,
    )

    out_path = args.out if args.out.is_absolute() else BASE_DIR / args.out
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        m.to_excel(w, sheet_name="Сравнение_60", index=False)
    print("Создан:", out_path)


if __name__ == "__main__":
    main()
