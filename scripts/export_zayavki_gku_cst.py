"""
Копирует (опционально) файл «Заявки {год}.xls» из папки ГКУ ЦСТ и строит свод:
  - Excel: несколько листов (месяцы, поставщик, контрольные итоги);
  - CSV: помесячный свод (удобно для diff и внешних сверок).

Исходник по умолчанию в репозитории:
  data/gku_cst/{год}/Заявки_{год}.xls

Пример первичной загрузки с диска E::
  python scripts/export_zayavki_gku_cst.py --year 2025 --import-from \"E:\\...\\Заявки 2025.xls\"

Связь с отчётом «Таганай — сводный за год» (PDF): см. docs/sverki/Zayavki_GKU_CST.md
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent.parent
DATA = BASE / "data" / "gku_cst"


def _month_from_row(row: pd.Series) -> str:
    for col in ("Время отправления", "Дата заявки"):
        v = row.get(col)
        if pd.isna(v):
            continue
        try:
            return pd.Timestamp(v).strftime("%Y-%m")
        except (ValueError, TypeError, OSError):
            continue
    return "без_даты"


def load_zayavki(xls: Path) -> pd.DataFrame:
    df = pd.read_excel(xls, sheet_name=0, engine="xlrd")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def build_svod(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = df.copy()
    df["_месяц"] = df.apply(_month_from_row, axis=1)

    po_mes = (
        df.groupby("_месяц", dropna=False)
        .size()
        .reset_index(name="заявок")
        .sort_values("_месяц")
    )

    col_post = "Поставщик"
    if col_post in df.columns:
        po_post = (
            df.groupby(col_post, dropna=False)
            .size()
            .reset_index(name="заявок")
            .sort_values("заявок", ascending=False)
        )
    else:
        po_post = pd.DataFrame(columns=[col_post, "заявок"])

    kontrol = pd.DataFrame(
        [
            ("всего_строк_в_файле", len(df)),
            ("с_временем_отправления", int(df["Время отправления"].notna().sum())),
            ("без_времени_отправления", int(df["Время отправления"].isna().sum())),
            ("уникальных_ФИО", df["ФИО получателя"].dropna().nunique()),
        ],
        columns=["показатель", "значение"],
    )
    return po_mes, po_post, kontrol


def main() -> None:
    p = argparse.ArgumentParser(description="Свод заявок ГКУ ЦСТ из Заявки {год}.xls")
    p.add_argument("--year", type=int, default=2025)
    p.add_argument(
        "--xls",
        type=str,
        default="",
        help="Путь к .xls (по умолчанию data/gku_cst/{year}/Заявки_{year}.xls)",
    )
    p.add_argument(
        "--import-from",
        type=str,
        default="",
        metavar="PATH",
        help="Скопировать указанный .xls в каталог data/gku_cst/{year}/ перед расчётом",
    )
    args = p.parse_args()
    year = args.year
    dest_dir = DATA / str(year)
    dest_dir.mkdir(parents=True, exist_ok=True)
    xls_path = Path(args.xls) if args.xls else dest_dir / f"Заявки_{year}.xls"

    if args.import_from:
        src = Path(args.import_from)
        if not src.is_file():
            sys.exit(f"Нет файла для копирования: {src}")
        shutil.copy2(src, xls_path)
        print(f"Скопировано: {src} -> {xls_path}")

    if not xls_path.is_file():
        sys.exit(
            f"Нет файла {xls_path}. Укажите --import-from или положите Заявки_{year}.xls в {dest_dir}"
        )

    df = load_zayavki(xls_path)
    po_mes, po_post, kontrol = build_svod(df)

    xlsx_out = dest_dir / f"zayavki_{year}_svod.xlsx"
    csv_mes = dest_dir / f"zayavki_{year}_po_mesyacam.csv"
    csv_post = dest_dir / f"zayavki_{year}_po_postavschiku.csv"

    with pd.ExcelWriter(xlsx_out, engine="openpyxl") as w:
        po_mes.to_excel(w, sheet_name="по_месяцам", index=False)
        po_post.to_excel(w, sheet_name="по_поставщику", index=False)
        kontrol.to_excel(w, sheet_name="контроль", index=False)

    po_mes.to_csv(csv_mes, index=False, encoding="utf-8-sig")
    po_post.to_csv(csv_post, index=False, encoding="utf-8-sig")

    print(f"Строк в исходнике: {len(df)}")
    print(f"Excel: {xlsx_out.relative_to(BASE)}")
    print(f"CSV:   {csv_mes.relative_to(BASE)}")
    print(f"       {csv_post.relative_to(BASE)}")


if __name__ == "__main__":
    main()
