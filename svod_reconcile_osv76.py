"""
Сверка физлиц из ОСВ по счёту 76 с листом «Персонал_2024_soc_taxi».

Ищет ФИО из колонки контрагента, сопоставляет с колонкой «ФИО» в Personal_2024_soc_taxi.xlsx
и проставляет: «уже в ФОТ» / «вне ФОТ» / «не ФЛ (организация)».

Исходник ОСВ:
  - по умолчанию: _extract_osv/Osv_schet_76_2024.xls (после rename_extract_osv_files.py);
  - или ваш .xlsx с тем же порядком колонок, как в 1С (A — наименование, обороты в D/E),
    плюс опционально последний столбец с ролью («водитель», «соц. работник» и т.п.).

Выход: Osv76_sverka_2024.xlsx

Запуск:
  python svod_reconcile_osv76.py
  python svod_reconcile_osv76.py --osv "C:\\path\\OSV_76_moi.xlsx" --sheet Лист1
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
EXTRACT_DIR = BASE_DIR / "_extract_osv"
PERSONAL_DEFAULT = BASE_DIR / "Personal_2024_soc_taxi.xlsx"
OSV76_DEFAULT = EXTRACT_DIR / "Osv_schet_76_2024.xls"
OUT_FILE = BASE_DIR / "Osv76_sverka_2024.xlsx"

ORG_MARKERS = (
    "ООО",
    "ПАО",
    "АО ",
    " АО",
    "САО",
    "ОПФР",
    "ИП ",
    " ИП",
    "РЕСО",
    "ГАЗПРОМ",
    "РОСНЕФТЬ",
    "РОССЕТИ",
    "БАНК",
)


def norm_fio(s: str) -> str:
    s = str(s).strip().upper().replace("Ё", "Е")
    s = re.sub(r"\s+", " ", s)
    return s


def is_org_or_special(name: str) -> bool:
    n = norm_fio(name)
    if not n or n == "76" or n.startswith("ИТОГ"):
        return True
    return any(m in n for m in ORG_MARKERS)


def load_personal_fio(path: Path, sheet: str) -> set[str]:
    if not path.is_file():
        raise FileNotFoundError(f"Нет файла персонала: {path}")
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    if "ФИО" not in df.columns:
        raise SystemExit(f"В листе «{sheet}» нет колонки «ФИО»")
    return {norm_fio(x) for x in df["ФИО"].dropna().astype(str) if str(x).strip()}


def _to_amount(value: object) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _detect_role_column(df: pd.DataFrame) -> int | None:
    """Столбец с ролью: заголовок «Роль» или значения водитель/соц./страховка в теле таблицы."""
    role_cell = re.compile(
        r"водител|соц\.\s*работ|соцработ|страховк|роль",
        flags=re.IGNORECASE,
    )
    # строка заголовка «Роль»
    for r in range(min(12, len(df))):
        for c in range(df.shape[1]):
            v = df.iloc[r, c]
            if isinstance(v, str) and "роль" in v.lower():
                return c
    # в теле: считаем попадания по столбцам в строках 8–40
    best_c, best_score = None, 0
    start = min(8, len(df) - 1)
    end = min(45, len(df))
    for c in range(1, df.shape[1]):
        score = 0
        for r in range(start, end):
            v = df.iloc[r, c]
            if isinstance(v, str) and role_cell.search(v):
                score += 1
        if score > best_score:
            best_c, best_score = c, score
    if best_score >= 2:
        return best_c
    if df.shape[1] > 7:
        return df.shape[1] - 1
    return None


def load_osv76_rows(path: Path, sheet: str | int | None) -> pd.DataFrame:
    kw: dict = {"header": None}
    if path.suffix.lower() == ".xls":
        kw["engine"] = "xlrd"
    else:
        kw["engine"] = "openpyxl"
    if sheet is not None:
        kw["sheet_name"] = sheet
    df = pd.read_excel(path, **kw)

    role_col = _detect_role_column(df)

    rows = []
    for i in range(len(df)):
        name = df.iloc[i, 0]
        if not isinstance(name, str):
            continue
        name = name.strip()
        if not name or name == "76" or name.upper().startswith("ИТОГ"):
            continue

        turn_dr = _to_amount(df.iloc[i, 3])
        turn_cr = _to_amount(df.iloc[i, 4])
        if turn_dr == 0 and turn_cr == 0:
            continue

        role = ""
        if role_col is not None and role_col < df.shape[1]:
            rv = df.iloc[i, role_col]
            if isinstance(rv, str):
                role = rv.strip()

        rows.append(
            {
                "Строка_исходник": i + 1,
                "Контрагент": name,
                "Оборот_Дт": turn_dr,
                "Оборот_Кт": turn_cr,
                "Роль_из_файла": role,
            }
        )
    return pd.DataFrame(rows)


def match_status(name: str, personal: set[str]) -> str:
    if is_org_or_special(name):
        return "не ФЛ (организация / служебная строка)"
    n = norm_fio(name)
    if n in personal:
        return "уже в ФОТ"
    # частичное: все слова из ОСВ входят в одну строку персонала
    words = [w for w in n.split() if len(w) > 1]
    if len(words) >= 2:
        for p in personal:
            if all(w in p for w in words):
                return "уже в ФОТ (совпадение по ФИО, проверить вручную)"
    return "вне ФОТ"


def main() -> None:
    ap = argparse.ArgumentParser(description="Сверка ОСВ-76 с Personal_2024_soc_taxi")
    ap.add_argument("--personal", type=Path, default=PERSONAL_DEFAULT, help="Файл с персоналом")
    ap.add_argument(
        "--osv",
        type=Path,
        default=OSV76_DEFAULT if OSV76_DEFAULT.is_file() else None,
        help="ОСВ 76 (.xls / .xlsx). По умолчанию _extract_osv/Osv_schet_76_2024.xls",
    )
    ap.add_argument("--sheet-personal", default="Персонал_2024_soc_taxi")
    ap.add_argument("--sheet-osv", default=None, help="Лист Excel с ОСВ (если не первый)")
    args = ap.parse_args()

    if args.osv is None or not args.osv.is_file():
        raise SystemExit(
            "Укажите путь к ОСВ 76: --osv \"...\" "
            f"(ожидается после переименования: {OSV76_DEFAULT})"
        )

    personal = load_personal_fio(args.personal, args.sheet_personal)
    detail = load_osv76_rows(args.osv, args.sheet_osv)

    detail["Статус_сверки_с_ФОТ"] = detail["Контрагент"].apply(lambda x: match_status(x, personal))

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as w:
        detail.to_excel(w, sheet_name="Сверка_76", index=False)
        pd.DataFrame(sorted(personal), columns=["ФИО_нормализовано_в_персонале"]).to_excel(
            w, sheet_name="Справочник_ФИО", index=False
        )

    print("ОСВ:", args.osv)
    print("Персонал:", args.personal)
    print("Создан:", OUT_FILE)
    print(detail["Статус_сверки_с_ФОТ"].value_counts().to_string())


if __name__ == "__main__":
    main()
