"""
Пересчёт «ЗП из учёта» по ОСВ 70 с очисткой и применение той же методики долей,
что в `svod_personal_core` (группы затрат, k_соцтакси для админов).

Счёт **76** подгружается **только для контроля** (как в `svod_reconcile_osv76.py`):
суммы по 76 **не суммируются** с базой 70 — иначе задвоение с выплатами/начислениями по 70.

База для умножения на долю (по умолчанию):
  --base kt  — оборот **Кредит** по 70 за период (типично движение начислений зарплаты);
  --base dt  — оборот **Дебет** (выплаты/удержания с 70);
  --base max — max(Дт, Кт) по строке после агрегации по ФИО.

«Технические» строки: контрагенты-организации (`ORG_MARKERS` из svod_reconcile_osv76),
служебные подписи, строки без признака ФЛО (`is_person_name`).

Примеры:
  python scripts/recalc_fot_osv70_osv76.py
  python scripts/recalc_fot_osv70_osv76.py --year 2025 \\
    --osv70 _extract_osv/Osv_schet_70_2025_vernaya_1c.xls \\
    --osv76 _extract_osv/Osv_schet_76_2025.xls

Имя `Fot_osv70_soc_taxi_2024.xlsx` не задаёт год источника: год задают только `--year`
и фактические пути ОСВ. При несоответствии (например, --year 2024 и файл ..._70_2025...)
скрипт завершится с ошибкой; обход — только `--allow-year-mismatch`.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from svod_personal_core import (  # noqa: E402
    COST_GROUP_ORDER,
    get_cost_group,
    get_participation,
    is_person_name,
)
from svod_reconcile_osv76 import (  # noqa: E402
    is_org_or_special,
    norm_fio,
)


def _to_amount(value: object) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _read_raw(path: Path) -> pd.DataFrame:
    kw: dict = {"header": None, "sheet_name": 0}
    if path.suffix.lower() == ".xls":
        kw["engine"] = "xlrd"
    else:
        kw["engine"] = "openpyxl"
    return pd.read_excel(path, **kw)


def load_osv_account_rows(path: Path, account_label: str) -> pd.DataFrame:
    """Строки ОСВ 70/76: колонки 1–6 — сальдо/обороты как в выгрузке 1С (см. svod_reconcile_osv76)."""
    df = _read_raw(path)
    acc = str(account_label).strip()
    rows: list[dict] = []
    for i in range(len(df)):
        name = df.iloc[i, 0]
        if not isinstance(name, str):
            continue
        name = name.strip()
        if not name or name.upper().startswith("ИТОГ"):
            continue
        if name == acc:
            continue

        sd_dt = _to_amount(df.iloc[i, 1])
        sd_kt = _to_amount(df.iloc[i, 2])
        to_dt = _to_amount(df.iloc[i, 3])
        to_kt = _to_amount(df.iloc[i, 4])
        ed_dt = _to_amount(df.iloc[i, 5])
        ed_kt = _to_amount(df.iloc[i, 6]) if df.shape[1] > 6 else 0.0

        if to_dt == 0 and to_kt == 0 and sd_dt == 0 and sd_kt == 0 and ed_dt == 0 and ed_kt == 0:
            continue

        rows.append(
            {
                "Строка_исходник": i + 1,
                "Контрагент": name,
                "Сальдо_н_Дт": sd_dt,
                "Сальдо_н_Кт": sd_kt,
                "Оборот_Дт": to_dt,
                "Оборот_Кт": to_kt,
                "Сальдо_к_Дт": ed_dt,
                "Сальдо_к_Кт": ed_kt,
            }
        )
    return pd.DataFrame(rows)


def classify_row(name: str, year: int) -> tuple[str, str]:
    """Возвращает (класс, причина)."""
    if is_org_or_special(name):
        return "организация_или_служебная", "ORG_MARKERS или служебная строка"
    if not is_person_name(name, year):
        return "не_ФЛ_в_реестре", "не похоже на ФИО сотрудника"
    return "ФЛ", ""


def pick_base(row: pd.Series, mode: str) -> float:
    dt, kt = float(row["Оборот_Дт"]), float(row["Оборот_Кт"])
    if mode == "kt":
        return kt
    if mode == "dt":
        return dt
    if mode == "max":
        return max(dt, kt)
    raise ValueError(mode)


def default_osv70(year: int) -> Path:
    base = BASE_DIR / "_extract_osv"
    if year != 2025:
        return base / f"Osv_schet_70_{year}.xls"
    candidates = [base / "Osv_schet_70_2025.xls", base / "Osv_schet_70_2025_vernaya_1c.xls"]
    best: Path | None = None
    best_rows = -1
    for c in candidates:
        if not c.is_file():
            continue
        try:
            n = len(_read_raw(c))
        except Exception:
            continue
        if n > best_rows:
            best_rows = n
            best = c
    return best if best is not None else base / "Osv_schet_70_2025.xls"


def default_osv76(year: int) -> Path:
    return BASE_DIR / "_extract_osv" / f"Osv_schet_76_{year}.xls"


def infer_year_from_osv_filename(path: Path) -> int | None:
    """
    Извлекает год из типового имени `Osv_schet_70_2024.xls` / `Osv_schet_76_2025.xls`.
    Если встречается несколько разных календарных лет — None (неоднозначно).
    """
    name = path.name
    m = re.search(r"_(?:70|76)_(\d{4})\b", name)
    if m:
        return int(m.group(1))
    m = re.search(r"за\s*(\d{4})\s*г", name, flags=re.I)
    if m:
        return int(m.group(1))
    years = sorted({int(x) for x in re.findall(r"\b(20[012][0-9]{2})\b", name)})
    if len(years) == 1:
        return years[0]
    return None


def assert_osv_year_matches(
    year: int, path: Path, label: str, allow_mismatch: bool
) -> tuple[int | None, str]:
    """Проверяет, что имя файла ОСВ не относится к другому календарному году."""
    inferred = infer_year_from_osv_filename(path)
    if inferred is None:
        return inferred, "год из имени не распознан — проверьте вручную"
    if inferred != year and not allow_mismatch:
        raise SystemExit(
            f"Ошибка: --year {year}, но {label} указывает на файл «{path.name}» "
            f"(по имени это {inferred} год). Исправьте путь или переименуйте выгрузку; "
            f"для осознанного исключения используйте --allow-year-mismatch."
        )
    if inferred != year:
        return inferred, f"ВНИМАНИЕ: год в имени {inferred}, а --year {year} (разрешено флагом)"
    return inferred, "OK"


def main() -> None:
    ap = argparse.ArgumentParser(description="ЗП из ОСВ 70 + очистка + доли соцтакси; ОСВ 76 — контроль")
    ap.add_argument("--year", type=int, default=2024, help="Год (пути к ОСВ по умолчанию)")
    ap.add_argument("--osv70", type=Path, default=None, help="ОСВ по счёту 70 (.xls/.xlsx)")
    ap.add_argument("--osv76", type=Path, default=None, help="ОСВ по счёту 76 (опционально)")
    ap.add_argument(
        "--base",
        choices=("kt", "dt", "max"),
        default="kt",
        help="Какой оборот по 70 брать как базу для долей (по умолчанию Кт — начисления)",
    )
    ap.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Выходной .xlsx (по умолчанию Fot_osv70_soc_taxi_{year}.xlsx в корне)",
    )
    ap.add_argument(
        "--allow-year-mismatch",
        action="store_true",
        help="Не прерывать расчёт, если год в имени ОСВ не совпадает с --year (только если осознанно)",
    )
    args = ap.parse_args()
    year = args.year

    osv70 = args.osv70
    if osv70 is None:
        osv70 = default_osv70(year)
    elif not osv70.is_absolute():
        osv70 = BASE_DIR / osv70

    osv76 = args.osv76
    if osv76 is None:
        osv76 = default_osv76(year)
    elif not osv76.is_absolute():
        osv76 = BASE_DIR / osv76

    out = args.out
    if out is None:
        out = BASE_DIR / f"Fot_osv70_soc_taxi_{year}.xlsx"
    elif not out.is_absolute():
        out = BASE_DIR / out

    if not osv70.is_file():
        raise SystemExit(f"Нет файла ОСВ 70: {osv70}")

    inf70, check70 = assert_osv_year_matches(year, osv70, "ОСВ 70", args.allow_year_mismatch)
    inf76: int | None = None
    check76 = ""
    if osv76.is_file():
        inf76, check76 = assert_osv_year_matches(year, osv76, "ОСВ 76", args.allow_year_mismatch)

    raw70 = load_osv_account_rows(osv70, "70")
    raw70["Счёт"] = "70"
    raw70["Класс"] = raw70["Контрагент"].apply(lambda n: classify_row(n, year)[0])
    raw70["Класс_примечание"] = raw70["Контрагент"].apply(lambda n: classify_row(n, year)[1])

    fl70 = raw70[raw70["Класс"] == "ФЛ"].copy()
    fl70["ФИО_норм"] = fl70["Контрагент"].map(norm_fio)
    if len(fl70) < 8:
        print(
            "Предупреждение: после фильтрации мало строк ФЛ в ОСВ 70 (",
            len(fl70),
            "). Возможно, неполная выгрузка — укажите полный файл через --osv70.",
        )
    agg = (
        fl70.groupby("ФИО_норм", as_index=False)
        .agg(
            {
                "Контрагент": "first",
                "Оборот_Дт": "sum",
                "Оборот_Кт": "sum",
                "Сальдо_н_Дт": "sum",
                "Сальдо_н_Кт": "sum",
                "Сальдо_к_Дт": "sum",
                "Сальдо_к_Кт": "sum",
            }
        )
    )
    agg["База_для_долей"] = agg.apply(lambda r: pick_base(r, args.base), axis=1)

    parts: list[tuple[str, float, bool]] = []
    groups: list[str] = []
    shares: list[float] = []
    on_soc: list[float] = []

    for _, row in agg.iterrows():
        fio = row["Контрагент"]
        p = get_participation(fio, year)
        g = get_cost_group(fio)
        base = float(row["База_для_долей"])
        share = p[1]
        parts.append(p)
        groups.append(g)
        shares.append(share)
        on_soc.append(base * share)

    agg["Тип_доли"] = [p[0] for p in parts]
    agg["Доля_участия"] = shares
    agg["apply_k"] = [p[2] for p in parts]
    agg["Группа_затрат"] = groups
    agg[f"На_соцтакси_{year}_только_70база"] = on_soc

    order_map = {g: i for i, g in enumerate(COST_GROUP_ORDER)}
    svod = agg.groupby("Группа_затрат", as_index=False)[f"На_соцтакси_{year}_только_70база"].sum()
    svod["__ord"] = svod["Группа_затрат"].map(order_map).fillna(99)
    svod = svod.sort_values("__ord").drop(columns="__ord")
    itog = float(svod[f"На_соцтакси_{year}_только_70база"].sum())
    svod = pd.concat(
        [
            svod,
            pd.DataFrame(
                [{"Группа_затрат": "ИТОГО (ОСВ70 * доля, без взносов)", f"На_соцтакси_{year}_только_70база": itog}]
            ),
        ],
        ignore_index=True,
    )

    raw76 = pd.DataFrame()
    if osv76.is_file():
        raw76 = load_osv_account_rows(osv76, "76")
        raw76["Счёт"] = "76"
        raw76["Класс"] = raw76["Контрагент"].apply(lambda n: classify_row(n, year)[0])
        raw76["Класс_примечание"] = raw76["Контрагент"].apply(lambda n: classify_row(n, year)[1])
    else:
        print("Предупреждение: нет файла ОСВ 76, лист контроля пропущен:", osv76)

    meta = pd.DataFrame(
        [
            ("логический_год_отчета_--year", year),
            ("osv70_путь", str(osv70)),
            ("osv70_год_из_имени_файла", "" if inf70 is None else str(inf70)),
            ("osv70_проверка_года", check70),
            ("osv76_путь", str(osv76) if osv76.is_file() else ""),
            ("osv76_год_из_имени_файла", "" if inf76 is None else str(inf76)),
            ("osv76_проверка_года", check76),
            ("base_mode", args.base),
            (
                "пояснение",
                "На_соцтакси — только Оборот * доля; взносы из ОСВ здесь не добавляются. "
                "Имя выходного файла не определяет год: источник — только пути osv70/osv76.",
            ),
            ("счёт76", "Не суммируется с 70 (протокол: сверка ФЛ, не себестоимость)."),
        ],
        columns=["ключ", "значение"],
    )

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        meta.to_excel(w, sheet_name="метаданные", index=False)
        raw70.to_excel(w, sheet_name="ОСВ70_все_строки", index=False)
        agg.to_excel(w, sheet_name="ОСВ70_ФЛ_агрегат", index=False)
        svod.to_excel(w, sheet_name="свод_по_группам", index=False)
        if not raw76.empty:
            raw76.to_excel(w, sheet_name="ОСВ76_контроль", index=False)
            fl76 = raw76[raw76["Класс"] == "ФЛ"].copy()
            fl76["ФИО_норм"] = fl76["Контрагент"].map(norm_fio)
            fl76a = (
                fl76.groupby("ФИО_норм", as_index=False)
                .agg(
                    {
                        "Контрагент": "first",
                        "Оборот_Дт": "sum",
                        "Оборот_Кт": "sum",
                    }
                )
                .rename(columns={"Оборот_Дт": "Оборот_Дт_76", "Оборот_Кт": "Оборот_Кт_76", "Контрагент": "ФИО_из_76"})
            )
            agg_m = agg[["ФИО_норм", "Контрагент", "База_для_долей", f"На_соцтакси_{year}_только_70база"]].rename(
                columns={"Контрагент": "ФИО_из_70"}
            )
            m = fl76a.merge(agg_m, on="ФИО_норм", how="outer")
            m.to_excel(w, sheet_name="сверка_ФЛ_70_76", index=False)

    print("ОСВ 70:", osv70)
    print("ОСВ 76:", osv76 if osv76.is_file() else "(нет)")
    print("База:", args.base)
    print("ИТОГО На_соцтакси (база 70 * доля):", f"{itog:,.2f}")
    print("Создан:", out)


if __name__ == "__main__":
    main()
