"""
Общая логика свода по ОСВ счёта 60: классификация контрагентов, группы для записки, Excel.

Точки входа по годам: `svod_osv60_2024.py`, `svod_osv60_2025.py` (вызывают `run_year`).
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
EXTRACT_DIR = BASE_DIR / "_extract_osv"

SOC_TAXI_REVENUE_SHARE = 0.78

COL_ON_SOC = "На_соцтакси"

# Внутренние группы (детализация)
G_TOP = "Топливо (ГСМ)"
G_ARENDA_AVTO = "Аренда автомобиля"
G_ARENDA_GARAZH = "Аренда гаража/стоянки"
G_REMONT = "Ремонт и обслуживание автомобиля"
G_MOYKA = "Мойка автомобилей"
G_STRAH = "Страхование автомобиля"
G_MED = "Медосмотр водителей"
G_KOMM = "Коммунальные услуги"
G_IT = "Связь, IT и офисные расходы"
G_PROCH = "Прочие эксплуатационные"


def _norm(s: str) -> str:
    return str(s).strip().upper()


BANK_MARKERS = (
    "БАНК",
    "ВТБ",
    "СБЕР",
    "СБЕРБАНК",
    "АЛЬФА-БАНК",
    "АЛЬФАБАНК",
    "ТИНЬКОФФ",
    "ГАЗПРОМБАНК",
    "РОСБАНК",
    "СОВКОМБАНК",
    "ОТП БАНК",
    "АК БАРС",
    "РАЙФФАЙЗЕН",
    "МОСКОМПРИВАТБАНК",
)


def is_bank_or_clearing(name: str) -> bool:
    n = _norm(name)
    return any(marker in n for marker in BANK_MARKERS)


def classify_counterparty(name: str) -> tuple[str, bool] | None:
    """
    Возвращает (внутренняя_группа, прямые_100_процентов) или None,
    если контрагент не входит в перечень для расчёта себестоимости соцтакси.
    """
    n = _norm(name)

    if is_bank_or_clearing(name):
        return None

    if "ЛИКАРД" in n or "ГАЗПРОМНЕФТ" in n or "ГАЗПРОМНЕФТЬ" in n:
        return G_TOP, True
    if "ГАЗПРОМ" in n:
        return G_KOMM, False

    if "ЛОПАКОВА" in n and ("НАТАЛ" in n or "НАТАЛИ" in n):
        return G_ARENDA_AVTO, True
    if "НОРЧАК" in n:
        return G_ARENDA_GARAZH, True

    if "РЕСО" in n:
        return G_STRAH, True

    if "БУДЬ" in n and "ЗДОРОВ" in n:
        return G_MED, True
    if "СИБИРСКОЕ" in n and "ЗДОРОВ" in n:
        return G_MED, True

    if "АЛЬТЕРНАТИВА" in n:
        return G_REMONT, True
    if "КАФФА" in n:
        return G_REMONT, True
    if "КОЛЕСА" in n:
        return G_REMONT, True
    if "МОМЗЕРОВ" in n:
        return G_REMONT, True
    if "ТИМЧЕНКО" in n:
        return G_REMONT, True
    if "ПЕТРАШ" in n:
        return G_REMONT, True

    if "МАМЕДОВ" in n:
        return G_MOYKA, True

    if "СЕВЭНКО" in n or "СЕВЕНКО" in n:
        return G_KOMM, False
    if "ЭК ВОСТОК" in n or "ЭКВОСТОК" in n:
        return G_KOMM, False
    if "ЯМАЛ" in n and "ЭКОЛОГ" in n:
        return G_KOMM, False
    if "СТАТУС" in n and "2" in n:
        return G_KOMM, False

    if n.startswith("КБ ") or " КБ " in n or n.endswith(" КБ"):
        return G_IT, False
    if "РЕГ." in n and "РУ" in n:
        return G_IT, False
    if "Т2" in n or "Т 2" in n:
        return G_IT, False
    if "КАНИНА" in n:
        return G_IT, False
    if "ЯРОШ" in n:
        return G_IT, False

    if "РЫБАЛКО" in n:
        return G_PROCH, False

    return None


def find_osv60_file(year: int) -> Path:
    preferred = EXTRACT_DIR / f"Osv_schet_60_{year}.xls"
    if preferred.is_file():
        return preferred
    if EXTRACT_DIR.is_dir():
        files = sorted(EXTRACT_DIR.glob(f"*60*{year}*.xls"))
        if files:
            return files[0]
    raise FileNotFoundError(
        f"Не найден файл ОСВ 60 за {year} год (*.xls). "
        f"Положите `Osv_schet_60_{year}.xls` в {EXTRACT_DIR} или запустите rename_extract_osv_files.py"
    )


def to_amount(value: object) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    s = str(value).replace("\xa0", "").replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_osv60_rows(path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = pd.read_excel(path, header=None, engine="xlrd")
    included = []
    excluded = []
    for i in range(len(df)):
        name = df.iloc[i, 0]
        if not isinstance(name, str):
            continue
        name = name.strip()
        if not name or name == "60" or name.upper().startswith("ИТОГ"):
            continue

        open_dr = to_amount(df.iloc[i, 1])
        open_cr = to_amount(df.iloc[i, 2])
        turn_dr = to_amount(df.iloc[i, 3])
        turn_cr = to_amount(df.iloc[i, 4])
        close_dr = to_amount(df.iloc[i, 5])
        close_cr = to_amount(df.iloc[i, 6])

        if turn_dr == 0 and turn_cr == 0:
            continue

        base = turn_cr if turn_cr else turn_dr

        classified = classify_counterparty(name)
        if classified is None:
            excluded.append(
                {
                    "Строка": i + 1,
                    "Контрагент": name,
                    "Оборот_Дт": turn_dr,
                    "Оборот_Кт": turn_cr,
                    "База_для_расчёта": base,
                    "Причина_исключения": "банк/расчёты"
                    if is_bank_or_clearing(name)
                    else "нет правила классификации для соцтакси",
                }
            )
            continue

        grp, direct = classified
        on_soc = base if direct else base * SOC_TAXI_REVENUE_SHARE

        included.append(
            {
                "Строка": i + 1,
                "Контрагент": name,
                "Сальдо_Дт_нач": open_dr,
                "Сальдо_Кт_нач": open_cr,
                "Оборот_Дт": turn_dr,
                "Оборот_Кт": turn_cr,
                "Сальдо_Дт_кон": close_dr,
                "Сальдо_Кт_кон": close_cr,
                "База_для_расчёта": base,
                "Внутренняя_группа": grp,
                "Прямые_100pct": direct,
                COL_ON_SOC: on_soc,
            }
        )
    return pd.DataFrame(included), pd.DataFrame(excluded)


def build_zapiska_table(detail: pd.DataFrame, year: int) -> pd.DataFrame:
    def sum_groups(labels: list[str]) -> float:
        m = detail["Внутренняя_группа"].isin(labels)
        return float(detail.loc[m, COL_ON_SOC].sum())

    z1 = sum_groups([G_ARENDA_AVTO, G_ARENDA_GARAZH])
    z2 = sum_groups([G_TOP])
    z3 = sum_groups([G_REMONT])
    z4 = sum_groups([G_STRAH])
    z5 = sum_groups([G_MED])
    z6 = sum_groups([G_MOYKA])
    z7 = sum_groups([G_KOMM, G_IT, G_PROCH])

    data = [
        ("1. Аренда специализированного транспортного средства и гаража", z1),
        ("2. Расходы на топливо (ГСМ)", z2),
        ("3. Текущий ремонт, обслуживание и автотовары", z3),
        ("4. Страхование транспортных средств (ОСАГО/иные виды)", z4),
        ("5. Медицинские осмотры водителей", z5),
        ("6. Мойка автомобилей и обеспечение санитарного состояния", z6),
        (
            "7. Прочие эксплуатационные расходы, отнесённые на социальное такси пропорционально доле выручки",
            z7,
        ),
    ]
    out = pd.DataFrame(data, columns=["Группа затрат (пояснительная записка)", "Сумма на соцтакси, руб."])
    total = float(out["Сумма на соцтакси, руб."].sum())
    out = pd.concat(
        [
            out,
            pd.DataFrame(
                [
                    {
                        "Группа затрат (пояснительная записка)": (
                            f"ИТОГО прочие расходы по счёту 60 на соцтакси за {year} год"
                        ),
                        "Сумма на соцтакси, руб.": total,
                    }
                ]
            ),
        ],
        ignore_index=True,
    )
    return out


def run_year(year: int) -> Path:
    """Строит `Osv60_soc_taxi_{year}.xlsx` из ОСВ 60 за указанный год."""
    src = find_osv60_file(year)
    out_file = BASE_DIR / f"Osv60_soc_taxi_{year}.xlsx"

    detail, skipped = load_osv60_rows(src)
    zapiska = build_zapiska_table(detail, year)

    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="Строки_ОСВ60", index=False)
        skipped.to_excel(writer, sheet_name="Не_включено_в_свод", index=False)
        zapiska.to_excel(writer, sheet_name="Группа_затрат_записка", index=False)

    print("Источник:", src)
    print("Создан файл:", out_file)
    print("Строк включено:", len(detail), "исключено:", len(skipped))
    print("Итого по счёту 60 на соцтакси:", round(float(zapiska.iloc[-1, 1]), 2), "руб.")
    return out_file
