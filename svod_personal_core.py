"""
Общая логика свода ФОТ + взносы + доли соцтакси по году (2024, 2025, …).

Вызывается из `svod_2024.py` / `svod_2025.py`. Исходники по умолчанию:
  Svodnyi-po-ZP-za-{year}-god.xlsx
  Nalogi-i-vznosy*.xlsx (для 2025 при наличии — сначала Nalogi-i-vznosy*2025*.xlsx)
"""

from __future__ import annotations

import os
from pathlib import Path

import pandas as pd

from config_allocation import get_soc_taxi_share

BASE_DIR = Path(__file__).resolve().parent

GROUP_DRIVERS_SOCIAL = "1. ФОТ водителей и соцработников"
GROUP_DISPATCHER = "2. ФОТ диспетчера"
GROUP_REPAIR = "3. ФОТ ремонтного персонала (Егоров, Шакалов, Родионов)"
GROUP_SOC_SERVICE_HEAD = "4. ФОТ руководителя соцслужбы (Лопакова Н.Ф.)"
GROUP_ORG_HEAD = "5. ФОТ руководителя организации (50% Богдановского)"
GROUP_ADMIN = "6. ФОТ бухгалтерии, IT и админ-персонала"
GROUP_OTHER = "0%. Другая деятельность (Роснефть/прочее)"
GROUP_TOTAL = "ИТОГО трудовые затраты по соцтакси"

COST_GROUP_RULES: dict[str, str] = {
    "Цуканов": GROUP_DRIVERS_SOCIAL,
    "Грачев": GROUP_DRIVERS_SOCIAL,
    "Дмитриев": GROUP_DRIVERS_SOCIAL,
    "Котелевский": GROUP_DRIVERS_SOCIAL,
    "Кульбатырова": GROUP_DISPATCHER,
    "Егоров": GROUP_REPAIR,
    "Родионов": GROUP_REPAIR,
    "Шакалов": GROUP_REPAIR,
    "Лопакова Наталия": GROUP_SOC_SERVICE_HEAD,
    "Богдановский": GROUP_ORG_HEAD,
    "Поплаухина": GROUP_ADMIN,
    "Голоушкин": GROUP_ADMIN,
    "Иванова": GROUP_ADMIN,
    "Исламова": GROUP_ADMIN,
    "Лопакова Светлана": GROUP_ADMIN,
}

COST_GROUP_ORDER = [
    GROUP_DRIVERS_SOCIAL,
    GROUP_DISPATCHER,
    GROUP_REPAIR,
    GROUP_SOC_SERVICE_HEAD,
    GROUP_ORG_HEAD,
    GROUP_ADMIN,
    GROUP_OTHER,
]


def _participation_rules_for_year(year: int) -> dict[str, tuple[str, float, bool]]:
    k = get_soc_taxi_share(year)
    return {
        "Цуканов": ("Прямые", 1.0, False),
        "Грачев": ("Прямые", 1.0, False),
        "Дмитриев": ("Прямые", 1.0, False),
        "Кульбатырова": ("Прямые", 1.0, False),
        "Егоров": ("Прямые", 1.0, False),
        "Родионов": ("Прямые", 1.0, False),
        "Шакалов": ("Прямые", 1.0, False),
        "Котелевский": ("Частично", 0.1, False),
        "Лопакова Наталия": ("Адм. (соцтакси)", 0.5, False),
        "Богдановский": ("Адм. общая", 0.5, True),
        "Поплаухина": ("Адм.", k, False),
        "Голоушкин": ("Адм.", k, False),
        "Иванова": ("Адм.", k, False),
        "Исламова": ("Адм.", k, False),
        "Лопакова Светлана": ("Адм.", k, False),
        "Зубцов": ("0%", 0.0, False),
        "Капков": ("0%", 0.0, False),
        "Першаков": ("0%", 0.0, False),
    }


def paths_for_year(year: int) -> dict[str, Path]:
    zp = BASE_DIR / f"Svodnyi-po-ZP-za-{year}-god.xlsx"
    return {
        "personal_file": zp,
        "personal_out": BASE_DIR / f"Personal_{year}_auto.xlsx",
        "vznosy_out": BASE_DIR / f"Vznosy_{year}_auto.xlsx",
        "itog_out": BASE_DIR / f"Itog_personal_vznosy_{year}.xlsx",
        "soc_taxi_out": BASE_DIR / f"Personal_{year}_soc_taxi.xlsx",
    }


def vznosy_files_for_year(year: int) -> list[Path]:
    override = os.environ.get("VZNOSY_GLOB")
    if override:
        return sorted(BASE_DIR.glob(override))
    if year != 2024:
        y_files = sorted(BASE_DIR.glob(f"Nalogi-i-vznosy*{year}*.xlsx"))
        if y_files:
            return y_files
    return sorted(BASE_DIR.glob("Nalogi-i-vznosy*.xlsx"))


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0)


def get_participation(fio: object, year: int) -> tuple[str, float, bool]:
    text = str(fio)
    rules = _participation_rules_for_year(year)
    for marker, rule in rules.items():
        if marker in text:
            return rule
    return "Другая деятельность (Роснефть/прочее)", 0.0, False


def get_cost_group(fio: object) -> str:
    text = str(fio)
    for marker, group in COST_GROUP_RULES.items():
        if marker in text:
            return group
    return GROUP_OTHER


def is_person_name(value: object, year: int) -> bool:
    text = str(value).strip()
    if not text or text in {"nan", "Сотрудник"}:
        return False
    stop_words = [
        str(year),
        "Налоги",
        "Период",
        "Организация",
        "Начислено",
        "НДФЛ",
        "Единый",
        "Итого",
        "Список",
    ]
    if any(word in text for word in stop_words):
        return False
    if "РКООИ" in text or "ТАГАНАЙ" in text.upper():
        return False
    return len(text.split()) >= 3


def extract_tax_sheet(df: pd.DataFrame, year: int) -> pd.DataFrame | None:
    df = df.copy()
    if df.empty or df.shape[1] < 5:
        return None
    text = " ".join(str(value) for value in df.to_numpy().ravel() if pd.notna(value))
    if "Налоги и взносы" not in text:
        return None
    part = df.iloc[:, [0, 1, 3, 4]].copy()
    part.columns = ["ФИО", "Начислено_по_взносам", "Единый_тариф", "Несч_случаи"]
    part = part[part["ФИО"].apply(lambda v: is_person_name(v, year))]
    part["Начислено_по_взносам"] = to_number(part["Начислено_по_взносам"])
    part["Единый_тариф"] = to_number(part["Единый_тариф"])
    part["Несч_случаи"] = to_number(part["Несч_случаи"])
    v_col = f"Взносы_{year}"
    part[v_col] = part["Единый_тариф"] + part["Несч_случаи"]
    return part[["ФИО", "Начислено_по_взносам", v_col]]


def read_tax_sheets(file: Path, year: int) -> list[pd.DataFrame]:
    all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    frames: list[pd.DataFrame] = []
    for sheet_name, df in all_sheets.items():
        part = extract_tax_sheet(df, year)
        if part is None:
            continue
        print(f"Найден лист с налогами и взносами: {file.name} / {sheet_name}")
        frames.append(part)
    return frames


def build_personal_summary(year: int, paths: dict[str, Path]) -> pd.DataFrame:
    personal_file = paths["personal_file"]
    personal_out = paths["personal_out"]
    if not personal_file.exists():
        raise FileNotFoundError(f"Не найден файл с зарплатой: {personal_file}")

    fio_col = "Фамилия, имя, отчество"
    sum_col = "Начислено всего"
    fot_col = f"ФОТ_{year}"

    all_sheets = pd.read_excel(personal_file, sheet_name=None)
    frames: list[pd.DataFrame] = []

    for sheet_name, df in all_sheets.items():
        df = normalize_columns(df)
        if fio_col not in df.columns or sum_col not in df.columns:
            print(f"Пропущен лист зарплаты без нужных колонок: {sheet_name}")
            continue
        part = df[[fio_col, sum_col]].copy()
        part.columns = ["ФИО", fot_col]
        part[fot_col] = to_number(part[fot_col])
        frames.append(part)

    if not frames:
        print("Обычные колонки зарплаты не найдены, пробую взять ФОТ из листов налогов.")
        frames = [
            frame[["ФИО", "Начислено_по_взносам"]].rename(columns={"Начислено_по_взносам": fot_col})
            for frame in read_tax_sheets(personal_file, year)
        ]

    if not frames:
        raise SystemExit("Не найдено ни одного листа с ФИО и начислениями.")

    full = pd.concat(frames, ignore_index=True)
    full = full.dropna(subset=["ФИО"])
    full = full[full[fot_col].notna()]
    summary = full.groupby("ФИО", as_index=False)[fot_col].sum()
    summary.to_excel(personal_out, index=False)
    return summary


def build_vznosy_summary(year: int, paths: dict[str, Path]) -> pd.DataFrame:
    vznosy_out = paths["vznosy_out"]
    v_col = f"Взносы_{year}"
    files = vznosy_files_for_year(year)
    if not files:
        raise FileNotFoundError(
            f"Не найдены файлы взносов (год {year}). Положите Nalogi-i-vznosy*.xlsx "
            "или задайте VZNOSY_GLOB."
        )
    frames: list[pd.DataFrame] = []
    for file in files:
        frames.extend(read_tax_sheets(file, year))
    if not frames:
        raise SystemExit("Не найдено ни одного файла взносов с нужными колонками.")
    full = pd.concat(frames, ignore_index=True)
    summary = full.groupby("ФИО", as_index=False).agg(
        {"Начислено_по_взносам": "sum", v_col: "sum"}
    )
    summary.to_excel(vznosy_out, index=False)
    return summary


def build_total_summary(
    personal: pd.DataFrame, vznosy: pd.DataFrame, year: int, paths: dict[str, Path]
) -> pd.DataFrame:
    fot_col = f"ФОТ_{year}"
    v_col = f"Взносы_{year}"
    total = personal.merge(vznosy[["ФИО", v_col]], on="ФИО", how="outer")
    total[fot_col] = total[fot_col].fillna(0)
    total[v_col] = total[v_col].fillna(0)
    total["Итого_ФОТ_плюс_взносы"] = total[fot_col] + total[v_col]
    total = total.sort_values("ФИО")
    total.to_excel(paths["itog_out"], index=False)
    return total


def build_soc_taxi_summary(total: pd.DataFrame, year: int, paths: dict[str, Path]) -> pd.DataFrame:
    fot_col = f"ФОТ_{year}"
    v_col = f"Взносы_{year}"
    on_col = f"На_соцтакси_{year}"
    k = get_soc_taxi_share(year)

    result = total.copy()
    participation = result["ФИО"].apply(lambda f: get_participation(f, year))
    result["Категория участия"] = participation.apply(lambda item: item[0])
    result["Доля участия в соцтакси"] = participation.apply(lambda item: item[1])
    result["Группа затрат"] = result["ФИО"].apply(get_cost_group)
    result["Применять долю выручки (k)"] = participation.apply(lambda item: item[2])

    result[on_col] = result["Итого_ФОТ_плюс_взносы"] * result["Доля участия в соцтакси"]
    mask = result["Применять долю выручки (k)"]
    result.loc[mask, on_col] = (
        result.loc[mask, "Итого_ФОТ_плюс_взносы"]
        * result.loc[mask, "Доля участия в соцтакси"]
        * k
    )

    result = result[
        [
            "ФИО",
            "Категория участия",
            "Доля участия в соцтакси",
            "Группа затрат",
            fot_col,
            v_col,
            "Итого_ФОТ_плюс_взносы",
            on_col,
        ]
    ]

    group_summary = (
        result.groupby("Группа затрат", as_index=False)[on_col]
        .sum()
        .set_index("Группа затрат")
        .reindex(COST_GROUP_ORDER, fill_value=0)
        .reset_index()
        .rename(columns={on_col: "Сумма ФОТ+взносы на соцтакси, руб."})
    )
    total_row = pd.DataFrame(
        [
            {
                "Группа затрат": GROUP_TOTAL,
                "Сумма ФОТ+взносы на соцтакси, руб.": group_summary[
                    group_summary["Группа затрат"] != GROUP_OTHER
                ]["Сумма ФОТ+взносы на соцтакси, руб."].sum(),
            }
        ]
    )
    group_summary = pd.concat([group_summary, total_row], ignore_index=True)

    sheet_p = f"Персонал_{year}_soc_taxi"
    sheet_g = f"Группа_затрат_{year}"
    with pd.ExcelWriter(paths["soc_taxi_out"], engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name=sheet_p, index=False)
        group_summary.to_excel(writer, sheet_name=sheet_g, index=False)

    return result


def run_personal_pipeline(year: int) -> None:
    paths = paths_for_year(year)
    print(f"Собираю ФОТ за {year} год...")
    personal = build_personal_summary(year, paths)
    print(f"Создан файл: {paths['personal_out']}")

    print(f"Собираю страховые взносы за {year} год...")
    vznosy = build_vznosy_summary(year, paths)
    print(f"Создан файл: {paths['vznosy_out']}")

    print("Собираю итоговую таблицу ФОТ + взносы...")
    total = build_total_summary(personal, vznosy, year, paths)
    print(f"Готово, итоговый файл создан: {paths['itog_out']}")

    print("Собираю таблицу персонала с долями участия в соцтакси...")
    build_soc_taxi_summary(total, year, paths)
    print(f"Готово, целевой файл создан: {paths['soc_taxi_out']}")
