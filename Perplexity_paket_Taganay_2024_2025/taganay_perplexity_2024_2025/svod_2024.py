from pathlib import Path

import pandas as pd

from config_allocation import get_soc_taxi_share

BASE_DIR = Path(__file__).resolve().parent

PERSONAL_FILE = BASE_DIR / "Svodnyi-po-ZP-za-2024-god.xlsx"
PERSONAL_OUT = BASE_DIR / "Personal_2024_auto.xlsx"

VZNOSY_PATTERN = "Nalogi-i-vznosy*.xlsx"
VZNOSY_OUT = BASE_DIR / "Vznosy_2024_auto.xlsx"

ITOG_OUT = BASE_DIR / "Itog_personal_vznosy_2024.xlsx"
PERSONAL_SOC_TAXI_OUT = BASE_DIR / "Personal_2024_soc_taxi.xlsx"

SOC_TAXI_REVENUE_SHARE_2024 = get_soc_taxi_share(2024)

GROUP_DRIVERS_SOCIAL = "1. ФОТ водителей и соцработников"
GROUP_DISPATCHER = "2. ФОТ диспетчера"
GROUP_REPAIR = "3. ФОТ ремонтного персонала (Егоров, Шакалов, Родионов)"
GROUP_SOC_SERVICE_HEAD = "4. ФОТ руководителя соцслужбы (Лопакова Н.Ф.)"
GROUP_ORG_HEAD = "5. ФОТ руководителя организации (50% Богдановского)"
GROUP_ADMIN = "6. ФОТ бухгалтерии, IT и админ-персонала"
GROUP_OTHER = "0%. Другая деятельность (Роснефть/прочее)"
GROUP_TOTAL = "ИТОГО трудовые затраты по соцтакси"

PARTICIPATION_RULES = {
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
    "Поплаухина": ("Адм.", SOC_TAXI_REVENUE_SHARE_2024, False),
    "Голоушкин": ("Адм.", SOC_TAXI_REVENUE_SHARE_2024, False),
    "Иванова": ("Адм.", SOC_TAXI_REVENUE_SHARE_2024, False),
    "Исламова": ("Адм.", SOC_TAXI_REVENUE_SHARE_2024, False),
    "Лопакова Светлана": ("Адм.", SOC_TAXI_REVENUE_SHARE_2024, False),
    "Зубцов": ("0%", 0.0, False),
    "Капков": ("0%", 0.0, False),
    "Першаков": ("0%", 0.0, False),
}

COST_GROUP_RULES = {
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


def get_participation(fio: object) -> tuple[str, float, bool]:
    text = str(fio)

    for marker, rule in PARTICIPATION_RULES.items():
        if marker in text:
            return rule

    return "Другая деятельность (Роснефть/прочее)", 0.0, False


def get_cost_group(fio: object) -> str:
    text = str(fio)

    for marker, group in COST_GROUP_RULES.items():
        if marker in text:
            return group

    return GROUP_OTHER


def is_person_name(value: object) -> bool:
    text = str(value).strip()
    if not text or text in {"nan", "Сотрудник"}:
        return False
    stop_words = [
        "2024",
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


def extract_tax_sheet(df: pd.DataFrame) -> pd.DataFrame | None:
    df = df.copy()

    if df.empty or df.shape[1] < 5:
        return None

    text = " ".join(str(value) for value in df.to_numpy().ravel() if pd.notna(value))
    if "Налоги и взносы" not in text:
        return None

    part = df.iloc[:, [0, 1, 3, 4]].copy()
    part.columns = ["ФИО", "Начислено_по_взносам", "Единый_тариф", "Несч_случаи"]
    part = part[part["ФИО"].apply(is_person_name)]
    part["Начислено_по_взносам"] = to_number(part["Начислено_по_взносам"])
    part["Единый_тариф"] = to_number(part["Единый_тариф"])
    part["Несч_случаи"] = to_number(part["Несч_случаи"])
    part["Взносы_2024"] = part["Единый_тариф"] + part["Несч_случаи"]
    return part[["ФИО", "Начислено_по_взносам", "Взносы_2024"]]


def read_tax_sheets(file: Path) -> list[pd.DataFrame]:
    all_sheets = pd.read_excel(file, sheet_name=None, header=None)
    frames = []

    for sheet_name, df in all_sheets.items():
        part = extract_tax_sheet(df)
        if part is None:
            continue

        print(f"Найден лист с налогами и взносами: {file.name} / {sheet_name}")
        frames.append(part)

    return frames


def build_personal_summary() -> pd.DataFrame:
    if not PERSONAL_FILE.exists():
        raise FileNotFoundError(f"Не найден файл с зарплатой: {PERSONAL_FILE}")

    fio_col = "Фамилия, имя, отчество"
    sum_col = "Начислено всего"

    all_sheets = pd.read_excel(PERSONAL_FILE, sheet_name=None)
    frames = []

    for sheet_name, df in all_sheets.items():
        df = normalize_columns(df)

        if fio_col not in df.columns or sum_col not in df.columns:
            print(f"Пропущен лист зарплаты без нужных колонок: {sheet_name}")
            continue

        part = df[[fio_col, sum_col]].copy()
        part.columns = ["ФИО", "ФОТ_2024"]
        part["ФОТ_2024"] = to_number(part["ФОТ_2024"])
        frames.append(part)

    if not frames:
        print("Обычные колонки зарплаты не найдены, пробую взять ФОТ из листов налогов.")
        frames = [
            frame[["ФИО", "Начислено_по_взносам"]].rename(
                columns={"Начислено_по_взносам": "ФОТ_2024"}
            )
            for frame in read_tax_sheets(PERSONAL_FILE)
        ]

    if not frames:
        raise SystemExit("Не найдено ни одного листа с ФИО и начислениями.")

    full = pd.concat(frames, ignore_index=True)
    full = full.dropna(subset=["ФИО"])
    full = full[full["ФОТ_2024"].notna()]

    summary = full.groupby("ФИО", as_index=False)["ФОТ_2024"].sum()
    summary.to_excel(PERSONAL_OUT, index=False)
    return summary


def build_vznosy_summary() -> pd.DataFrame:
    files = sorted(BASE_DIR.glob(VZNOSY_PATTERN))

    if not files:
        raise FileNotFoundError(f"Не найдены файлы взносов по маске: {VZNOSY_PATTERN}")

    frames = []

    for file in files:
        frames.extend(read_tax_sheets(file))

    if not frames:
        raise SystemExit("Не найдено ни одного файла взносов с нужными колонками.")

    full = pd.concat(frames, ignore_index=True)
    summary = full.groupby("ФИО", as_index=False).agg(
        {
            "Начислено_по_взносам": "sum",
            "Взносы_2024": "sum",
        }
    )
    summary.to_excel(VZNOSY_OUT, index=False)
    return summary


def build_total_summary(personal: pd.DataFrame, vznosy: pd.DataFrame) -> pd.DataFrame:
    total = personal.merge(vznosy[["ФИО", "Взносы_2024"]], on="ФИО", how="outer")
    total["ФОТ_2024"] = total["ФОТ_2024"].fillna(0)
    total["Взносы_2024"] = total["Взносы_2024"].fillna(0)
    total["Итого_ФОТ_плюс_взносы"] = total["ФОТ_2024"] + total["Взносы_2024"]
    total = total.sort_values("ФИО")
    total.to_excel(ITOG_OUT, index=False)
    return total


def build_soc_taxi_summary(total: pd.DataFrame) -> pd.DataFrame:
    result = total.copy()
    participation = result["ФИО"].apply(get_participation)

    result["Категория участия"] = participation.apply(lambda item: item[0])
    result["Доля участия в соцтакси"] = participation.apply(lambda item: item[1])
    result["Группа затрат"] = result["ФИО"].apply(get_cost_group)
    result["Применять долю выручки 78%"] = participation.apply(lambda item: item[2])

    result["На_соцтакси_2024"] = (
        result["Итого_ФОТ_плюс_взносы"] * result["Доля участия в соцтакси"]
    )

    mask = result["Применять долю выручки 78%"]
    result.loc[mask, "На_соцтакси_2024"] = (
        result.loc[mask, "Итого_ФОТ_плюс_взносы"]
        * result.loc[mask, "Доля участия в соцтакси"]
        * SOC_TAXI_REVENUE_SHARE_2024
    )

    result = result[
        [
            "ФИО",
            "Категория участия",
            "Доля участия в соцтакси",
            "Группа затрат",
            "ФОТ_2024",
            "Взносы_2024",
            "Итого_ФОТ_плюс_взносы",
            "На_соцтакси_2024",
        ]
    ]

    group_summary = (
        result.groupby("Группа затрат", as_index=False)["На_соцтакси_2024"].sum()
        .set_index("Группа затрат")
        .reindex(COST_GROUP_ORDER, fill_value=0)
        .reset_index()
        .rename(columns={"На_соцтакси_2024": "Сумма ФОТ+взносы на соцтакси, руб."})
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

    with pd.ExcelWriter(PERSONAL_SOC_TAXI_OUT, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="Персонал_2024_soc_taxi", index=False)
        group_summary.to_excel(writer, sheet_name="Группа_затрат_2024", index=False)

    return result


def main() -> None:
    print("Собираю ФОТ за 2024 год...")
    personal = build_personal_summary()
    print(f"Создан файл: {PERSONAL_OUT}")

    print("Собираю страховые взносы за 2024 год...")
    vznosy = build_vznosy_summary()
    print(f"Создан файл: {VZNOSY_OUT}")

    print("Собираю итоговую таблицу ФОТ + взносы...")
    total = build_total_summary(personal, vznosy)
    print(f"Готово, итоговый файл создан: {ITOG_OUT}")

    print("Собираю таблицу персонала с долями участия в соцтакси...")
    build_soc_taxi_summary(total)
    print(f"Готово, целевой файл создан: {PERSONAL_SOC_TAXI_OUT}")


if __name__ == "__main__":
    main()
