"""
Общая логика свода по ОСВ счёта 60: загрузка, группы, записка, запись Excel.

Публичный CLI: `svod_osv60.py --year …` (см. также функции в этом модуле).
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from config_allocation import get_soc_taxi_share
from config_osv60 import (
    G_ARENDA_AVTO,
    G_ARENDA_GARAZH,
    G_IT,
    G_KOMM,
    G_MED,
    G_MOYKA,
    G_PROCH,
    G_REMONT,
    G_STRAH,
    G_TOP,
    classify_by_article,
)

BASE_DIR = Path(__file__).resolve().parent
EXTRACT_DIR = BASE_DIR / "_extract_osv"

COL_ON_SOC = "На_соцтакси"


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
    Классификация по наименованию контрагента (сырая выгрузка 1С без статьи).
    Возвращает (внутренняя_группа, прямые_100_процентов) или None.
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


def classify_row(
    counterparty: str,
    article: str | None,
) -> tuple[str, bool] | None:
    """Сначала статья расходов (если есть), иначе контрагент."""
    if is_bank_or_clearing(counterparty):
        return None
    by_art = classify_by_article(article)
    if by_art is not None:
        return by_art
    return classify_counterparty(counterparty)


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


def _cell_article(df: pd.DataFrame, row: int, article_col_idx: int | None) -> str | None:
    if article_col_idx is None:
        return None
    if article_col_idx < 0 or article_col_idx >= df.shape[1]:
        return None
    v = df.iloc[row, article_col_idx]
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s else None


def guess_article_column(df: pd.DataFrame) -> int | None:
    """
    Для выгрузок с доп. столбцом «Статья расходов» (как «верная» ОСВ из бухгалтерии).
    Ищем столбец справа от стандартных 7 числовых колонок с текстовыми пояснениями.
    """
    if df.shape[1] <= 7:
        return None
    best_c: int | None = None
    best_score = 0.0
    for c in range(7, min(df.shape[1], 14)):
        texts = 0
        total = 0
        for i in range(8, min(len(df), 40)):
            name = df.iloc[i, 0]
            if not isinstance(name, str) or not name.strip():
                continue
            nu = name.strip().upper()
            if nu == "60" or nu.startswith("ИТОГ"):
                continue
            v = df.iloc[i, c]
            total += 1
            if isinstance(v, str) and len(v.strip()) > 1:
                t = v.strip()
                if not re.fullmatch(r"[\d\s.,+-]+", t.replace(",", ".")):
                    texts += 1
        if total == 0:
            continue
        score = texts / total
        if score > best_score:
            best_score = score
            best_c = c
    if best_c is not None and best_score >= 0.35:
        return best_c
    return None


def load_osv60_rows(
    path: Path,
    article_col_idx: int | None = None,
    year: int = 2024,
) -> tuple[pd.DataFrame, pd.DataFrame, int | None]:
    df = pd.read_excel(path, header=None, engine="xlrd")
    k_indirect = get_soc_taxi_share(year)
    effective_article_col = article_col_idx
    if effective_article_col is None:
        effective_article_col = guess_article_column(df)
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
        article = _cell_article(df, i, effective_article_col)

        classified = classify_row(name, article)
        if classified is None:
            excluded.append(
                {
                    "Строка": i + 1,
                    "Контрагент": name,
                    "Статья_расходов": article or "",
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
        # Прямые расходы соцтакси — 100% базы; косвенные — по доле приходов по соцуслугам (см. config_allocation)
        on_soc = base if direct else base * k_indirect

        included.append(
            {
                "Строка": i + 1,
                "Контрагент": name,
                "Статья_расходов": article or "",
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
    return pd.DataFrame(included), pd.DataFrame(excluded), effective_article_col


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


def build_svod_dataframe(detail: pd.DataFrame) -> pd.DataFrame:
    """
    Свод по внутренним группам: сумма оборота по кредиту и сумма, отнесённая на соцтакси.
    """
    if detail.empty:
        return pd.DataFrame(columns=["Внутренняя_группа", "Сумма_оборот_Кт", "Сумма_на_соцтакси"])
    return (
        detail.groupby("Внутренняя_группа", as_index=False)
        .agg(Сумма_оборот_Кт=("Оборот_Кт", "sum"), Сумма_на_соцтакси=(COL_ON_SOC, "sum"))
        .sort_values("Внутренняя_группа", ignore_index=True)
    )


def write_osv60_workbook(
    output_path: Path,
    detail: pd.DataFrame,
    skipped: pd.DataFrame,
    svod: pd.DataFrame,
    zapiska: pd.DataFrame,
) -> None:
    """Четыре листа: детально, свод по группам, исключённые строки, строки для пояснительной."""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="Строки_ОСВ60", index=False)
        if svod is not None and len(svod) > 0:
            svod.to_excel(writer, sheet_name="Свод_по_группам", index=False)
        skipped.to_excel(writer, sheet_name="Не_включено_в_свод", index=False)
        zapiska.to_excel(writer, sheet_name="Группа_затрат_записка", index=False)


def run_with_paths(
    year: int,
    osv_path: Path | None = None,
    out_path: Path | None = None,
    article_col_idx: int | None = None,
) -> Path:
    """Строит Excel-свод по счёту 60 за указанный год."""
    src = osv_path if osv_path is not None else find_osv60_file(year)
    if not src.is_file():
        raise FileNotFoundError(f"Нет файла ОСВ: {src}")
    out_file = out_path if out_path is not None else (BASE_DIR / f"Osv60_soc_taxi_{year}.xlsx")

    detail, skipped, used_article_col = load_osv60_rows(
        src, article_col_idx=article_col_idx, year=year
    )
    svod = build_svod_dataframe(detail)
    zapiska = build_zapiska_table(detail, year)
    write_osv60_workbook(out_file, detail, skipped, svod, zapiska)

    print_osv60_run_summary(
        src, out_file, used_article_col, article_col_idx, detail, skipped, zapiska
    )
    return out_file


def print_osv60_run_summary(
    src: Path,
    out_file: Path,
    used_article_col: int | None,
    article_col_idx: int | None,
    detail: pd.DataFrame,
    skipped: pd.DataFrame,
    zapiska: pd.DataFrame,
) -> None:
    print("Источник:", src)
    print("Создан файл:", out_file)
    if used_article_col is not None:
        src_note = "задана" if article_col_idx is not None else "определена автоматически"
        print(f"Колонка «Статья расходов» ({src_note}), 0-based индекс:", used_article_col)
    print("Строк включено:", len(detail), "исключено:", len(skipped))
    print("Итого по счёту 60 на соцтакси:", round(float(zapiska.iloc[-1, 1]), 2), "руб.")
