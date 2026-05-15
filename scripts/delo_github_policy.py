# -*- coding: utf-8 -*-
"""Политика GitHub: крупные PDF локально/Яндекс.Диск; в git — сводные ведомости."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import yaml
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent.parent
REGISTERS_DIR = BASE / "data" / "delo_registers"
OPIS_PATH = BASE / "docs" / "obosnovanie2025" / "Опись_приложений_шаблон.xlsx"
KATALOG_PATH = BASE / "Дело3а-78-2026Таганай" / "Каталог_приложений.md"
YANDEX_DIR = BASE / "docs" / "delo"
YANDEX_EXAMPLE = YANDEX_DIR / "yandex_disk_links.example.yaml"
YANDEX_LINKS = YANDEX_DIR / "yandex_disk_links.yaml"

OPIS_SHEET = "Опись приложений"
OPIS_HEADER_ROW = 9
OPIS_COL_YANDEX = 9  # I
PLACEHOLDER = "https://disk.yandex.ru/…"

# app_id -> ключ в yaml (None = только GitHub, без отдельного архива)
YANDEX_KEYS: dict[str, str | None] = {
    "11.1": "postavshchiki",
    "11.2": None,
    "12.1": None,
    "12.2": None,
    "12.3": "raznaryadki_2024",
    "12.4": "raznaryadki_2025",
    "13.1": "putevye_listy_2024",
    "13.2": None,
}

REGISTER_FILES = {
    "11.2": "11.2_Сводная_ведомость_поставщиков.xlsx",
    "12.1": "12.1_Сводная_ведомость_разнарядок_2024.xlsx",
    "12.2": "12.2_Сводная_ведомость_разнарядок_2025.xlsx",
    "13.2": "13.2_Сводная_ведомость_путевые_листы_2024.xlsx",
}


def load_yandex_links() -> dict[str, str]:
    path = YANDEX_LINKS if YANDEX_LINKS.is_file() else YANDEX_EXAMPLE
    if not path.is_file():
        return {}
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return {str(k): str(v).strip() for k, v in (data.get("archives") or {}).items() if v}


def link_for_app(app_id: str, links: dict[str, str] | None = None) -> str:
    links = links if links is not None else load_yandex_links()
    key = YANDEX_KEYS.get(app_id)
    if not key:
        return ""
    url = links.get(key, "").strip()
    return url or PLACEHOLDER


def register_path(app_id: str) -> Path:
    name = REGISTER_FILES[app_id]
    return REGISTERS_DIR / name


def ensure_registers_dir() -> Path:
    REGISTERS_DIR.mkdir(parents=True, exist_ok=True)
    return REGISTERS_DIR


def copy_register_to_case(src: Path, case_subdir: str) -> Path | None:
    """Копия в папку дела (локально); в git не попадает."""
    if not src.is_file():
        return None
    case_dir = BASE / "Дело3а-78-2026Таганай" / case_subdir
    case_dir.mkdir(parents=True, exist_ok=True)
    dest = case_dir / src.name
    import shutil

    shutil.copy2(src, dest)
    return dest


def _row_by_filename(ws, filename: str) -> int | None:
    for r in range(OPIS_HEADER_ROW + 1, ws.max_row + 30):
        if ws.cell(r, 7).value == filename:
            return r
    return None


def _row_by_app(ws, app_id: str) -> int | None:
    for r in range(OPIS_HEADER_ROW + 1, ws.max_row + 30):
        e = ws.cell(r, 5).value
        if e and str(e).strip() == app_id:
            return r
    return None


def ensure_opis_yandex_column() -> None:
    if not OPIS_PATH.is_file():
        return
    wb = load_workbook(OPIS_PATH)
    ws = wb[OPIS_SHEET]
    ws.cell(OPIS_HEADER_ROW, OPIS_COL_YANDEX, "Ссылка Яндекс.Диск")
  # row 7 hint
    hint = ws.cell(7, 2).value or ""
    if "Яндекс" not in str(hint):
        ws.cell(
            7,
            2,
            str(hint).rstrip(".")
            + " Колонка I — публичная ссылка на архив PDF (если файл не в репозитории).",
        )
    wb.save(OPIS_PATH)


def upsert_opis_archive_rows() -> None:
    """Строки для архивов только на Диске (11.1, 12.3, 12.4)."""
    if not OPIS_PATH.is_file():
        return
    wb = load_workbook(OPIS_PATH)
    ws = wb[OPIS_SHEET]
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    links = load_yandex_links()

    archives = [
        (
            "11.1",
            "11 Поставщики",
            "Архив документов поставщиков (исходные PDF, папки 1–37)",
            "0",
            "(только Яндекс.Диск; см. кол. I)",
            "postavshchiki",
        ),
        (
            "12.3",
            "12 Разнарядки",
            "Исходные PDF разнарядок за 2024 год (сканы)",
            "629",
            "(только Яндекс.Диск)",
            "raznaryadki_2024",
        ),
        (
            "12.4",
            "12 Разнарядки",
            "Исходные PDF разнарядок за 2025 год (по дням)",
            "530",
            "(только Яндекс.Диск)",
            "raznaryadki_2025",
        ),
        (
            "13.1",
            "13 Путевые листы",
            "Исходные PDF путевых листов за 2024 год",
            "0",
            "(только Яндекс.Диск; см. кол. I)",
            "putevye_listy_2024",
        ),
    ]

    insert_at = ws.max_row
    while insert_at > OPIS_HEADER_ROW and not ws.cell(insert_at, 2).value:
        insert_at -= 1
    insert_at += 1

    for app, section, title, sheets, fname, ykey in archives:
        row = _row_by_app(ws, app)
        if row is None:
            row = insert_at
            insert_at += 1
        ws.cell(row, 2, section)
        ws.cell(row, 3, title)
        ws.cell(row, 4, today)
        ws.cell(row, 5, app)
        ws.cell(row, 6, sheets)
        ws.cell(row, 7, fname)
        url = links.get(ykey, "").strip() or PLACEHOLDER
        ws.cell(row, OPIS_COL_YANDEX, url)

    # Старая строка с ZIP (11.1) — не в git
    for r in range(OPIS_HEADER_ROW + 1, ws.max_row + 5):
        g = ws.cell(r, 7).value
        if not g:
            continue
        gs = str(g)
        if "11.1_" in gs and ".zip" in gs.lower():
            ws.cell(r, 3, "(снято с git — архив на Яндекс.Диск, прил. 11.1)")
            ws.cell(r, 5, "—")
            ws.cell(r, 6, "0")
            ws.cell(r, 7, "—")
            ws.cell(r, OPIS_COL_YANDEX, "см. строку 11.1 ниже")
        if "части .001" in str(ws.cell(r, 5).value or ""):
            ws.cell(r, 5, "—")
            ws.cell(r, 6, "0")
            ws.cell(r, 7, "—")
            ws.cell(r, OPIS_COL_YANDEX, "см. прил. 11.1")

    wb.save(OPIS_PATH)


def apply_yandex_links_to_opis() -> int:
    """Проставляет ссылки из yaml в колонку I для всех app_id из YANDEX_KEYS."""
    if not OPIS_PATH.is_file():
        return 0
    ensure_opis_yandex_column()
    links = load_yandex_links()
    wb = load_workbook(OPIS_PATH)
    ws = wb[OPIS_SHEET]
    updated = 0

    for app_id, ykey in YANDEX_KEYS.items():
        if ykey is None:
            continue
        url = link_for_app(app_id, links)
        row = _row_by_app(ws, app_id)
        if row:
            ws.cell(row, OPIS_COL_YANDEX, url)
            updated += 1

    # сводные в git — пустая или поясняющая ссылка на архив
    for app_id, reg_file in REGISTER_FILES.items():
        row = _row_by_filename(ws, reg_file)
        if row:
            ref = {"11.2": "11.1", "12.1": "12.3", "12.2": "12.4", "13.2": "13.1"}.get(
                app_id
            )
            note = f"исходники: приложение {ref}" if ref else ""
            if note and not ws.cell(row, OPIS_COL_YANDEX).value:
                ws.cell(row, OPIS_COL_YANDEX, note)

    wb.save(OPIS_PATH)
    return updated


def update_register_opis_row(app_id: str, title: str, sheets: int) -> None:
    fname = REGISTER_FILES.get(app_id)
    if not fname or not OPIS_PATH.is_file():
        return
    wb = load_workbook(OPIS_PATH)
    ws = wb[OPIS_SHEET]
    row = _row_by_filename(ws, fname)
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    if app_id.startswith("11"):
        section = "11 Поставщики"
    elif app_id.startswith("13"):
        section = "13 Путевые листы"
    else:
        section = "12 Разнарядки"
    if row is None:
        row = ws.max_row + 1
        while row > OPIS_HEADER_ROW and not ws.cell(row, 2).value:
            row -= 1
        row += 1
    ws.cell(row, 2, section)
    ws.cell(row, 3, title)
    ws.cell(row, 4, today)
    ws.cell(row, 5, app_id)
    ws.cell(row, 6, sheets)
    ws.cell(row, 7, fname)
    ref = {"11.2": "11.1", "12.1": "12.3", "12.2": "12.4", "13.2": "13.1"}.get(app_id)
    if ref:
        ws.cell(row, OPIS_COL_YANDEX, f"исходники PDF: приложение {ref} (Яндекс.Диск)")
    wb.save(OPIS_PATH)
