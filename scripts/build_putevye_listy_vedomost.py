# -*- coding: utf-8 -*-
"""Сводная ведомость PDF «Путевые листы 2024» → data/delo_registers/13.2_*.xlsx + опись."""

from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from delo_github_policy import (  # noqa: E402
    apply_yandex_links_to_opis,
    copy_register_to_case,
    ensure_registers_dir,
    update_register_opis_row,
    upsert_opis_archive_rows,
    REGISTER_FILES,
)
from sozdat_vedomost_pdf import build_vedomost  # noqa: E402

DOCS_FOLDER = BASE / "docs" / "Путевые листы 2024"
CASE_FOLDER = BASE / "Дело3а-78-2026Таганай" / "13_Путевые_листы"
APP_ID = "13.2"


def _source_folder() -> Path:
    """Приоритет: папка дела (рекурсивно все PDF), иначе docs."""
    if CASE_FOLDER.is_dir():
        pdfs = list(CASE_FOLDER.rglob("*.pdf")) + list(CASE_FOLDER.rglob("*.PDF"))
        if pdfs:
            return CASE_FOLDER
    if DOCS_FOLDER.is_dir():
        return DOCS_FOLDER
    raise FileNotFoundError(
        f"Нет PDF в {CASE_FOLDER} и нет каталога {DOCS_FOLDER}"
    )


def main() -> None:
    folder = _source_folder()

    ensure_registers_dir()
    out = ensure_registers_dir() / REGISTER_FILES[APP_ID]

    build_vedomost(
        folder,
        out,
        "ВЕДОМОСТЬ ДОКУМЕНТОВ ПУТЕВЫХ ЛИСТОВ 2024",
        group_label="Папка / группа",
        mirror_dir=None,
    )
    copy_register_to_case(out, "13_Путевые_листы")

    upsert_opis_archive_rows()
    update_register_opis_row(
        APP_ID,
        "Сводная ведомость путевых листов за 2024 год (PDF)",
        2,
    )
    apply_yandex_links_to_opis()
    print(f"OK: {out}")


if __name__ == "__main__":
    main()
