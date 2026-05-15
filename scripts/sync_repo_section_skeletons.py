# -*- coding: utf-8 -*-
"""
В корне репозитория: папки разделов с README.md и Сводная_ведомость.xlsx
(содержимое — по фактическим файлам учёта; README в инвентаризацию не входит).

Плюс 11_Поставщики и подпапка 11.1_Поставщики с теми же ведомостями.

Запуск: python scripts/sync_repo_section_skeletons.py
"""

from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))

from delo_case_vedomosti import (  # noqa: E402
    OUTPUT_NAME,
    _write_anchor,
    build_pdf_page_map,
)
from sozdat_vedomost_pdf import iter_inventory_files  # noqa: E402

SECTION_FOLDERS = [
    "01_Иск",
    "03_Учредительные",
    "05_ЕКЖЯ",
    "07_Кадры",
    "08_Аренда",
]


def main() -> None:
    disk: dict[str, dict] = {}
    for name in SECTION_FOLDERS:
        d = BASE / name
        d.mkdir(exist_ok=True)
        readme = d / "README.md"
        if not readme.exists():
            readme.write_bytes(b"")
        paths = list(iter_inventory_files(d))
        pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
        page_cache = build_pdf_page_map(pdfs, BASE, disk)
        _write_anchor(d, BASE, paths, page_cache)
        print(f"OK {name} -> {OUTPUT_NAME}")

    p11 = BASE / "11_Поставщики"
    p111 = p11 / "11.1_Поставщики"
    p11.mkdir(parents=True, exist_ok=True)
    p111.mkdir(parents=True, exist_ok=True)
    for pth in (p11 / "README.md", p111 / "README.md"):
        if not pth.exists():
            pth.write_bytes(b"")
    for anchor in (p11, p111):
        paths = list(iter_inventory_files(anchor))
        pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
        page_cache = build_pdf_page_map(pdfs, BASE, disk)
        _write_anchor(anchor, BASE, paths, page_cache)
        print(f"OK {anchor.relative_to(BASE)} -> {OUTPUT_NAME}")


if __name__ == "__main__":
    main()
