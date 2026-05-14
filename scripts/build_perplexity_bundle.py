# -*- coding: utf-8 -*-
"""
Сборка ZIP для загрузки в Perplexity: исходники, выходы расчётов, код и конфиги.

Запуск из корня проекта:
  python scripts/build_perplexity_bundle.py

Результат: Perplexity_paket_Taganay_2024_2025.zip в корне репозитория.
"""

from __future__ import annotations

import zipfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
ARCHIVE_NAME = "Perplexity_paket_Taganay_2024_2025.zip"
PREFIX = "taganay_perplexity_2024_2025"

PY_FILES = [
    "svod_2024.py",
    "svod_osv60.py",
    "svod_osv60_core.py",
    "config_osv60.py",
    "config_allocation.py",
    "compare_osv60.py",
    "svod_reconcile_osv76.py",
    "svod_personal.py",
    "svod_vznosy.py",
    "debug_svod.py",
]

DOC_FILES = ["README.md", "requirements.txt", "beseda.md"]


def _collect_files() -> list[tuple[Path, str]]:
    """Пары (абсолютный путь, путь внутри архива)."""
    out: list[tuple[Path, str]] = []

    def add(rel: str, must_exist: bool = True) -> None:
        p = ROOT / rel
        if not p.is_file():
            if must_exist:
                raise FileNotFoundError(rel)
            return
        out.append((p, f"{PREFIX}/{rel}"))

    for name in DOC_FILES:
        add(name, must_exist=(name != "beseda.md"))

    for name in PY_FILES:
        add(name)

    scripts_dir = ROOT / "scripts"
    if scripts_dir.is_dir():
        for p in sorted(scripts_dir.glob("*.py")):
            rel = p.relative_to(ROOT).as_posix()
            out.append((p, f"{PREFIX}/{rel}"))

    extract = ROOT / "_extract_osv"
    if extract.is_dir():
        for p in sorted(extract.iterdir()):
            if p.is_file():
                rel = p.relative_to(ROOT).as_posix()
                out.append((p, f"{PREFIX}/{rel}"))

    docs_dir = ROOT / "docs"
    if docs_dir.is_dir():
        for p in sorted(docs_dir.rglob("*")):
            if p.is_file() and p.suffix.lower() in (".pdf", ".md"):
                rel = p.relative_to(ROOT).as_posix()
                out.append((p, f"{PREFIX}/{rel}"))

    for pattern in ("*.xlsx", "*.xls", "*.pdf"):
        for p in sorted(ROOT.glob(pattern)):
            rel = p.relative_to(ROOT).as_posix()
            if rel.startswith("_") or "Perplexity_paket" in rel:
                continue
            out.append((p, f"{PREFIX}/{rel}"))

    seen: set[str] = set()
    dedup: list[tuple[Path, str]] = []
    for pair in out:
        if pair[1] in seen:
            continue
        seen.add(pair[1])
        dedup.append(pair)
    return dedup


def _manifest_lines(entries: list[tuple[Path, str]]) -> str:
    lines = [
        "Пакет для Perplexity: Таганай — своды 2024/2025, сравнение, пояснительная записка (данные и код).",
        f"Дата сборки: {date.today().isoformat()}",
        "",
        "Состав (пути внутри архива):",
        "",
    ]
    for _, arc in sorted(entries, key=lambda x: x[1]):
        lines.append(arc)
    lines.extend(
        [
            "",
            "Кратко по назначению:",
            "- Svodnyi-po-ZP-za-2024-god.xlsx / .pdf — исходная сводная по ЗП;",
            "- Nalogi-i-vznosy*.xlsx — взносы;",
            "- Personal_2024_soc_taxi.xlsx — ФОТ+взносы, доли соцтакси, лист Группа_затрат_2024;",
            "- Osv60_soc_taxi_YYYY.xlsx — прочие расходы по сч. 60, лист Группа_затрат_записка;",
            "- Osv76_sverka_YYYY.xlsx — сверка ОСВ 76 с персоналом;",
            "- Svod_60_2024_2025.xlsx — сравнение групп записки (compare_osv60);",
            "- Svod_2024_2025_iz_arhiva_OSV_i_svod.xlsx — копия свода из папки отчётов (если синхронизировали);",
            "- _extract_osv/Osv_schet_* — выгрузки ОСВ из 1С;",
            "- config_allocation.py — доли приходов по соцуслугам (косвенные: 2024=0.78, 2025=0.76);",
            "- README.md — воспроизводимый pipeline;",
            "- docs/** — PDF и описания для раздела «Документальное обоснование».",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> Path:
    entries = _collect_files()
    manifest = _manifest_lines(entries)
    out_zip = ROOT / ARCHIVE_NAME

    with zipfile.ZipFile(out_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{PREFIX}/MANIFEST.txt", manifest.encode("utf-8"))
        for abs_path, arcname in entries:
            zf.write(abs_path, arcname)

    print("Создан архив:", out_zip)
    print("Файлов внутри (без MANIFEST):", len(entries))
    return out_zip


if __name__ == "__main__":
    main()
