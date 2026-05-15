# -*- coding: utf-8 -*-
"""Аудит готовности приложений к отчётам / делу 3а-78/2026."""

from __future__ import annotations

import json
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

# (id, название, путь от BASE, обязательно для отчётов)
ITEMS = [
    ("02.1", "Обоснование 2024 PDF", "docs/obosnovanie2024/obosnovanie2024.pdf", True),
    ("02.2", "Обоснование 2025 PDF", "docs/obosnovanie2025/obosnovanie2025.pdf", True),
    ("02.3", "Отчёт №3 сравнительный", "docs/obosnovanie2025/Отчёт_№3_Сравнительный.pdf", False),
    ("04.1-04.5", "Соглашения 2025", "2025_uslugi", True),
    ("06.1", "ОСВ 60 2025", "Дело3а-78-2026Таганай/06_Финансы", False),
    ("06.3", "ВТБ 2024", "docs/Inbox_для_дела_3а-78-2026/06.3_VTB_BankStatement_2024.pdf", True),
    ("06.3b", "ВТБ 2025", "docs/Inbox_для_дела_3а-78-2026/06.3_VTB_BankStatement_2025.pdf", True),
    ("06.3sber", "Сбер 2024", "docs/bank_vypiski/sber", True),
    ("06.4", "Сверка банк-свод", "06.4_Сверка_банк_свод_2024-2025.xlsx", True),
    ("09.1-09.9", "Нормативка", "docs/Inbox_для_дела_3а-78-2026", True),
    ("10.1-10.5", "Рыночные зарплаты", "docs/Inbox_для_дела_3а-78-2026", True),
    ("11.2", "Свод поставщиков", "data/delo_registers/11.2_Сводная_ведомость_поставщиков.xlsx", True),
    ("11.2pdf", "Ведомость PDF поставщики", "docs/Поставщики", False),
    ("12.1", "Свод разнарядки 2024", "data/delo_registers/12.1_Сводная_ведомость_разнарядок_2024.xlsx", True),
    ("12.2", "Свод разнарядки 2025", "data/delo_registers/12.2_Сводная_ведомость_разнарядок_2025.xlsx", True),
    ("yandex", "Ссылки Яндекс", "docs/delo/yandex_disk_links.yaml", True),
    ("opis", "Опись", "docs/obosnovanie2025/Опись_приложений_шаблон.xlsx", True),
    ("Personal24", "Personal 2024", "Personal_2024_soc_taxi.xlsx", True),
    ("Personal25", "Personal 2025", "Personal_2025_soc_taxi.xlsx", False),
    ("01", "Иск", "Дело3а-78-2026Таганай/01_Иск", False),
    ("03", "Учредительные", "Дело3а-78-2026Таганай/03_Учредительные", False),
    ("05", "ЕКЖЯ", "Дело3а-78-2026Таганай/05_ЕКЖЯ", False),
    ("07", "Кадры", "Дело3а-78-2026Таганай/07_Кадры", False),
    ("08", "Аренда VW", "Дело3а-78-2026Таганай/08_Аренда", False),
    ("10osv", "ОСВ 70 ФОТ", "Fot_osv70_soc_taxi_2025", False),
]

SECTIONS = list(f"{i:02d}_" for i in range(1, 13))


def check_path(rel: str) -> dict:
    p = BASE / rel
    if p.is_file():
        return {"status": "ok", "detail": f"файл {p.stat().st_size // 1024} КБ"}
    if p.is_dir():
        files = list(p.rglob("*"))
        n_files = sum(1 for f in files if f.is_file())
        if n_files:
            return {"status": "ok", "detail": f"папка, {n_files} файлов"}
        return {"status": "empty", "detail": "папка пуста"}
    # glob by name in base
    if "/" not in rel and "*" not in rel:
        hits = list(BASE.rglob(rel.split("/")[-1]))
        if hits:
            return {"status": "ok", "detail": f"найдено: {len(hits)} — {hits[0].relative_to(BASE)}"}
    return {"status": "missing", "detail": "не найдено"}


def main() -> None:
    case = BASE / "Дело3а-78-2026Таганай"
    sections = {}
    if case.is_dir():
        for d in case.iterdir():
            if d.is_dir():
                n = sum(1 for _ in d.rglob("*") if _.is_file())
                sections[d.name] = n

    results = []
    for item_id, name, path, required in ITEMS:
        r = check_path(path)
        r["id"] = item_id
        r["name"] = name
        r["required"] = required
        results.append(r)

    # yandex filled?
    ypath = BASE / "docs/delo/yandex_disk_links.yaml"
    yandex_ok = False
    if ypath.is_file():
        import yaml

        data = yaml.safe_load(ypath.read_text(encoding="utf-8")) or {}
        arch = data.get("archives") or {}
        filled = [k for k, v in arch.items() if v and "disk.yandex" in str(v)]
        yandex_ok = len(filled) >= 1
        for r in results:
            if r["id"] == "yandex":
                r["status"] = "ok" if filled else "partial"
                r["detail"] = f"заполнено ключей: {len(filled)}/3 ({', '.join(filled) or 'нет URL'})"

    # 11.2 vedomost PDF format
    post = list((BASE / "docs/Поставщики").glob("Ведомость_PDF_*.xlsx")) if (BASE / "docs/Поставщики").is_dir() else []
    reg11 = BASE / "data/delo_registers/11.2_Сводная_ведомость_поставщиков.xlsx"
    if post and not reg11.is_file():
        results.append(
            {
                "id": "11.2warn",
                "name": "Ведомость PDF поставщики (не в registers)",
                "required": True,
                "status": "partial",
                "detail": f"есть {post[-1].name}, нет data/delo_registers/11.2_...",
            }
        )

    out = {
        "sections_in_case": sections,
        "items": results,
        "ok": sum(1 for r in results if r["status"] == "ok"),
        "missing_required": [
            r["id"] for r in results if r["required"] and r["status"] in ("missing", "empty")
        ],
        "partial": [r["id"] for r in results if r["status"] == "partial"],
    }
    out_path = BASE / "docs" / "delo" / "audit_prilozheniya.json"
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
