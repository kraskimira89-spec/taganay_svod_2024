# -*- coding: utf-8 -*-
"""
Копирует PDF договоров из локальной папки в docs/dogovory/<категория>/ по ключевым словам в имени файла.

Источник по умолчанию: переменная окружения DOGOVORY_SOURCE или путь ниже.

Запуск из корня проекта:
  python scripts/sync_dogovory.py

Пример:
  set DOGOVORY_SOURCE=E:\\ЦИП Таганай\\Договоры
  python scripts/sync_dogovory.py
"""

from __future__ import annotations

import os
import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST_ROOT = BASE / "docs" / "dogovory"

# Категория -> подстроки в имени файла (без учёта регистра)
MAPPING: list[tuple[str, tuple[str, ...]]] = [
    ("arenda", ("vw", "caddy", "volkswagen", "фолькс", "аренда", "lopakova", "лопакова", "norchak", "норчак", "garazh", "гараж", "01-20", "01_20")),
    ("toplivo", ("likard", "ликард", "gazprom", "газпром", "топлив", "gsm", "азс")),
    ("remont", ("alternativ", "альтернатив", "kaff", "кафф", "timchenko", "тимченко", "petrash", "петраш", "ремонт", "сто", "автотех")),
    ("strakhovanie", ("reso", "ресо", "осаго", "страх", "полис")),
    ("medosmotr", ("bud", "буд", "zdorov", "здоров", "sibir", "сибир", "медосмотр", "медицин")),
    ("svyaz_it", ("t2", "т2", "mobile", "мобайл", "print", "принт", "картридж", "reg.ru", "regru", "kb", "клиент", "it", "связ")),
    ("kommunalka", ("sev", "севэн", "sevenko", "vostok", "восток", "ek ", "энерг", "yamal", "ямаль", "эко", "хвс", "тепло", "мусор")),
    ("kp_avtotransport", ("gazel", "газель", "next", "кп_", "kp_", "коммерч", "автотранспорт")),
]


def categorize(name: str) -> str:
    lower = name.lower()
    for cat, keys in MAPPING:
        if any(k in lower for k in keys):
            return cat
    return "other"


def main() -> None:
    src_root = Path(os.environ.get("DOGOVORY_SOURCE", r"E:\ЦИП Таганай\Договоры")).expanduser()
    if not src_root.is_dir():
        raise SystemExit(
            f"Нет папки-источника: {src_root}. Задайте DOGOVORY_SOURCE=полный_путь к папке с PDF."
        )
    copied = 0
    for path in sorted(src_root.rglob("*.pdf")):
        if not path.is_file():
            continue
        cat = categorize(path.name)
        dest_dir = DEST_ROOT / cat
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / path.name
        if dest.is_file():
            n = 1
            while (dest_dir / f"{path.stem}_{n}{path.suffix}").is_file():
                n += 1
            dest = dest_dir / f"{path.stem}_{n}{path.suffix}"
        shutil.copy2(path, dest)
        print(f"{path.name} -> {cat}/")
        copied += 1
    if copied == 0:
        raise SystemExit(f"В {src_root} не найдено PDF. Проверьте путь.")
    print("Всего скопировано:", copied, "файлов в", DEST_ROOT.relative_to(BASE))


if __name__ == "__main__":
    main()
