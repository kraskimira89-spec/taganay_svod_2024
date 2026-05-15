# -*- coding: utf-8 -*-
"""
Раскладывает содержимое docs/ДСЗН 2025 по тарифу соцтакси/Суд ЯНАО
в Дело3а-78-2026Таганай/01_Иск/ по тематическим подпапкам 01.x.

Запуск:
  python scripts/import_sud_yanao_to_isk.py
"""

from __future__ import annotations

import shutil
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
SRC = BASE / "docs" / "ДСЗН 2025 по тарифу соцтакси" / "Суд ЯНАО"
DST = BASE / "Дело3а-78-2026Таганай" / "01_Иск"


def classification_rel(src_root: Path, rel: Path) -> Path:
    """Для *.sig с соседним PDF с тем же именем — классифицируем как PDF (лежат в одной папке)."""
    if rel.suffix.lower() != ".sig":
        return rel
    name_lower = rel.name.lower()
    if name_lower.endswith(".pdf.sig"):
        paired = rel.with_name(rel.stem)
    else:
        paired = rel.with_name(rel.stem + ".pdf")
    if (src_root / paired).is_file():
        return paired
    return rel


def _bank_payment_filename(name_lower: str) -> bool:
    """Типовые маски выписок/подтверждений из интернет-банка и биллинга (латиница в имени)."""
    return (
        "payment_commission" in name_lower
        or "advanced_payment" in name_lower
        or "on_schet" in name_lower
        or "vtb_payment" in name_lower
    )


def _narjad_zakaz_filename(name_lower: str) -> bool:
    return "наряд-заказ" in name_lower or (
        "наряд" in name_lower and "заказ" in name_lower
    )


def bucket(rel: Path) -> str:
    """rel — относительный путь от SRC (Posix-стиль в строке для проверок)."""
    s = rel.as_posix().lower()
    n = rel.name.lower()
    parts_lower = " ".join(p.lower() for p in rel.parts)

    if "attachments" in parts_lower or "attachment_taganai" in parts_lower:
        return "01.7_Вложения_ГАС"
    if any("бесед" in p for p in rel.parts):
        return "01.9_Черновики_беседы"
    if "заявки" in s.split("/"):
        return "01.6_Данные_таблицы"
    if _bank_payment_filename(n):
        return "01.13_Банковские_платежи_и_счета"
    if _narjad_zakaz_filename(n):
        return "01.14_Наряд_заказы_СТО"
    if "для суда" in parts_lower:
        return "01.11_Комплект_Для_суда"
    if "attachments" in n or ("attachment" in n and n.endswith(".zip")):
        return "01.7_Вложения_ГАС"
    if "путев" in parts_lower:
        return "01.12_Путевые_листы_и_транспорт"
    if any("материал" in p.lower() for p in rel.parts):
        if any(
            x in parts_lower for x in ("подтвержд", "дела", "дело", "гас")
        ):
            return "01.10_Пакеты_документов_ГАС"
    if "правительств" in parts_lower:
        return "01.8_НПА_методические_материалы"
    if "заседан" in parts_lower or "заседан" in n:
        return "01.3_Определения_и_акты_суда"
    if "89-" in s or "_89-" in s or "89_" in s:
        return "01.3_Определения_и_акты_суда"
    if "итог" in n and ".202" in n:
        return "01.3_Определения_и_акты_суда"
    if any("комплект" in p.lower() for p in rel.parts):
        return "01.1_Исковые_документы_и_комплекты"
    if "ходатай" in n:
        return "01.2_Ходатайства"
    if "отзыв" in n:
        return "01.4_Отзывы"
    if "госпошлин" in n or "квитанц" in n:
        return "01.5_Госпошлина"
    if n.endswith((".csv", ".xlsx")):
        return "01.6_Данные_таблицы"

    # НПА / методичка (не судебные акты по делу)
    npa_markers = (
        "442",
        "557",
        "1285",
        "1306",
        "129-",
        "129_",
        "202-",
        "202_",
        "federalnyj",
        "federal",
        "приказ_дсзн",
        "постановление правительства рф",
        "постановление правительства янао",
        "cbr_press",
    )
    if any(m in n for m in npa_markers):
        return "01.8_НПА_методические_материалы"
    if "постановление" in n and "суд" not in n and not any(
        x in n for x in ("89-07", "89-02", "3а-", "3a-")
    ):
        if any(x in n for x in ("557", "1285", "1306", "129", "442")):
            return "01.8_НПА_методические_материалы"

    if "определен" in n:
        return "01.3_Определения_и_акты_суда"
    if "жалоб" in n:
        return "01.1_Исковые_документы_и_комплекты"
    if "исков" in n or "иск " in n or "иск." in n:
        return "01.1_Исковые_документы_и_комплекты"
    if "административн" in n and "заяв" in n:
        return "01.1_Исковые_документы_и_комплекты"
    if "соглашен" in n and "22-25" in n:
        return "01.8_НПА_методические_материалы"
    if "20-26" in n or "22-25" in n:
        return "01.1_Исковые_документы_и_комплекты"

    if n.endswith((".png", ".rar")) or "shpargalka" in n or "plan_ustnykh" in n:
        return "01.9_Черновики_и_служебное"
    if n.startswith("00_") or "opis_dokumentov" in n:
        return "01.9_Черновики_и_служебное"

    if n.endswith(".sig"):
        return "01.0_ЭЦП"

    return "01.9_Прочее_процессуальное"


def main() -> None:
    if not SRC.is_dir():
        raise FileNotFoundError(SRC)

    DST.mkdir(parents=True, exist_ok=True)
    for child in list(DST.iterdir()):
        if child.is_dir() and (
            child.name.startswith("01.") or child.name.startswith("00_")
        ):
            shutil.rmtree(child, ignore_errors=True)
        elif child.name == "README_структура_01_Иск.md":
            child.unlink(missing_ok=True)

    copied = 0
    for path in sorted(SRC.rglob("*")):
        if not path.is_file():
            continue
        if path.name == "Thumbs.db":
            continue
        rel = path.relative_to(SRC)
        b = bucket(classification_rel(SRC, rel))
        dest = DST / b / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, dest)
        copied += 1

    readme = DST / "README_структура_01_Иск.md"
    readme.write_text(
        """# Раздел 01 — Иск (импорт из «Суд ЯНАО»)

Исходная папка: `docs/ДСЗН 2025 по тарифу соцтакси/Суд ЯНАО`

Файлы **скопированы** (не перемещены) в подпапки по назначению:

| Подпапка | Содержание |
|----------|------------|
| `01.0_ЭЦП` | Подписи `.sig` без парного PDF в той же папке исходника; иначе `.sig` лежит рядом с PDF |
| `01.1_Исковые_документы_и_комплекты` | Иск, жалобы, комплекты приложений |
| `01.2_Ходатайства` | Ходатайства |
| `01.3_Определения_и_акты_суда` | Определения суда, номера 89-… |
| `01.4_Отзывы` | Отзывы |
| `01.5_Госпошлина` | Госпошлина, квитанции |
| `01.6_Данные_таблицы` | CSV, XLSX, папка «Заявки» |
| `01.7_Вложения_ГАС` | Attachments*, архивы zip ГАС |
| `01.8_НПА_методические_материалы` | 442-ФЗ, 557–П, приказы, постановления Правительства и т.п. |
| `01.10_Пакеты_документов_ГАС` | Вложенные пакеты «к материалам дела», подтверждения ГАС |
| `01.11_Комплект_Для_суда` | Файлы под деревом пути «для суда», кроме отфильтрованных банком/нарядом ниже |
| `01.12_Путевые_листы_и_транспорт` | Путевые листы и сопутствующие документы в древе исходника |
| `01.13_Банковские_платежи_и_счета` | По имени: `VTB_Payment_commission`, `advanced_payment`, `ON_SCHET` (приоритетнее «для суда») |
| `01.14_Наряд_заказы_СТО` | Наряд-заказы (ремонт и аналоги) по имени файла |
| `01.9_*` | Черновики, беседы, прочее |

Пересборка: `python scripts/import_sud_yanao_to_isk.py` (копирует поверх заново).

""",
        encoding="utf-8",
    )
    print(f"OK: скопировано файлов: {copied} -> {DST}")


if __name__ == "__main__":
    main()
