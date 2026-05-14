"""
Копирует выгрузки ОСВ за 2025 год из папки бухгалтерии в отдельный каталог репозитория
`_extract_osv_2025` (не путать с `_extract_osv`, куда кладут «боевые» имена для скриптов).

Выбор источника при нескольких кандидатах на один логический файл:
  1) в имени есть подстрока «верно» (в т.ч. «вернопо…» содержит «верно»);
  2) иначе более новая дата изменения;
  3) при равенстве — лексикографически последнее имя (стабильный тай-брейк).

Источник: E:\\ЦИП Таганай\\6-БУХГАЛТЕРИЯ\\Оборотно-сальдовые ведомости
Переопределение: BUHGALTERIYA_OSV_DIR

Запуск из корня репозитория:
  python scripts/sync_osv_2025_folder.py
"""

from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
DEST = BASE / "_extract_osv_2025"
DEFAULT_SRC = Path(r"E:\ЦИП Таганай\6-БУХГАЛТЕРИЯ\Оборотно-сальдовые ведомости")


def has_verno(name: str) -> bool:
    return "верно" in name.casefold()


def score_path(p: Path) -> tuple[int, float, str]:
    st = p.stat()
    return (1 if has_verno(p.name) else 0, st.st_mtime, p.name)


def pick_best(candidates: list[Path], label: str) -> Path | None:
    if not candidates:
        print(f"  {label}: нет файлов")
        return None
    candidates = sorted(candidates, key=score_path, reverse=True)
    best = candidates[0]
    v, m, _ = score_path(best)
    print(
        f"  {label}: {best.name!r} "
        f"(«верно»={'да' if v else 'нет'}, {datetime.fromtimestamp(m)})"
    )
    return best


def main() -> None:
    src = Path(os.environ.get("BUHGALTERIYA_OSV_DIR", str(DEFAULT_SRC)))
    if not src.is_dir():
        raise SystemExit(f"Нет папки-источника: {src}")

    DEST.mkdir(parents=True, exist_ok=True)
    allf = [p for p in src.iterdir() if p.is_file()]

    def by_suffix(sfx: str) -> list[Path]:
        return [p for p in allf if p.suffix.casefold() == sfx.casefold()]

    xls = by_suffix(".xls")
    xlsx = by_suffix(".xlsx")

    # Уточнение для 60/70/76/62: не брать файл «январь 2024 — декабрь 2025» как ОСВ только 2025
    def acc_clean(paths: list[Path], acc: str) -> list[Path]:
        res = []
        for p in paths:
            n = p.name.casefold()
            if "январ" in n and "2024" in n and "2025" in n:
                continue
            if acc in p.name:
                res.append(p)
        return res

    c70 = acc_clean([p for p in xls if "70" in p.name and "2025" in p.name], "70")
    c60 = acc_clean([p for p in xls if "60" in p.name and "2025" in p.name], "60")
    c76 = acc_clean([p for p in xls if "76" in p.name and "2025" in p.name], "76")
    c62 = acc_clean([p for p in xls if "62" in p.name and "2025" in p.name], "62")

    print(f"Источник: {src}\nНазначение: {DEST}\n")
    print("Выбор файлов по счетам (.xls):")
    p70 = pick_best(c70, "70")
    p60 = pick_best(c60, "60")
    p76 = pick_best(c76, "76")
    p62 = pick_best(c62, "62")

    mapping: list[tuple[Path, str]] = []
    if p70:
        mapping.append((p70, "Osv_schet_70_2025.xls"))
    if p60:
        mapping.append((p60, "Osv_schet_60_2025.xls"))
    if p76:
        mapping.append((p76, "Osv_schet_76_2025.xls"))
    if p62:
        mapping.append((p62, "Osv_schet_62_2025.xls"))

    # Длинная ОСВ 60: xlsx с 2024 и 2025 в имени
    long60 = [p for p in xlsx if "60" in p.name and "2024" in p.name and "2025" in p.name]
    pl = pick_best(long60, "60 (сквозная 2024–2025, .xlsx)")
    if pl:
        mapping.append((pl, "Osv_schet_60_2024-2025.xlsx"))

    # Полная ОСВ за календарный 2025 (без привязки к номеру счёта в фильтре)
    full2025: list[Path] = []
    for p in xlsx:
        n = p.name.casefold()
        if "2025" not in n:
            continue
        if "2024" in n and "2025" in n:
            continue
        if "ведом" not in n and "оборот" not in n:
            continue
        if "по счету" in n or "по счёту" in n:
            continue
        full2025.append(p)
    pf = pick_best(full2025, "ОСВ за 2025 г. (целиком, .xlsx)")
    if pf:
        mapping.append((pf, "Oborotnosaldovaya_vedomost_2025.xlsx"))

    lines: list[str] = []
    for srcp, outn in mapping:
        outp = DEST / outn
        shutil.copy2(srcp, outp)
        lines.append(f"{outn}  <-  {srcp.name}")

    manifest = DEST / "_manifest_sources.txt"
    manifest.write_text(
        "\n".join(
            [
                f"Синхронизация: {datetime.now().isoformat(timespec='seconds')}",
                f"Источник: {src}",
                "",
                *lines,
                "",
            ]
        ),
        encoding="utf-8",
    )
    print(f"\nСкопировано файлов: {len(mapping)}")
    print(f"Манифест: {manifest}")


if __name__ == "__main__":
    main()
