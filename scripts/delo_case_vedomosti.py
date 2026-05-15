# -*- coding: utf-8 -*-
"""
Сводные ведомости во всех вложенных папках дела Дело3а-78-2026Таганай.

В **каждую** подпапку дела (включая пустые) записывается `Сводная_ведомость.xlsx`
(листы «Ведомость файлов», «Итоги по группам»). При пустом поддереве —
шапка и нулевые итоги.

Колонка «путь» — от корня дела (см. `document_records_for_paths` + case_root).

Запуск:
  python scripts/delo_case_vedomosti.py              # полная пересборка
  python scripts/delo_case_vedomosti.py --watch        # автообновление (watchdog)

Для --watch: pip install watchdog
"""

from __future__ import annotations

import argparse
import os
import sys
import threading
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CASE_ROOT_DEFAULT = BASE / "Дело3а-78-2026Таганай"
OUTPUT_NAME = "Сводная_ведомость.xlsx"
SUBTITLE = "Дело № 3а-78/2026 · Суд ЯНАО · РКООИ ЦИП «Таганай»"

sys.path.insert(0, str(Path(__file__).resolve().parent))

from sozdat_vedomost_pdf import (  # noqa: E402
    create_xlsx_inventory,
    document_records_for_paths,
    index_anchor_to_files,
    iter_inventory_files,
    pdf_page_count,
)

SKIP_SUBDIR_NAMES = frozenset({".git", "__pycache__", ".svn"})


def iter_all_case_dirs(case_root: Path) -> list[Path]:
    """Все каталоги под корнем дела (сам корень включительно), без служебных веток."""
    cr = case_root.resolve()
    out: list[Path] = []
    for dirpath, dirnames, _files in os.walk(cr, topdown=True):
        dirnames[:] = [
            d
            for d in dirnames
            if d not in SKIP_SUBDIR_NAMES and not d.startswith(".")
        ]
        out.append(Path(dirpath).resolve())
    return sorted(out, key=lambda x: (len(x.parts), str(x).casefold()))


def _title_for_anchor(anchor: Path, case_root: Path) -> str:
    case_root = case_root.resolve()
    anchor = anchor.resolve()
    if anchor == case_root:
        label = "корень дела"
    else:
        label = anchor.relative_to(case_root).as_posix()
    return f"ВЕДОМОСТЬ ФАЙЛОВ · {label}"


def _write_anchor(
    anchor: Path,
    case_root: Path,
    paths: list[Path],
    page_cache: dict[Path, int | str],
) -> None:
    out = anchor / OUTPUT_NAME
    uniq = sorted(set(paths), key=lambda p: str(p).lower())
    records = document_records_for_paths(uniq, anchor, page_cache, case_root)
    create_xlsx_inventory(
        records,
        anchor,
        out,
        _title_for_anchor(anchor, case_root),
        "Путь от корня дела",
        SUBTITLE,
    )


def rebuild_all(case_root: Path) -> tuple[int, int]:
    """Полная пересборка. Возвращает (число папок с ведомостью, учётных файлов)."""
    case_root = case_root.resolve()
    if not case_root.is_dir():
        raise FileNotFoundError(case_root)

    all_files = list(iter_inventory_files(case_root))
    page_cache: dict[Path, int | str] = {}
    for p in all_files:
        if p.suffix.lower() == ".pdf":
            page_cache[p] = pdf_page_count(p)

    buckets = index_anchor_to_files(case_root, all_files)
    all_dirs = iter_all_case_dirs(case_root)

    for anchor in all_dirs:
        paths = buckets.get(anchor, [])
        _write_anchor(anchor, case_root, paths, page_cache)

    return len(all_dirs), len(all_files)


def _skip_watch_event(path: Path) -> bool:
    n = path.name.lower()
    if n == OUTPUT_NAME.lower():
        return True
    if n.startswith("~$"):
        return True
    return False


def watch_case(case_root: Path, debounce_s: float = 2.5) -> None:
    try:
        from watchdog.events import FileSystemEventHandler
        from watchdog.observers import Observer
    except ImportError:
        print("Установите watchdog: pip install watchdog")
        sys.exit(1)

    case_root = case_root.resolve()
    lock = threading.Lock()
    timer: threading.Timer | None = None

    def job() -> None:
        try:
            n, nf = rebuild_all(case_root)
            print(
                f"[{OUTPUT_NAME}] папок с ведомостью: {n}, "
                f"учётных файлов в деле: {nf}"
            )
        except Exception as exc:
            print(f"Ошибка пересборки: {exc}")

    def schedule() -> None:
        nonlocal timer
        with lock:
            if timer is not None:
                timer.cancel()
            timer = threading.Timer(debounce_s, lambda: job())
            timer.daemon = True
            timer.start()

    class _Handler(FileSystemEventHandler):
        def on_any_event(self, event):  # type: ignore[override]
            if event.is_directory:
                schedule()
                return
            path = Path(event.src_path)
            if _skip_watch_event(path):
                return
            schedule()

    obs = Observer()
    obs.schedule(_Handler(), str(case_root), recursive=True)
    obs.start()
    print(f"Наблюдение: {case_root} (debounce {debounce_s} с, Ctrl+C — выход)")
    try:
        n, nf = rebuild_all(case_root)
        print(f"Стартовая пересборка: папок {n}, учётных файлов {nf}")
        while obs.is_alive():
            time.sleep(0.5)
    except KeyboardInterrupt:
        obs.stop()
    obs.join(timeout=5)


def main() -> None:
    ap = argparse.ArgumentParser(description="Сводные ведомости по дереву дела")
    ap.add_argument(
        "--root",
        type=Path,
        default=CASE_ROOT_DEFAULT,
        help="Корень папки дела",
    )
    ap.add_argument(
        "--watch",
        action="store_true",
        help="Автообновление при изменениях (pip install watchdog)",
    )
    args = ap.parse_args()
    root = args.root if args.root.is_absolute() else BASE / args.root

    if args.watch:
        watch_case(root)
        return
    n, nf = rebuild_all(root)
    print(f"OK: папок с ведомостью: {n}, учётных файлов: {nf} -> {root}")


if __name__ == "__main__":
    main()
