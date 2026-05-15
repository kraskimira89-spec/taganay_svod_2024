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

Для --watch: pip install watchdog. После стартовой полной пересборки при изменениях
обновляются только сводные ведомости по цепочке папок от затронутого пути до корня дела;
страницы PDF берутся из кэша `.delo_vedomosti_pdf_pages.json`, если файл не менялся.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from collections.abc import Iterable
import threading
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
CASE_ROOT_DEFAULT = BASE / "Дело3а-78-2026Таганай"
OUTPUT_NAME = "Сводная_ведомость.xlsx"
PDF_PAGES_CACHE_NAME = ".delo_vedomosti_pdf_pages.json"
SUBTITLE = "Дело № 3а-78/2026 · Суд ЯНАО · РКООИ ЦИП «Таганай»"

sys.path.insert(0, str(Path(__file__).resolve().parent))

from sozdat_vedomost_pdf import (  # noqa: E402
    INVENTORY_SUFFIXES,
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
    uniq = sorted({p.resolve() for p in paths}, key=lambda p: str(p).lower())
    records = document_records_for_paths(uniq, anchor, page_cache, case_root)
    create_xlsx_inventory(
        records,
        anchor,
        out,
        _title_for_anchor(anchor, case_root),
        "Путь от корня дела",
        SUBTITLE,
    )


def _pdf_cache_path(case_root: Path) -> Path:
    return case_root.resolve() / PDF_PAGES_CACHE_NAME


def load_pdf_pages_disk(case_root: Path) -> dict[str, dict]:
    """На диске: относительный posix-путь → {mtime_ns, size, pages}."""
    path = _pdf_cache_path(case_root)
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return dict(data.get("pdfs", {}))
    except Exception:
        return {}


def save_pdf_pages_disk(case_root: Path, pdfs: dict[str, dict]) -> None:
    path = _pdf_cache_path(case_root)
    path.write_text(
        json.dumps({"version": 1, "pdfs": pdfs}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def build_pdf_page_map(
    pdf_paths: Iterable[Path],
    case_root: Path,
    disk: dict[str, dict],
) -> dict[Path, int | str]:
    """Словарь resolved Path → число страниц; disk обновляется для новых/изменённых PDF."""
    cr = case_root.resolve()
    out: dict[Path, int | str] = {}
    for p in pdf_paths:
        pr = p.resolve()
        rel = pr.relative_to(cr).as_posix()
        st = pr.stat()
        ent = disk.get(rel)
        if (
            ent
            and int(ent.get("mtime_ns", -1)) == st.st_mtime_ns
            and int(ent.get("size", -1)) == st.st_size
        ):
            out[pr] = ent["pages"]
        else:
            pages = pdf_page_count(pr)
            disk[rel] = {
                "mtime_ns": st.st_mtime_ns,
                "size": st.st_size,
                "pages": pages,
            }
            out[pr] = pages
    return out


def dirty_anchors_for_rel(case_root: Path, rel_posix: str) -> set[Path]:
    """
    Каталоги, чья сводная ведомость зависит от изменения по rel_posix
    (файл учёта или каталог в дереве дела): цепочка предков до корня дела.
    """
    cr = case_root.resolve()
    rel_path = Path(rel_posix.replace("\\", "/"))
    parts = rel_path.parts
    out: set[Path] = {cr}
    if not parts:
        return out
    last_suffix = Path(parts[-1]).suffix.lower()
    is_accounting_file = last_suffix in INVENTORY_SUFFIXES
    dir_parts = parts[:-1] if is_accounting_file else parts
    cur = cr
    for part in dir_parts:
        cur = (cur / part).resolve()
        out.add(cur)
    return out


def rebuild_incremental(
    case_root: Path,
    dirty_rels: set[str],
    disk: dict[str, dict],
) -> int:
    """Перезаписать ведомости только по цепочке якорей от изменённых путей. Число якорей."""
    cr = case_root.resolve()
    anchors: set[Path] = set()
    for rel in dirty_rels:
        anchors |= dirty_anchors_for_rel(cr, rel)
    for anchor in sorted(anchors, key=lambda p: (len(p.parts), str(p).casefold())):
        paths = list(iter_inventory_files(anchor))
        pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
        page_cache = build_pdf_page_map(pdfs, cr, disk)
        _write_anchor(anchor, cr, paths, page_cache)
    return len(anchors)


def rebuild_all(case_root: Path) -> tuple[int, int]:
    """Полная пересборка. Возвращает (число папок с ведомостью, учётных файлов)."""
    case_root = case_root.resolve()
    if not case_root.is_dir():
        raise FileNotFoundError(case_root)

    disk = load_pdf_pages_disk(case_root)
    all_files = list(iter_inventory_files(case_root))
    pdf_files = [p for p in all_files if p.suffix.lower() == ".pdf"]
    cur_rels = {p.resolve().relative_to(case_root).as_posix() for p in pdf_files}
    for key in list(disk.keys()):
        if key not in cur_rels:
            del disk[key]

    page_cache = build_pdf_page_map(pdf_files, case_root, disk)
    save_pdf_pages_disk(case_root, disk)

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
    dirty: set[str] = set()

    def job() -> None:
        try:
            with lock:
                batch = set(dirty)
                dirty.clear()
            if not batch:
                return
            disk = load_pdf_pages_disk(case_root)
            n_anc = rebuild_incremental(case_root, batch, disk)
            save_pdf_pages_disk(case_root, disk)
            print(
                f"[{OUTPUT_NAME}] инкремент: каталогов-якорей {n_anc}, "
                f"событий по путям {len(batch)}"
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
            paths: list[Path] = [Path(event.src_path)]
            if getattr(event, "dest_path", None):
                paths.append(Path(event.dest_path))
            touched = False
            for path in paths:
                if _skip_watch_event(path):
                    continue
                try:
                    rel = path.resolve().relative_to(case_root).as_posix()
                except (ValueError, OSError):
                    continue
                with lock:
                    dirty.add(rel)
                touched = True
            if touched:
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
