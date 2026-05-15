# -*- coding: utf-8 -*-
"""
Сводная ведомость по PDF (рекурсивный обход папки).

Листы: «Ведомость PDF», «Итоги по группам».

Пример:
  python scripts/sozdat_vedomost_pdf.py "docs/Разнорядки 2025" \\
    --title "ВЕДОМОСТЬ ДОКУМЕНТОВ РАЗНАРЯДОК 2025" \\
    --group-label "Месяц / направление" \\
    --output data/delo_registers/12.2_Сводная_ведомость_разнарядок_2025.xlsx
"""

from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent

ARIAL = "Arial"
C_BRAND = "01696F"
C_ROSE = "F5E3EC"

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from pypdf import PdfReader
except ImportError:
    import subprocess

    for pkg in ("openpyxl", "pypdf"):
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from pypdf import PdfReader

THIN = Side(style="thin", color="B0B0B0")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill(start_color=C_BRAND, end_color=C_BRAND, fill_type="solid")
HEADER_FONT = Font(name=ARIAL, size=11, bold=True, color="FFFFFF")


def scan_folder(root_folder: Path) -> list[dict]:
    root = root_folder.resolve()
    if not root.is_dir():
        raise FileNotFoundError(root)

    pdf_files = sorted(set(root.rglob("*.pdf")) | set(root.rglob("*.PDF")))
    if not pdf_files:
        raise FileNotFoundError(f"PDF не найдены: {root}")

    print(f"Сканирую: {root}")
    print(f"Найдено PDF: {len(pdf_files)}")

    records: list[dict] = []
    errors = 0
    total = len(pdf_files)

    for i, pdf_path in enumerate(pdf_files, 1):
        try:
            stat = pdf_path.stat()
            size_kb = round(stat.st_size / 1024, 1)
            mtime = datetime.fromtimestamp(stat.st_mtime)
            rel = pdf_path.relative_to(root)
            parts = rel.parts
            if len(parts) > 1:
                subfolder = " / ".join(parts[:-1])
            else:
                subfolder = "(корень)"

            pages: int | str = "?"
            try:
                pages = len(PdfReader(str(pdf_path)).pages)
            except Exception:
                errors += 1

            records.append(
                {
                    "name": pdf_path.name,
                    "subfolder": subfolder,
                    "size_kb": size_kb,
                    "mtime": mtime,
                    "pages": pages,
                    "full_path": str(pdf_path),
                }
            )
            if i % 50 == 0 or i == total:
                print(f"  {i}/{total} ({i * 100 // total}%)")
        except Exception as exc:
            print(f"  ошибка {pdf_path.name}: {exc}")
            errors += 1

    print(f"Готово: {len(records)} файлов, ошибок: {errors}")
    return records


def create_xlsx(
    records: list[dict],
    root_folder: Path,
    output_path: Path,
    title: str,
    group_label: str,
    subtitle: str,
) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ведомость PDF"

    ws["B2"] = title
    ws["B2"].font = Font(name=ARIAL, size=16, bold=True, color=C_BRAND)
    ws.merge_cells("B2:G2")
    ws["B3"] = subtitle
    ws["B3"].font = Font(name=ARIAL, size=11, italic=True, color="7A7974")
    ws.merge_cells("B3:G3")
    ws["B4"] = f"Источник: {root_folder}"
    ws["B4"].font = Font(name=ARIAL, size=10, italic=True, color="7A7974")
    ws.merge_cells("B4:G4")

    total_kb = sum(r["size_kb"] for r in records)
    total_pages = sum(r["pages"] for r in records if isinstance(r["pages"], int))

    ws["B6"], ws["C6"] = "ВСЕГО ФАЙЛОВ:", len(records)
    ws["D6"], ws["E6"] = "ВСЕГО СТРАНИЦ:", total_pages
    ws["F6"], ws["G6"] = "ОБЩИЙ РАЗМЕР:", f"{round(total_kb / 1024, 1)} МБ"
    ws["H6"], ws["I6"] = "ДАТА:", datetime.now()
    ws["I6"].number_format = "DD.MM.YYYY"
    for c in ("B6", "D6", "F6", "H6"):
        ws[c].font = Font(name=ARIAL, size=11, bold=True, color=C_BRAND)
        ws[c].alignment = Alignment(horizontal="right")

    headers = [
        ("A", "№\nп/п", 6),
        ("B", group_label, 28),
        ("C", "Имя файла", 50),
        ("D", "Размер, КБ", 12),
        ("E", "Дата файла", 14),
        ("F", "Стр.", 8),
        ("G", "Полный путь", 60),
    ]
    for col, htitle, width in headers:
        cell = ws[f"{col}8"]
        cell.value = htitle
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL
        ws.column_dimensions[col].width = width
    ws.row_dimensions[8].height = 32

    sorted_recs = sorted(records, key=lambda r: (r["subfolder"].lower(), r["name"].lower()))
    for i, rec in enumerate(sorted_recs, 1):
        r = 8 + i
        ws[f"A{r}"] = i
        ws[f"A{r}"].alignment = Alignment(horizontal="center")
        ws[f"A{r}"].border = BORDER_ALL
        ws[f"B{r}"] = rec["subfolder"]
        ws[f"B{r}"].font = Font(name=ARIAL, size=10, bold=True)
        ws[f"B{r}"].border = BORDER_ALL
        ws[f"C{r}"] = rec["name"]
        ws[f"C{r}"].font = Font(name="Consolas", size=9)
        ws[f"C{r}"].border = BORDER_ALL
        ws[f"D{r}"] = rec["size_kb"]
        ws[f"D{r}"].number_format = "#,##0.0"
        ws[f"D{r}"].alignment = Alignment(horizontal="right")
        ws[f"D{r}"].border = BORDER_ALL
        ws[f"E{r}"] = rec["mtime"]
        ws[f"E{r}"].number_format = "DD.MM.YYYY HH:MM"
        ws[f"E{r}"].border = BORDER_ALL
        ws[f"F{r}"] = rec["pages"]
        if isinstance(rec["pages"], int):
            ws[f"F{r}"].number_format = "0"
        ws[f"F{r}"].alignment = Alignment(horizontal="center")
        ws[f"F{r}"].border = BORDER_ALL
        if rec["pages"] == "?":
            ws[f"F{r}"].fill = PatternFill(start_color=C_ROSE, end_color=C_ROSE, fill_type="solid")
        ws[f"G{r}"] = rec["full_path"]
        ws[f"G{r}"].font = Font(name="Consolas", size=8, color="7A7974")
        ws[f"G{r}"].border = BORDER_ALL
        ws.row_dimensions[r].height = 22

    ws.freeze_panes = "A9"
    last_row = 8 + len(sorted_recs)
    ws.auto_filter.ref = f"A8:G{last_row}"

    ws2 = wb.create_sheet("Итоги по группам")
    ws2["B2"] = f"ИТОГИ ПО {group_label.upper()}"
    ws2["B2"].font = Font(name=ARIAL, size=16, bold=True, color=C_BRAND)
    ws2.merge_cells("B2:H2")

    headers2 = [
        ("B", "№", 5),
        ("C", group_label, 35),
        ("D", "Файлов", 12),
        ("E", "Стр. (сумма)", 14),
        ("F", "Размер, МБ", 14),
        ("G", "Период от", 18),
        ("H", "Период до", 18),
    ]
    for col, htitle, width in headers2:
        cell = ws2[f"{col}4"]
        cell.value = htitle
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL
        ws2.column_dimensions[col].width = width

    groups: dict[str, dict] = defaultdict(
        lambda: {"files": 0, "kb": 0.0, "pages": 0, "earliest": None, "latest": None}
    )
    for rec in records:
        g = groups[rec["subfolder"]]
        g["files"] += 1
        g["kb"] += rec["size_kb"]
        if isinstance(rec["pages"], int):
            g["pages"] += rec["pages"]
        mt = rec["mtime"]
        if g["earliest"] is None or mt < g["earliest"]:
            g["earliest"] = mt
        if g["latest"] is None or mt > g["latest"]:
            g["latest"] = mt

    for i, (subfolder, stats) in enumerate(
        sorted(groups.items(), key=lambda x: -x[1]["files"]), 1
    ):
        r = 4 + i
        ws2[f"B{r}"] = i
        ws2[f"C{r}"] = subfolder
        ws2[f"D{r}"] = stats["files"]
        ws2[f"E{r}"] = stats["pages"]
        ws2[f"F{r}"] = round(stats["kb"] / 1024, 2)
        ws2[f"G{r}"] = stats["earliest"]
        ws2[f"H{r}"] = stats["latest"]
        for c in "G", "H":
            ws2[f"{c}{r}"].number_format = "DD.MM.YYYY"
        for c in "BCDEFGH":
            ws2[f"{c}{r}"].border = BORDER_ALL

    total_r = 4 + len(groups) + 1
    ws2[f"C{total_r}"] = "ИТОГО:"
    ws2[f"D{total_r}"] = f"=SUM(D5:D{total_r - 1})"
    ws2[f"E{total_r}"] = f"=SUM(E5:E{total_r - 1})"
    ws2[f"F{total_r}"] = f"=SUM(F5:F{total_r - 1})"
    ws2.freeze_panes = "A5"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path


def build_vedomost(
    folder: Path,
    output: Path,
    title: str,
    group_label: str = "Месяц / направление",
    subtitle: str = "Дело № 3а-78/2026 · Суд ЯНАО · РКООИ ЦИП «Таганай»",
    mirror_dir: Path | None = None,
) -> Path:
    records = scan_folder(folder)
    out = create_xlsx(records, folder, output, title, group_label, subtitle)
    if mirror_dir:
        mirror_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        mirror = mirror_dir / f"Ведомость_PDF_{folder.name.replace(' ', '_')}_{stamp}.xlsx"
        import shutil

        shutil.copy2(out, mirror)
        print(f"Копия: {mirror}")

    total_pages = sum(r["pages"] for r in records if isinstance(r["pages"], int))
    groups = Counter(r["subfolder"] for r in records)
    print(f"Файл: {out} ({out.stat().st_size // 1024} КБ)")
    print(f"PDF: {len(records)}, стр.: {total_pages}, групп: {len(groups)}")
    return out


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", type=Path, help="Папка с PDF")
    ap.add_argument("--output", "-o", type=Path, required=True)
    ap.add_argument("--title", "-t", required=True)
    ap.add_argument("--group-label", default="Месяц / направление")
    ap.add_argument("--mirror-dir", type=Path, default=None)
    args = ap.parse_args()
    folder = args.folder if args.folder.is_absolute() else BASE / args.folder
    output = args.output if args.output.is_absolute() else BASE / args.output
    mirror = None
    if args.mirror_dir:
        mirror = args.mirror_dir if args.mirror_dir.is_absolute() else BASE / args.mirror_dir
    build_vedomost(folder, output, args.title, args.group_label, mirror_dir=mirror)


if __name__ == "__main__":
    main()
