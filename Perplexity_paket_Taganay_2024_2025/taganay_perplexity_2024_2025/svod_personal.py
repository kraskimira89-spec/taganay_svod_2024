import pandas as pd
from pathlib import Path

base_dir = Path(r"C:\taganay_svod_2024")
src_file = base_dir / "Svodnyi-po-ZP-za-2024-god.xlsx"

all_sheets = pd.read_excel(src_file, sheet_name=None)

frames = []

for name, df in all_sheets.items():
    df.columns = [str(c).strip() for c in df.columns]

    # ← сюда подставляем ТОЧНЫЕ названия колонок, которые ты увидел в debug_svod.py
    fio_col = "Фамилия, имя, отчество"
    sum_col = "Начислено всего"

    # проверяем, есть ли такие колонки на листе
    if fio_col not in df.columns or sum_col not in df.columns:
        continue

    part = df[[fio_col, sum_col]].copy()
    part.columns = ["ФИО", "Начислено"]
    frames.append(part)

if not frames:
    raise SystemExit("Не найдено ни одного листа с колонками ФИО и Начислено")

full = pd.concat(frames, ignore_index=True)
full = full.dropna(subset=["ФИО"])
full = full[full["Начислено"].notna()]

svod = full.groupby("ФИО", as_index=False)["Начислено"].sum()
svod = svod.rename(columns={"Начислено": "ФОТ_2024"})

out_file = base_dir / "Personal_2024_auto.xlsx"
svod.to_excel(out_file, index=False)

print("Готово, файл создан:", out_file)