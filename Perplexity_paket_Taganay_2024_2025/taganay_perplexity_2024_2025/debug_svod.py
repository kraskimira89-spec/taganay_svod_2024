import pandas as pd
from pathlib import Path

base_dir = Path(r"C:\taganay_svod_2024")
src_file = base_dir / "Svodnyi-po-ZP-za-2024-god.xlsx"

all_sheets = pd.read_excel(src_file, sheet_name=None)

print("Найдено листов:", len(all_sheets))
for name, df in all_sheets.items():
    print("\n--- Лист:", name, "---")
    df.columns = [str(c).strip() for c in df.columns]
    print("Колонки:")
    for col in df.columns:
        print("  -", col)