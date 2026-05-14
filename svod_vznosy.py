import pandas as pd
from pathlib import Path


def to_number(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0)


def is_person_name(value):
    text = str(value).strip()
    if not text or text in {"nan", "Сотрудник"}:
        return False
    stop_words = [
        "2024",
        "Налоги",
        "Период",
        "Организация",
        "Начислено",
        "НДФЛ",
        "Единый",
        "Итого",
        "Список",
    ]
    if any(word in text for word in stop_words):
        return False
    if "РКООИ" in text or "ТАГАНАЙ" in text.upper():
        return False
    return len(text.split()) >= 3


# 1. Путь к папке и файлу
base_dir = Path(r"C:\taganay_svod_2024")
src_file = base_dir / "Nalogi-i-vznosy-1.xlsx"  # здесь имя твоего файла

# 2. Читаем первый лист без шапки: после распознавания PDF заголовки часто внутри листа
df = pd.read_excel(src_file, header=None)

# 3. Оставляем нужные столбцы: ФИО, начислено, единый тариф, несчастные случаи
part = df.iloc[:, [0, 1, 3, 4]].copy()
part.columns = ["ФИО", "Начислено", "Единый_тариф", "Несч_случаи"]
part = part[part["ФИО"].apply(is_person_name)]
part["Начислено"] = to_number(part["Начислено"])
part["Единый_тариф"] = to_number(part["Единый_тариф"])
part["Несч_случаи"] = to_number(part["Несч_случаи"])

# 6. Считаем сумму взносов по сотруднику за этот период
part["Взносы_за_период"] = part["Единый_тариф"].fillna(0) + part["Несч_случаи"].fillna(0)

# 7. Группируем по ФИО (если вдруг один человек встречается дважды в этом файле)
svod = part.groupby("ФИО", as_index=False).agg({
    "Начислено": "sum",
    "Взносы_за_период": "sum"
})

# 8. Сохраняем результат в новый файл
out_file = base_dir / "Vznosy_2024_aprel.xlsx"
svod.to_excel(out_file, index=False)

print("Готово, файл создан:", out_file)