# Сводные ведомости для GitHub (дело 3а-78/2026)

В репозиторий **загружаются только** эти Excel-файлы (реестры), а не сами архивы PDF.

| Файл | Приложение | Содержание |
|------|------------|------------|
| `11.2_Сводная_ведомость_поставщиков.xlsx` | 11.2 | Реестр папок поставщиков |
| `12.1_Сводная_ведомость_разнарядок_2024.xlsx` | 12.1 | Реестр PDF разнарядок 2024 |
| `12.2_Сводная_ведомость_разнарядок_2025.xlsx` | 12.2 | Реестр PDF разнарядок 2025 |

**Исходные PDF** (папки `docs/Поставщики`, `docs/Разнорядки 20xx`) хранятся локально и на **Яндекс.Диске**. Ссылки — в `docs/delo/yandex_disk_links.yaml` и в колонке **I** описи `docs/obosnovanie2025/Опись_приложений_шаблон.xlsx`.

Пересборка:

```bash
python scripts/build_postavshchiki_svod.py
python scripts/build_raznaryadki_svod.py
python scripts/setup_delo_github_policy.py
python scripts/apply_yandex_links_to_opis.py
```
