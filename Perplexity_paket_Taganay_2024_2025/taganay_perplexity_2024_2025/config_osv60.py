"""
Правила групп затрат по счёту 60 для соцтакси.

- Константы `G_*` — канонические названия внутренних групп (листы Excel, записка).
- `ARTICLE_KEYWORD_RULES` — сопоставление подстрок из «Статьи расходов» (если колонка есть в файле).
  Порядок важен: первое совпадение выигрывает.
- Короткие ключи (например «сто») сопоставляются как отдельное слово, чтобы не цеплять «стоимость».
"""

from __future__ import annotations

import re

# --- Внутренние группы (как в своде и записке) ---
G_TOP = "Топливо (ГСМ)"
G_ARENDA_AVTO = "Аренда автомобиля"
G_ARENDA_GARAZH = "Аренда гаража/стоянки"
G_REMONT = "Ремонт и обслуживание автомобиля"
G_MOYKA = "Мойка автомобилей"
G_STRAH = "Страхование автомобиля"
G_MED = "Медосмотр водителей"
G_KOMM = "Коммунальные услуги"
G_IT = "Связь, IT и офисные расходы"
G_PROCH = "Прочие эксплуатационные"

# (подстрока_в_нижнем_регистре, внутренняя_группа, прямые_100pct)
# is_direct: True — 100% на соцтакси; False — × доля приходов по соцуслугам (см. `config_allocation.get_soc_taxi_share`)
ARTICLE_KEYWORD_RULES: list[tuple[str, str, bool]] = [
    # Прямые авторасходы
    ("топливо лукойл", G_TOP, True),
    ("топливо", G_TOP, True),
    ("аренда автомобиля", G_ARENDA_AVTO, True),
    ("аренда гаража", G_ARENDA_GARAZH, True),
    ("мойка автомобиля", G_MOYKA, True),
    ("ремонт автомобиля", G_REMONT, True),
    ("масло для автомобиля", G_REMONT, True),
    ("запчасти для автомобиля", G_REMONT, True),
    ("запачасти", G_REMONT, True),  # опечатка в учётной выгрузке
    ("автотовары", G_REMONT, True),
    ("шины", G_REMONT, True),
    ("сто", G_REMONT, True),  # отдельное слово — см. _article_has_keyword
    # Страхование
    ("страховки на автомобиль", G_STRAH, True),
    # Медосмотры
    ("предрейсовый, послерейсовый осмотр", G_MED, True),
    ("послерейсовый осмотр", G_MED, True),
    ("предрейсовый", G_MED, True),
    ("медосмотр", G_MED, True),
    # Коммуналка и эксплуатация помещений
    ("хвс, водоотведение", G_KOMM, False),
    ("водоотведение, тепловая", G_KOMM, False),
    ("тепловая энергия", G_KOMM, False),
    ("электроэнергия", G_KOMM, False),
    ("выввоз мусора", G_KOMM, False),
    ("выврз мусора", G_KOMM, False),  # опечатка в выгрузке
    ("вывоз мусора", G_KOMM, False),
    ("содержание и текущий ремонт мира", G_KOMM, False),
    # Связь, IT, офис
    ("мобильная связь", G_IT, False),
    ("программное обеспечение", G_IT, False),
    ("пограмное обеспечение", G_IT, False),  # опечатка: «пограмное»
    ("пограмное", G_IT, False),
    ("заправка картриджа", G_IT, False),
    ("заправк", G_IT, False),  # «заправкка …» и опечатки
    ("ремонт мфу", G_IT, False),
    # Прочее
    ("охранная сигнализация", G_PROCH, False),
]


def _article_has_keyword(text_lower: str, key_lower: str) -> bool:
    """Для коротких ключей ищем отдельное слово/аббревиатуру."""
    if len(key_lower) <= 4:
        return bool(re.search(rf"(^|[\s,;])({re.escape(key_lower)})([\s,;]|$)", text_lower))
    return key_lower in text_lower


def classify_by_article(article: str | None) -> tuple[str, bool] | None:
    """
    Классификация по тексту статьи расходов.
    Возвращает (внутренняя_группа, прямые_100pct) или None, если правило не сработало.
    """
    if article is None or not isinstance(article, str):
        return None
    text = article.lower().strip()
    if not text:
        return None
    for key, group, is_direct in ARTICLE_KEYWORD_RULES:
        if _article_has_keyword(text, key.lower()):
            return group, is_direct
    return None
