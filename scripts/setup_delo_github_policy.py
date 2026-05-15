# -*- coding: utf-8 -*-
"""Один раз: колонка «Яндекс.Диск» в описи + строки архивов 11.1, 12.3, 12.4."""

from delo_github_policy import (
    apply_yandex_links_to_opis,
    ensure_opis_yandex_column,
    upsert_opis_archive_rows,
)


def main() -> None:
    ensure_opis_yandex_column()
    upsert_opis_archive_rows()
    n = apply_yandex_links_to_opis()
    print(f"Опись обновлена (ссылок/строк: {n}). Заполните docs/delo/yandex_disk_links.yaml")


if __name__ == "__main__":
    main()
