# -*- coding: utf-8 -*-
"""Проставить ссылки из docs/delo/yandex_disk_links.yaml в опись (колонка I)."""

from delo_github_policy import apply_yandex_links_to_opis, load_yandex_links


def main() -> None:
    links = load_yandex_links()
    filled = sum(1 for v in links.values() if v and "disk.yandex" in v)
    n = apply_yandex_links_to_opis()
    print(f"Обновлено строк: {n}; заполненных URL в yaml: {filled}")


if __name__ == "__main__":
    main()
