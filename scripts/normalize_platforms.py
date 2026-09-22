#!/usr/bin/env python3
"""
Canonical platform-name mapping shared by the BI extraction and the
spend/Metrika merge, so all three sources land on the same platform key.

Names in the BI export come in many casings/spellings/format-suffixes
for the same real platform (e.g. "Cian ТГБ", "Cian Брендирование" -> cian).
Update ALIASES as new variants show up in future exports rather than
special-casing them downstream.
"""

import re

# raw name (lowercased, trimmed) -> canonical platform key
ALIASES = {
    # Циан
    "cian": "cian", "cian тгб": "cian", "cian брендирование": "cian",

    # Яндекс — все форматы медийки/видео на Яндексе считаем одной площадкой;
    # yandex_go и yandexmarket — отдельные продукты, не сливаем с "yandex"
    "yandex": "yandex", "video yandex": "yandex", "media yandex": "yandex",
    "yandex_go": "yandex_go",
    "yandexmarket": "yandexmarket",
    "yd_rtb": "yandex_rtb",

    # Novostroy семья ресурсов — разные сайты, не сливать друг с другом
    "novostroy-m.ru медийка": "novostroy-m",
    "novostroy_m": "novostroy-m",
    "novostroy-m": "novostroy-m",
    "novostroy-gid медийка": "novostroy-gid",
    "novostroy-gid": "novostroy-gid",
    "novostroycity": "novostroy-city",

    "pronovostroyki": "pronovostroiki",
    "pronovostroiki": "pronovostroiki",

    "adheads": "adheads",

    "move": "move", "move.ru": "move",

    "mts": "mts", "tbank": "tbank", "alfabank": "alfabank",

    "ozon": "ozon",

    "irn": "irn", "irn тгб": "irn",

    "realty": "realty", "cre": "cre",

    "mskguru": "mskguru",
    "rbk медийка": "rbc", "рбк медийка": "rbc",
    "гдеэтотдом медийка": "gdeetotdom", "гдеэтотдом": "gdeetotdom",

    "avito": "avito",

    "mobidriven": "mobidriven",
    "aviasales": "aviasales",
    "thecottage": "thecottage",
    "slickjump": "slickjump",
    "turbotarget": "turbotarget",
    "onetarget": "onetarget",
    "q.bid": "qbid", "qbid": "qbid",
    "dcreo": "dcreo",
    "getintent": "getintent",
    "byyd": "byyd",
    "roxot": "roxot",
    "da": "da",

    "не указан": "unmapped", "не передан в выгрузке": "unmapped",

    # --- варианты из выгрузки расходов (Ежемесячная статистика MI // Level) ---
    "mts.ru olv": "mts",
    "market yandex": "yandexmarket",
    "т-банк": "tbank",
    "яндекс go": "yandex_go",
    "media.yandex": "yandex",
    "media.yandex прайм баннер": "yandex",
    "promopages.yandex.ru": "yandex",
    "video yandex (connected tv)": "yandex",
    "дзен": "yandex",
    "яндекс погода": "yandex",
    "adsheads": "adheads",
    "альфа-мобайл": "alfabank",
    "avito": "avito",
    "wildberries": "wildberries",
    "пятерочка": "pyaterochka",
    "2gis": "2gis",
    "vk ads": "vk_ads",
    "rutube": "rutube",
    # программатик-вендоры/DSP — отдельные площадки, не имеют прямых лидов в BI
    "adspector olv": "adspector",
    "astralab": "astralab",
    "buzzoola": "buzzoola",
    "marketcall": "marketcall",
    "amberdata": "amberdata",
    "vox": "vox",
    "first data": "first_data",
    "сми2": "smi2",
    "soloway": "soloway",
    "tiburon": "tiburon",
    "innovation lab": "innovation_lab",
    "digital alliance": "digital_alliance",
    "digital alliance online cinema pack": "digital_alliance",
    "digital alliance in-stream": "digital_alliance",
}

SUFFIX_STRIP = [" медийка", " тгб", " брендирование"]


def normalize_platform(raw: str) -> str:
    if raw is None:
        return "unmapped"
    key = raw.strip().lower()
    if key in ALIASES:
        return ALIASES[key]
    for suf in SUFFIX_STRIP:
        if key.endswith(suf):
            key = key[: -len(suf)].strip()
            if key in ALIASES:
                return ALIASES[key]
    # fallback: cleaned raw string, flagged for manual review via caller
    return re.sub(r"\s+", "_", key)


if __name__ == "__main__":
    tests = ["Cian ТГБ", "cian", "Novostroy-M.ru медийка", "yandex_go", "Move", "move.ru", "Не указан"]
    for t in tests:
        print(f"{t!r:35} -> {normalize_platform(t)}")
