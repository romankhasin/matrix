#!/usr/bin/env python3
"""
Сборка данных для дашборда: расход + конверсии (сделки / встречи / звонки) + постклик,
в разрезе месяц × блок ЖК × площадка.

Блок ЖК определяется по токену в начале UTM-кампании (sp_yandex_..., LM_Медийка_...),
у расходов — по колонке «Проект» (см. PROJECT_BLOCK).

Конверсии и модели атрибуции (модели у разных типов РАЗНЫЕ, это подписано в интерфейсе):
  - Сделка  : «Участие» / «Участие (уник.)»
  - Встреча : «Участие» / «Участие (уник.)»
  - Звонок  : «Первый клик 90 дн» («1 клик 90 дн» / «Первый клик (разрыв 90 дн)») —
              для звонков других моделей в выгрузках нет
Строки сделок/встреч с моделью «первый клик» отбрасываются (та же логика, что раньше:
не смешивать окна атрибуции). Две августовские выгрузки (ПБА — медийка, PBI — программатик)
дополняют друг друга и суммируются, как и раньше со сделками.

Usage:
    python3 build_dashboard_data.py \
        --conv-mayjul "<Свод_ПБА_май-июль_2026.xlsx>" --conv-aug "<Свод_конверсий BI_август_2026.xlsx>" \
        --spend ../data/spend_source.xlsx --metrika-dir ../data \
        --metrika-utm "<Метки UTM ЖК.csv>" "<Метки UTM Коммерция.csv>" \
        --budget "<Бюджеты свод 25-26.xlsx>" \
        --out ../data/dashboard_data.json --inject ../dashboard/index.html ../docs/index.html
"""

import argparse
import csv
import datetime
import json
import re
import warnings
from collections import defaultdict

import openpyxl

from normalize_platforms import ALIASES, normalize_platform

warnings.filterwarnings("ignore")

BLOCKS = [
    ("comfort", "Комфорт"),
    ("comfortplus", "Комфорт+"),
    ("business", "Бизнес"),
    ("deluxe", "Делюкс"),
    ("commerce", "Коммерция"),
    ("brand", "Бренд"),
    ("none", "Без блока"),
]

# токен UTM-кампании -> блок
TOKEN_BLOCK = {
    # бизнес
    "ak": "business", "lb": "business", "lv": "business", "lvol-rnn": "business", "lvol": "business",
    "pc": "business", "pr": "business",
    # комфорт
    "gvl": "comfort", "lmech-rspb": "comfort", "lmech": "comfort", "ng": "comfort",
    "sel": "comfort", "sp": "comfort",
    "ln": "comfort", "ls": "comfort", "ls3": "comfort",  # Нагатинская, Стрешнево
    # комфорт+
    "zv": "comfortplus", "lm": "comfortplus",
    # делюкс
    "sav": "deluxe", "sav17": "deluxe", "sav27": "deluxe",
    # коммерция (Воронцовская, Нижегородская W, Свободы)
    "vr": "commerce", "lwnizh-rmsk": "commerce", "lwnizh": "commerce", "lnizh-rmsk": "commerce",
    "lwng": "commerce", "lwsv": "commerce",
    # бренд (Левел Групп, остатки, премиум, регионы, зонтик)
    "lg": "brand",
}

# название ЖК из «ЖК <название> | ...» в кампании — запасной путь, если токен не распознан
NAME_BLOCK = {
    "причальный": "business", "мичуринский": "comfortplus", "волга-нн": "business",
    "саввинская 17": "deluxe", "саввинская 27": "deluxe",
    "стрешнево": "comfort", "нагатинская": "comfort",
}

# опечатки и варианты написания проекта -> каноническое название
PROJECT_ALIAS = {
    "южнопоротовая": "южнопортовая",
    "level group": "левел групп", "levelgroup": "левел групп",
    "саввинские": "саввинская",
}

# колонка «Проект» в файле расходов -> блок
PROJECT_BLOCK = {
    "мичуринский": "comfortplus", "звенигородская": "comfortplus",
    "южнопортовая": "comfort", "лесной": "comfort", "мечникова": "comfort",
    "нижегородская": "comfort", "селигерская": "comfort",
    "нагатинская": "comfort", "стрешнево": "comfort",
    "павелецкая сити": "business", "войковская": "business", "волга": "business",
    "академическая": "business", "бауманская": "business", "причальный": "business",
    "саввинская": "deluxe", "саввинская 17": "deluxe", "саввинская 27": "deluxe",
    "воронцовская": "commerce", "work воронцовская": "commerce",
    "нижегородская work": "commerce", "work нижегородская": "commerce",
    "свободы": "commerce", "work свободы": "commerce",
    "левел групп": "brand", "левел групп (немарк)": "brand", "остатки": "brand",
    "регионы": "brand", "премиальная коллекция": "brand", "level зонтик москва": "brand",
}

MONTHS = [
    ("may2026", "Май 2026", "Май 2026", "may"),
    ("jun2026", "Июнь 2026", "Июнь 2026", "jun"),
    ("jul2026", "Июль 2026", "Июль 2026", "jul"),
    ("aug2026", "Август 2026", "Август 2026", "august"),
]
METRIKA_FILE = {"may": "metrika_may2026.csv", "jun": "metrika_jun2026.csv",
                "jul": "metrika_jul2026.csv", "august": "metrika_august2026.csv"}

COST_CANDIDATES = ["Расход с НДС и АК с учетом компенсации", "Расход с НДС и АК", "Расход с НДС", "Расход без НДС"]
ZW = re.compile(r"[​‌‍﻿]")

# кампании старых лет (например «..._oct23», «..._sep24») отбрасываются: их сделки/встречи/звонки
# не отражают текущее размещение и искажают рейтинг актуальных кампаний
STALE_CAMPAIGN_YEARS = re.compile(r"_(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(23|24)(?:[_\W]|$)", re.I)


def is_stale_campaign(campaign):
    return bool(STALE_CAMPAIGN_YEARS.search(campaign or ""))


def num(v):
    return v if isinstance(v, (int, float)) else 0.0


# ---------------------------------------------------------------- конверсии

def conv_kind(ctype, model):
    """('deal'|'meet'|'call' | None). None — строка не входит в выбранные модели."""
    model = (model or "").lower()
    if ctype == "Звонок":
        return "call" if ("клик" in model) else None
    if "участие" not in model:
        return None
    return {"Сделка": "deal", "Встреча": "meet"}.get(ctype)


def campaign_token(campaign):
    s = ZW.sub("", campaign or "").replace("Comagic |||", "").strip()
    seg = s.split("|")[-1].strip() if "|" in s else s
    first = re.split(r"[_ /]", seg)[0].lower()
    if first in TOKEN_BLOCK:
        return first
    # токен внутри «otclick~cpm~ng_otclick...» / «redllama~cpm~sav_...»
    m = re.search(r"~([a-z0-9-]+)_", seg.lower())
    if m and m.group(1) in TOKEN_BLOCK:
        return m.group(1)
    return first


def campaign_block(campaign, project_in_name):
    tok = campaign_token(campaign)
    if tok in TOKEN_BLOCK:
        return TOKEN_BLOCK[tok], tok
    for text in (project_in_name or "", campaign or ""):
        m = re.search(r"ЖК\s+([^|]+)", text)
        if m and m.group(1).strip().lower() in NAME_BLOCK:
            return NAME_BLOCK[m.group(1).strip().lower()], tok
    m = re.search(r"ЖК\s+([^|]+?)\s*\|", campaign or "")
    if m and m.group(1).strip().lower() in NAME_BLOCK:
        return NAME_BLOCK[m.group(1).strip().lower()], tok
    return "none", tok


def campaign_platform(campaign):
    """Площадка из названия кампании, если колонка «Площадка» не передана."""
    s = ZW.sub("", campaign or "")
    if "Comagic |||" in s:
        parts = s.replace("Comagic |||", "").strip().split("_")
        if len(parts) >= 3 and parts[1].lower() in ("медийка", "программатик"):
            return normalize_platform(parts[2])
        if len(parts) >= 2:
            return normalize_platform(parts[1])
    segs = [x.strip() for x in s.split("|")]
    if len(segs) >= 4:
        return normalize_platform(segs[1])
    return "unmapped"


def load_conversions(path, header_first_row, offset, month_fixed=None):
    """offset=1 для файла май–июль (первая колонка «Месяц» сдвигает тип и модель), 0 для августа."""
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))[header_first_row:]
    out = []
    for r in rows:
        if not r[0]:
            continue
        month = month_fixed or r[0]
        ctype, model, platform, campaign, pname = r[2 + offset], r[3 + offset], r[8], r[9], r[10]
        conv = r[12] or 0
        kind = conv_kind(ctype, model)
        if not kind or is_stale_campaign(campaign):
            continue
        canon = normalize_platform(platform)
        if platform in (None, "Не указан", "Не передан в выгрузке") or canon == "unmapped":
            canon = campaign_platform(campaign)
        block, tok = campaign_block(campaign, pname)
        out.append({"month": month, "kind": kind, "platform": canon, "block": block,
                    "token": tok, "conv": conv, "campaign": campaign})
    return out


# ---------------------------------------------------------------- расходы

def load_spend(path, sheet):
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    header = [h.strip() if isinstance(h, str) else h for h in rows[0]]
    idx = {n: header.index(n) for n in header if n}
    cost_cols = [idx[c] for c in COST_CANDIDATES if c in idx]
    out, unknown_projects = [], set()
    for r in rows[1:]:
        if not r or not r[idx["Площадка"]] or r[idx["Площадка"]] == "ИТОГО":
            continue
        project = (r[idx["Проект"]] or "").strip()
        block = PROJECT_BLOCK.get(project.lower())
        if block is None:
            unknown_projects.add(project)
            block = "none"
        cost = 0.0
        for ci in cost_cols:
            cost = num(r[ci])
            if cost:
                break
        out.append({"platform": normalize_platform(r[idx["Площадка"]]), "block": block, "project": project,
                    "cost": cost, "impr": num(r[idx["Показы"]]), "clicks": num(r[idx["Клики"]])})
    return out, unknown_projects


RU_MONTHS = {"январь": 1, "февраль": 2, "март": 3, "апрель": 4, "май": 5, "июнь": 6,
             "июль": 7, "август": 8, "сентябрь": 9, "октябрь": 10, "ноябрь": 11, "декабрь": 12}
MONTH_SLUG = {1: "jan", 2: "feb", 3: "mar", 4: "apr", 5: "may", 6: "jun", 7: "jul", 8: "aug",
              9: "sep", 10: "oct", 11: "nov", 12: "dec"}
VAT = 1.2  # расход в бюджетном своде без НДС -> приводим к «с НДС»


def to_num(v):
    """Число из ячейки: '312300.00₽', '1 041 582', '134220,23', '-' -> float."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if v is None:
        return 0.0
    t = re.sub(r"[^\d,.\-]", "", str(v)).replace(",", ".")
    try:
        return float(t) if t not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def load_deals_dated(path):
    """Сделки с точной датой закрытия («участия_встречи_звонки_pbi_2025_2026_свод.xlsx», лист
    «участия в сделках»): 6997 строк, каждая — «Сделка», модель «Участие (уник.)». Встреч и звонков
    в файле нет, несмотря на название. Блок ЖК берём из UTM-кампании (campaign_block), а не из
    колонки «Класс ЖК» в файле — она классифицирует иначе (например, Мичуринский там «Комфорт»,
    Звенигородская — «Бизнес»), чтобы не расходиться с остальной разбивкой на блоки."""
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True).worksheets[1]
    out = defaultdict(lambda: defaultdict(int))
    for r in ws.iter_rows(values_only=True):
        platform_raw, date, campaign = r[4], r[2], r[5]
        if not platform_raw or not isinstance(date, datetime.datetime) or is_stale_campaign(campaign):
            continue
        canon = normalize_platform(platform_raw)
        if canon == "unmapped":
            canon = campaign_platform(campaign)
        block, _ = campaign_block(campaign, None)
        conv = r[6] if isinstance(r[6], (int, float)) else 1
        key = f"{MONTH_SLUG[date.month]}{date.year}"
        out[key][(block, canon)] += conv
    return out


# --- «PBi для стратегии тест.xlsx»: сделки, распиханные по колонкам без единой раскладки
# (Excel сам разбил UTM-кампанию на токены построчно, порядок и число колонок гуляют) ---
FORMAT_MAP = {
    "tgb": "tgb",
    "branding": "branding", "brand": "branding",
    "video": "video", "olv": "video", "ctv": "video",
    "banner": "banner", "banners": "banner", "richmedia": "banner",
}
MONTH_TOKEN = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "june": 6,
               "jul": 7, "july": 7, "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12}
MONTH_TOKEN_RE = re.compile(r"^([a-z]+)(\d{2})$", re.I)


def flatten_tokens(row):
    """Разбирает строку на токены: значения колонок (кроме последней — счётчика), дополнительно
    режет по '|' и '~' — Excel иногда кладёт в одну ячейку кусок необработанной строки целиком."""
    toks = []
    for cell in row[:-1]:
        if cell is None:
            continue
        for chunk in str(cell).split("|"):
            for piece in chunk.split("~"):
                piece = piece.strip()
                if piece:
                    toks.append(piece)
    return toks


def token_block(toks):
    t0 = toks[0].strip().lower() if toks else ""
    if t0 in TOKEN_BLOCK:
        return TOKEN_BLOCK[t0], t0
    for t in toks:
        tl = t.strip().lower()
        if tl in TOKEN_BLOCK:
            return TOKEN_BLOCK[tl], tl
    return "none", t0


def token_platform(toks, known_platforms):
    """Ищет площадку по токенам: сначала пары токенов (novostroy+m -> novostroy-m,
    yandex+go -> yandex_go), затем одиночные — и то и другое только среди уже известных
    (уже занесённых в бюджеты/расходы) канонических названий."""
    for i in range(len(toks) - 1):
        for sep in ("-", "_"):
            canon = normalize_platform(f"{toks[i]}{sep}{toks[i + 1]}")
            if canon in known_platforms:
                return canon
    for t in toks:
        if t.strip().lower() in TOKEN_BLOCK:
            continue
        canon = normalize_platform(t)
        if canon in known_platforms:
            return canon
    return None


def token_format(toks):
    for t in toks:
        f = FORMAT_MAP.get(t.strip().lower())
        if f:
            return f
    return None


def token_month_year(toks):
    for t in toks:
        m = MONTH_TOKEN_RE.match(t.strip())
        if m and m.group(1).lower() in MONTH_TOKEN:
            return 2000 + int(m.group(2)), MONTH_TOKEN[m.group(1).lower()]
    return None


def load_deals_tokenized(path, known_platforms):
    """Возвращает (deals_by_month, fmt_by_month): deals_by_month — month_key -> {(block,platform): count},
    fmt_by_month — month_key -> {(block,platform): {"video":n,"banner":n,"tgb":n,"branding":n,"?":n}}.
    Строки с токеном года 2023/2024 отбрасываются (STALE_CAMPAIGN_YEARS — та же логика, что для остальных
    источников конверсий)."""
    ws = openpyxl.load_workbook(path, data_only=True, read_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))[1:]
    deals = defaultdict(lambda: defaultdict(int))
    fmts = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    skipped_plat = skipped_month = skipped_stale = 0
    for r in rows:
        if not any(x is not None for x in r):
            continue
        toks = flatten_tokens(r)
        conv = r[-1] if isinstance(r[-1], (int, float)) else 1
        my = token_month_year(toks)
        if my is None:
            skipped_month += conv
            continue
        year, month = my
        if year in (2023, 2024):
            skipped_stale += conv
            continue
        platform = token_platform(toks, known_platforms)
        if platform is None:
            skipped_plat += conv
            continue
        block, _ = token_block(toks)
        fmt = token_format(toks) or "?"
        key = f"{MONTH_SLUG[month]}{year}"
        deals[key][(block, platform)] += conv
        fmts[key][(block, platform)][fmt] += conv
    print(f"[deals-tokenized] распознано {sum(sum(v.values()) for v in deals.values())}, "
          f"без площадки {skipped_plat}, без месяца {skipped_month}, отброшено (2023/2024) {skipped_stale}")
    return {k: dict(v) for k, v in deals.items()}, {k: {kk: dict(vv) for kk, vv in v.items()} for k, v in fmts.items()}


def load_budget(path, skip_keys=(), deals_by_month=None, fmt_by_month=None):
    """Бюджетный свод («Бюджеты свод 25-26»): листы «<Месяц> <год>», столбцы D/E/G = Расход/показы/клики.
    Листы с другой раскладкой (недельные рабочие) пропускаются. Расход умножается на 1,2 (с НДС)."""
    deals_by_month = deals_by_month or {}
    fmt_by_month = fmt_by_month or {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out, unknown = {}, set()
    seen_keys = set()
    for ws in wb.worksheets:
        parts = ws.title.strip().lower().split()
        if len(parts) != 2 or parts[0] not in RU_MONTHS or not parts[1].isdigit():
            continue
        month, year = RU_MONTHS[parts[0]], int(parts[1])
        rows = list(ws.iter_rows(values_only=True))
        h = [str(x).strip().lower() if x else "" for x in rows[0]]
        if not (h[3] == "расход" and h[4] == "показы" and h[6] == "клики"):
            raise SystemExit(f"Лист {ws.title!r}: в D/E/G не Расход/показы/клики: {h[3:7]}")
        acc = defaultdict(lambda: [0.0, 0.0, 0.0])
        for r in rows[1:]:
            if r[0] is None or not str(r[0]).strip():
                continue
            project = str(r[0]).strip().lower()
            project = PROJECT_ALIAS.get(project, project)
            block = PROJECT_BLOCK.get(project)
            if block is None:
                unknown.add(str(r[0]).strip())
                block = "none"
            a = acc[(block, normalize_platform(str(r[2]).strip()))]
            a[0] += to_num(r[3]) * VAT
            a[1] += to_num(r[4])
            a[2] += to_num(r[6])
        key = f"{MONTH_SLUG[month]}{year}"
        seen_keys.add(key)
        if key in skip_keys:
            continue
        rows_out = {(b, p): {"b": b, "p": p, "cost": round(v[0], 2), "impr": int(v[1]), "clicks": int(v[2]),
                             "deal": 0, "meet": 0, "call": 0} for (b, p), v in acc.items()}
        for (b, p), deals in deals_by_month.get(key, {}).items():
            rows_out.setdefault((b, p), {"b": b, "p": p, "cost": 0.0, "impr": 0, "clicks": 0, "deal": 0, "meet": 0, "call": 0})
            rows_out[(b, p)]["deal"] += deals
        for (b, p), fmt in fmt_by_month.get(key, {}).items():
            if (b, p) in rows_out:
                rows_out[(b, p)]["fmt"] = fmt
        out[key] = {"year": year, "month": month, "rows": [rows_out[k] for k in sorted(rows_out)]}
    # месяцы со сделками, но без листа в бюджетном своде (сейчас — январь 2025): отдельная запись,
    # только сделки, без расхода/показов/кликов
    for key, bp_deals in deals_by_month.items():
        if key in seen_keys or key in skip_keys:
            continue
        m = re.match(r"([a-z]{3})(\d{4})", key)
        month = {v: k for k, v in MONTH_SLUG.items()}[m.group(1)]
        month_fmt = fmt_by_month.get(key, {})
        out[key] = {"year": int(m.group(2)), "month": month,
                    "rows": [{"b": b, "p": p, "cost": 0.0, "impr": 0, "clicks": 0, "deal": d, "meet": 0, "call": 0,
                             **({"fmt": month_fmt[(b, p)]} if month_fmt.get((b, p)) else {})}
                             for (b, p), d in sorted(bp_deals.items())]}
    return out, unknown


def load_metrika(path):
    post = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if row["platform"] == "unmapped":
                continue
            post[row["platform"]] = {"visits": float(row["visits"]),
                                     "bounce": float(row["bounce_rate"]),
                                     "time": float(row["avg_time_on_site_sec"])}
    return post


MONTH_KEY = {"2026-05": "may2026", "2026-06": "jun2026", "2026-07": "jul2026", "2026-08": "aug2026"}


def load_metrika_utm(paths):
    """Метрика «Метки UTM» (UTM Source × UTM Campaign × месяц визита), оба счётчика.
    Блок берётся из токена UTM-кампании; отказы/время — визит-взвешенные по (месяц, блок, площадка).
    Сюда попадают только кампании с токеном ЖК в названии — визиты без токена (например, часть
    Яндекса: Карты, Дзен, Промопейджес) в выгрузке отсутствуют."""
    acc = defaultdict(lambda: [0.0, 0.0, 0.0])
    for path in paths:
        rows = list(csv.reader(open(path, encoding="utf-8-sig")))
        for r in rows[2:]:  # строка 1 — заголовок, строка 2 — «Итого и средние»
            if not r or not r[0] or r[2][:7] not in MONTH_KEY:
                continue
            visits = float(r[3])
            h, m, sec = (int(x) for x in r[7].split(":"))
            block, _ = campaign_block(r[1], None)
            a = acc[(MONTH_KEY[r[2][:7]], block, normalize_platform(r[0]))]
            a[0] += visits
            a[1] += float(r[6]) * visits * 100
            a[2] += (h * 3600 + m * 60 + sec) * visits
    out = defaultdict(dict)
    for (month, block, platform), (v, bw, tw) in acc.items():
        if platform == "unmapped":
            continue
        out[month][f"{block}|{platform}"] = {"visits": int(v), "bounce": round(bw / v, 2), "time": round(tw / v, 1)}
    return out


def campaign_month_year_fallback(toks, fallback_year):
    """Запасной разбор месяца/года — только для метрики без отдельного столбца месяца:
    диапазоны «jan25-may25» (берём первый месяц) и голый месяц без года («discount_feb»,
    год берётся из имени файла — периода отчёта)."""
    for t in toks:
        for part in t.strip().split("-"):
            m = MONTH_TOKEN_RE.match(part)
            if m and m.group(1).lower() in MONTH_TOKEN:
                return 2000 + int(m.group(2)), MONTH_TOKEN[m.group(1).lower()]
    if fallback_year is not None:
        for t in toks:
            tl = t.strip().lower()
            if tl in MONTH_TOKEN:
                return fallback_year, MONTH_TOKEN[tl]
    return None


def load_metrika_campaign(paths, known_platforms):
    """Метрика «Метки UTM» без отдельных колонок площадки/месяца — только UTM Campaign, визиты,
    отказы, время на сайте (xlsx, отчёт уже отфильтрован по периоду отчёта, например
    «Метки UTM-2025-01-01-2025-03-31.xlsx»). Блок, площадка и месяц/год достаются из токенов
    самой кампании — так же, как в load_deals_tokenized. Строки без определяемого месяца или
    площадки, и с годом 2023/2024, отбрасываются.
    Возвращает (post_by_month, postb_by_month) — те же формы, что load_metrika()/load_metrika_utm(),
    только оба сразу из одного источника (постклик по площадке в целом и по блок×площадка)."""
    acc_plat = defaultdict(lambda: [0.0, 0.0, 0.0])
    acc_block = defaultdict(lambda: [0.0, 0.0, 0.0])
    skipped_month = skipped_plat = skipped_stale = 0
    for path in paths:
        # период отчёта (первый год в имени файла) — запасной год для токенов вида «..._discount_feb»
        # без цифр года, и для диапазонов «jan25-may25» (берём первый месяц диапазона)
        fname_year_m = re.search(r"(\d{4})-\d{2}-\d{2}", path)
        fallback_year = int(fname_year_m.group(1)) if fname_year_m else None
        ws = openpyxl.load_workbook(path, data_only=True, read_only=True).worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        header = [str(h).strip() if h else "" for h in rows[0]]
        if header[:4] != ["UTM Campaign", "Визиты", "Отказы", "Время на сайте"]:
            raise SystemExit(f"{path}: неожиданные колонки {header[:4]}")
        for r in rows[1:]:
            campaign = r[0]
            if not campaign or not isinstance(r[1], (int, float)):
                continue
            toks = str(campaign).split("_")
            my = token_month_year(toks) or campaign_month_year_fallback(toks, fallback_year)
            if my is None:
                skipped_month += r[1]
                continue
            year, month = my
            if year in (2023, 2024):
                skipped_stale += r[1]
                continue
            platform = token_platform(toks, known_platforms)
            if platform is None:
                skipped_plat += r[1]
                continue
            block, _ = campaign_block(str(campaign), None)
            visits = float(r[1])
            bounce_pct = float(r[2]) * 100
            hh, mm, ss = (int(x) for x in str(r[3]).split(":"))
            time_sec = hh * 3600 + mm * 60 + ss
            key = f"{MONTH_SLUG[month]}{year}"
            for acc, k in ((acc_plat, (key, platform)), (acc_block, (key, f"{block}|{platform}"))):
                a = acc[k]
                a[0] += visits; a[1] += bounce_pct * visits; a[2] += time_sec * visits
    print(f"[metrika-campaign] визитов распознано {sum(a[0] for a in acc_plat.values()):,.0f}, "
          f"без площадки {skipped_plat:,.0f}, без месяца {skipped_month:,.0f}, отброшено (2023/2024) {skipped_stale:,.0f}")
    post = defaultdict(dict)
    for (key, platform), (v, bw, tw) in acc_plat.items():
        post[key][platform] = {"visits": int(v), "bounce": round(bw / v, 2), "time": round(tw / v, 1)}
    postb = defaultdict(dict)
    for (key, bp), (v, bw, tw) in acc_block.items():
        postb[key][bp] = {"visits": int(v), "bounce": round(bw / v, 2), "time": round(tw / v, 1)}
    return dict(post), dict(postb)


def inject(html_path, payload):
    s = open(html_path, encoding="utf-8").read()
    a, b = "/*DATA:START*/", "/*DATA:END*/"
    i, j = s.index(a) + len(a), s.index(b)
    s = s[:i] + "\nconst DATA = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n" + s[j:]
    open(html_path, "w", encoding="utf-8").write(s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--conv-mayjul", help="Свод_ПБА_май-июль_2026.xlsx — если не передан, май/июнь/июль 2026 идут только с бюджетом (без конверсий)")
    ap.add_argument("--conv-aug", help="Свод_конверсий BI_август_2026.xlsx — если не передан, август 2026 идёт только с бюджетом")
    ap.add_argument("--spend", required=True)
    ap.add_argument("--metrika-dir", help="если не передан, постклик (отказы/время) не подключается")
    ap.add_argument("--metrika-utm", nargs="*", default=[], help="выгрузки Метрики «Метки UTM» с UTM Campaign (оба счётчика)")
    ap.add_argument("--budget", help="бюджетный свод за прошлые месяцы (Бюджеты свод 25-26.xlsx): только расход, показы, клики")
    ap.add_argument("--deals-dated", help="сделки с точной датой (участия_встречи_звонки_pbi_2025_2026_свод.xlsx): заменяет сделки мая-июня 2026 и добавляет сделки к месяцам до мая 2026")
    ap.add_argument("--deals-tokenized", help="сделки с разложенными по колонкам UTM-токенами, включая формат (PBi для стратегии тест.xlsx): площадка определяется по уже занесённым в бюджеты/расходы названиям")
    ap.add_argument("--metrika-campaign", nargs="*", default=[], help="Метрика без колонок площадки/месяца — только UTM Campaign/визиты/отказы/время (Метки UTM-<период>.xlsx, можно несколько файлов за разные периоды)")
    ap.add_argument("--out", required=True)
    ap.add_argument("--inject", nargs="*", default=[])
    args = ap.parse_args()

    convs = []
    if args.conv_mayjul:
        convs += load_conversions(args.conv_mayjul, 9, 1)
    if args.conv_aug:
        convs += load_conversions(args.conv_aug, 5, 0, "Август 2026")
    have_convs = bool(args.conv_mayjul or args.conv_aug)

    deals_dated = load_deals_dated(args.deals_dated) if args.deals_dated else {}
    REPLACE_DEAL_MONTHS = {"may2026", "jun2026"}  # для этих месяцев берём сделки из deals_dated (точная дата) вместо convs
    REPLACE_MONTH_LABELS = {"Май 2026": "may2026", "Июнь 2026": "jun2026"}
    if deals_dated:
        convs = [c for c in convs if not (c["kind"] == "deal" and REPLACE_MONTH_LABELS.get(c["month"]) in REPLACE_DEAL_MONTHS)]

    known_platforms = None

    def get_known_platforms():
        # «уже занесённые площадки» — по названию, а не по формальному алиасу: всё, что уже встречается
        # в расходах (spend_source.xlsx) и в бюджетном своде за этот же запуск
        nonlocal known_platforms
        if known_platforms is None:
            known_platforms = set(ALIASES.values())
            for _, label, _, _ in MONTHS:
                spend_rows, _ = load_spend(args.spend, label)
                known_platforms.update(r["platform"] for r in spend_rows)
            if args.budget:
                harvest, _ = load_budget(args.budget, skip_keys=set())
                for m in harvest.values():
                    known_platforms.update(r["p"] for r in m["rows"])
        return known_platforms

    deals_tokenized, fmt_tokenized = {}, {}
    if args.deals_tokenized:
        deals_tokenized, fmt_tokenized = load_deals_tokenized(args.deals_tokenized, get_known_platforms())

    metrika_campaign_post, metrika_campaign_postb = {}, {}
    if args.metrika_campaign:
        metrika_campaign_post, metrika_campaign_postb = load_metrika_campaign(args.metrika_campaign, get_known_platforms())

    # сводим deals_dated и deals_tokenized в одну надбавку к сделкам (суммируются, если оба покрывают месяц)
    deal_overrides = defaultdict(lambda: defaultdict(int))
    for src in (deals_dated, deals_tokenized):
        for k, v in src.items():
            for bp, c in v.items():
                deal_overrides[k][bp] += c

    postb = load_metrika_utm(args.metrika_utm) if args.metrika_utm else {}
    payload = {"blocks": [{"key": k, "label": l} for k, l in BLOCKS], "months": {}}
    report = []
    for key, label, src_month, slug in MONTHS:
        spend, unknown = load_spend(args.spend, label)
        acc = defaultdict(lambda: {"cost": 0.0, "impr": 0.0, "clicks": 0.0, "deal": 0, "meet": 0, "call": 0})
        for r in spend:
            a = acc[(r["block"], r["platform"])]
            a["cost"] += r["cost"]; a["impr"] += r["impr"]; a["clicks"] += r["clicks"]
        for c in convs:
            if c["month"] == src_month:
                acc[(c["block"], c["platform"])][c["kind"]] += c["conv"]
        for (b, p), deals in deal_overrides.get(key, {}).items():
            acc[(b, p)]["deal"] += deals
        rows = [{"b": b, "p": p, "cost": round(v["cost"], 2), "impr": int(v["impr"]), "clicks": int(v["clicks"]),
                 "deal": v["deal"], "meet": v["meet"], "call": v["call"],
                 **({"fmt": fmt_tokenized[key][(b, p)]} if fmt_tokenized.get(key, {}).get((b, p)) else {})}
                for (b, p), v in sorted(acc.items())]
        month_num = {"may": 5, "jun": 6, "jul": 7, "august": 8}[slug]
        post = metrika_campaign_post.get(key) or (load_metrika(f"{args.metrika_dir}/{METRIKA_FILE[slug]}") if args.metrika_dir else {})
        postb_month = {**postb.get(key, {}), **metrika_campaign_postb.get(key, {})}
        is_full = have_convs or key in metrika_campaign_postb  # рейтинг строится, если есть Метрика (хоть откуда)
        payload["months"][key] = {"label": label.lower(), "year": 2026, "month": month_num, "full": is_full, "rows": rows,
                                  "post": post, "postb": postb_month}
        tot = {k: sum(r[k] for r in rows) for k in ("cost", "deal", "meet", "call")}
        report.append((label, tot, unknown))

    budget_report = []
    if args.budget:
        budget, unknown_projects = load_budget(args.budget, skip_keys=set(payload["months"]), deals_by_month=deal_overrides, fmt_by_month=fmt_tokenized)
        for key, m in sorted(budget.items(), key=lambda kv: (kv[1]["year"], kv[1]["month"])):
            name = next(n for n, i in RU_MONTHS.items() if i == m["month"])
            has_metrika = key in metrika_campaign_postb
            payload["months"][key] = {"label": f"{name} {m['year']}", "year": m["year"], "month": m["month"],
                                      "full": has_metrika, "rows": m["rows"],
                                      "post": metrika_campaign_post.get(key, {}), "postb": metrika_campaign_postb.get(key, {})}
            budget_report.append((key, sum(r["cost"] for r in m["rows"]), sum(r["impr"] for r in m["rows"]),
                                  sum(r["clicks"] for r in m["rows"])))
        if unknown_projects:
            print("Проекты бюджетного свода без блока:", sorted(unknown_projects))
    payload["months"] = dict(sorted(payload["months"].items(), key=lambda kv: (kv[1]["year"], kv[1]["month"])))

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)

    print("Итоги по месяцам (после фильтров):")
    for label, tot, unknown in report:
        print(f"  {label}: расход {tot['cost']:,.0f}, сделки {tot['deal']}, встречи {tot['meet']}, звонки {tot['call']}"
              + (f"  | проекты без блока в расходах: {sorted(unknown)}" if unknown else ""))
    for key, cost, impr, clicks in budget_report:
        print(f"  бюджет {key}: расход с НДС {cost:,.0f}, показы {impr:,}, клики {clicks:,}")
    print("Конверсии по блокам (все месяцы):")
    byb = defaultdict(lambda: defaultdict(int))
    for c in convs:
        byb[c["block"]][c["kind"]] += c["conv"]
    for b, l in BLOCKS:
        print(f"  {l}: {dict(byb[b])}")
    toks = defaultdict(int)
    for c in convs:
        if c["block"] == "none":
            toks[c["token"]] += c["conv"]
    print("Токены без блока:", dict(sorted(toks.items(), key=lambda kv: -kv[1])))
    unm = defaultdict(int)
    for c in convs:
        if c["platform"] == "unmapped":
            unm[c["kind"]] += c["conv"]
    print("Конверсии без определённой площадки:", dict(unm))

    for p in args.inject:
        inject(p, payload)
        print("-> data injected:", p)


if __name__ == "__main__":
    main()
