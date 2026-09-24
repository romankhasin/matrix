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
        --out ../data/dashboard_data.json --inject ../dashboard/index.html ../docs/index.html
"""

import argparse
import csv
import json
import re
import warnings
from collections import defaultdict

import openpyxl

from normalize_platforms import normalize_platform

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
}

# колонка «Проект» в файле расходов -> блок
PROJECT_BLOCK = {
    "мичуринский": "comfortplus", "звенигородская": "comfortplus",
    "южнопортовая": "comfort", "лесной": "comfort", "мечникова": "comfort",
    "нижегородская": "comfort", "селигерская": "comfort",
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
        if not kind:
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


def inject(html_path, payload):
    s = open(html_path, encoding="utf-8").read()
    a, b = "/*DATA:START*/", "/*DATA:END*/"
    i, j = s.index(a) + len(a), s.index(b)
    s = s[:i] + "\nconst DATA = " + json.dumps(payload, ensure_ascii=False, separators=(",", ":")) + ";\n" + s[j:]
    open(html_path, "w", encoding="utf-8").write(s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--conv-mayjul", required=True)
    ap.add_argument("--conv-aug", required=True)
    ap.add_argument("--spend", required=True)
    ap.add_argument("--metrika-dir", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--inject", nargs="*", default=[])
    args = ap.parse_args()

    convs = load_conversions(args.conv_mayjul, 9, 1) + load_conversions(args.conv_aug, 5, 0, "Август 2026")

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
        rows = [{"b": b, "p": p, "cost": round(v["cost"], 2), "impr": int(v["impr"]), "clicks": int(v["clicks"]),
                 "deal": v["deal"], "meet": v["meet"], "call": v["call"]}
                for (b, p), v in sorted(acc.items())]
        payload["months"][key] = {"label": label.lower(), "rows": rows,
                                  "post": load_metrika(f"{args.metrika_dir}/{METRIKA_FILE[slug]}")}
        tot = {k: sum(r[k] for r in rows) for k in ("cost", "deal", "meet", "call")}
        report.append((label, tot, unknown))

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False)

    print("Итоги по месяцам (после фильтров):")
    for label, tot, unknown in report:
        print(f"  {label}: расход {tot['cost']:,.0f}, сделки {tot['deal']}, встречи {tot['meet']}, звонки {tot['call']}"
              + (f"  | проекты без блока в расходах: {sorted(unknown)}" if unknown else ""))
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
