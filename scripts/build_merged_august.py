#!/usr/bin/env python3
"""
Ad-hoc merge for August 2026: spend (from 'Ежемесячная статистика MI // Level',
sheet 'Август 2026') + BI deals (data/bi_deals_august2026.csv).

Metrika (bounce rate / time on site) is NOT included yet — pending API access.
cpa is computed only where BI attributed at least one deal to that platform;
otherwise it's left blank rather than shown as a misleading zero/infinite value.
"""

import csv
import warnings

import openpyxl

from normalize_platforms import normalize_platform

warnings.filterwarnings("ignore")

SPEND_XLSX = "../data/spend_source.xlsx"
SPEND_SHEET = "Август 2026"
BI_CSV = "../data/bi_deals_august2026.csv"
OUT_CSV = "../data/merged_august2026.csv"


def load_spend():
    wb = openpyxl.load_workbook(SPEND_XLSX, data_only=True)
    ws = wb[SPEND_SHEET]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    data_rows = [r for r in rows if r[0] and r[0] != "ИТОГО" and r[2] is not None]

    agg = {}
    for r in data_rows:
        platform, project, impr, clicks, cost_no_vat, cost_vat, cost_vat_ak = r[0], r[1], r[2], r[3], r[4], r[5], r[6]
        canon = normalize_platform(platform)
        a = agg.setdefault(canon, {
            "platform": canon, "raw_names": set(), "projects": set(),
            "impressions": 0.0, "clicks": 0.0, "cost": 0.0,
        })
        def num(v):
            return v if isinstance(v, (int, float)) else 0

        a["raw_names"].add(platform)
        a["projects"].add(project)
        a["impressions"] += num(impr)
        a["clicks"] += num(clicks)
        a["cost"] += num(cost_vat_ak) or num(cost_vat) or num(cost_no_vat)
    return agg


def load_bi():
    bi = {}
    with open(BI_CSV, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            bi[row["platform"]] = int(row["deals"])
    return bi


def main():
    spend = load_spend()
    bi = load_bi()

    all_platforms = set(spend) | set(bi)
    out_rows = []
    for p in all_platforms:
        s = spend.get(p, {"cost": 0, "impressions": 0, "clicks": 0, "raw_names": set(), "projects": set()})
        deals = bi.get(p, 0)
        cost = s["cost"]
        impr = s["impressions"]
        clicks = s["clicks"]
        cpm = (cost / impr * 1000) if impr > 0 else None
        cpc = (cost / clicks) if clicks > 0 else None
        cpa = (cost / deals) if deals > 0 and cost > 0 else None
        out_rows.append({
            "platform": p,
            "cost": round(cost, 2),
            "impressions": int(impr),
            "clicks": int(clicks),
            "cpm": round(cpm, 2) if cpm else "",
            "cpc": round(cpc, 2) if cpc else "",
            "deals_bi": deals,
            "cpa": round(cpa, 2) if cpa else "",
            "has_spend": cost > 0,
            "has_deals": deals > 0,
            "raw_spend_names": "; ".join(sorted(s["raw_names"])),
        })

    out_rows.sort(key=lambda r: (-r["cost"], -r["deals_bi"]))

    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
        writer.writeheader()
        writer.writerows(out_rows)

    both = [r for r in out_rows if r["has_spend"] and r["has_deals"]]
    spend_only = [r for r in out_rows if r["has_spend"] and not r["has_deals"]]
    deals_only = [r for r in out_rows if r["has_deals"] and not r["has_spend"]]

    print(f"Итого площадок: {len(out_rows)}")
    print(f"  есть и расход, и сделки (можно считать CPA): {len(both)} -> {[r['platform'] for r in both]}")
    print(f"  есть расход, но 0 сделок в BI: {len(spend_only)}")
    print(f"  есть сделки в BI, но 0 расхода в этом файле: {len(deals_only)} -> {[r['platform'] for r in deals_only]}")
    print(f"-> {OUT_CSV}")


if __name__ == "__main__":
    main()
