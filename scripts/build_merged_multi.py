#!/usr/bin/env python3
"""
Generalized version of build_merged_august.py for an arbitrary month.

Spend sheet layout varies slightly month to month (May has extra
postclick/CTR columns, June/July match August's compact layout), so
columns are located by header name instead of a fixed index.

Usage:
    python3 build_merged_multi.py --spend ../data/spend_source_may_jul2026.xlsx \
        --sheet "Июнь 2026" --bi ../data/bi_deals_jun2026.csv \
        --out ../data/merged_jun2026.csv
"""

import argparse
import csv
import warnings

import openpyxl

from normalize_platforms import normalize_platform

warnings.filterwarnings("ignore")

COST_CANDIDATES = ["Расход с НДС и АК с учетом компенсации", "Расход с НДС и АК", "Расход с НДС", "Расход без НДС"]


def num(v):
    return v if isinstance(v, (int, float)) else 0.0


def load_spend(path, sheet):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(min_row=1, values_only=True))
    header = [h.strip() if isinstance(h, str) else h for h in rows[0]]
    idx = {name: header.index(name) for name in header if name}

    platform_i = idx["Площадка"]
    project_i = idx["Проект"]
    impr_i = idx.get("Показы")
    clicks_i = idx.get("Клики")
    cost_cols = [idx[c] for c in COST_CANDIDATES if c in idx]

    agg = {}
    for r in rows[1:]:
        if not r or not r[platform_i] or r[platform_i] == "ИТОГО":
            continue
        platform = r[platform_i]
        canon = normalize_platform(platform)
        a = agg.setdefault(canon, {
            "platform": canon, "raw_names": set(), "projects": set(),
            "impressions": 0.0, "clicks": 0.0, "cost": 0.0,
        })
        a["raw_names"].add(platform)
        if project_i is not None and r[project_i]:
            a["projects"].add(r[project_i])
        if impr_i is not None:
            a["impressions"] += num(r[impr_i])
        if clicks_i is not None:
            a["clicks"] += num(r[clicks_i])
        cost = 0.0
        for ci in cost_cols:
            cost = num(r[ci])
            if cost:
                break
        a["cost"] += cost
    return agg


def load_bi(path):
    bi = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            bi[row["platform"]] = int(row["deals"])
    return bi


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--spend", required=True)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--bi", required=True)
    parser.add_argument("--out", required=True)
    args = parser.parse_args()

    spend = load_spend(args.spend, args.sheet)
    bi = load_bi(args.bi)

    all_platforms = set(spend) | set(bi)
    out_rows = []
    for p in all_platforms:
        s = spend.get(p, {"cost": 0, "impressions": 0, "clicks": 0, "raw_names": set(), "projects": set()})
        deals = bi.get(p, 0)
        cost, impr, clicks = s["cost"], s["impressions"], s["clicks"]
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

    with open(args.out, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
        writer.writeheader()
        writer.writerows(out_rows)

    both = [r for r in out_rows if r["has_spend"] and r["has_deals"]]
    spend_only = [r for r in out_rows if r["has_spend"] and not r["has_deals"]]
    deals_only = [r for r in out_rows if r["has_deals"] and not r["has_spend"]]

    print(f"[{args.sheet}] Итого площадок: {len(out_rows)}")
    print(f"  есть и расход, и сделки (можно считать CPA): {len(both)} -> {[r['platform'] for r in both]}")
    print(f"  есть расход, но 0 сделок в BI: {len(spend_only)}")
    print(f"  есть сделки в BI, но 0 расхода в этом файле: {len(deals_only)} -> {[r['platform'] for r in deals_only]}")
    print(f"-> {args.out}")


if __name__ == "__main__":
    main()
