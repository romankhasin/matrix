#!/usr/bin/env python3
"""
Merge three data sources into a single per-platform table:
  1. Ad spend / impressions / clicks (from CSV export, e.g. Google Sheets/Drive)
  2. Yandex Metrika (visits, bounce rate, time on site, conversions)
  3. BI export (calls / meetings), keyed by UTM source/campaign

Output: data/merged.csv with one row per (utm_source, utm_campaign) pair
and all metrics needed by the ranker.

Usage:
    python3 merge_data.py \
        --spend data/spend.csv \
        --metrika data/metrika_zhk.csv data/metrika_commerce.csv \
        --bi data/bi_export.csv \
        --out data/merged.csv

Expected column names (rename your exports to match, or adjust the
COLUMN_MAP dictionaries below):

  spend.csv:    utm_source, utm_campaign, cost, impressions, clicks
  metrika_*.csv: utm_source, utm_campaign, visits, bounceRate,
                 avgVisitDurationSeconds, pageDepth, <goal columns...>
  bi_export.csv: utm_source, utm_campaign, calls, meetings, deals (optional)
"""

import argparse
import csv
from collections import defaultdict


def read_csv(path):
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def norm_key(row, source_col="utm_source", campaign_col="utm_campaign"):
    src = (row.get(source_col) or "").strip().lower()
    camp = (row.get(campaign_col) or "").strip().lower()
    return (src, camp)


def to_float(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--spend", required=True)
    parser.add_argument("--metrika", nargs="+", required=True)
    parser.add_argument("--bi", required=True)
    parser.add_argument("--out", default="data/merged.csv")
    args = parser.parse_args()

    merged = defaultdict(lambda: {
        "utm_source": "", "utm_campaign": "",
        "cost": 0.0, "impressions": 0.0, "clicks": 0.0,
        "visits": 0.0, "bounce_rate": 0.0, "avg_time_on_site": 0.0, "page_depth": 0.0,
        "goal_conversions": 0.0,
        "calls": 0.0, "meetings": 0.0, "deals": 0.0,
    })

    # 1. Spend
    for row in read_csv(args.spend):
        key = norm_key(row)
        m = merged[key]
        m["utm_source"], m["utm_campaign"] = row.get("utm_source", ""), row.get("utm_campaign", "")
        m["cost"] += to_float(row.get("cost"))
        m["impressions"] += to_float(row.get("impressions"))
        m["clicks"] += to_float(row.get("clicks"))

    # 2. Metrika (one or more counters)
    for path in args.metrika:
        for row in read_csv(path):
            key = norm_key(row)
            m = merged[key]
            m["utm_source"], m["utm_campaign"] = row.get("utm_source", ""), row.get("utm_campaign", "")
            m["visits"] += to_float(row.get("ym:s:visits") or row.get("visits"))
            # bounce rate / time on site are rates -> weighted average by visits done later;
            # for now just take the row value (assumes one row per key per file)
            m["bounce_rate"] = to_float(row.get("ym:s:bounceRate") or row.get("bounceRate"))
            m["avg_time_on_site"] = to_float(row.get("ym:s:avgVisitDurationSeconds") or row.get("avgVisitDurationSeconds"))
            m["page_depth"] = to_float(row.get("ym:s:pageDepth") or row.get("pageDepth"))
            # sum any goalXreaches columns as total goal conversions
            for col, val in row.items():
                if col.startswith("ym:s:goal") and col.endswith("reaches"):
                    m["goal_conversions"] += to_float(val)

    # 3. BI export
    for row in read_csv(args.bi):
        key = norm_key(row)
        m = merged[key]
        m["utm_source"], m["utm_campaign"] = row.get("utm_source", ""), row.get("utm_campaign", "")
        m["calls"] += to_float(row.get("calls"))
        m["meetings"] += to_float(row.get("meetings"))
        m["deals"] += to_float(row.get("deals"))

    # Derived metrics
    out_rows = []
    for (src, camp), m in merged.items():
        target_conversions = m["calls"] + m["meetings"]  # main KPI per the brief
        cpa = m["cost"] / target_conversions if target_conversions > 0 else None
        cpm = (m["cost"] / m["impressions"] * 1000) if m["impressions"] > 0 else None
        cpc = (m["cost"] / m["clicks"]) if m["clicks"] > 0 else None
        ctr = (m["clicks"] / m["impressions"] * 100) if m["impressions"] > 0 else None

        out_rows.append({
            "platform": m["utm_source"] or src,
            "campaign": m["utm_campaign"] or camp,
            "cost": round(m["cost"], 2),
            "impressions": int(m["impressions"]),
            "clicks": int(m["clicks"]),
            "cpm": round(cpm, 2) if cpm is not None else "",
            "cpc": round(cpc, 2) if cpc is not None else "",
            "ctr_pct": round(ctr, 2) if ctr is not None else "",
            "visits": int(m["visits"]),
            "bounce_rate_pct": round(m["bounce_rate"], 2),
            "avg_time_on_site_sec": round(m["avg_time_on_site"], 1),
            "page_depth": round(m["page_depth"], 2),
            "site_goal_conversions": int(m["goal_conversions"]),
            "calls": int(m["calls"]),
            "meetings": int(m["meetings"]),
            "deals": int(m["deals"]),
            "target_conversions": int(target_conversions),
            "cpa": round(cpa, 2) if cpa is not None else "",
        })

    out_rows.sort(key=lambda r: (r["platform"], r["campaign"]))

    with open(args.out, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()) if out_rows else [])
        writer.writeheader()
        writer.writerows(out_rows)

    print(f"Merged {len(out_rows)} platform/campaign rows -> {args.out}")


if __name__ == "__main__":
    main()
