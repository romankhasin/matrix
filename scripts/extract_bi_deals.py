#!/usr/bin/env python3
"""
Extract target conversions (deals, "участие" attribution model) from the
client's BI export and aggregate them by canonical platform name.

Filter applied (per agreed logic):
  - Тип конверсии == "Сделка"
  - Модель атрибуции contains "Участие" (excludes "Первый клик 90 дн" / "1 клик 90 дн"
    to avoid mixing attribution windows / double counting with meetings export)

Usage:
    python3 extract_bi_deals.py --in "<путь к xlsx>" --sheet "Сводные данные" --out data/bi_deals.csv
"""

import argparse
import csv
import warnings

import openpyxl

from normalize_platforms import normalize_platform

warnings.filterwarnings("ignore")

COLS = [
    "source_file", "source_tab", "conv_type", "attribution_model",
    "period", "conv_date", "project", "platform_group", "platform_raw",
    "campaign", "project_in_campaign_name", "format", "conversions",
]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="infile", required=True)
    parser.add_argument("--sheet", default="Сводные данные")
    parser.add_argument("--header-row", type=int, default=5, help="1-indexed row with column headers")
    parser.add_argument("--out", default="data/bi_deals.csv")
    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.infile, data_only=True)
    ws = wb[args.sheet]
    rows = list(ws.iter_rows(min_row=args.header_row + 1, values_only=True))

    filtered = [
        dict(zip(COLS, r))
        for r in rows
        if r[2] == "Сделка" and r[3] and "Участие" in r[3]
    ]

    print(f"Всего строк в файле: {len(rows)}")
    print(f"После фильтра (Сделка + модель 'Участие'): {len(filtered)}, "
          f"сумма конверсий: {sum(r['conversions'] or 0 for r in filtered)}")

    agg = {}
    unmapped = set()
    for r in filtered:
        canon = normalize_platform(r["platform_raw"])
        if canon == "unmapped" or canon == (r["platform_raw"] or "").strip().lower().replace(" ", "_"):
            if canon not in ("unmapped",) and r["platform_raw"] not in [None, "Не указан"]:
                unmapped.add((r["platform_raw"], canon))
        agg.setdefault(canon, {"platform": canon, "deals": 0, "raw_names": set()})
        agg[canon]["deals"] += r["conversions"] or 0
        agg[canon]["raw_names"].add(r["platform_raw"])

    out_rows = sorted(agg.values(), key=lambda x: -x["deals"])

    with open(args.out, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["platform", "deals", "raw_names_included"])
        for r in out_rows:
            writer.writerow([r["platform"], r["deals"], "; ".join(sorted(r["raw_names"]))])

    print(f"-> {args.out} ({len(out_rows)} площадок)")
    if unmapped:
        print("\nВНИМАНИЕ: эти названия площадок не найдены в словаре ALIASES "
              "(normalize_platforms.py) и остались как есть — проверьте вручную:")
        for raw, canon in sorted(unmapped):
            print(f"  {raw!r} -> {canon!r}")


if __name__ == "__main__":
    main()
