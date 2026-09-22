#!/usr/bin/env python3
"""
Extract target conversions (deals, "Участие" attribution model) from a
*multi-month* BI export ("Свод конверсий за N месяцев") and aggregate them
by (month, canonical platform).

Same filter as extract_bi_deals.py, adapted to the wider column layout of
the combined export (extra "Источник файла" / "Источник (вкладка)" columns
shift Тип конверсии / Модель атрибуции by one position):

  - Тип конверсии == "Сделка"
  - Модель атрибуции contains "Участие" (excludes "1 клик 90 дн" / "Первый
    клик (разрыв 90 дн)")

This does NOT change attribution methodology vs. the August single-month
export: both exclusively use "Участие (уник.)" for Сделка-type rows, so the
filter resolves identically.

Usage:
    python3 extract_bi_deals_multi.py --in "<путь к xlsx>" --sheet "Сводные данные" \
        --header-row 10 --out-prefix data/bi_deals
"""

import argparse
import csv
import warnings
from collections import defaultdict

import openpyxl

from normalize_platforms import normalize_platform, ALIASES, SUFFIX_STRIP


def is_truly_unmapped(raw: str) -> bool:
    """True only if normalize_platform had to fall back to the raw-cleanup
    regex, i.e. the name isn't in ALIASES even after suffix stripping."""
    if raw is None:
        return False
    key = raw.strip().lower()
    if key in ALIASES:
        return False
    for suf in SUFFIX_STRIP:
        if key.endswith(suf) and key[: -len(suf)].strip() in ALIASES:
            return False
    return True

warnings.filterwarnings("ignore")

# 0-indexed columns in the combined multi-month export
COLS = [
    "month", "source_file", "source_tab", "conv_type", "attribution_model",
    "conv_date", "project", "platform_group", "platform_raw",
    "campaign", "project_in_campaign_name", "format", "conversions",
]

MONTH_SLUG = {
    "Январь": "jan", "Февраль": "feb", "Март": "mar", "Апрель": "apr",
    "Май": "may", "Июнь": "jun", "Июль": "jul", "Август": "aug",
    "Сентябрь": "sep", "Октябрь": "oct", "Ноябрь": "nov", "Декабрь": "dec",
}


def month_to_slug(month_str: str) -> str:
    # "Май 2026" -> "may2026"
    parts = (month_str or "").strip().split()
    if len(parts) == 2 and parts[0] in MONTH_SLUG:
        return f"{MONTH_SLUG[parts[0]]}{parts[1]}"
    return (month_str or "unknown").strip().lower().replace(" ", "_")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--in", dest="infile", required=True)
    parser.add_argument("--sheet", default="Сводные данные")
    parser.add_argument("--header-row", type=int, default=10, help="1-indexed row with column headers")
    parser.add_argument("--out-prefix", default="data/bi_deals")
    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.infile, data_only=True)
    ws = wb[args.sheet]
    rows = list(ws.iter_rows(min_row=args.header_row + 1, values_only=True))

    filtered = [
        dict(zip(COLS, r))
        for r in rows
        if len(r) > 4 and r[3] == "Сделка" and r[4] and "Участие" in r[4]
    ]

    print(f"Всего строк в файле: {len(rows)}")
    print(f"После фильтра (Сделка + модель, содержащая 'Участие'): {len(filtered)}, "
          f"сумма конверсий: {sum(r['conversions'] or 0 for r in filtered)}")

    # sanity: confirm no plain "Участие" sneaks in alongside "Участие (уник.)"
    models_seen = {r["attribution_model"] for r in filtered}
    if len(models_seen) > 1:
        print(f"ВНИМАНИЕ: для 'Сделка' встретилось несколько моделей атрибуции: {models_seen} "
              f"— проверьте, не задвояются ли конверсии.")

    by_month = defaultdict(list)
    for r in filtered:
        by_month[r["month"]].append(r)

    all_unmapped = set()
    written_files = []

    for month, month_rows in sorted(by_month.items(), key=lambda kv: kv[0]):
        agg = {}
        for r in month_rows:
            canon = normalize_platform(r["platform_raw"])
            if r["platform_raw"] not in (None, "Не указан", "Не передан в выгрузке") and is_truly_unmapped(r["platform_raw"]):
                all_unmapped.add((r["platform_raw"], canon))
            agg.setdefault(canon, {"platform": canon, "deals": 0, "raw_names": set()})
            agg[canon]["deals"] += r["conversions"] or 0
            agg[canon]["raw_names"].add(r["platform_raw"])

        out_rows = sorted(agg.values(), key=lambda x: -x["deals"])
        out_path = f"{args.out_prefix}_{month_to_slug(month)}.csv"
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["platform", "deals", "raw_names_included"])
            for r in out_rows:
                writer.writerow([r["platform"], r["deals"], "; ".join(sorted(n for n in r["raw_names"] if n))])

        total = sum(r["deals"] for r in out_rows)
        print(f"-> {out_path} ({len(out_rows)} площадок, {total} сделок) [{month}]")
        written_files.append(out_path)

    if all_unmapped:
        print("\nВНИМАНИЕ: эти названия площадок не найдены в словаре ALIASES "
              "(normalize_platforms.py) и остались как есть — проверьте вручную:")
        for raw, canon in sorted(all_unmapped):
            print(f"  {raw!r} -> {canon!r}")


if __name__ == "__main__":
    main()
