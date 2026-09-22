#!/usr/bin/env python3
"""
Fetch data from Yandex Metrika Reporting API for one or more counters
and save the results as CSV files, broken down by UTM source/campaign.

Usage:
    export METRIKA_TOKEN="y0__..."
    python3 fetch_metrika.py --date1 2026-08-01 --date2 2026-08-31

Counters are configured in COUNTERS below.
"""

import argparse
import csv
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

API_URL = "https://api-metrika.yandex.net/stat/v1/data"

# Счётчики Яндекс.Метрики: имя -> ID
COUNTERS = {
    "zhk": 53197618,        # Жилые комплексы
    "commerce": 100470605,  # Коммерция
}

# Метрики, которые нас интересуют для постклик-качества + конверсий
METRICS = [
    "ym:s:visits",
    "ym:s:bounceRate",           # % отказов
    "ym:s:avgVisitDurationSeconds",  # среднее время на сайте
    "ym:s:pageDepth",
    "ym:s:goal<goal_id>conversionRate",  # заменяется ниже на конкретные цели, либо считаем сумму по всем целям
]

# Группировка по UTM source/campaign
DIMENSIONS = ["ym:s:lastSignUTMSource", "ym:s:lastSignUTMCampaign"]


def fetch(counter_id: int, token: str, date1: str, date2: str, metrics, dimensions, retries=3):
    params = {
        "ids": counter_id,
        "metrics": ",".join(metrics),
        "dimensions": ",".join(dimensions),
        "date1": date1,
        "date2": date2,
        "accuracy": "full",
        "limit": 100000,
    }
    url = f"{API_URL}?{urllib.parse.urlencode(params)}"
    req = urllib.request.Request(url, headers={"Authorization": f"OAuth {token}"})

    last_err = None
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                import json
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="ignore")
            last_err = f"HTTP {e.code}: {body}"
            if e.code in (429, 503):
                time.sleep(2 ** attempt)
                continue
            raise RuntimeError(last_err)
        except urllib.error.URLError as e:
            last_err = str(e)
            time.sleep(2 ** attempt)
    raise RuntimeError(f"Failed after {retries} attempts: {last_err}")


def get_conversion_goals(counter_id: int, token: str):
    """Fetch the list of goals configured for a counter."""
    import json
    url = f"https://api.metrika.yandex.net/management/v1/counter/{counter_id}/goals"
    req = urllib.request.Request(url, headers={"Authorization": f"OAuth {token}"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    return data.get("goals", [])


def save_csv(rows, headers, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)
    print(f"  -> {path} ({len(rows)} rows)")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--date1", required=True, help="YYYY-MM-DD")
    parser.add_argument("--date2", required=True, help="YYYY-MM-DD")
    parser.add_argument("--out", default="data/metrika")
    args = parser.parse_args()

    token = os.environ.get("METRIKA_TOKEN")
    if not token:
        print("ERROR: set METRIKA_TOKEN env var", file=sys.stderr)
        sys.exit(1)

    for name, counter_id in COUNTERS.items():
        print(f"Counter {name} ({counter_id}):")

        # 1. Goals list (needed to build per-goal conversion metrics)
        try:
            goals = get_conversion_goals(counter_id, token)
        except Exception as e:
            print(f"  WARNING: could not fetch goals: {e}")
            goals = []

        goal_metrics = [f"ym:s:goal{g['id']}conversionRate" for g in goals]
        goal_metrics += [f"ym:s:goal{g['id']}reaches" for g in goals]

        base_metrics = [
            "ym:s:visits",
            "ym:s:bounceRate",
            "ym:s:avgVisitDurationSeconds",
            "ym:s:pageDepth",
        ]
        metrics = base_metrics + goal_metrics

        try:
            result = fetch(counter_id, token, args.date1, args.date2, metrics, DIMENSIONS)
        except Exception as e:
            print(f"  ERROR fetching data: {e}")
            continue

        data = result.get("data", [])
        rows = []
        for item in data:
            dims = [d.get("name", "(none)") for d in item.get("dimensions", [])]
            vals = item.get("metrics", [])
            rows.append(dims + vals)

        headers = ["utm_source", "utm_campaign"] + metrics
        out_path = f"{args.out}_{name}.csv"
        save_csv(rows, headers, out_path)

        # save goal id -> name mapping for later joins
        if goals:
            goal_map_path = f"{args.out}_{name}_goals.csv"
            save_csv(
                [[g["id"], g["name"]] for g in goals],
                ["goal_id", "goal_name"],
                goal_map_path,
            )


if __name__ == "__main__":
    main()
