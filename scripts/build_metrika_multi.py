#!/usr/bin/env python3
"""
Combine the zhk + commerce Metrika UTM aggregates (already normalized to
canonical platform keys via normalize_platforms.py inside the browser JS
extraction) into one data/metrika_<month>2026.csv per month, weighting
bounce/time/depth by visits across the two counters -- same shape as
data/metrika_august2026.csv.

Input files: data/metrika_raw2/{zhk,commerce}_<month>2026_agg.tsv
Columns: platform, visits, bounce_rate, avg_time_on_site_sec, page_depth
(already canonical-platform-aggregated per counter; here we just merge
the two counters together).

Usage:
    python3 build_metrika_multi.py --month may2026
"""
import argparse
import csv

BASE = "/home/claude/matrix/data"


def load(path):
    rows = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            parts = line.split("\t")
            if len(parts) < 5:
                continue
            platform, visits, bounce, time_s, depth = parts[:5]
            rows[platform] = {
                "visits": int(visits),
                "bounce": float(bounce),
                "time": float(time_s),
                "depth": float(depth),
            }
    return rows


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", required=True, help="e.g. may2026")
    args = parser.parse_args()

    zhk = load(f"{BASE}/metrika_raw2/zhk_{args.month}_agg.tsv")
    commerce = load(f"{BASE}/metrika_raw2/commerce_{args.month}_agg.tsv")

    agg = {}
    for source, counter in [(zhk, "zhk"), (commerce, "commerce")]:
        for platform, d in source.items():
            a = agg.setdefault(platform, {"visits": 0, "bounceW": 0.0, "timeW": 0.0, "depthW": 0.0, "raw": set()})
            a["visits"] += d["visits"]
            a["bounceW"] += d["bounce"] * d["visits"]
            a["timeW"] += d["time"] * d["visits"]
            a["depthW"] += d["depth"] * d["visits"]
            a["raw"].add(counter)

    out_path = f"{BASE}/metrika_{args.month}.csv"
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["platform", "visits", "bounce_rate", "avg_time_on_site_sec", "page_depth", "counters"])
        for p in sorted(agg, key=lambda k: -agg[k]["visits"]):
            a = agg[p]
            w.writerow([
                p, a["visits"],
                round(a["bounceW"] / a["visits"], 2),
                round(a["timeW"] / a["visits"], 1),
                round(a["depthW"] / a["visits"], 2),
                ";".join(sorted(a["raw"])),
            ])

    print(f"platforms: {len(agg)} -> {out_path}")


if __name__ == "__main__":
    main()
