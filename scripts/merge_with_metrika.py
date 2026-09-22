#!/usr/bin/env python3
"""
Add Metrika postclick columns (bounce_rate, avg_time_on_site_sec, metrika_visits)
to data/merged_august2026.csv, matched by canonical platform key.
Run after build_merged_august.py and parse_metrika.py.
"""
import csv

BASE = "../data"

metrika = {}
with open(f"{BASE}/metrika_august2026.csv", encoding="utf-8") as f:
    for row in csv.DictReader(f):
        metrika[row["platform"]] = row

with open(f"{BASE}/merged_august2026.csv", encoding="utf-8") as f:
    reader = csv.DictReader(f)
    rows = list(reader)
    base_fields = [fn for fn in reader.fieldnames if fn not in ("bounce_rate", "avg_time_on_site_sec", "metrika_visits")]

new_fields = base_fields + ["bounce_rate", "avg_time_on_site_sec", "metrika_visits"]
matched = 0
for r in rows:
    m = metrika.get(r["platform"])
    if m:
        r["bounce_rate"] = m["bounce_rate"]
        r["avg_time_on_site_sec"] = m["avg_time_on_site_sec"]
        r["metrika_visits"] = m["visits"]
        matched += 1
    else:
        r["bounce_rate"] = ""
        r["avg_time_on_site_sec"] = ""
        r["metrika_visits"] = ""

with open(f"{BASE}/merged_august2026.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=new_fields)
    w.writeheader()
    w.writerows(rows)

print(f"matched {matched}/{len(rows)} platforms with metrika postclick data -> {BASE}/merged_august2026.csv")
