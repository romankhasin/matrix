#!/usr/bin/env python3
"""
Parse Yandex Metrika "Метки UTM" report exports (tab-separated dumps read
from the report table in the UI — API access is blocked from this
environment) for both counters, aggregate visits/bounce/time-on-site by
canonical platform, and write data/metrika_august2026.csv.

Columns per line, tab-separated:
name, visits, visits%, visitors, visitors%, bounce%, page_depth, time_on_site
"""
import re, csv, sys

def parse_time(s):
    s = s.strip()
    if s in ("< 1 с", "0 с", "0 c"):
        return 0.5 if "<" in s else 0.0
    total = 0.0
    m_min = re.search(r"(\d+)\s*м", s)
    m_sec = re.search(r"(\d+)\s*[cс]", s)
    if m_min:
        total += int(m_min.group(1)) * 60
    if m_sec:
        total += int(m_sec.group(1))
    return total

def parse_num(s):
    return float(s.replace(" ", " ").replace(" ", "").replace(",", "."))

def load_file(path):
    rows = []
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line.strip():
                continue
            parts = line.split("\t")
            if len(parts) < 8:
                continue
            name = parts[0]
            if name == "Итого и средние":
                continue
            visits = int(parts[1].replace(" ", "").replace(" ", ""))
            bounce = parse_num(parts[5].replace("%", "").strip())
            depth = parse_num(parts[6])
            time_on_site = parse_time(parts[7])
            rows.append({"name": name, "visits": visits, "bounce": bounce, "depth": depth, "time": time_on_site})
    return rows

BASE = "/home/claude/matrix/data"
zhk = load_file(f"{BASE}/metrika_raw_zhk_august2026.tsv")
commerce = load_file(f"{BASE}/metrika_raw_commerce_august2026.tsv")

sys.path.insert(0, "/home/claude/matrix/scripts")
from normalize_platforms import normalize_platform

agg = {}
for source_rows, counter in [(zhk, "zhk"), (commerce, "commerce")]:
    for r in source_rows:
        c = "unmapped" if r["name"] == "Не определено" else normalize_platform(r["name"])
        a = agg.setdefault(c, {"visits": 0, "bounce_weighted": 0.0, "time_weighted": 0.0, "depth_weighted": 0.0, "raw": set()})
        a["visits"] += r["visits"]
        a["bounce_weighted"] += r["bounce"] * r["visits"]
        a["time_weighted"] += r["time"] * r["visits"]
        a["depth_weighted"] += r["depth"] * r["visits"]
        a["raw"].add(f"{counter}:{r['name']}")

out = {}
for platform, a in agg.items():
    if a["visits"] == 0:
        continue
    out[platform] = {
        "visits": a["visits"],
        "bounce_rate": round(a["bounce_weighted"] / a["visits"], 2),
        "avg_time_on_site_sec": round(a["time_weighted"] / a["visits"], 1),
        "page_depth": round(a["depth_weighted"] / a["visits"], 2),
        "raw_utm_sources": "; ".join(sorted(a["raw"])),
    }

with open(f"{BASE}/metrika_august2026.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["platform", "visits", "bounce_rate", "avg_time_on_site_sec", "page_depth", "raw_utm_sources"])
    for p in sorted(out, key=lambda k: -out[k]["visits"]):
        d = out[p]
        w.writerow([p, d["visits"], d["bounce_rate"], d["avg_time_on_site_sec"], d["page_depth"], d["raw_utm_sources"]])

print(f"platforms with metrika data: {len(out)}")
