"""Aggregate %LOCALAPPDATA%\\KashiharaCity\\pptcom\\com_usage.jsonl (one line per executed COM op,
written by ComProxyEngine) into a per-(type, member) table.

usage: python usage_report.py [--src addin|rig] [--since 2026-09-28] [--csv out.csv]

Columns: calls / ok / failed / blocked, plus the most recent error text for failures.
The type is the COM interface of the object the member was called on (from ITypeInfo).
"""
import argparse
import collections
import csv
import io
import json
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
LOG = os.path.expandvars(r"%LOCALAPPDATA%\KashiharaCity\pptcom\com_usage.jsonl")

ap = argparse.ArgumentParser()
ap.add_argument("--src")
ap.add_argument("--since")
ap.add_argument("--csv")
args = ap.parse_args()

stats = collections.defaultdict(lambda: {"calls": 0, "ok": 0, "failed": 0, "blocked": 0, "last_error": ""})
total = 0
for line in open(LOG, encoding="utf-8"):
    line = line.strip()
    if not line:
        continue
    r = json.loads(line)
    if args.src and r.get("src") != args.src:
        continue
    if args.since and r.get("ts", "") < args.since:
        continue
    total += 1
    s = stats[(r.get("type", "?"), r.get("member", ""))]
    s["calls"] += 1
    if r.get("ok"):
        s["ok"] += 1
    elif r.get("blocked"):
        s["blocked"] += 1
        s["last_error"] = r.get("error", "")
    else:
        s["failed"] += 1
        s["last_error"] = r.get("error", "")

rows = sorted(stats.items(), key=lambda kv: -kv[1]["calls"])
types = sorted({t for t, _ in stats})
print("ops: %d  distinct type.member: %d  types: %d" % (total, len(stats), len(types)))
print("%-28s %-26s %6s %6s %6s %6s  %s" % ("type", "member", "calls", "ok", "fail", "block", "last error"))
for (t, m), s in rows:
    print("%-28s %-26s %6d %6d %6d %6d  %s" % (t, m, s["calls"], s["ok"], s["failed"], s["blocked"], s["last_error"][:70]))
print("\ntypes touched: " + ", ".join(types))

if args.csv:
    with open(args.csv, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["type", "member", "calls", "ok", "failed", "blocked", "last_error"])
        for (t, m), s in rows:
            w.writerow([t, m, s["calls"], s["ok"], s["failed"], s["blocked"], s["last_error"]])
