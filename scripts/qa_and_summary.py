"""Join Claude Code classifications to source data, run QA checks, and print spend summaries.

Usage: .venv/bin/python scripts/qa_and_summary.py
Outputs: output/vendors_classified.csv, output/qa_report.txt
"""
import csv
import collections
import openpyxl

SRC = "data/vendor_source.xlsx"
CLS = "data/classifications.psv"
ALLOWED_DEPTS = None  # read from Config tab
ALLOWED_ACTIONS = {"Terminate", "Consolidate", "Optimize"}
VAGUE = ["business services provider", "services provider", "various", "misc"]

wb = openpyxl.load_workbook(SRC)
ALLOWED_DEPTS = {r[0] for r in wb["Config"].iter_rows(min_row=2, values_only=True) if r[0]}
ws = wb["Vendor Analysis Assessment"]
src = {i: (r[0], r[2]) for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2) if r[0]}

with open(CLS) as f:
    cls = {int(r["row"]): r for r in csv.DictReader(f, delimiter="|")}

issues = []
# 1. Coverage: every source row classified exactly once, no extras
missing = sorted(set(src) - set(cls))
extra = sorted(set(cls) - set(src))
if missing: issues.append(f"Missing classification for rows: {missing}")
if extra: issues.append(f"Classification for non-existent rows: {extra}")
# 2. Valid values
for row, c in cls.items():
    if c["department"] not in ALLOWED_DEPTS: issues.append(f"Row {row}: invalid department {c['department']}")
    if c["action"] not in ALLOWED_ACTIONS: issues.append(f"Row {row}: invalid action {c['action']}")
    d = c["description"].strip()
    if len(d) < 15: issues.append(f"Row {row}: description too short")
    if any(v == d.lower() for v in VAGUE): issues.append(f"Row {row}: vague description")
    if len(d) > 160: issues.append(f"Row {row}: description longer than one line ({len(d)} chars)")
# 3. Spend reconciliation
total = sum(c for _, c in src.values())
cls_total = sum(src[r][1] for r in cls if r in src)

lines = [f"Vendors in source: {len(src)}", f"Vendors classified: {len(cls)}",
         f"Source total spend: ${total:,.2f}", f"Classified total spend: ${cls_total:,.2f}",
         f"QA issues: {len(issues)}"] + issues

def agg(key):
    d = collections.defaultdict(lambda: [0, 0.0])
    for r, c in cls.items():
        d[c[key]][0] += 1; d[c[key]][1] += src[r][1]
    return sorted(d.items(), key=lambda x: -x[1][1])

for key in ["department", "action", "category", "confidence"]:
    lines.append(f"\n== Spend by {key} ==")
    for k, (n, s) in agg(key):
        lines.append(f"{k:28s} {n:4d} vendors  ${s:>13,.0f}  {s/total:6.1%}")

with open("output/vendors_classified.csv", "w", newline="") as f:
    w = csv.writer(f)
    w.writerow(["Row", "Vendor Name", "Department", "Last 12 months Cost (USD)", "Description", "Suggestion", "Category", "Confidence"])
    for r in sorted(src):
        c = cls[r]
        w.writerow([r, src[r][0], c["department"], round(src[r][1], 2), c["description"], c["action"], c["category"], c["confidence"]])

open("output/qa_report.txt", "w").write("\n".join(lines) + "\n")
print("\n".join(lines))
