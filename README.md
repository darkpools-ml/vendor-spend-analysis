# Vendor Spend Analysis - Crossover Assessment

Author: Obedience Kuguyo

Analysis of 386 vendors ($7.89M, last 12 months) for an acquired company: department, description and Terminate / Consolidate / Optimize recommendation per vendor, the Top 3 savings opportunities, methodology and a CEO/CFO memo. Done with Claude Code.

## Folder structure

| Path | What it is |
|---|---|
| `data/vendor_source.xlsx` | Original spreadsheet, exported unchanged from the Google Sheet provided |
| `data/vendors_raw.csv` | Vendor name + spend extracted from the source (row-numbered) |
| `data/classifications.psv` | Claude Code output: department, action, confidence, category, description per source row |
| `scripts/qa_and_summary.py` | Joins classifications to source, runs QA checks, prints spend by department/action/category |
| `scripts/build_workbook.py` | Calculates Top-3 savings and writes the completed workbook (all 4 tabs) |
| `docs/methodology.md` | Methodology tab text: approach, tools, prompts, QA evidence |
| `docs/executive_memo.md` | One-page CEO/CFO memo (Recommendations tab) |
| `output/vendors_classified.csv` | Final per-vendor table |
| `output/qa_report.txt` | QA results and spend breakdowns (validation evidence) |
| `output/savings_model.txt` | Savings calculation with every assumption |
| `output/Vendor_Analysis_Assessment_completed.xlsx` | Completed workbook, imported into Google Sheets for submission |

## How to reproduce

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl
.venv/bin/python scripts/qa_and_summary.py
.venv/bin/python scripts/build_workbook.py
```

## Process

1. **Extract**: downloaded the provided Google Sheet as xlsx; read the department list from the `Config` tab.
2. **Classify (Claude Code)**: for every vendor Claude Code assigned a department (from the Config list only), a specific one-line description, an action, a spend category and a confidence rating.
3. **QA**: `qa_and_summary.py` checks coverage (386/386), allowed values, description length/specificity and that total spend reconciles exactly ($7,887,360.40). Result: 0 issues. The top 30 vendors (~80% of spend) were spot-checked by hand.
4. **Analyse**: aggregated spend by category to find concentration and duplication.
5. **Size savings**: conservative % assumptions per category (see `output/savings_model.txt`).
6. **Report**: Top 3 Opportunities, Methodology and CEO/CFO memo written into the workbook.

## Key results

- Top 3 opportunities: ~$1.9M annual savings (24% of spend)
  1. Salesforce rightsizing and renegotiation: ~$863K
  2. Office footprint consolidation: ~$510K
  3. Advisor and consultant rationalisation: ~$524K
- Additional travel and tail-spend savings: ~$177K

## Limitations

Descriptions are inferred from vendor names and public knowledge, not contracts. 62 low-confidence vendors (3% of spend) need confirmation with Accounts Payable before action.

## Submitted Google Sheet

https://docs.google.com/spreadsheets/d/1kfViOnIedy0QDn2xTQbFwq6Hb4oK7AWzsRXYmYKIJLg/edit
