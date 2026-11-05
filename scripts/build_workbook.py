"""Compute Top-3 savings from the classified data and write the completed workbook.

Usage: .venv/bin/python scripts/build_workbook.py
Input:  data/vendor_source.xlsx, output/vendors_classified.csv
Output: output/Vendor_Analysis_Assessment_completed.xlsx, output/savings_model.txt
"""
import csv
import openpyxl
from openpyxl.styles import Alignment, Font

rows = list(csv.DictReader(open("output/vendors_classified.csv")))
for r in rows:
    r["cost"] = float(r["Last 12 months Cost (USD)"])
TOTAL = sum(r["cost"] for r in rows)

def spend(pred):
    return sum(r["cost"] for r in rows if pred(r))

# ---- Savings model (assumptions documented inline) ----
sf = spend(lambda r: r["Category"] == "CRM")
sf_consult = spend(lambda r: r["Category"] == "Salesforce Consulting")
opp1 = 0.25 * sf + 0.40 * sf_consult   # 25% licence rightsizing/renewal, 40% consulting scope cut

office = spend(lambda r: r["Category"] == "Office Space")
opp2 = 0.50 * office                   # exit/downsize to one flex hub per country

audit = spend(lambda r: r["Category"] == "Audit & Tax")
deal = spend(lambda r: r["Category"] == "Deal Advisory")
consult = spend(lambda r: r["Category"] in ("Consulting", "Contractor"))
opp3 = 0.30 * audit + 1.00 * deal + 0.60 * consult  # one audit/tax firm; deal fees do not recur; cut ad-hoc consultants

travel = spend(lambda r: r["Category"] == "Travel")
tail = spend(lambda r: r["Category"] in ("Events", "Events / Catering", "Gifts", "Office Supplies"))
other = 0.20 * travel + 0.70 * tail

model = f"""Total 12-month spend: ${TOTAL:,.0f}
Opp 1 Salesforce: CRM ${sf:,.0f} x25% + Cloudcrossing ${sf_consult:,.0f} x40% = ${opp1:,.0f}
Opp 2 Offices: office space ${office:,.0f} x50% = ${opp2:,.0f}
Opp 3 Advisors: audit/tax ${audit:,.0f} x30% + deal advisory ${deal:,.0f} x100% + consultants/contractors ${consult:,.0f} x60% = ${opp3:,.0f}
Top 3 total: ${opp1+opp2+opp3:,.0f} ({(opp1+opp2+opp3)/TOTAL:.1%} of spend)
Other (not in Top 3): travel ${travel:,.0f} x20% + tail spend ${tail:,.0f} x70% = ${other:,.0f}
Grand total identified: ${opp1+opp2+opp3+other:,.0f} ({(opp1+opp2+opp3+other)/TOTAL:.1%})
"""
open("output/savings_model.txt", "w").write(model)
print(model)

def k(x):  # $ rounded to nearest $1k
    return f"${round(x/1000):,}K"

# ---- Write workbook ----
wb = openpyxl.load_workbook("data/vendor_source.xlsx")
ws = wb["Vendor Analysis Assessment"]
by_row = {int(r["Row"]): r for r in rows}
for i, r in by_row.items():
    assert ws.cell(i, 1).value == r["Vendor Name"], f"row mismatch at {i}"
    ws.cell(i, 2).value = r["Department"]
    ws.cell(i, 4).value = r["Description"]
    ws.cell(i, 5).value = r["Suggestion"]
ws.cell(1, 6).value = "Category"
ws.cell(1, 7).value = "Confidence"
for i, r in by_row.items():
    ws.cell(i, 6).value = r["Category"]
    ws.cell(i, 7).value = r["Confidence"]

top = wb["Top 3 Opportunities"]
top["D1"] = "Estimated annual savings (USD)"
opps = [
    ("Salesforce Licence Rightsizing & Renegotiation",
     f"Salesforce is ${sf/1e6:.2f}M (40% of all vendor spend) - by far the largest line. Run a seat-utilisation audit "
     f"(remove inactive/duplicate users, downgrade edition where Sales Cloud features are unused), then renegotiate at "
     f"renewal using the parent company's enterprise agreement. Cut Cloudcrossing consulting (${sf_consult/1e3:,.0f}K) to "
     f"a capped run-and-maintain scope. Assumes 25% licence saving and 40% consulting saving.", opp1),
    ("Office Footprint Consolidation",
     f"${office/1e6:.2f}M is spread across 11 office/coworking vendors in London, Zagreb (3), Chennai (2), Sydney, "
     f"Singapore and the US. Move to one flexible hub per country, exit duplicate Zagreb and India sites at lease break, "
     f"and fold into the parent's real estate where it overlaps. Assumes 50% reduction.", opp2),
    ("Professional Services & Advisor Rationalisation",
     f"Audit/tax work is split across 11 firms (${audit/1e3:,.0f}K incl. BDO, Grant Thornton, PwC, Crowe). Deal-related "
     f"fees (${deal/1e3:,.0f}K: RSM, Houlihan Lokey, Intralinks, Vector) end post-close, and ${consult/1e3:,.0f}K goes to "
     f"37 ad-hoc consultants/contractors. Move to the parent's single audit/tax firm and approved-panel model. Assumes 30% "
     f"audit/tax saving, 100% deal-fee run-off, 60% consultant reduction.", opp3),
]
for n, (title, expl, sav) in enumerate(opps, start=2):
    top.cell(n, 2).value = title
    top.cell(n, 3).value = expl
    top.cell(n, 4).value = round(sav, -3)
    top.cell(n, 4).number_format = '"$"#,##0'
    top.cell(n, 3).alignment = Alignment(wrap_text=True, vertical="top")
top.cell(5, 2).value = "Total (Top 3)"
top.cell(5, 4).value = "=SUM(D2:D4)"
top.cell(5, 4).number_format = '"$"#,##0'
top.cell(5, 2).font = Font(bold=True)
top.column_dimensions["B"].width = 40
top.column_dimensions["C"].width = 100
top.column_dimensions["D"].width = 22

meth = wb["Methodology"]
meth_text = open("docs/methodology.md").read()
for i, para in enumerate(meth_text.strip().split("\n\n"), start=2):
    meth.cell(i, 1).value = para.strip()
    meth.cell(i, 1).alignment = Alignment(wrap_text=True, vertical="top")
meth.column_dimensions["A"].width = 120

memo = wb["CEOCFO Recommendations"]
memo_text = open("docs/executive_memo.md").read()
memo["A2"] = memo_text.strip()
memo["A2"].alignment = Alignment(wrap_text=True, vertical="top")
memo.column_dimensions["A"].width = 120

wb.save("output/Vendor_Analysis_Assessment_completed.xlsx")
print("Saved output/Vendor_Analysis_Assessment_completed.xlsx")
