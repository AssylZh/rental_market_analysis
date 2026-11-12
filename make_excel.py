"""Build rental_market_analysis.xlsx from data/metro_summary.csv (created by analysis.ipynb).

Run: python make_excel.py
"""
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

F = "Arial"
BLUE = Font(name=F, color="0000FF")
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
NORMAL = Font(name=F)
TITLE = Font(name=F, bold=True, size=14)
NOTE = Font(name=F, italic=True, color="595959", size=9)
HEADER = Font(name=F, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
KEY_FILL = PatternFill("solid", fgColor="E2EFDA")
USD = '$#,##0;($#,##0);"-"'
PCT = "0.0%"
SIGNED_PCT = "+0.0%;-0.0%;0.0%"

m = pd.read_csv("data/metro_summary.csv")
N = len(m)
FIRST, LAST = 5, 4 + N           # data rows on Metro_Data and Cash_Flow

wb = Workbook()


def header(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=t)
        c.font, c.fill = HEADER, HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32


def cell(ws, ref, value, font=NORMAL, fmt=None, fill=None):
    ws[ref] = value
    ws[ref].font = font
    if fmt:
        ws[ref].number_format = fmt
    if fill:
        ws[ref].fill = fill


def widths(ws, ws_widths):
    for col, w in ws_widths.items():
        ws.column_dimensions[col].width = w


# ------------------------------------------------------------------ README
ws = wb.active
ws.title = "README"
cell(ws, "A1", "U.S. Rental Property Market Analysis", TITLE)
cell(ws, "A2", "Zillow Research data (ZHVI home values, ZORI rents), 100 largest U.S. metros, September 2026", NOTE)
lines = [
    ("Metro_Data", "Home values and rents now, 1 year ago and 5 years ago, with price-to-rent ratio and growth formulas"),
    ("Region_Summary", "Averages by Census region (pivot-style summary) with charts"),
    ("Assumptions", "Mortgage rate, down payment, taxes, insurance, maintenance, vacancy, management: edit here"),
    ("Cash_Flow", "Monthly cash flow if you buy a typical home in each metro and rent it out"),
    ("Forecast_12M", "12-month cash-flow forecast for any metro: pick one from the dropdown"),
    ("", ""),
    ("Blue text", "Data from Zillow or an input you can change"),
    ("Black text", "Formula"),
    ("Green text", "Formula pulling from another tab"),
    ("Yellow fill", "Cells to edit"),
    ("", ""),
    ("Source", "zillow.com/research/data: ZHVI All Homes (smoothed, seasonally adjusted) and ZORI All Homes Plus "
               "Multifamily (smoothed), Metro level. Mortgage rate: Freddie Mac Primary Mortgage Market Survey, "
               "Nov 5, 2026."),
]
for i, (a, b) in enumerate(lines, start=4):
    cell(ws, f"A{i}", a, BOLD if a and a not in ("Blue text", "Black text", "Green text", "Yellow fill") else NORMAL)
    cell(ws, f"B{i}", b)
    ws[f"B{i}"].alignment = Alignment(wrap_text=True, vertical="top")
ws["A10"].font, ws["A12"].font = BLUE, GREEN
ws["A13"].fill = INPUT_FILL
widths(ws, {"A": 18, "B": 100})

# -------------------------------------------------------------- Metro_Data
ws = wb.create_sheet("Metro_Data")
cell(ws, "A1", "Metro data: Zillow home values and rents", TITLE)
cell(ws, "A2", "Blue = Zillow data (Sep 2026, Sep 2025, Sep 2021). Black = formulas.", NOTE)
header(ws, 4, ["Metro", "State", "Region", "Size rank", "Home value (now)", "Monthly rent (now)",
               "Home value 1y ago", "Home value 5y ago", "Rent 1y ago", "Rent 5y ago",
               "Price-to-rent ratio", "Home value growth 1y", "Home value growth 5y",
               "Rent growth 1y", "Rent growth 5y"])
for i, r in enumerate(m.itertuples(index=False), start=FIRST):
    cell(ws, f"A{i}", r.metro, BLUE)
    cell(ws, f"B{i}", r.state, BLUE)
    cell(ws, f"C{i}", r.region, BLUE)
    cell(ws, f"D{i}", int(r.size_rank), BLUE)
    for col, v in zip("EFGHIJ", (r.home_value, r.monthly_rent, r.value_1y, r.value_5y, r.rent_1y, r.rent_5y)):
        cell(ws, f"{col}{i}", round(float(v)), BLUE, USD)
    cell(ws, f"K{i}", f"=E{i}/(F{i}*12)", fmt="0.0")
    cell(ws, f"L{i}", f"=E{i}/G{i}-1", fmt=SIGNED_PCT)
    cell(ws, f"M{i}", f"=E{i}/H{i}-1", fmt=SIGNED_PCT)
    cell(ws, f"N{i}", f"=F{i}/I{i}-1", fmt=SIGNED_PCT)
    cell(ws, f"O{i}", f"=F{i}/J{i}-1", fmt=SIGNED_PCT)
ws.freeze_panes = "B5"
ws.auto_filter.ref = f"A4:O{LAST}"
widths(ws, {"A": 30, "B": 7, "C": 11, "D": 9, **{c: 13 for c in "EFGHIJ"}, **{c: 12 for c in "KLMNO"}})
cell(ws, f"A{LAST + 2}", "Tip: to build a PivotTable, click any cell in this table, then Insert > PivotTable "
                         "(e.g. Rows = Region, Values = Average of Price-to-rent ratio).", NOTE)

MD = lambda col: f"Metro_Data!${col}${FIRST}:${col}${LAST}"

# ---------------------------------------------------------- Region_Summary
ws = wb.create_sheet("Region_Summary")
cell(ws, "A1", "Summary by region (averages across metros)", TITLE)
cell(ws, "A2", "Formulas (COUNTIF / AVERAGEIF) over Metro_Data, so they update if the data changes.", NOTE)
header(ws, 4, ["Region", "Metros", "Avg home value", "Avg monthly rent", "Avg price-to-rent",
               "Avg home growth 5y", "Avg rent growth 5y", "Avg home growth 1y", "Avg rent growth 1y"])
regions = ["Northeast", "Midwest", "South", "West"]
for i, reg in enumerate(regions + ["All metros"], start=5):
    cell(ws, f"A{i}", reg, BOLD)
    if reg == "All metros":
        cell(ws, f"B{i}", f"=SUM(B5:B8)", BOLD, "0")
        for col, src, fmt in (("C", "E", USD), ("D", "F", USD), ("E", "K", "0.0"), ("F", "M", SIGNED_PCT),
                              ("G", "O", SIGNED_PCT), ("H", "L", SIGNED_PCT), ("I", "N", SIGNED_PCT)):
            cell(ws, f"{col}{i}", f"=AVERAGE({MD(src)})", BOLD, fmt, KEY_FILL)
        continue
    cell(ws, f"B{i}", f'=COUNTIF({MD("C")},A{i})', GREEN, "0")
    for col, src, fmt in (("C", "E", USD), ("D", "F", USD), ("E", "K", "0.0"), ("F", "M", SIGNED_PCT),
                          ("G", "O", SIGNED_PCT), ("H", "L", SIGNED_PCT), ("I", "N", SIGNED_PCT)):
        cell(ws, f"{col}{i}", f'=AVERAGEIF({MD("C")},$A{i},{MD(src)})', GREEN, fmt)
widths(ws, {"A": 14, **{c: 14 for c in "BCDEFGHI"}})

ch = BarChart()
ch.type = "col"
ch.title = "Average price-to-rent ratio by region"
ch.add_data(Reference(ws, min_col=5, min_row=4, max_row=8), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=8))
ch.legend = None
ch.y_axis.title = "Home value ÷ yearly rent"
ch.series[0].graphicalProperties.solidFill = "2A78D6"
ch.height, ch.width = 7, 13
ch.x_axis.delete = ch.y_axis.delete = False
ws.add_chart(ch, "A12")

ch2 = BarChart()
ch2.type = "col"
ch2.title = "5-year growth: home values vs. rents"
ch2.add_data(Reference(ws, min_col=6, max_col=7, min_row=4, max_row=8), titles_from_data=True)
ch2.set_categories(Reference(ws, min_col=1, min_row=5, max_row=8))
ch2.y_axis.number_format = "0%"
ch2.series[0].graphicalProperties.solidFill = "2A78D6"
ch2.series[1].graphicalProperties.solidFill = "EB6834"
ch2.legend.position = "b"
ch2.height, ch2.width = 7, 13
ch2.x_axis.delete = ch2.y_axis.delete = False
ws.add_chart(ch2, "F12")

# ------------------------------------------------------------- Assumptions
ws = wb.create_sheet("Assumptions")
cell(ws, "A1", "Rental property assumptions (edit the yellow cells)", TITLE)
header(ws, 3, ["Assumption", "Value", "Note"])
assumptions = [
    ("Mortgage rate (30-year fixed)", 0.0718, "0.00%", "Freddie Mac weekly average, Nov 5, 2026"),
    ("Loan term (years)", 30, "0", ""),
    ("Down payment", 0.20, PCT, "Typical investor loan"),
    ("Closing costs (% of price)", 0.03, PCT, "Paid upfront; used for cash-on-cash return"),
    ("Property tax (% of value per year)", 0.011, "0.00%", "U.S. average; varies a lot by state"),
    ("Insurance (% of value per year)", 0.005, "0.00%", "Higher in FL, TX, CA"),
    ("Maintenance (% of value per year)", 0.01, "0.00%", "Common rule of thumb"),
    ("Vacancy (% of rent)", 0.05, PCT, "About 2-3 weeks empty per year"),
    ("Property management (% of rent)", 0.08, PCT, "Typical management fee"),
    ("Cost inflation (per year, for forecast)", 0.03, PCT, "Used in Forecast_12M"),
]
for i, (lab, v, fmt, note) in enumerate(assumptions, start=4):
    cell(ws, f"A{i}", lab)
    cell(ws, f"B{i}", v, BLUE, fmt, INPUT_FILL)
    cell(ws, f"C{i}", note, NOTE)
widths(ws, {"A": 40, "B": 12, "C": 48})
A = {k: f"Assumptions!$B${i}" for i, k in enumerate(
    ["rate", "years", "down", "closing", "tax", "ins", "maint", "vac", "mgmt", "infl"], start=4)}

# --------------------------------------------------------------- Cash_Flow
ws = wb.create_sheet("Cash_Flow")
cell(ws, "A1", "Monthly cash flow: buy a typical home in each metro and rent it out", TITLE)
cell(ws, "A2", "All costs per month. Change inputs on the Assumptions tab.", NOTE)
header(ws, 4, ["Metro", "Home value", "Monthly rent", "Mortgage payment", "Property tax", "Insurance",
               "Maintenance", "Vacancy", "Management", "Total costs", "Monthly cash flow",
               "Cap rate", "Cash-on-cash return", "Rank (best = 1)"])
for i in range(FIRST, LAST + 1):
    cell(ws, f"A{i}", f"=Metro_Data!A{i}", GREEN)
    cell(ws, f"B{i}", f"=Metro_Data!E{i}", GREEN, USD)
    cell(ws, f"C{i}", f"=Metro_Data!F{i}", GREEN, USD)
    cell(ws, f"D{i}", f"=PMT({A['rate']}/12,{A['years']}*12,-B{i}*(1-{A['down']}))", fmt=USD)
    cell(ws, f"E{i}", f"=B{i}*{A['tax']}/12", fmt=USD)
    cell(ws, f"F{i}", f"=B{i}*{A['ins']}/12", fmt=USD)
    cell(ws, f"G{i}", f"=B{i}*{A['maint']}/12", fmt=USD)
    cell(ws, f"H{i}", f"=C{i}*{A['vac']}", fmt=USD)
    cell(ws, f"I{i}", f"=C{i}*{A['mgmt']}", fmt=USD)
    cell(ws, f"J{i}", f"=SUM(D{i}:I{i})", fmt=USD)
    cell(ws, f"K{i}", f"=C{i}-J{i}", BOLD, USD)
    cell(ws, f"L{i}", f"=(C{i}-SUM(E{i}:I{i}))*12/B{i}", fmt="0.0%")
    cell(ws, f"M{i}", f"=K{i}*12/(B{i}*({A['down']}+{A['closing']}))", fmt="0.0%")
    cell(ws, f"N{i}", f"=RANK(K{i},$K${FIRST}:$K${LAST})", fmt="0")
ws.freeze_panes = "B5"
ws.auto_filter.ref = f"A4:N{LAST}"
widths(ws, {"A": 30, **{c: 12 for c in "BCDEFGHIJ"}, "K": 14, "L": 10, "M": 13, "N": 10})
# Summary box
cell(ws, "P4", "Summary", BOLD)
summary = [
    ("Metros with positive cash flow", f"=COUNTIF($K${FIRST}:$K${LAST},\">0\")", "0"),
    ("Median monthly cash flow", f"=MEDIAN($K${FIRST}:$K${LAST})", USD),
    ("Best metro", f"=INDEX($A${FIRST}:$A${LAST},MATCH(MAX($K${FIRST}:$K${LAST}),$K${FIRST}:$K${LAST},0))", None),
    ("Best monthly cash flow", f"=MAX($K${FIRST}:$K${LAST})", USD),
    ("Worst metro", f"=INDEX($A${FIRST}:$A${LAST},MATCH(MIN($K${FIRST}:$K${LAST}),$K${FIRST}:$K${LAST},0))", None),
    ("Worst monthly cash flow", f"=MIN($K${FIRST}:$K${LAST})", USD),
    ("Mortgage rate used", f"={A['rate']}", "0.00%"),
]
for j, (lab, f, fmt) in enumerate(summary, start=5):
    cell(ws, f"P{j}", lab)
    cell(ws, f"Q{j}", f, BOLD, fmt, KEY_FILL)
cell(ws, "P13", "Try it: set the mortgage rate to 4% on the Assumptions tab and watch these update.", NOTE)
widths(ws, {"O": 3, "P": 30, "Q": 20})
# Conditional formatting: green/red cash flow
from openpyxl.formatting.rule import CellIsRule
ws.conditional_formatting.add(f"K{FIRST}:K{LAST}", CellIsRule(operator="greaterThan", formula=["0"],
                              font=Font(name=F, bold=True, color="008000")))
ws.conditional_formatting.add(f"K{FIRST}:K{LAST}", CellIsRule(operator="lessThan", formula=["0"],
                              font=Font(name=F, bold=True, color="C00000")))

# ------------------------------------------------------------ Forecast_12M
ws = wb.create_sheet("Forecast_12M")
cell(ws, "A1", "12-month cash-flow forecast for one metro", TITLE)
cell(ws, "A2", "Pick a metro in the yellow cell (dropdown). Rent grows at that metro's last-12-month rate; "
               "costs grow with inflation; the mortgage payment is fixed.", NOTE)
cell(ws, "A4", "Metro", BOLD)
best = m.loc[m["monthly_cash_flow"].idxmax(), "metro"]
cell(ws, "B4", best, Font(name=F, bold=True, color="0000FF"), fill=INPUT_FILL)
dv = DataValidation(type="list", formula1=f"=Metro_Data!$A${FIRST}:$A${LAST}", allow_blank=False)
ws.add_data_validation(dv)
dv.add("B4")
match = f"MATCH($B$4,Metro_Data!$A${FIRST}:$A${LAST},0)"
inputs = [
    ("Home value", f"=INDEX(Metro_Data!$E${FIRST}:$E${LAST},{match})", USD),
    ("Starting monthly rent", f"=INDEX(Metro_Data!$F${FIRST}:$F${LAST},{match})", USD),
    ("Rent growth (last 12 months)", f"=INDEX(Metro_Data!$N${FIRST}:$N${LAST},{match})", SIGNED_PCT),
    ("Monthly mortgage payment", f"=PMT({A['rate']}/12,{A['years']}*12,-B5*(1-{A['down']}))", USD),
    ("Monthly tax + insurance + maintenance", f"=B5*({A['tax']}+{A['ins']}+{A['maint']})/12", USD),
]
for j, (lab, f, fmt) in enumerate(inputs, start=5):
    cell(ws, f"A{j}", lab)
    cell(ws, f"B{j}", f, GREEN, fmt)

header(ws, 11, ["Month", "Rent", "Vacancy + management", "Mortgage", "Tax + insurance + maintenance",
                "Total costs", "Cash flow", "Cumulative cash flow"])
for k in range(1, 13):
    r = 11 + k
    cell(ws, f"A{r}", k, fmt="0")
    cell(ws, f"B{r}", f"=$B$6*(1+$B$7)^((A{r}-1)/12)", fmt=USD)
    cell(ws, f"C{r}", f"=B{r}*({A['vac']}+{A['mgmt']})", fmt=USD)
    cell(ws, f"D{r}", "=$B$8", fmt=USD)
    cell(ws, f"E{r}", f"=$B$9*(1+{A['infl']})^((A{r}-1)/12)", fmt=USD)
    cell(ws, f"F{r}", f"=C{r}+D{r}+E{r}", fmt=USD)
    cell(ws, f"G{r}", f"=B{r}-F{r}", BOLD, USD)
    cell(ws, f"H{r}", f"=G{r}" if k == 1 else f"=H{r-1}+G{r}", fmt=USD)
cell(ws, "A24", "12-month total", BOLD)
for col in "BCDEFG":
    cell(ws, f"{col}24", f"=SUM({col}12:{col}23)", BOLD, USD, KEY_FILL)
cell(ws, "A26", "Rent needed to break even (month 1)", BOLD)
cell(ws, "B26", f"=($B$8+$B$9)/(1-{A['vac']}-{A['mgmt']})", BOLD, USD, KEY_FILL)
cell(ws, "C26", "Mortgage + fixed costs, grossed up for vacancy and management", NOTE)
widths(ws, {"A": 38, "B": 26, "C": 14, "D": 12, "E": 16, "F": 12, "G": 12, "H": 14})

lc = LineChart()
lc.title = "Cumulative cash flow over 12 months"
lc.add_data(Reference(ws, min_col=8, min_row=11, max_row=23), titles_from_data=True)
lc.set_categories(Reference(ws, min_col=1, min_row=12, max_row=23))
lc.legend = None
lc.y_axis.number_format = "$#,##0"
lc.x_axis.title = "Month"
lc.series[0].graphicalProperties.line.solidFill = "2A78D6"
lc.series[0].graphicalProperties.line.width = 22000
lc.series[0].smooth = False
lc.x_axis.delete = lc.y_axis.delete = False
lc.x_axis.tickLblPos = "low"   # keep month labels at the bottom when values are negative
lc.height, lc.width = 7, 15
ws.add_chart(lc, "J4")

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
wb.save("rental_market_analysis.xlsx")
print("Saved rental_market_analysis.xlsx")
