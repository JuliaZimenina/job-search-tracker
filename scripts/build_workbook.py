"""Build Job_Search_Tracker.xlsx from the synthetic dataset.

The workbook is formula-driven: Python only writes input values, formulas,
formatting and charts. Every metric is calculated by Excel itself, so the
file keeps working when new rows are added by hand.

Usage:
    python scripts/generate_synthetic_data.py
    python scripts/build_workbook.py               # writes Job_Search_Tracker.xlsx
"""
from __future__ import annotations

import csv
import sys
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.title import title_maker
from openpyxl.drawing.text import CharacterProperties, ParagraphProperties
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "synthetic_applications.csv"
OUT = ROOT / "Job_Search_Tracker.xlsx"

FIRST, CAPACITY = 5, 400                 # data rows 5..404
LAST = FIRST + CAPACITY - 1
AS_OF_DEMO = dt.date(2026, 11, 2)        # frozen report date for the demo
START_DEMO = dt.date(2026, 8, 17)

FONT = "Arial"
NAVY, BLUE, GREY = "1F3864", "2F5597", "F2F2F2"
INPUT_FILL = PatternFill("solid", fgColor="FFF9E6")
CALC_FILL = PatternFill("solid", fgColor=GREY)
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
TILE_FILL = PatternFill("solid", fgColor="DEEAF6")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", fgColor=color)


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# --------------------------------------------------------------------------- #
# Columns of the Applications table: (header, kind, width, hint, formula)
# kind: input | list | formula.  Formulas are written for row 5.
# --------------------------------------------------------------------------- #
COLUMNS = [
    ("#", "formula", 5, "", '=IF(B5="","",ROW()-4)'),
    ("Position", "input", 30, "job title", None),
    ("Company", "input", 24, "", None),
    ("Applied on", "input", 12, "YYYY-MM-DD", None),
    ("Response date", "input", 12, "first reply", None),
    ("Days since applied", "formula", 10, "", '=IF(D5="","",AsOfDate-D5)'),
    ("Days to response", "formula", 10, "", '=IF(OR(D5="",E5=""),"",E5-D5)'),
    ("Status", "list", 20, "current stage", None),
    ("Signal", "formula", 12, "",
     '=IF(OR(B5="",D5=""),"",IF(IFERROR(INDEX(FlagClosed,MATCH(H5,StatusList,0)),0)=1,"Closed",'
     'IF(J5="Yes","Responded",IF(F5>=SilentAfter,"Silent",IF(F5>=RemindAfter,"Follow up","Waiting")))))'),
    ("Responded", "formula", 12, "Yes/No",
     '=IF(OR(B5="",D5=""),"",IF(OR(E5<>"",IFERROR(INDEX(FlagResponse,MATCH(H5,StatusList,0)),0)=1),"Yes","No"))'),
    ("Interview date", "input", 12, "next interview", None),
    ("Interviewed", "formula", 12, "Yes/No",
     '=IF(OR(B5="",D5=""),"",IF(OR(K5<>"",IFERROR(INDEX(FlagInterview,MATCH(H5,StatusList,0)),0)=1),"Yes","No"))'),
    ("Contact person", "input", 16, "recruiter", None),
    ("Contact details", "input", 24, "e-mail / phone", None),
    ("Source", "list", 16, "", None),
    ("Work format", "list", 11, "", None),
    ("Location", "input", 18, "", None),
    ("Salary from", "input", 11, "number", None),
    ("Salary to", "input", 11, "optional", None),
    ("Currency", "list", 9, "", None),
    ("Period", "list", 10, "", None),
    ("EUR per month", "formula", 11, "midpoint",
     '=IF(OR(R5="",T5=""),"",IFERROR(IF(S5="",R5,(R5+S5)/2)*INDEX(RateToEUR,MATCH(T5,CurrencyList,0))'
     '/IF(U5="Yearly",12,1),""))'),
    ("Job link", "input", 30, "", None),
    ("Notes", "input", 40, "stages, impressions", None),
    ("Sort key", "formula", 9, "helper", '=IF(OR(I5="Silent",I5="Follow up"),F5+ROW()/100000,"")'),
]
CSV_TO_COL = {"Position": "B", "Company": "C", "Applied on": "D", "Response date": "E",
              "Status": "H", "Interview date": "K", "Contact person": "M",
              "Contact details": "N", "Source": "O", "Work format": "P", "Location": "Q",
              "Salary from": "R", "Salary to": "S", "Currency": "T", "Period": "U",
              "Job link": "W", "Notes": "X"}
RANGE_NAMES = {"r_pos": "Position", "r_comp": "Company", "r_applied": "Applied on",
               "r_resp_date": "Response date", "r_days": "Days since applied",
               "r_days_resp": "Days to response", "r_status": "Status", "r_signal": "Signal",
               "r_responded": "Responded", "r_interview": "Interviewed", "r_source": "Source",
               "r_format": "Work format", "r_salary": "EUR per month", "r_key": "Sort key"}

STATUSES = [  # status, responded, closed, interview
    ("Applied", 0, 0, 0), ("HR screening", 1, 0, 1), ("Technical interview", 1, 0, 1),
    ("Test assignment", 1, 0, 1), ("Final interview", 1, 0, 1), ("Offer", 1, 0, 1),
    ("Rejected", 1, 1, 0), ("No response", 0, 1, 0), ("Withdrawn", 0, 1, 0), ("On hold", 1, 0, 0),
]
FORMATS = ["Remote", "On-site", "Hybrid"]
SOURCES = ["LinkedIn", "Company website", "Jobs.cz", "Recruiter", "Referral", "Other"]
CURRENCIES = [("EUR", 1), ("USD", 0.86), ("CZK", 0.041), ("PLN", 0.235), ("GBP", 1.15)]
PERIODS = ["Monthly", "Yearly"]


def build_settings(wb):
    ws = wb.create_sheet("Settings")
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Settings"
    ws["B2"].font = f(16, True, NAVY)
    ws["B3"] = "Yellow cells are editable. Lists feed the drop-downs on the Applications sheet."
    ws["B3"].font = f(9, italic=True, color="595959")

    params = [
        ("Follow-up after (days without response)", 7, "RemindAfter", None),
        ("Long silence after (days without response)", 21, "SilentAfter", None),
        ("Interview alert (days ahead)", 3, "InterviewAlert", None),
        ("Analysis start date (week grid)", START_DEMO, "StartDate", "yyyy-mm-dd"),
        ("Report date override (leave blank = today)", AS_OF_DEMO, None, "yyyy-mm-dd"),
        ("Report date used in calculations", '=IF(C9="",TODAY(),C9)', "AsOfDate", "yyyy-mm-dd"),
    ]
    ws["B4"], ws["C4"] = "Parameter", "Value"
    for i, (label, val, name, fmt) in enumerate(params):
        r = 5 + i
        ws[f"B{r}"], ws[f"C{r}"] = label, val
        ws[f"B{r}"].font = f()
        c = ws[f"C{r}"]
        c.font = f(bold=True)
        c.fill = CALC_FILL if isinstance(val, str) else INPUT_FILL
        c.border = BOX
        c.alignment = Alignment(horizontal="center")
        if fmt:
            c.number_format = fmt
        if name:
            add_name(wb, name, f"Settings!$C${r}")
    ws["B11"] = ("The demo freezes the report date at 2026-11-02 so that the synthetic data "
                 "always shows the same picture. Clear C9 for live use.")
    ws["B11"].font = f(8, italic=True, color="595959")
    ws["B11"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B11:C13")

    def block(col, headers, rows, name_map):
        for j, h in enumerate(headers):
            ws.cell(4, col + j, h)
        for i, row in enumerate(rows):
            for j, v in enumerate(row if isinstance(row, tuple) else (row,)):
                c = ws.cell(5 + i, col + j, v)
                c.fill, c.border, c.font = INPUT_FILL, BOX, f()
                if j:
                    c.alignment = Alignment(horizontal="center")
        last = 4 + len(rows)
        for j, name in name_map.items():
            L = get_column_letter(col + j)
            add_name(wb, name, f"Settings!${L}$5:${L}${last}")

    block(5, ["Status", "Responded", "Closed", "Interview"], STATUSES,
          {0: "StatusList", 1: "FlagResponse", 2: "FlagClosed", 3: "FlagInterview"})
    block(10, ["Work format"], FORMATS, {0: "FormatList"})
    block(12, ["Source"], SOURCES, {0: "SourceList"})
    block(14, ["Currency", "1 unit = … EUR"], CURRENCIES, {0: "CurrencyList", 1: "RateToEUR"})
    block(17, ["Salary period"], PERIODS, {0: "PeriodList"})
    ws["N11"] = "Indicative rates; update manually."
    ws["N11"].font = f(8, italic=True, color="595959")
    ws["E16"] = "Flags (1/0) define how each status is counted in the Responded, Signal and Interviewed columns."
    ws["E16"].font = f(8, italic=True, color="595959")

    for col in (5, 6, 7, 8, 10, 12, 14, 15, 17):
        h = ws.cell(4, col)
        h.font, h.fill, h.border = f(bold=True, color="FFFFFF"), HEAD_FILL, BOX
        h.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in ("B4", "C4"):
        ws[c].font, ws[c].fill, ws[c].border = f(bold=True, color="FFFFFF"), HEAD_FILL, BOX
    widths = {"A": 2, "B": 44, "C": 13, "D": 3, "E": 20, "F": 11, "G": 9, "H": 10, "I": 3,
              "J": 13, "K": 3, "L": 17, "M": 3, "N": 10, "O": 14, "P": 3, "Q": 14}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.row_dimensions[4].height = 30


def build_applications(wb, records):
    ws = wb.create_sheet("Applications")
    ws["A1"] = "APPLICATIONS"
    ws["A1"].font = f(16, True, NAVY)
    ws["E1"] = "Grey columns are calculated; yellow columns are entered by hand. Add new rows directly below the last record."
    ws["E1"].font = f(9, italic=True, color="595959")
    for j, (head, kind, width, hint, formula) in enumerate(COLUMNS, start=1):
        L = get_column_letter(j)
        ws.column_dimensions[L].width = width
        t = ws.cell(2, j, {"input": "input", "list": "input (list)", "formula": "formula"}[kind])
        t.font = f(8, color="595959")
        t.fill = CALC_FILL if kind == "formula" else INPUT_FILL
        t.alignment = Alignment(horizontal="center")
        ws.cell(3, j, hint).font = f(8, italic=True, color="7F7F7F")
        ws.cell(3, j).alignment = Alignment(horizontal="center")
        h = ws.cell(4, j, head)
        h.font = f(bold=True, color="FFFFFF")
        h.fill = HEAD_FILL
        h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for r in range(FIRST, LAST + 1):
            c = ws.cell(r, j)
            c.font = f()
            if formula:
                c.value = _shift(formula, r)
                c.fill = CALC_FILL
            if head in ("Applied on", "Response date", "Interview date"):
                c.number_format = "yyyy-mm-dd"
            elif head in ("Salary from", "Salary to", "EUR per month"):
                c.number_format = "#,##0"
            elif head == "Sort key":
                c.number_format = "0.00000"
            if kind != "input" or head in ("Applied on", "Response date", "Interview date"):
                c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[4].height = 32

    for i, rec in enumerate(records):
        r = FIRST + i
        for key, L in CSV_TO_COL.items():
            v = rec[key]
            if v == "":
                continue
            if key in ("Applied on", "Response date", "Interview date"):
                v = dt.date.fromisoformat(v)
            elif key in ("Salary from", "Salary to"):
                v = float(v)
            ws[f"{L}{r}"] = v

    # Excel table with calculated columns (new rows inherit the formulas)
    tab = Table(displayName="tblApplications", ref=f"A4:{get_column_letter(len(COLUMNS))}{LAST}")
    cols = []
    for j, (head, kind, *_rest) in enumerate(COLUMNS, start=1):
        tc = TableColumn(id=j, name=head)
        if _rest[-1]:
            tc.calculatedColumnFormula = TableFormula(attr_text=_rest[-1].lstrip("="))
        cols.append(tc)
    tab.tableColumns = cols
    tab._initialise_columns = lambda: None
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(tab)
    for name, col in RANGE_NAMES.items():
        add_name(wb, name, f"tblApplications[{col}]")

    # Data validation
    def dv(**kw):
        v = DataValidation(allow_blank=True, showErrorMessage=True, **kw)
        ws.add_data_validation(v)
        return v
    for L, lst in (("H", "StatusList"), ("O", "SourceList"), ("P", "FormatList"),
                   ("T", "CurrencyList"), ("U", "PeriodList")):
        dv(type="list", formula1=lst).add(f"{L}{FIRST}:{L}{LAST}")
    dd = dv(type="date", operator="greaterThanOrEqual", formula1="43831",
       error="Enter a date (YYYY-MM-DD) after 2020-01-01.")
    dd.add(f"D{FIRST}:D{LAST}")
    dd.add(f"K{FIRST}:K{LAST}")
    dv(type="custom", formula1=f"AND(ISNUMBER(E{FIRST}),E{FIRST}>=D{FIRST})",
       error="Response date must be a date on or after the application date.").add(f"E{FIRST}:E{LAST}")
    dv(type="decimal", operator="greaterThanOrEqual", formula1="0").add(f"R{FIRST}:S{LAST}")

    # Conditional formatting
    rng = lambda a, b: f"{a}{FIRST}:{b}{LAST}"
    cf = ws.conditional_formatting
    cf.add(rng("I", "I"), FormulaRule(formula=[f'$I{FIRST}="Silent"'], fill=fill("F8CBAD"), font=f(bold=True, color="9C0006")))
    cf.add(rng("I", "I"), FormulaRule(formula=[f'$I{FIRST}="Follow up"'], fill=fill("FFE699"), font=f(color="7F6000")))
    cf.add(rng("I", "I"), FormulaRule(formula=[f'$I{FIRST}="Responded"'], fill=fill("C6EFCE"), font=f(color="006100")))
    cf.add(rng("B", "C"), FormulaRule(formula=[f'$H{FIRST}="Offer"'], fill=fill("C6EFCE"), font=f(bold=True, color="006100")))
    cf.add(rng("B", "X"), FormulaRule(formula=[f'$I{FIRST}="Closed"'], font=f(color="A6A6A6")))
    cf.add(rng("H", "H"), FormulaRule(formula=[f'$H{FIRST}="Rejected"'], font=f(color="C00000")))
    cf.add(rng("H", "H"), FormulaRule(formula=[f'INDEX(FlagInterview,MATCH($H{FIRST},StatusList,0))=1'], font=f(bold=True, color=BLUE)))
    cf.add(rng("K", "K"), FormulaRule(
        formula=[f'AND($K{FIRST}<>"",$K{FIRST}>=AsOfDate,$K{FIRST}-AsOfDate<=InterviewAlert)'],
        fill=fill("FFEB9C"), font=f(bold=True, color="9C5700")))
    cf.add(rng("G", "G"), ColorScaleRule(start_type="min", start_color="C6EFCE", mid_type="percentile",
                                         mid_value=50, mid_color="FFEB9C", end_type="max", end_color="F8CBAD"))
    ws.freeze_panes = "D5"
    return ws


def _shift(formula: str, r: int) -> str:
    """Rewrite a row-5 formula for row r (only relative refs to row 5 exist)."""
    import re
    return re.sub(r"(?<![A-Za-z_])([A-Z])5(?!\d)", lambda m: f"{m.group(1)}{r}", formula)


def build_calculations(wb):
    ws = wb.create_sheet("Calculations")
    ws["A1"] = "CALCULATIONS (feeds the dashboard and charts; no manual input)"
    ws["A1"].font = f(12, True, NAVY)
    heads = {"A4": "Funnel", "B4": "Count", "D4": "Status", "E4": "Applications",
             "G4": "Work format", "H4": "Applications", "J4": "Week of", "K4": "Applications",
             "L4": "Responses", "M4": "Monday", "O4": "Days waiting", "P4": "Applications",
             "R4": "Source", "S4": "Applications", "T4": "Responses", "U4": "Response rate",
             "W4": "Rank", "X4": "Sort key", "Y4": "Row in table"}
    for c, v in heads.items():
        ws[c] = v
        ws[c].font, ws[c].fill = f(bold=True, color="FFFFFF"), HEAD_FILL
        ws[c].alignment = Alignment(horizontal="center", wrap_text=True)

    funnel = [("Applications", "=COUNT(r_applied)"), ("Responses", '=COUNTIF(r_responded,"Yes")'),
              ("Interviews", '=COUNTIF(r_interview,"Yes")'), ("Offers", '=COUNTIF(r_status,"Offer")')]
    for i, (k, v) in enumerate(funnel):
        ws[f"A{5+i}"], ws[f"B{5+i}"] = k, v
    for i in range(len(STATUSES)):
        ws[f"D{5+i}"] = f"=Settings!E{5+i}"
        ws[f"E{5+i}"] = f"=COUNTIF(r_status,D{5+i})"
    for i in range(len(FORMATS)):
        ws[f"G{5+i}"] = f"=Settings!J{5+i}"
        ws[f"H{5+i}"] = f"=COUNTIF(r_format,G{5+i})"

    # Rolling 12-week grid ending with the current week
    ws["L2"], ws["M2"] = "Current week (Mon):", "=AsOfDate-WEEKDAY(AsOfDate,2)+1"
    ws["M2"].number_format = "yyyy-mm-dd"
    ws["L2"].font = f(9, italic=True)
    ws["M5"] = "=IF(StartDate-WEEKDAY(StartDate,2)+1>$M$2,\"\",MAX(StartDate-WEEKDAY(StartDate,2)+1,$M$2-77))"
    months = ",".join(f'"{m}"' for m in ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul",
                                         "Aug", "Sep", "Oct", "Nov", "Dec"])
    for r in range(5, 17):
        if r > 5:
            ws[f"M{r}"] = f'=IF(M{r-1}="","",IF(M{r-1}+7<=$M$2,M{r-1}+7,""))'
        ws[f"M{r}"].number_format = "yyyy-mm-dd"
        ws[f"J{r}"] = f'=IF(M{r}="","",CHOOSE(MONTH(M{r}),{months})&" "&DAY(M{r}))'
        ws[f"K{r}"] = f'=IF(M{r}="","",COUNTIFS(r_applied,">="&M{r},r_applied,"<"&(M{r}+7)))'
        ws[f"L{r}"] = f'=IF(M{r}="","",COUNTIFS(r_resp_date,">="&M{r},r_resp_date,"<"&(M{r}+7)))'

    buckets = [("0–6", 0, 7), ("7–13", 7, 14), ("14–20", 14, 21), ("21–30", 21, 31), ("31+", 31, 100000)]
    for i, (lab, lo, hi) in enumerate(buckets):
        ws[f"O{5+i}"] = lab
        ws[f"P{5+i}"] = (f'=COUNTIFS(r_days,">={lo}",r_days,"<{hi}",r_responded,"No",'
                         f'r_signal,"<>Closed")')
    for i in range(len(SOURCES)):
        r = 5 + i
        ws[f"R{r}"] = f"=Settings!L{r}"
        ws[f"S{r}"] = f"=COUNTIF(r_source,R{r})"
        ws[f"T{r}"] = f'=COUNTIFS(r_source,R{r},r_responded,"Yes")'
        ws[f"U{r}"] = f"=IF(S{r}=0,0,T{r}/S{r})"
        ws[f"U{r}"].number_format = "0%"
    for i in range(10):
        r = 5 + i
        ws[f"W{r}"] = i + 1
        ws[f"X{r}"] = f'=IFERROR(LARGE(r_key,W{r}),"")'
        ws[f"Y{r}"] = f'=IF(X{r}="","",MATCH(X{r},r_key,0))'
    for row in ws.iter_rows(min_row=5, max_row=16):
        for c in row:
            c.font = f()
    for L in "ABCDEFGHIJKLMNOPQRSTUVWXY":
        ws.column_dimensions[L].width = 12
    for L in ("A", "D", "R"):
        ws.column_dimensions[L].width = 19
    ws.column_dimensions["C"].width = ws.column_dimensions["F"].width = 3
    ws.column_dimensions["I"].width = ws.column_dimensions["N"].width = 3
    ws.column_dimensions["Q"].width = ws.column_dimensions["V"].width = 3
    return ws


def tile(ws, col, row, label, formula, fmt="0"):
    L1, L2 = get_column_letter(col), get_column_letter(col + 1)
    ws.merge_cells(f"{L1}{row}:{L2}{row}")
    ws.merge_cells(f"{L1}{row+1}:{L2}{row+2}")
    lab = ws[f"{L1}{row}"]
    lab.value = label
    lab.font = f(9, True, "595959")
    lab.alignment = Alignment(horizontal="center", vertical="center")
    val = ws[f"{L1}{row+1}"]
    val.value = formula
    val.font = f(20, True, NAVY)
    val.number_format = fmt
    val.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(row, row + 3):
        for c in (col, col + 1):
            ws.cell(r, c).fill = TILE_FILL


def set_title(ch, text):
    """Title above the plot area (overlay off), so it never covers axes or bars."""
    t = title_maker(text)
    t.overlay = False
    t.tx.rich.p[0].pPr = ParagraphProperties(defRPr=CharacterProperties(sz=1200, b=True))
    for run in t.tx.rich.p[0].r:
        run.rPr = CharacterProperties(sz=1200, b=True)
    ch.title = t


def bar(title, cats, data, horizontal=False, color="2F5597", height=7.5, width=14.2):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    set_title(ch, title)
    ch.style = 10
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.height, ch.width = height, width
    ch.legend = None
    ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                   showLegendKey=False, showPercent=False)
    ch.series[0].graphicalProperties.solidFill = color
    ch.series[0].graphicalProperties.line.solidFill = color
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None
    ch.gapWidth = 60
    if horizontal:
        ch.x_axis.scaling.orientation = "maxMin"   # first category on top
        ch.y_axis.crosses = "max"                   # keeps the value axis at the bottom
    return ch


def build_dashboard(wb, calc):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    for L in "BCDEFGHIJKLM":
        ws.column_dimensions[L].width = 13
    ws.column_dimensions["A"].width = 2
    ws.merge_cells("B1:M1")
    ws["B1"] = "Job Search Dashboard: Data Analyst"
    ws["B1"].font = f(18, True, NAVY)
    ws["B2"] = "Data as of:"
    ws["B2"].font = f(9, True)
    ws["C2"] = "=AsOfDate"
    ws["C2"].number_format = "yyyy-mm-dd"
    ws["C2"].font = f(9, True)
    ws.merge_cells("E2:M2")
    ws["E2"] = "Demo dataset: all records are synthetic and generated for illustration."
    ws["E2"].font = f(9, italic=True, color="C00000")
    ws["E2"].alignment = Alignment(horizontal="right")

    kpis1 = [("Total applications", "=COUNT(r_applied)", "0"),
             ("Responses received", '=COUNTIF(r_responded,"Yes")', "0"),
             ("Response rate", "=IF(B5=0,0,D5/B5)", "0%"),
             ("Interviews", '=COUNTIF(r_interview,"Yes")', "0"),
             ("Offers", '=COUNTIF(r_status,"Offer")', "0"),
             ("Long silence", '=COUNTIF(r_signal,"Silent")', "0")]
    kpis2 = [("Active applications", '=B5-COUNTIF(r_signal,"Closed")', "0"),
             ("Avg days to response", '=IFERROR(AVERAGE(r_days_resp),"-")', "0.0"),
             ("Applied, last 7 days", '=COUNTIFS(r_applied,">="&(AsOfDate-6),r_applied,"<="&AsOfDate)', "0"),
             ("Applied, last 30 days", '=COUNTIFS(r_applied,">="&(AsOfDate-29),r_applied,"<="&AsOfDate)', "0"),
             ("Rejections", '=COUNTIF(r_status,"Rejected")', "0"),
             ("Avg salary, EUR/month", '=IFERROR(AVERAGE(r_salary),"-")', "#,##0")]
    for i, (lab, frm, fmt) in enumerate(kpis1):
        tile(ws, 2 + 2 * i, 4, lab, frm, fmt)
    for i, (lab, frm, fmt) in enumerate(kpis2):
        tile(ws, 2 + 2 * i, 8, lab, frm, fmt)
    ws.conditional_formatting.add("L5:M6", FormulaRule(formula=["$L$5>0"], fill=fill("F8CBAD"), font=f(20, True, "9C0006")))
    ws.conditional_formatting.add("L5:M6", FormulaRule(formula=["$L$5=0"], fill=fill("C6EFCE"), font=f(20, True, "006100")))

    # Attention list
    ws.merge_cells("B12:H12")
    ws["B12"] = "Needs attention: applications without a response (top 10, longest wait first)"
    ws.merge_cells("J12:M12")
    ws["J12"] = "Source effectiveness"
    for c in ("B12", "J12"):
        ws[c].font = f(11, True, NAVY)
    for rng_, txt in (("B13:D13", "Position"), ("E13:F13", "Company"), ("G13:G13", "Days"),
                      ("H13:H13", "Signal"), ("J13:J13", "Source"), ("K13:K13", "Applications"),
                      ("L13:L13", "Responses"), ("M13:M13", "Response rate")):
        if ":" in rng_ and rng_.split(":")[0] != rng_.split(":")[1]:
            ws.merge_cells(rng_)
        c = ws[rng_.split(":")[0]]
        c.value = txt
        c.font, c.fill = f(9, True, "FFFFFF"), HEAD_FILL
        c.alignment = Alignment(horizontal="center")
    for c in "CDF":
        ws[f"{c}13"].fill = HEAD_FILL
    for i in range(10):
        r, k = 14 + i, f"Calculations!$Y${5+i}"
        ws.merge_cells(f"B{r}:D{r}")
        ws.merge_cells(f"E{r}:F{r}")
        ws[f"B{r}"] = f'=IF({k}="","",INDEX(r_pos,{k}))'
        ws[f"E{r}"] = f'=IF({k}="","",INDEX(r_comp,{k}))'
        ws[f"G{r}"] = f'=IF({k}="","",INDEX(r_days,{k}))'
        ws[f"H{r}"] = f'=IF({k}="","",INDEX(r_signal,{k}))'
        for c in "BEGH":
            ws[f"{c}{r}"].font = f(9)
            ws[f"{c}{r}"].border = Border(bottom=THIN)
        for c in "CDF":
            ws[f"{c}{r}"].border = Border(bottom=THIN)
        ws[f"G{r}"].alignment = ws[f"H{r}"].alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add("G14:H23", FormulaRule(formula=['$H14="Silent"'], font=f(9, True, "9C0006"), fill=fill("F8CBAD")))
    ws.conditional_formatting.add("G14:H23", FormulaRule(formula=['$H14="Follow up"'], font=f(9, color="7F6000"), fill=fill("FFE699")))

    for i in range(len(SOURCES)):
        r = 14 + i
        for c, src, fmt in (("J", "R", "@"), ("K", "S", "0"), ("L", "T", "0"), ("M", "U", "0%")):
            ws[f"{c}{r}"] = f"=Calculations!{src}{5+i}"
            ws[f"{c}{r}"].number_format = fmt
            ws[f"{c}{r}"].font = f(9)
            ws[f"{c}{r}"].border = Border(bottom=THIN)
            if c != "J":
                ws[f"{c}{r}"].alignment = Alignment(horizontal="center")
    ws["J20"], ws["K20"], ws["L20"], ws["M20"] = "Total", "=SUM(K14:K19)", "=SUM(L14:L19)", "=IF(K20=0,0,L20/K20)"
    ws["M20"].number_format = "0%"
    for c in "JKLM":
        ws[f"{c}20"].font = f(9, True)
        if c != "J":
            ws[f"{c}20"].alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add("M14:M19", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                         end_value=1, color="9DC3E6"))

    ws.merge_cells("B25:M25")
    ws["B25"] = "Analytics"
    ws["B25"].font = f(13, True, "FFFFFF")
    ws["B25"].fill = HEAD_FILL

    ws.add_chart(bar("Application funnel", Reference(calc, min_col=1, min_row=5, max_row=8),
                     Reference(calc, min_col=2, min_row=4, max_row=8), horizontal=True), "B27")
    ws.add_chart(bar("Applications by current status", Reference(calc, min_col=4, min_row=5, max_row=14),
                     Reference(calc, min_col=5, min_row=4, max_row=14), horizontal=True,
                     color="5B9BD5"), "H27")
    wk = bar("Weekly activity (last 12 weeks)", Reference(calc, min_col=10, min_row=5, max_row=16),
             Reference(calc, min_col=11, max_col=12, min_row=4, max_row=16), width=29.6)
    wk.series[1].graphicalProperties.solidFill = "A9D18E"
    wk.series[1].graphicalProperties.line.solidFill = "A9D18E"
    from openpyxl.chart.legend import Legend
    wk.legend = Legend()
    wk.legend.position = "b"
    wk.legend.overlay = False
    ws.add_chart(wk, "B43")
    ws.add_chart(bar("Open applications without a response, by days waiting",
                     Reference(calc, min_col=15, min_row=5, max_row=9),
                     Reference(calc, min_col=16, min_row=4, max_row=9), color="ED7D31"), "B59")
    pie = PieChart()
    set_title(pie, "Applications by work format")
    pie.legend.overlay = False
    pie.add_data(Reference(calc, min_col=8, min_row=4, max_row=7), titles_from_data=True)
    pie.set_categories(Reference(calc, min_col=7, min_row=5, max_row=7))
    pie.dataLabels = DataLabelList(showPercent=True, showVal=False, showSerName=False,
                                   showCatName=False, showLegendKey=False)
    pie.height, pie.width = 7.5, 14.2
    ws.add_chart(pie, "H59")
    return ws


def main():
    with DATA.open(encoding="utf-8") as fh:
        records = list(csv.DictReader(fh))
    wb = Workbook()
    wb.remove(wb.active)
    build_settings(wb)
    build_applications(wb, records)
    calc = build_calculations(wb)
    build_dashboard(wb, calc)
    wb._sheets = [wb["Dashboard"], wb["Applications"], wb["Calculations"], wb["Settings"]]
    wb.active = 0
    wb.properties.creator = "Job Search Tracker (pet project)"
    wb.properties.title = "Job Search Tracker: synthetic demo"
    wb.calculation.fullCalcOnLoad = True    # Excel computes all formulas when the file opens
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else OUT
    wb.save(out)
    print(f"{len(records)} records -> {out}")


if __name__ == "__main__":
    main()
