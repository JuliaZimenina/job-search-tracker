"""Cross-check the workbook's calculated values against an independent Python model.

The workbook must contain calculated values, so open it in Excel once and save
it (or recalculate it with LibreOffice) before running this script.

Usage:
    python scripts/validate_workbook.py Job_Search_Tracker.xlsx
"""
from __future__ import annotations

import datetime as dt
import statistics as st
import sys

import openpyxl

FIRST, LAST = 5, 404


def as_date(v):
    return v.date() if isinstance(v, dt.datetime) else v


def main(path: str) -> int:
    wb = openpyxl.load_workbook(path, data_only=True)
    s = wb["Settings"]
    as_of = as_date(s["C10"].value)
    if as_of is None:
        print("No calculated values found. Open and save the file in Excel first.")
        return 2
    remind, silent = s["C5"].value, s["C6"].value
    flags = {s[f"E{r}"].value: (s[f"F{r}"].value, s[f"G{r}"].value, s[f"H{r}"].value)
             for r in range(5, 15)}
    rates = {s[f"N{r}"].value: s[f"O{r}"].value for r in range(5, 10)}

    ws = wb["Applications"]
    model, errors = [], []
    for r in range(FIRST, LAST + 1):
        pos, d, e = ws[f"B{r}"].value, as_date(ws[f"D{r}"].value), as_date(ws[f"E{r}"].value)
        if not (pos and d):
            continue
        status, k = ws[f"H{r}"].value, ws[f"K{r}"].value
        resp_f, closed_f, int_f = flags.get(status, (0, 0, 0))
        days = (as_of - d).days
        responded = "Yes" if e or resp_f else "No"
        signal = ("Closed" if closed_f else "Responded" if responded == "Yes"
                  else "Silent" if days >= silent else "Follow up" if days >= remind else "Waiting")
        lo, hi, cur, per = (ws[f"{c}{r}"].value for c in "RSTU")
        eur = None
        if lo and cur:
            eur = (lo if not hi else (lo + hi) / 2) * rates[cur] / (12 if per == "Yearly" else 1)
        row = dict(F=days, G=(e - d).days if e else None, I=signal, J=responded,
                   L="Yes" if k or int_f else "No", V=eur)
        for col, want in row.items():
            got = ws[f"{col}{r}"].value
            got = None if got == "" else got
            ok = (got == want) or (isinstance(got, (int, float)) and isinstance(want, (int, float))
                                   and abs(got - want) < 1e-6)
            if not ok:
                errors.append(f"Applications!{col}{r}: workbook={got!r} model={want!r}")
        row.update(status=status, source=ws[f"O{r}"].value, d=d)
        model.append(row)

    dash = wb["Dashboard"]
    expected = {
        "B5": len(model),
        "D5": sum(m["J"] == "Yes" for m in model),
        "F5": sum(m["J"] == "Yes" for m in model) / len(model) if model else 0,
        "H5": sum(m["L"] == "Yes" for m in model),
        "J5": sum(m["status"] == "Offer" for m in model),
        "L5": sum(m["I"] == "Silent" for m in model),
        "B9": sum(m["I"] != "Closed" for m in model),
        "D9": st.mean([m["G"] for m in model if m["G"] is not None]),
        "F9": sum(0 <= (as_of - m["d"]).days <= 6 for m in model),
        "H9": sum(0 <= (as_of - m["d"]).days <= 29 for m in model),
        "J9": sum(m["status"] == "Rejected" for m in model),
        "L9": st.mean([m["V"] for m in model if m["V"] is not None]),
    }
    for cell, want in expected.items():
        got = dash[cell].value
        if not isinstance(got, (int, float)) or abs(got - want) > 1e-6:
            errors.append(f"Dashboard!{cell}: workbook={got!r} model={want!r}")

    waiting = sorted((m for m in model if m["I"] in ("Silent", "Follow up")),
                     key=lambda m: m["F"], reverse=True)[:10]
    shown = [dash[f"G{r}"].value for r in range(14, 24) if dash[f"G{r}"].value not in (None, "")]
    if shown != [m["F"] for m in waiting]:
        errors.append(f"Dashboard top-10 days: workbook={shown} model={[m['F'] for m in waiting]}")

    print(f"Checked {len(model)} applications, {len(expected)} KPIs and the top-10 list.")
    if errors:
        print(f"{len(errors)} mismatch(es):")
        print("\n".join("  " + e for e in errors))
        return 1
    print("All values match the independent model.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "Job_Search_Tracker.xlsx"))
