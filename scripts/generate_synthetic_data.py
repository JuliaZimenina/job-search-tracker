"""Generate the synthetic job-application dataset used in the demo workbook.

All records are artificial: company names, people, e-mail addresses, links and
salaries are invented. The generator is seeded, so the output is reproducible.

Usage:
    python scripts/generate_synthetic_data.py      # writes data/synthetic_applications.csv
"""
from __future__ import annotations

import csv
import datetime as dt
import random
from pathlib import Path

SEED = 27
AS_OF = dt.date(2026, 11, 2)       # snapshot date of the demo
START = dt.date(2026, 8, 17)       # first week of the search (Monday)
N_APPLICATIONS = 64

POSITIONS = [
    "Junior Data Analyst", "Data Analyst", "BI Analyst", "Reporting Analyst",
    "Product Analyst", "Business Data Analyst", "Power BI Developer",
    "Data Quality Analyst", "Marketing Analyst", "Operations Analyst",
    "Financial Data Analyst", "Data Governance Analyst", "SQL Analyst",
]
COMPANIES = [
    "Northwind Analytics", "Bluepeak Logistics", "Vltava Retail Group",
    "Kestrel Fintech", "Orbisoft", "Helix Health Data", "Granite Insurance",
    "Lumen Energy", "Mapleton Consulting", "Quarry Labs", "Tidewater Commerce",
    "Aurora Telecom", "Brightpath Education", "Cobalt Mobility", "Fernway Foods",
    "Pinecrest Media", "Redwood Payments", "Summit Travel", "Zenith Manufacturing",
    "Larkspur Software", "Oakridge Pharma", "Meridian Freight",
]
FIRST_NAMES = ["Anna", "Petra", "Martin", "Lukas", "Eva", "Tomas", "Klara",
               "Jakub", "Sofia", "Daniel", "Marta", "Oliver", "Lena", "Pavel"]
LAST_INITIALS = list("BDHKLMNPRSTVZ")
SOURCES_W = [("LinkedIn", 34), ("Company website", 18), ("Jobs.cz", 18),
             ("Recruiter", 14), ("Referral", 8), ("Other", 8)]
LOCATIONS = [  # (location, currency, period, low, high, format weights)
    ("Prague", "CZK", "Monthly", 50000, 85000),
    ("Brno", "CZK", "Monthly", 45000, 75000),
    ("Warsaw", "PLN", "Monthly", 9000, 16000),
    ("Berlin", "EUR", "Yearly", 48000, 65000),
    ("Barcelona", "EUR", "Yearly", 30000, 42000),
    ("London", "GBP", "Yearly", 35000, 52000),
    ("Remote (EU)", "EUR", "Monthly", 2800, 4200),
    ("Remote (US company)", "USD", "Yearly", 55000, 75000),
]
LOCATION_W = [30, 10, 8, 8, 6, 5, 25, 8]
FORMATS_W = [("Remote", 40), ("Hybrid", 45), ("On-site", 15)]


def wchoice(rng: random.Random, pairs):
    items, weights = zip(*pairs)
    return rng.choices(items, weights=weights, k=1)[0]


def applied_date(rng: random.Random) -> dt.date:
    """Activity grows over time: later weeks get more applications."""
    span = (AS_OF - START).days
    x = rng.random() ** 0.8          # skew towards recent dates
    return START + dt.timedelta(days=round(x * span))


def build() -> list[dict]:
    rng = random.Random(SEED)
    rows = []
    for _ in range(N_APPLICATIONS):
        d = applied_date(rng)
        age = (AS_OF - d).days
        loc, cur, per, lo, hi = rng.choices(LOCATIONS, weights=LOCATION_W, k=1)[0]
        source = wchoice(rng, SOURCES_W)
        fmt = "Remote" if loc.startswith("Remote") else wchoice(rng, FORMATS_W[1:])
        resp_p = {"Referral": 0.8, "Recruiter": 0.75}.get(source, 0.42)

        status, resp, interview, note = "Applied", None, None, ""
        if rng.random() < resp_p and age >= 1:
            resp = d + dt.timedelta(days=min(age, max(1, int(rng.expovariate(1 / 5)))))
            since = (AS_OF - resp).days
            r = rng.random()
            if r < 0.45:
                status = "Rejected"
                note = rng.choice(["Template rejection e-mail",
                                   "Rejected after CV review",
                                   "Position filled internally",
                                   "More years of experience required"])
            elif r < 0.55:
                status = "On hold"
                note = "Hiring paused until next quarter"
            else:
                # progress through the pipeline depending on time since response
                stage = rng.choices(range(5), weights=[4, 3, 4, 2, 1], k=1)[0]
                stage = min(stage, since // 5)       # later stages need time
                status = ["HR screening", "Test assignment", "Technical interview",
                          "Final interview", "Offer"][stage]
                if status != "HR screening":
                    interview = resp + dt.timedelta(days=rng.randint(3, 8))
                    if status in ("Technical interview", "Final interview") and rng.random() < 0.5:
                        interview = AS_OF + dt.timedelta(days=rng.randint(1, 5))  # upcoming
                    interview = min(interview, AS_OF + dt.timedelta(days=5))
                note = {"HR screening": "Intro call with recruiter scheduled",
                        "Test assignment": "Take-home: SQL + Power BI case, 3 days",
                        "Technical interview": "SQL live coding and dashboard review",
                        "Final interview": "Meeting with the hiring manager",
                        "Offer": "Verbal offer received, written offer pending"}[status]
                if rng.random() < 0.25 and status in ("Technical interview", "Test assignment"):
                    status, note = "Rejected", "Rejected after technical stage"
                # an open stage older than three weeks is resolved, not left hanging
                last_touch = interview or resp
                if status in ("HR screening", "Test assignment", "Technical interview",
                              "Final interview") and last_touch < AS_OF - dt.timedelta(days=21):
                    if interview:
                        status, note = "Rejected", f"Rejected after {status.lower()}"
                    else:
                        status, note = "No response", "No follow-up after the screening call"
        elif age > 35 and rng.random() < 0.5:
            status, note = "No response", "Closed after 5+ weeks of silence"
        elif rng.random() < 0.04:
            status, note = "Withdrawn", "Withdrew: relocation required"

        salary_known = source in ("Recruiter", "Jobs.cz") or rng.random() < 0.45
        s_from = s_to = None
        if salary_known:
            s_from = round(rng.uniform(lo, (lo + hi) / 2), -2 if per == "Monthly" else -3)
            s_to = round(rng.uniform(s_from, hi), -2 if per == "Monthly" else -3) if rng.random() < 0.7 else None

        contact = email = ""
        if source in ("Recruiter", "Referral") or resp:
            fn, li = rng.choice(FIRST_NAMES), rng.choice(LAST_INITIALS)
            contact = f"{fn} {li}."
            email = f"{fn.lower()}.{li.lower()}@example.com"

        rows.append({
            "Position": rng.choice(POSITIONS),
            "Company": rng.choice(COMPANIES),
            "Applied on": d, "Response date": resp, "Status": status,
            "Interview date": interview, "Contact person": contact, "Contact details": email,
            "Source": source, "Work format": fmt, "Location": loc,
            "Salary from": s_from, "Salary to": s_to,
            "Currency": cur if salary_known else None,
            "Period": per if salary_known else None,
            "Job link": f"https://example.com/jobs/{rng.randint(10000, 99999)}",
            "Notes": note,
        })
    rows.sort(key=lambda r: r["Applied on"])
    return rows


def main() -> None:
    rows = build()
    out = Path(__file__).resolve().parents[1] / "data" / "synthetic_applications.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        for r in rows:
            w.writerow({k: (v.isoformat() if isinstance(v, dt.date) else ("" if v is None else v))
                        for k, v in r.items()})
    print(f"{len(rows)} synthetic applications -> {out}")


if __name__ == "__main__":
    main()
