import random
from datetime import date, timedelta
from itertools import product
from pathlib import Path

from openpyxl import Workbook

random.seed(42)

INPUT_DIR = Path("input")
INPUT_DIR.mkdir(exist_ok=True)

WEEK_START = date(2026, 9, 28)  # a Monday
SUBS = [
    "Apex Electrical", "Northline Plumbing", "Maple HVAC", "Cornerstone Concrete",
    "Summit Framing", "BrightPath Drywall", "Lakeshore Roofing", "Ironbridge Steel",
    "Greenline Landscaping", "Harbour Painting",
]
PROJECTS = ["P-101", "P-102", "P-103"]
WORKERS = ["J. Smith", "A. Khan", "M. Rossi", "L. Chen", "D. Patel", "S. Okafor"]
TASKS = ["Install", "Prep work", "Site cleanup", "Inspection", "Material handling", "Rework"]
RATES = [45, 52, 58, 65, 72]

KEYS = ["date", "project", "worker", "hours", "rate", "description"]
HEADERS = [
    ["Date", "Project Code", "Worker", "Hours", "Rate", "Description"],
    ["Work Date", "Project #", "Employee", "Hrs", "Hourly Rate", "Notes"],
    ["DATE", "Job Code", "Name", "Hours Worked", "Rate/hr", "Task"],
    ["Day", "Job #", "Crew Member", "Hrs Worked", "$/hr", "Work Done"],
    ["Date", "Project", "Worker", "Hours", "Rate per hour", "Description"],
]
DATE_FORMATS = ["%Y-%m-%d", "%b %d, %Y", "%d-%b-%Y", None]  # None = real Excel date


def build_rows():
    days = [WEEK_START + timedelta(days=d) for d in range(5)]
    combos = random.sample(list(product(WORKERS, days)), random.randint(8, 14))
    rate = random.choice(RATES)
    return [
        {
            "date": day, "project": random.choice(PROJECTS), "worker": worker,
            "hours": random.choice([4, 6, 7.5, 8, 8, 8, 9, 10]), "rate": rate,
            "description": random.choice(TASKS),
        }
        for worker, day in combos
    ]


def inject_problems(i, rows):
    if i == 0:
        for r in rows:
            r["rate"] = f"${r['rate']:.2f}"      # rate stored as text
    if i == 2:
        rows[0]["hours"] = "8 hrs"                # text hours
        rows[1]["hours"] = 16                     # impossible day
    if i == 3:
        rows.append(dict(rows[0]))                # duplicate entry
    if i == 4:
        rows[2]["hours"] = None                   # missing hours
    if i == 5:
        rows[1]["project"] = "P-1O2"              # typo: letter O
    if i == 7:
        rows[1]["project"] = "p102"               # fixable format
    if i == 9:
        rows[0]["date"] = rows[0]["date"] - timedelta(days=14)  # wrong week


def write_file(i, sub):
    rows = build_rows()
    inject_problems(i, rows)
    labels = dict(zip(KEYS, HEADERS[i % len(HEADERS)]))
    order = KEYS[:]
    random.shuffle(order)
    fmt = DATE_FORMATS[i % len(DATE_FORMATS)]

    wb = Workbook()
    ws = wb.active
    ws.title = "Timesheet"
    ws.append([sub])
    ws.append(["Weekly timesheet"])
    for _ in range(i % 3):
        ws.append([])
    ws.append([labels[k] for k in order])

    for idx, r in enumerate(rows):
        if i == 8 and idx == 4:
            ws.append([])                         # random blank row
        ws.append([
            (r[k] if fmt is None else r[k].strftime(fmt)) if k == "date" else r[k]
            for k in order
        ])
    if i == 6:
        ws.append(["TOTAL"])                      # junk total row

    name = sub.lower().replace(" ", "_")
    wb.save(INPUT_DIR / f"{name}_week_2026-09-28.xlsx")


if __name__ == "__main__":
    for i, sub in enumerate(SUBS):
        write_file(i, sub)
    print(f"Created {len(SUBS)} messy files in {INPUT_DIR}/")