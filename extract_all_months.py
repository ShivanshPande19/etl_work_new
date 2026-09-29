#!/usr/bin/env python3
"""
Robust multi-month extractor for MASTER SALES REPORT -2.xlsx.

Each zone sheet stacks month blocks vertically:
  - column A holds the MONTH NAME on the first row of a block
  - column C holds the real weekday name for each day
  - column B's per-row date is unreliable (month-encoded), but its YEAR
    on the block's first row is correct (Nov/Dec 2025, Jan-Sep 2026)

This module returns, per zone:
  brands  -> union of brands with any non-zero sale across all months
  months  -> ordered list of {name, year, month_num, rows:[{date, day, values}]}
"""
import datetime as dt
from openpyxl import load_workbook

SRC = "MASTER SALES REPORT -2.xlsx"

MONTHS = {
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
    "JULY": 7, "AUGUST": 8, "AUG": 8, "SEPTEMBER": 9, "SEP": 9, "OCTOBER": 10,
    "OCT": 10, "NOVEMBER": 11, "NOV": 11, "DECEMBER": 12, "DEC": 12,
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "JUN": 6, "JUL": 7,
}
WEEKDAYS = {"MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY",
            "SATURDAY", "SUNDAY"}

ZONE_SHEETS = ["ALPHA 1", "CENTRAL 50", "BENNETT ", "DISTRICT 9"]


def month_key(a):
    if a is None:
        return None
    key = str(a).strip().upper()
    return MONTHS.get(key)


def brand_columns(ws):
    max_col = ws.max_column
    headers = {c: (str(ws.cell(1, c).value).strip() if ws.cell(1, c).value else "")
               for c in range(1, max_col + 1)}
    total_candidates = [c for c, h in headers.items() if "TOTAL" in h.upper()]
    total_col = max(total_candidates) if total_candidates else max_col
    cols = [(c, headers[c]) for c in range(4, total_col) if headers[c]]
    return cols, total_col


def extract_zone(ws):
    cols, total_col = brand_columns(ws)
    # 1) locate month block starts
    starts = []
    for r in range(2, ws.max_row + 1):
        mk = month_key(ws.cell(r, 1).value)
        if mk:
            yr = None
            bval = ws.cell(r, 2).value
            if isinstance(bval, dt.datetime):
                yr = bval.year
            starts.append((r, mk, yr))
    # 2) read each block until next start / non-weekday / blank
    start_rows = [s[0] for s in starts]
    months = []
    for idx, (r0, mnum, yr) in enumerate(starts):
        r_end = start_rows[idx + 1] if idx + 1 < len(start_rows) else ws.max_row + 1
        rows = []
        r = r0
        while r < r_end:
            day = ws.cell(r, 3).value
            dname = str(day).strip().upper() if day is not None else ""
            if dname not in WEEKDAYS:
                # stop at first non-day row inside the block (totals/blank)
                if rows:
                    break
                r += 1
                continue
            vals = {}
            for c, name in cols:
                v = ws.cell(r, c).value
                vals[name] = float(v) if isinstance(v, (int, float)) else 0.0
            rows.append({"day": dname.title(), "values": vals})
            r += 1
        if not rows:
            continue
        year = yr if yr else 2026
        for i, row in enumerate(rows):
            try:
                d = dt.date(year, mnum, i + 1)
            except ValueError:
                d = dt.date(year, mnum, 1)
            row["date"] = d
            # source day-name column has occasional bugs; derive from real date
            row["day"] = d.strftime("%A")
        months.append({"name": None, "year": year, "month_num": mnum, "rows": rows})
    # 3) union of active brands across all months
    active = []
    for _, name in cols:
        if any(row["values"].get(name, 0) for m in months for row in m["rows"]):
            active.append(name)
    # nice month name
    for m in months:
        m["name"] = dt.date(2000, m["month_num"], 1).strftime("%B").upper()
    return active, months


def load_all():
    src = load_workbook(SRC, data_only=True)
    result = {}
    for zone in ZONE_SHEETS:
        if zone not in src.sheetnames:
            continue
        brands, months = extract_zone(src[zone])
        result[zone] = {"brands": brands, "months": months}
    return result


if __name__ == "__main__":
    data = load_all()
    for zone, d in data.items():
        print(f"\n### {zone.strip()}  ({len(d['brands'])} brands)")
        print("   brands:", [b.strip() for b in d["brands"]])
        z_grand = 0
        for m in d["months"]:
            g = sum(sum(r["values"].values()) for r in m["rows"])
            z_grand += g
            print(f"   {m['name']:10s} {m['year']}  days={len(m['rows']):2d}  "
                  f"total=Rs {int(round(g)):>12,d}")
        print(f"   {'ZONE GRAND TOTAL':10s}        Rs {int(round(z_grand)):>12,d}")
