#!/usr/bin/env python3
"""
Build a premium, zone-wise September sales report from 'MASTER SALES REPORT -2.xlsx'.

- One formatted sheet per zone (day-wise rows x brand-wise columns)
- Daily total per day, brand-wise column totals, and zone grand total (ALL recomputed,
  because the source TOTAL columns contain formula errors)
- A summary/analytics block per zone (brand contribution %, best day, daily average, etc.)
- A cover 'SUMMARY' sheet comparing all zones
"""

import datetime as dt
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "MASTER SALES REPORT -2.xlsx"
OUT = "SEPTEMBER 2026 - ZONE SALES REPORT.xlsx"
YEAR = 2026

# Where each zone's September block starts, and its layout
# layout: first brand column index, total column index (1-based)
ZONES = {
    "ALPHA 1":    {"start": 317},
    "CENTRAL 50": {"start": 317},
    "BENNETT ":   {"start": 66},
}

# ---------- palette ----------
DARK   = "0F2D3F"   # title band
BLUE   = "16537E"   # header / grand total
LTBLUE = "D6E4F0"   # brand-total row / accents
BAND   = "F2F7FB"   # zebra band
WEEK   = "FCEEEC"   # weekend tint
WHITE  = "FFFFFF"
GREY   = "8A9BA8"
BORDERC= "D9E2EC"

INR = '"\u20b9"#,##0'
PCT = '0.0%'

thin = Side(style="thin", color=BORDERC)
box = Border(left=thin, right=thin, top=thin, bottom=thin)


def cell(ws, r, c, value=None, *, bold=False, size=11, color="1A2027",
         fill=None, align="center", fmt=None, border=True, italic=False, wrap=False):
    cl = ws.cell(r, c, value)
    cl.font = Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)
    cl.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fill:
        cl.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        cl.number_format = fmt
    if border:
        cl.border = box
    return cl


def extract(ws, start):
    """Return (brands list, rows) where rows = list of dicts {date, day, values{brand:val}}."""
    max_col = ws.max_column
    # header row is row 1; TOTAL column is last header cell containing 'TOTAL'
    headers = {c: (str(ws.cell(1, c).value).strip() if ws.cell(1, c).value else "")
               for c in range(1, max_col + 1)}
    total_col = max(c for c, h in headers.items() if "TOTAL" in h.upper())
    # brand columns: from col 4 (D) up to total_col-1, with a non-empty header
    brand_cols = [(c, headers[c]) for c in range(4, total_col) if headers[c]]

    rows = []
    r = start
    while r <= ws.max_row:
        day = ws.cell(r, 3).value      # DAY column
        date = ws.cell(r, 2).value     # DATE column
        if (day in (None, "")) and (date in (None, "")):
            break  # end of daily block (blank rows / totals row)
        if day in (None, ""):
            r += 1
            continue
        vals = {}
        for c, name in brand_cols:
            v = ws.cell(r, c).value
            vals[name] = float(v) if isinstance(v, (int, float)) else 0.0
        rows.append({"day": str(day).strip().title(), "values": vals})
        r += 1

    # assign clean dates Sept 1..N in row order
    for i, row in enumerate(rows):
        row["date"] = dt.date(YEAR, 9, i + 1)

    # keep only brands with any non-zero value in September
    active = [name for _, name in brand_cols
              if any(row["values"].get(name, 0) for row in rows)]
    return active, rows


def build_zone_sheet(wb, zone_name, brands, rows):
    ws = wb.create_sheet(zone_name.strip())
    ncol = 2 + len(brands) + 1  # Date, Day, brands..., Daily Total
    last_col_letter = get_column_letter(ncol)

    # ---- Title band ----
    ws.merge_cells(f"A1:{last_col_letter}1")
    t = ws.cell(1, 1, f"{zone_name.strip()}  \u2014  ZONE SALES REPORT")
    t.font = Font(name="Calibri", bold=True, size=20, color=WHITE)
    t.alignment = Alignment(horizontal="center", vertical="center")
    t.fill = PatternFill("solid", fgColor=DARK)
    ws.row_dimensions[1].height = 34

    ws.merge_cells(f"A2:{last_col_letter}2")
    s = ws.cell(2, 1, f"September {YEAR}   \u2022   Day-wise & Brand-wise Sales   \u2022   Amounts in \u20b9")
    s.font = Font(name="Calibri", size=11, italic=True, color=WHITE)
    s.alignment = Alignment(horizontal="center", vertical="center")
    s.fill = PatternFill("solid", fgColor=BLUE)
    ws.row_dimensions[2].height = 20

    # ---- Header row ----
    hr = 4
    cell(ws, hr, 1, "DATE", bold=True, color=WHITE, fill=BLUE)
    cell(ws, hr, 2, "DAY", bold=True, color=WHITE, fill=BLUE)
    for j, b in enumerate(brands):
        cell(ws, hr, 3 + j, b.strip().upper(), bold=True, color=WHITE, fill=BLUE, wrap=True)
    cell(ws, hr, ncol, "DAILY TOTAL", bold=True, color=WHITE, fill=BLUE, wrap=True)
    ws.row_dimensions[hr].height = 30

    # ---- Data rows ----
    brand_totals = {b: 0.0 for b in brands}
    grand = 0.0
    best_day = None
    best_val = -1
    r = hr + 1
    for i, row in enumerate(rows):
        is_weekend = row["day"] in ("Saturday", "Sunday")
        rowfill = WEEK if is_weekend else (BAND if i % 2 else WHITE)
        cell(ws, r, 1, row["date"].strftime("%d %b"), fill=rowfill,
             bold=is_weekend)
        cell(ws, r, 2, row["day"], fill=rowfill, color=(BLUE if is_weekend else "44515C"),
             bold=is_weekend)
        dtot = 0.0
        for j, b in enumerate(brands):
            v = row["values"].get(b, 0.0)
            dtot += v
            brand_totals[b] += v
            cell(ws, r, 3 + j, v if v else None, fmt=INR, fill=rowfill,
                 color=("C6CDD4" if not v else "1A2027"))
        cell(ws, r, ncol, dtot, fmt=INR, bold=True, fill=rowfill, color=BLUE)
        grand += dtot
        if dtot > best_val:
            best_val, best_day = dtot, row
        r += 1

    # ---- Brand totals row ----
    cell(ws, r, 1, "TOTAL", bold=True, color=WHITE, fill=BLUE)
    cell(ws, r, 2, "", fill=BLUE)
    for j, b in enumerate(brands):
        cell(ws, r, 3 + j, brand_totals[b], fmt=INR, bold=True, fill=LTBLUE, color=DARK)
    cell(ws, r, ncol, grand, fmt=INR, bold=True, fill=BLUE, color=WHITE)
    ws.row_dimensions[r].height = 22
    totals_row = r

    # ---- Contribution % row ----
    r += 1
    cell(ws, r, 1, "SHARE %", bold=True, color=WHITE, fill=DARK)
    cell(ws, r, 2, "", fill=DARK)
    for j, b in enumerate(brands):
        share = (brand_totals[b] / grand) if grand else 0
        cell(ws, r, 3 + j, share, fmt=PCT, bold=True, fill="EAF1F8", color=BLUE)
    cell(ws, r, ncol, (1 if grand else 0), fmt=PCT, bold=True, fill=DARK, color=WHITE)
    ws.row_dimensions[r].height = 20

    # ---- Column widths ----
    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 12
    for j in range(len(brands)):
        ws.column_dimensions[get_column_letter(3 + j)].width = 15
    ws.column_dimensions[last_col_letter].width = 16

    ws.freeze_panes = ws.cell(hr + 1, 3)
    ws.sheet_view.showGridLines = False

    # ---- Summary / analytics block ----
    sr = r + 2
    ws.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=ncol)
    h = ws.cell(sr, 1, "ZONE SUMMARY  \u2014  September " + str(YEAR))
    h.font = Font(name="Calibri", bold=True, size=13, color=WHITE)
    h.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    h.fill = PatternFill("solid", fgColor=DARK)
    ws.row_dimensions[sr].height = 24

    ndays = len(rows)
    top_brand = max(brand_totals, key=brand_totals.get)
    metrics = [
        ("Zone Grand Total", f"\u20b9{grand:,.0f}"),
        ("Operating Days", f"{ndays} days"),
        ("Average Daily Sales", f"\u20b9{grand/ndays:,.0f}" if ndays else "\u20b90"),
        ("Best Sales Day", f"{best_day['date'].strftime('%d %b')} ({best_day['day']}) \u2014 \u20b9{best_val:,.0f}"),
        ("Top Brand", f"{top_brand.strip()} \u2014 \u20b9{brand_totals[top_brand]:,.0f} ({brand_totals[top_brand]/grand*100:.1f}%)"),
        ("Active Brands", f"{len(brands)}"),
    ]
    mr = sr + 1
    for label, val in metrics:
        cl = cell(ws, mr, 1, label, bold=True, align="left", fill="EAF1F8", color=DARK)
        cl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.merge_cells(start_row=mr, start_column=2, end_row=mr, end_column=ncol)
        v = ws.cell(mr, 2, val)
        v.font = Font(name="Calibri", size=11, color="1A2027")
        v.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        v.border = box
        for cc in range(3, ncol + 1):
            ws.cell(mr, cc).border = box
        mr += 1

    return {"grand": grand, "brand_totals": brand_totals, "ndays": ndays,
            "top_brand": top_brand.strip(), "best": (best_day['date'], best_val),
            "brands": [b.strip() for b in brands]}


def build_summary_sheet(wb, results):
    ws = wb.create_sheet("SUMMARY", 0)
    ws.sheet_view.showGridLines = False
    ncol = 6
    lastc = get_column_letter(ncol)
    ws.merge_cells(f"A1:{lastc}1")
    t = ws.cell(1, 1, "SEPTEMBER 2026  \u2014  ALL-ZONE SALES SUMMARY")
    t.font = Font(name="Calibri", bold=True, size=20, color=WHITE)
    t.alignment = Alignment(horizontal="center", vertical="center")
    t.fill = PatternFill("solid", fgColor=DARK)
    ws.row_dimensions[1].height = 36

    ws.merge_cells(f"A2:{lastc}2")
    s = ws.cell(2, 1, "Amounts in \u20b9  \u2022  Totals recomputed from daily brand-wise data")
    s.font = Font(name="Calibri", size=11, italic=True, color=WHITE)
    s.alignment = Alignment(horizontal="center", vertical="center")
    s.fill = PatternFill("solid", fgColor=BLUE)

    hr = 4
    heads = ["ZONE", "OPERATING DAYS", "ACTIVE BRANDS", "AVG DAILY SALES", "TOP BRAND", "ZONE GRAND TOTAL"]
    for j, h in enumerate(heads):
        cell(ws, hr, j + 1, h, bold=True, color=WHITE, fill=BLUE, wrap=True)
    ws.row_dimensions[hr].height = 30

    r = hr + 1
    grand_all = 0
    for i, (zone, d) in enumerate(results.items()):
        fill = BAND if i % 2 else WHITE
        cell(ws, r, 1, zone.strip(), bold=True, align="left", fill=fill, color=DARK)
        cell(ws, r, 2, d["ndays"], fill=fill)
        cell(ws, r, 3, len(d["brands"]), fill=fill)
        cell(ws, r, 4, d["grand"] / d["ndays"] if d["ndays"] else 0, fmt=INR, fill=fill)
        cell(ws, r, 5, d["top_brand"], align="left", fill=fill)
        cell(ws, r, 6, d["grand"], fmt=INR, bold=True, fill=fill, color=BLUE)
        grand_all += d["grand"]
        r += 1

    cell(ws, r, 1, "ALL ZONES", bold=True, color=WHITE, fill=DARK, align="left")
    for cc in range(2, 6):
        cell(ws, r, cc, "", fill=DARK)
    cell(ws, r, 6, grand_all, fmt=INR, bold=True, fill=DARK, color=WHITE)
    ws.row_dimensions[r].height = 24

    # note about District 9
    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    n = ws.cell(r, 1, "Note: DISTRICT 9 has no September data in the source workbook "
                      "(only July & August), so it is not included.")
    n.font = Font(name="Calibri", size=10, italic=True, color=GREY)
    n.alignment = Alignment(horizontal="left", vertical="center")

    widths = [16, 16, 15, 18, 22, 20]
    for j, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(j + 1)].width = w


def main():
    src = load_workbook(SRC, data_only=True)
    out = Workbook()
    out.remove(out.active)

    results = {}
    for zone, meta in ZONES.items():
        ws = src[zone]
        brands, rows = extract(ws, meta["start"])
        print(f"[{zone.strip()}] brands={[b.strip() for b in brands]} days={len(rows)}")
        results[zone] = build_zone_sheet(out, zone, brands, rows)

    build_summary_sheet(out, results)
    out.save(OUT)
    print("Saved:", OUT)
    # verification print
    for zone, d in results.items():
        chk = sum(d["brand_totals"].values())
        print(f"  {zone.strip():12s} grand={d['grand']:,.0f} sum(brands)={chk:,.0f} match={abs(chk-d['grand'])<1e-6}")


if __name__ == "__main__":
    main()
