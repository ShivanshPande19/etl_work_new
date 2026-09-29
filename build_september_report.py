#!/usr/bin/env python3
"""
Build a premium, zone-wise September sales report from 'MASTER SALES REPORT -2.xlsx'.

Design language: near-black + red accent + neutral greys, generous spacing,
right-aligned figures (financial standard), thin rules, and the Indian
numbering system (lakh/crore grouping, e.g. Rs 4,35,705).

Output:
- One styled sheet per zone (day-wise rows x brand-wise columns)
- Daily total column, brand-wise totals row, brand share % row
- A per-zone analytics block (grand total, best day, daily average, top brand ...)
- A cover 'SUMMARY' sheet comparing all zones
All totals are recomputed from the daily data (source TOTAL columns had errors).
"""

import datetime as dt
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

SRC = "MASTER SALES REPORT -2.xlsx"
OUT = "SEPTEMBER 2026 - ZONE SALES REPORT.xlsx"
YEAR = 2026
FONT = "Segoe UI"

ZONES = {
    "ALPHA 1":    {"start": 317},
    "CENTRAL 50": {"start": 317},
    "BENNETT ":   {"start": 66},
}

# ---------- palette (soothing: teal-slate + soft sage + warm neutral) ----------
INK    = "34515B"   # deep muted teal-slate (title / total band)
INK2   = "4A6D76"   # muted teal (column header fill)
REDF   = "4F7D6E"   # soft sage accent  (fills on white)  [grand total, highlights]
REDT   = "9BC3B6"   # light sage accent (text on the dark band)
GREY   = "8B969D"   # secondary text (cool gray)
GREYD  = "5C6B72"   # darker secondary
INKTX  = "2F3E45"   # primary text (dark slate)
LINE   = "E7E2D9"   # hairline borders (soft warm)
BAND   = "F7F5F0"   # zebra (warm off-white)
FILL2  = "EDF2EF"   # section / total-col fill (pale sage-gray)
WHITE  = "FFFFFF"

# Indian numbering format: <1 lakh -> thousands; lakh band; crore band
INR = '[>=10000000]"\u20b9"##\\,##\\,##\\,##0;[>=100000]"\u20b9"##\\,##\\,##0;"\u20b9"#,##0'
PCT = '0.0%'

hair = Side(style="thin", color=LINE)
box = Border(left=hair, right=hair, top=hair, bottom=hair)
bottom_only = Border(bottom=Side(style="thin", color=LINE))


def inr(n):
    """Indian-grouped currency string, e.g. 3514180 -> 'Rs 35,14,180'."""
    n = int(round(n))
    sign = "-" if n < 0 else ""
    s = str(abs(n))
    if len(s) > 3:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        grouped = ",".join(parts) + "," + last3
    else:
        grouped = s
    return f"{sign}\u20b9{grouped}"


def sty(cl, *, bold=False, size=11, color=INKTX, fill=None, align="center",
        fmt=None, border=box, italic=False, wrap=False, indent=0, font=FONT):
    cl.font = Font(name=font, bold=bold, size=size, color=color, italic=italic)
    cl.alignment = Alignment(horizontal=align, vertical="center",
                             wrap_text=wrap, indent=indent)
    if fill:
        cl.fill = PatternFill("solid", fgColor=fill)
    if fmt:
        cl.number_format = fmt
    if border is not None:
        cl.border = border
    return cl


def extract(ws, start):
    max_col = ws.max_column
    headers = {c: (str(ws.cell(1, c).value).strip() if ws.cell(1, c).value else "")
               for c in range(1, max_col + 1)}
    total_col = max(c for c, h in headers.items() if "TOTAL" in h.upper())
    brand_cols = [(c, headers[c]) for c in range(4, total_col) if headers[c]]

    rows = []
    r = start
    while r <= ws.max_row:
        day = ws.cell(r, 3).value
        date = ws.cell(r, 2).value
        if (day in (None, "")) and (date in (None, "")):
            break
        if day in (None, ""):
            r += 1
            continue
        vals = {}
        for c, name in brand_cols:
            v = ws.cell(r, c).value
            vals[name] = float(v) if isinstance(v, (int, float)) else 0.0
        rows.append({"day": str(day).strip().title(), "values": vals})
        r += 1
    for i, row in enumerate(rows):
        row["date"] = dt.date(YEAR, 9, i + 1)
    active = [name for _, name in brand_cols
              if any(row["values"].get(name, 0) for row in rows)]
    return active, rows


def set_widths(ws, widths):
    for idx, w in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = w


def title_band(ws, ncol, zone, subtitle):
    last = get_column_letter(ncol)
    # Row 1: big title with red first letter
    ws.merge_cells(f"A1:{last}1")
    zt = zone.strip()
    rich = CellRichText([
        TextBlock(InlineFont(rFont=FONT, b=True, sz=22, color="FF" + REDT), zt[0]),
        TextBlock(InlineFont(rFont=FONT, b=True, sz=22, color="FFFFFFFF"),
                  zt[1:] + "      SEPTEMBER 2026"),
    ])
    c = ws.cell(1, 1); c.value = rich
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.fill = PatternFill("solid", fgColor=INK)
    ws.row_dimensions[1].height = 42
    # Row 2: subtitle
    ws.merge_cells(f"A2:{last}2")
    sty(ws.cell(2, 1), size=10.5, color="C9C9C9", align="left", fill=INK,
        border=None, indent=1).value = subtitle
    ws.row_dimensions[2].height = 20
    # Row 3: red accent rule
    ws.merge_cells(f"A3:{last}3")
    ws.cell(3, 1).fill = PatternFill("solid", fgColor=REDF)
    ws.row_dimensions[3].height = 4
    # Row 4: spacer
    ws.row_dimensions[4].height = 8


def build_zone_sheet(wb, zone, brands, rows):
    ws = wb.create_sheet(zone.strip())
    ws.sheet_view.showGridLines = False
    ncol = 2 + len(brands) + 1
    dtot_col = ncol
    subtitle = "Day-wise & Brand-wise Sales   \u2022   Amounts in \u20b9 (INR)   \u2022   Recomputed from daily data"
    title_band(ws, ncol, zone, subtitle)

    # ---- header ----
    hr = 5
    sty(ws.cell(hr, 1), bold=True, size=10.5, color=WHITE, fill=INK2, wrap=True).value = "DATE"
    sty(ws.cell(hr, 2), bold=True, size=10.5, color=WHITE, fill=INK2, wrap=True).value = "DAY"
    for j, b in enumerate(brands):
        sty(ws.cell(hr, 3 + j), bold=True, size=10, color=WHITE, fill=INK2,
            wrap=True).value = b.strip().upper()
    sty(ws.cell(hr, dtot_col), bold=True, size=10.5, color=WHITE, fill=REDF,
        wrap=True).value = "DAILY TOTAL"
    ws.row_dimensions[hr].height = 34

    # ---- data ----
    brand_tot = {b: 0.0 for b in brands}
    grand = 0.0
    best = None
    best_val = -1
    r = hr + 1
    for i, row in enumerate(rows):
        wknd = row["day"] in ("Saturday", "Sunday")
        rf = BAND if i % 2 else WHITE
        sty(ws.cell(r, 1), size=10.5, color=INKTX, fill=rf,
            align="center").value = row["date"].strftime("%d %b")
        sty(ws.cell(r, 2), size=10.5, bold=wknd, color=(REDF if wknd else GREYD),
            fill=rf, align="center").value = row["day"]
        dsum = 0.0
        for j, b in enumerate(brands):
            v = row["values"].get(b, 0.0)
            dsum += v
            brand_tot[b] += v
            ccell = ws.cell(r, 3 + j)
            if v:
                sty(ccell, size=10.5, color=INKTX, fill=rf, align="right",
                    fmt=INR, indent=1).value = v
            else:
                sty(ccell, size=10.5, color="C4C4C4", fill=rf,
                    align="center").value = "\u2013"
        sty(ws.cell(r, dtot_col), bold=True, size=10.5, color=INKTX, fill=FILL2,
            align="right", fmt=INR, indent=1).value = dsum
        grand += dsum
        if dsum > best_val:
            best_val, best = dsum, row
        ws.row_dimensions[r].height = 19
        r += 1

    # ---- totals row ----
    sty(ws.cell(r, 1), bold=True, size=11, color=WHITE, fill=INK,
        align="center").value = "TOTAL"
    sty(ws.cell(r, 2), fill=INK, color=WHITE, border=box).value = ""
    for j, b in enumerate(brands):
        sty(ws.cell(r, 3 + j), bold=True, size=10.5, color=WHITE, fill=INK,
            align="right", fmt=INR, indent=1).value = brand_tot[b]
    sty(ws.cell(r, dtot_col), bold=True, size=12, color=WHITE, fill=REDF,
        align="right", fmt=INR, indent=1).value = grand
    ws.row_dimensions[r].height = 26
    tot_row = r

    # ---- share % row ----
    r += 1
    sty(ws.cell(r, 1), bold=True, size=9.5, color=GREYD, fill=FILL2,
        align="center").value = "SHARE %"
    sty(ws.cell(r, 2), fill=FILL2, border=box).value = ""
    top_b = max(brand_tot, key=brand_tot.get)
    for j, b in enumerate(brands):
        share = (brand_tot[b] / grand) if grand else 0
        is_top = (b == top_b)
        sty(ws.cell(r, 3 + j), bold=is_top, size=10, color=(REDF if is_top else GREYD),
            fill=FILL2, align="right", fmt=PCT, indent=1).value = share
    sty(ws.cell(r, dtot_col), bold=True, size=10, color=INKTX, fill=FILL2,
        align="right", fmt=PCT, indent=1).value = 1 if grand else 0
    ws.row_dimensions[r].height = 18

    # ---- widths ----
    widths = {1: 11, 2: 12}
    for j, b in enumerate(brands):
        longest_word = max((len(w) for w in b.strip().split()), default=6)
        numw = max(len(inr(brand_tot[b])), 9)
        widths[3 + j] = max(longest_word + 3, numw + 3, 14)
    widths[dtot_col] = max(len(inr(grand)) + 4, 15)
    set_widths(ws, widths)

    ws.freeze_panes = ws.cell(hr + 1, 3)

    # ---- summary block ----
    sr = r + 2
    last = get_column_letter(ncol)
    ws.merge_cells(f"A{sr}:{last}{sr}")
    sty(ws.cell(sr, 1), bold=True, size=12.5, color=WHITE, fill=INK,
        align="left", border=None, indent=1).value = "ZONE SUMMARY  \u2014  September 2026"
    ws.row_dimensions[sr].height = 26
    # red rule under summary heading
    ws.merge_cells(f"A{sr+1}:{last}{sr+1}")
    ws.cell(sr + 1, 1).fill = PatternFill("solid", fgColor=REDF)
    ws.row_dimensions[sr + 1].height = 3

    ndays = len(rows)
    metrics = [
        ("Zone Grand Total", inr(grand)),
        ("Operating Days", f"{ndays} days"),
        ("Average Daily Sales", inr(grand / ndays) if ndays else "\u20b90"),
        ("Best Sales Day", f"{best['date'].strftime('%d %b')} ({best['day']})  \u2022  {inr(best_val)}"),
        ("Top Brand", f"{top_b.strip()}  \u2022  {inr(brand_tot[top_b])}  ({brand_tot[top_b]/grand*100:.1f}%)"),
        ("Active Brands", str(len(brands))),
    ]
    mr = sr + 2
    for i, (label, val) in enumerate(metrics):
        lab = ws.cell(mr, 1)
        sty(lab, bold=True, size=10.5, color=GREYD, fill=FILL2, align="left", indent=1).value = label
        ws.merge_cells(start_row=mr, start_column=2, end_row=mr, end_column=ncol)
        vv = ws.cell(mr, 2)
        # highlight the grand-total value in red
        sty(vv, bold=(i == 0), size=(12 if i == 0 else 10.5),
            color=(REDF if i == 0 else INKTX), align="left", indent=1).value = val
        for cc in range(3, ncol + 1):
            ws.cell(mr, cc).border = box
        ws.row_dimensions[mr].height = 21
        mr += 1

    # footer note
    mr += 1
    ws.merge_cells(start_row=mr, start_column=1, end_row=mr, end_column=ncol)
    sty(ws.cell(mr, 1), size=9, italic=True, color=GREY, align="left",
        border=None, indent=1).value = (
        "Weekends shown in red \u2022 \u2013 denotes no recorded sale \u2022 "
        "all totals recomputed from daily brand-wise figures")

    return {"grand": grand, "brand_tot": brand_tot, "ndays": ndays,
            "top_brand": top_b.strip(), "best": (best['date'], best_val),
            "brands": [b.strip() for b in brands]}


def build_summary_sheet(wb, results):
    ws = wb.create_sheet("SUMMARY", 0)
    ws.sheet_view.showGridLines = False
    ncol = 6
    title_band(ws, ncol, "SUMMARY",
               "All-Zone Sales Overview   \u2022   Amounts in \u20b9 (INR)   \u2022   Totals recomputed from daily data")
    # fix the title text (title_band prints "SUMMARY  SEPTEMBER 2026"); override row1
    ws.unmerge_cells("A1:F1")
    ws.merge_cells("A1:F1")
    rich = CellRichText([
        TextBlock(InlineFont(rFont=FONT, b=True, sz=22, color="FF" + REDT), "S"),
        TextBlock(InlineFont(rFont=FONT, b=True, sz=22, color="FFFFFFFF"),
                  "EPTEMBER 2026      ALL-ZONE SALES SUMMARY"),
    ])
    c = ws.cell(1, 1); c.value = rich
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.fill = PatternFill("solid", fgColor=INK)

    hr = 5
    heads = ["ZONE", "OPERATING DAYS", "ACTIVE BRANDS", "AVG DAILY SALES",
             "TOP BRAND", "ZONE GRAND TOTAL"]
    fills = [INK2, INK2, INK2, INK2, INK2, REDF]
    for j, (h, f) in enumerate(zip(heads, fills)):
        sty(ws.cell(hr, j + 1), bold=True, size=10, color=WHITE, fill=f,
            wrap=True).value = h
    ws.row_dimensions[hr].height = 34

    r = hr + 1
    grand_all = 0
    for i, (zone, d) in enumerate(results.items()):
        rf = BAND if i % 2 else WHITE
        sty(ws.cell(r, 1), bold=True, size=11, color=INKTX, align="left",
            fill=rf, indent=1).value = zone.strip()
        sty(ws.cell(r, 2), size=10.5, color=INKTX, fill=rf).value = f"{d['ndays']} days"
        sty(ws.cell(r, 3), size=10.5, color=INKTX, fill=rf).value = len(d["brands"])
        sty(ws.cell(r, 4), size=10.5, color=INKTX, fill=rf, align="right",
            fmt=INR, indent=1).value = d["grand"] / d["ndays"] if d["ndays"] else 0
        sty(ws.cell(r, 5), size=10.5, color=INKTX, fill=rf, align="left",
            indent=1).value = d["top_brand"]
        sty(ws.cell(r, 6), bold=True, size=11, color=REDF, fill=rf, align="right",
            fmt=INR, indent=1).value = d["grand"]
        grand_all += d["grand"]
        ws.row_dimensions[r].height = 24
        r += 1

    sty(ws.cell(r, 1), bold=True, size=11.5, color=WHITE, fill=INK,
        align="left", indent=1).value = "ALL ZONES"
    for cc in range(2, 6):
        sty(ws.cell(r, cc), fill=INK, border=box).value = ""
    sty(ws.cell(r, 6), bold=True, size=12.5, color=WHITE, fill=REDF, align="right",
        fmt=INR, indent=1).value = grand_all
    ws.row_dimensions[r].height = 28

    r += 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    sty(ws.cell(r, 1), size=9.5, italic=True, color=GREY, align="left",
        border=None, indent=1).value = (
        "Note: DISTRICT 9 has no September data in the source workbook "
        "(only July & August), so it is not included.")

    set_widths(ws, {1: 16, 2: 16, 3: 15, 4: 18, 5: 24, 6: 20})
    ws.freeze_panes = ws.cell(hr + 1, 1)


def main():
    src = load_workbook(SRC, data_only=True)
    out = Workbook()
    out.remove(out.active)
    results = {}
    for zone, meta in ZONES.items():
        brands, rows = extract(src[zone], meta["start"])
        print(f"[{zone.strip()}] brands={[b.strip() for b in brands]} days={len(rows)}")
        results[zone] = build_zone_sheet(out, zone, brands, rows)
    build_summary_sheet(out, results)
    out.save(OUT)
    print("Saved:", OUT)
    total_all = 0
    for zone, d in results.items():
        chk = sum(d["brand_tot"].values())
        total_all += d["grand"]
        print(f"  {zone.strip():12s} grand={inr(d['grand'])}  sum(brands)={inr(chk)}  ok={abs(chk-d['grand'])<1e-6}")
    print("  ALL ZONES   ", inr(total_all))


if __name__ == "__main__":
    main()
