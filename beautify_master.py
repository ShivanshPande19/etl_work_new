#!/usr/bin/env python3
"""
Beautify the existing master workbook *in place style-wise* (data preserved).

- Keeps every sheet, every row, all sales numbers and all TOTAL formulas.
- Adds a premium, soothing look: teal-slate header + vertical month bands,
  soft sage/warm-neutral banding, thin borders, Indian-rupee number format,
  emphasised monthly TOTAL rows, weekend highlight, frozen header + labels.
- Cleans up the DATE column: regenerates correct, uniform dates from each
  month block (verified against the DAY column) because the source dates were
  entered inconsistently (day/month swapped for the 1st-12th, text after).

Writes a new file so the original stays untouched.
"""

import datetime as dt
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = "MASTER SALES REPORT -2.xlsx"
OUT = "MASTER SALES REPORT -2 (formatted).xlsx"
FONT = "Segoe UI"

# soothing palette (matches the reports)
INK   = "34515B"   # month band / strong bands
INK2  = "4A6D76"   # header row
INKTX = "2F3E45"   # primary text
GREYD = "5C6B72"
LINE  = "D9D3C8"   # soft warm hairline
BAND  = "F7F5F0"   # zebra
TOTF  = "DCEAE4"   # monthly total row (pale sage)
WKND  = "F1EDE4"   # weekend tint
SAGE  = "4F7D6E"   # sage accent (weekend day text / totals)
WHITE = "FFFFFF"

# Indian rupee; >=1 lakh grouped as lakh, else thousands, zero shown blank
INR = '[>=100000]"\u20b9"##\\,##\\,##0;[>0]"\u20b9"#,##0;""'

MONTHS = {m.upper(): i for i, m in enumerate(
    ["january","february","march","april","may","june","july",
     "august","september","october","november","december"], start=1)}
MONTHS.update({m[:3].upper(): i for m, i in list(MONTHS.items())})
DAYNUM = {d: i for i, d in enumerate(
    ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"])}

hair = Side(style="thin", color=LINE)
box = Border(left=hair, right=hair, top=hair, bottom=hair)
thick = Side(style="medium", color=INK2)


def month_of(v):
    if v is None:
        return None
    key = str(v).strip().upper()
    return MONTHS.get(key)


def is_blank(ws, r, ncol):
    return all(ws.cell(r, c).value in (None, "") for c in range(1, ncol + 1))


def last_used(ws, ncol):
    last = 1
    for r in range(1, ws.max_row + 1):
        if not is_blank(ws, r, ncol):
            last = r
    return last


def beautify(ws):
    # total column = last header containing TOTAL
    ncol = ws.max_column
    total_col = max((c for c in range(1, ncol + 1)
                     if ws.cell(1, c).value and "TOTAL" in str(ws.cell(1, c).value).upper()),
                    default=ncol)
    last = last_used(ws, total_col)
    ws.sheet_view.showGridLines = False

    # ---- header row ----
    labels = {1: "MONTH", 2: "DATE", 3: "DAY"}
    for c in range(1, total_col + 1):
        cell = ws.cell(1, c)
        if c in labels:
            cell.value = labels[c]
        elif cell.value is not None:
            cell.value = str(cell.value).strip().upper()
        cell.font = Font(name=FONT, bold=True, size=10.5, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill("solid", fgColor=INK2)
        cell.border = Border(left=hair, right=hair, top=hair, bottom=thick)
    ws.row_dimensions[1].height = 30

    # ---- month blocks ----
    month_rows = [(r, month_of(ws.cell(r, 1).value))
                  for r in range(2, last + 1) if month_of(ws.cell(r, 1).value)]
    blocks = []
    for i, (sr, mon) in enumerate(month_rows):
        er = (month_rows[i + 1][0] - 1) if i + 1 < len(month_rows) else last
        # trim trailing fully-blank rows
        while er > sr and is_blank(ws, er, total_col):
            er -= 1
        blocks.append((sr, er, mon, str(ws.cell(sr, 1).value).strip()))

    mismatches = 0
    for sr, er, mon, label in blocks:
        # year from a datetime cell in the block (its year is reliable)
        year = None
        for r in range(sr, er + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, (dt.datetime, dt.date)):
                year = v.year
                break
            if isinstance(v, str):
                m = re.match(r"^\s*(\d{1,2})/(\d{1,2})/(\d{2})\s*$", v)
                if m:
                    year = 2000 + int(m.group(3))
                    break
        # style + regenerate rows
        day_i = 0
        zebra_i = 0
        for r in range(sr, er + 1):
            day = ws.cell(r, 3).value
            daily = day not in (None, "")
            has_vals = any(ws.cell(r, c).value not in (None, "")
                           for c in range(4, total_col + 1))
            is_total = (not daily) and has_vals
            wknd = daily and str(day).strip().upper() in ("SATURDAY", "SUNDAY")

            if daily and year:
                day_i += 1
                try:
                    d = dt.date(year, mon, day_i)
                    if DAYNUM.get(str(day).strip().upper()) != d.weekday():
                        mismatches += 1
                    ws.cell(r, 2).value = d
                    ws.cell(r, 2).number_format = "dd mmm yy"
                except ValueError:
                    pass

            # row fill
            if is_total:
                rowfill = TOTF
            elif wknd:
                rowfill = WKND
            else:
                rowfill = BAND if zebra_i % 2 else WHITE
            if daily:
                zebra_i += 1

            for c in range(2, total_col + 1):
                cell = ws.cell(r, c)
                bold = is_total
                if c == 2:  # DATE
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name=FONT, size=10, bold=bold, color=INKTX)
                elif c == 3:  # DAY
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name=FONT, size=10, bold=(bold or wknd),
                                     color=(SAGE if wknd else GREYD))
                else:  # numbers
                    cell.alignment = Alignment(horizontal="right", vertical="center", indent=1)
                    cell.font = Font(name=FONT, size=10, bold=bold,
                                     color=(INK if is_total else INKTX))
                    if isinstance(cell.value, (int, float)) or is_total or daily:
                        cell.number_format = INR
                cell.fill = PatternFill("solid", fgColor=rowfill)
                top = thick if is_total else hair
                cell.border = Border(left=hair, right=hair, top=top, bottom=hair)

        # ---- vertical month band (col A merged over the block) ----
        ws.merge_cells(start_row=sr, start_column=1, end_row=er, end_column=1)
        band = ws.cell(sr, 1)
        band.value = label.upper()
        band.font = Font(name=FONT, bold=True, size=12, color=WHITE)
        band.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)
        band.fill = PatternFill("solid", fgColor=INK)
        for r in range(sr, er + 1):
            ws.cell(r, 1).border = Border(left=hair, right=hair, top=hair, bottom=hair)

    # ---- widths & freeze ----
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 11
    ws.column_dimensions["C"].width = 12
    for c in range(4, total_col):
        w = max((len(str(ws.cell(1, c).value or "")) for _ in (0,)), default=8)
        ws.column_dimensions[get_column_letter(c)].width = max(w + 2, 13)
    ws.column_dimensions[get_column_letter(total_col)].width = 15
    ws.freeze_panes = ws.cell(2, 4)  # keep header + MONTH/DATE/DAY visible
    return mismatches


def main():
    wb = load_workbook(SRC, data_only=False)  # keep formulas
    total_mismatch = 0
    for ws in wb.worksheets:
        m = beautify(ws)
        total_mismatch += m
        print(f"  styled {ws.title!r:14s} date/day mismatches: {m}")
    wb.save(OUT)
    print("Saved:", OUT, "| total date mismatches:", total_mismatch)


if __name__ == "__main__":
    main()
