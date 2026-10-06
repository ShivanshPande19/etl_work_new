#!/usr/bin/env python3
"""Azimuth Food Street - brand separator pages.

Layout (per the latest instruction):
  * double navy border frame
  * TOP   : "EAT TRUCK LOVE BY AZIMUTH"  (+ rule)
  * BELOW : "Zone: <Business Unit>"
  * MIDDLE: the brand name (big, bold) with a thick underline bar
  * BOTTOM: "File No." / "Date" blanks

One separate page per brand.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

OUT = "Separator_Pages.pdf"

# ---- palette ----
NAVY = HexColor("#1f3a63")
NAVY_DARK = HexColor("#17304f")
GREY = HexColor("#6c6c6c")
LINE_GREY = HexColor("#b9b9b9")

W, H = A4

TOP_HEADER = "EAT TRUCK LOVE BY AZIMUTH"

# ---- zones (Business Units) and their brands, in order ----
ZONES = [
    {
        "zone": "Big Bear BU",
        "brands": [
            "OG PIZZA",
            "BUKHARA2BEIJING",
            "THE MOMO BOX",
            "BUN BURGER FRIES",
            "THE COFFEE VAULT",
            "KOBY'S",
            "BURRITO BRO'S",
        ],
    },
    {
        "zone": "Melting Pot BU",
        "brands": ["BIGG BEAR", "EGGSPERT", "THE CHATPATA AFFAIR", "CROWNEST"],
    },
]


def spaced_centred(c, cx, y, text, font, size, color, char_space):
    """Draw letter-spaced text centred on cx (manual tracking)."""
    c.setFont(font, size)
    c.setFillColor(color)
    widths = [c.stringWidth(ch, font, size) for ch in text]
    total = sum(widths) + char_space * max(len(text) - 1, 0)
    x = cx - total / 2.0
    for ch, w in zip(text, widths):
        c.drawString(x, y, ch)
        x += w + char_space
    return total


def fit_font_size(c, text, font, max_width, start=58, min_size=18):
    size = start
    while size > min_size and c.stringWidth(text, font, size) > max_width:
        size -= 1
    return size


def draw_page(c, zone, title):
    # ---------- double border ----------
    c.setStrokeColor(NAVY)
    c.setLineWidth(2.2)
    c.rect(15 * mm, 15 * mm, W - 30 * mm, H - 30 * mm, stroke=1, fill=0)
    c.setLineWidth(0.8)
    c.rect(18 * mm, 18 * mm, W - 36 * mm, H - 36 * mm, stroke=1, fill=0)

    # ---------- top header ----------
    hdr_width = spaced_centred(c, W / 2, H - 50 * mm, TOP_HEADER,
                               "Helvetica-Bold", 17, NAVY, 1.6)
    half = max(hdr_width / 2 + 6 * mm, 46 * mm)
    c.setStrokeColor(LINE_GREY)
    c.setLineWidth(0.7)
    c.line(W / 2 - half, H - 54 * mm, W / 2 + half, H - 54 * mm)

    # zone line, just below the header rule
    spaced_centred(c, W / 2, H - 61 * mm, f"Zone: {zone}",
                   "Helvetica", 12, GREY, 2.0)

    # ---------- middle: big brand name ----------
    max_w = W - 36 * mm - 24 * mm
    size = fit_font_size(c, title, "Helvetica-Bold", max_w, start=60, min_size=20)
    c.setFont("Helvetica-Bold", size)
    c.setFillColor(NAVY_DARK)
    title_y = H / 2 - 8 * mm
    c.drawCentredString(W / 2, title_y, title)

    bar_w = min(c.stringWidth(title, "Helvetica-Bold", size) + 6 * mm, max_w)
    c.setFillColor(NAVY)
    c.rect(W / 2 - bar_w / 2, title_y - 9 * mm, bar_w, 2.4 * mm, stroke=0, fill=1)

    # ---------- bottom: File No. / Date ----------
    c.setFont("Helvetica", 11)
    c.setFillColor(GREY)
    base_y = 34 * mm
    c.drawString(30 * mm, base_y, "File No.:")
    c.setStrokeColor(LINE_GREY)
    c.setLineWidth(0.7)
    c.line(48 * mm, base_y - 1 * mm, 95 * mm, base_y - 1 * mm)
    c.drawString(115 * mm, base_y, "Date:")
    c.line(128 * mm, base_y - 1 * mm, 180 * mm, base_y - 1 * mm)

    c.showPage()


def main():
    c = canvas.Canvas(OUT, pagesize=A4)
    c.setTitle("Azimuth Food Street - Brand Separators")
    n = 0
    for z in ZONES:
        for brand in z["brands"]:
            draw_page(c, z["zone"], brand)
            n += 1
    c.save()
    print("Wrote", OUT, "with", n, "pages")


if __name__ == "__main__":
    main()
