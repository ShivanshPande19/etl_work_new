#!/usr/bin/env python3
"""Azimuth Food Street - File Distribution separator pages.

Design matches the supplied template:
  * double navy border frame
  * "AZIMUTH FOOD STREET" header + rule + "FILE DISTRIBUTION" sub-line
  * a small letter-spaced label (the Business Unit)
  * a big bold navy title (the brand) with a thick underline bar
  * "File No." / "Date" blanks at the bottom

One separate page is produced for every brand, grouped by Business Unit.
Each BU also gets its own divider page.
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

OUT = "Separator_Pages.pdf"

# ---- palette (from the template) ----
NAVY = HexColor("#1f3a63")
NAVY_DARK = HexColor("#17304f")
GREY = HexColor("#6c6c6c")
LINE_GREY = HexColor("#b9b9b9")

W, H = A4

# ---- Business Units and their brands (in order) ----
BUS = [
    {
        "bu": "BIG BEAR BU",
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
        "bu": "MELTING POT BU",
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


def fit_font_size(c, text, font, max_width, start=58, min_size=20):
    """Largest size (<= start) whose rendered width fits max_width."""
    size = start
    while size > min_size and c.stringWidth(text, font, size) > max_width:
        size -= 1
    return size


def draw_page(c, label, title):
    # ---------- double border ----------
    c.setStrokeColor(NAVY)
    c.setLineWidth(2.2)
    c.rect(15 * mm, 15 * mm, W - 30 * mm, H - 30 * mm, stroke=1, fill=0)
    c.setLineWidth(0.8)
    c.rect(18 * mm, 18 * mm, W - 36 * mm, H - 36 * mm, stroke=1, fill=0)

    # ---------- top header ----------
    spaced_centred(c, W / 2, H - 50 * mm, "AZIMUTH FOOD STREET",
                   "Helvetica-Bold", 18, NAVY, 2.0)
    c.setStrokeColor(LINE_GREY)
    c.setLineWidth(0.7)
    c.line(W / 2 - 48 * mm, H - 54 * mm, W / 2 + 48 * mm, H - 54 * mm)
    spaced_centred(c, W / 2, H - 60 * mm, "FILE DISTRIBUTION",
                   "Helvetica", 11, GREY, 3.0)

    # ---------- middle: label + big title ----------
    spaced_centred(c, W / 2, H / 2 + 6 * mm, label,
                   "Helvetica", 15, GREY, 6.0)

    max_w = W - 36 * mm - 24 * mm  # inside inner frame, with padding
    size = fit_font_size(c, title, "Helvetica-Bold", max_w, start=60, min_size=22)
    c.setFont("Helvetica-Bold", size)
    c.setFillColor(NAVY_DARK)
    title_y = H / 2 - 18 * mm
    c.drawCentredString(W / 2, title_y, title)

    # thick underline bar under the title
    bar_w = min(c.stringWidth(title, "Helvetica-Bold", size) + 6 * mm, max_w)
    c.setFillColor(NAVY)
    c.rect(W / 2 - bar_w / 2, title_y - 9 * mm, bar_w, 2.4 * mm, stroke=0, fill=1)

    # ---------- bottom: File No. / Date ----------
    c.setFont("Helvetica", 11)
    c.setFillColor(GREY)
    base_y = 34 * mm
    # File No.
    c.drawString(30 * mm, base_y, "File No.:")
    c.setStrokeColor(LINE_GREY)
    c.setLineWidth(0.7)
    c.line(48 * mm, base_y - 1 * mm, 95 * mm, base_y - 1 * mm)
    # Date
    c.drawString(115 * mm, base_y, "Date:")
    c.line(128 * mm, base_y - 1 * mm, 180 * mm, base_y - 1 * mm)

    c.showPage()


def main():
    c = canvas.Canvas(OUT, pagesize=A4)
    c.setTitle("Azimuth Food Street - File Distribution Separators")
    for item in BUS:
        # BU divider page
        draw_page(c, "BUSINESS UNIT", item["bu"])
        # one page per brand
        for brand in item["brands"]:
            draw_page(c, item["bu"], brand)
    c.save()
    pages = sum(1 + len(i["brands"]) for i in BUS)
    print("Wrote", OUT, "with", pages, "pages")


if __name__ == "__main__":
    main()
