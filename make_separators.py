#!/usr/bin/env python3
"""Generate section-divider / separator pages for the Azimuth Food Street booklet."""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

OUT = "Separator_Pages.pdf"

HEADER = "Eat Truck Love by Azimuth Business on Wheels"

# Business Units and their brands, in order.
BUS = [
    {
        "bu": "Melting Pot BU",
        "brands": ["Bigg Bear", "Eggspert", "The Chatpata Affair", "Crownest"],
    },
    {
        "bu": "Big Bear BU",
        "brands": [
            "OG Pizza",
            "Bukhara2Beijing",
            "The Momo Box",
            "Bun Burger Fries",
            "The Coffee Vault",
            "Koby's",
            "Burrito Bro's",
        ],
    },
]

# Palette
DARK = HexColor("#1b2a4a")
ACCENT = HexColor("#d4891f")
GREY = HexColor("#555555")

W, H = A4


def draw_page(c, bu, brands):
    # background band at top
    c.setFillColor(DARK)
    c.rect(0, H - 42 * mm, W, 42 * mm, stroke=0, fill=1)

    # top header text (company line)
    c.setFillColor(HexColor("#ffffff"))
    c.setFont("Helvetica-Bold", 17)
    c.drawCentredString(W / 2, H - 24 * mm, HEADER)

    # thin accent rule under header band
    c.setFillColor(ACCENT)
    c.rect(0, H - 44 * mm, W, 2 * mm, stroke=0, fill=1)

    # middle: BU name (big)
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", 46)
    c.drawCentredString(W / 2, H / 2 + 30 * mm, bu)

    # decorative short rule under BU name
    rule_w = 60 * mm
    c.setFillColor(ACCENT)
    c.rect((W - rule_w) / 2, H / 2 + 20 * mm, rule_w, 1.6 * mm, stroke=0, fill=1)

    # "Brands" label
    c.setFillColor(GREY)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, H / 2 + 2 * mm, "Brands")

    # brand list, centered
    c.setFillColor(DARK)
    c.setFont("Helvetica", 20)
    y = H / 2 - 14 * mm
    for i, b in enumerate(brands, start=1):
        c.drawCentredString(W / 2, y, f"{i}.  {b}")
        y -= 11 * mm

    # footer accent bar
    c.setFillColor(ACCENT)
    c.rect(0, 0, W, 10 * mm, stroke=0, fill=1)

    c.showPage()


def main():
    c = canvas.Canvas(OUT, pagesize=A4)
    c.setTitle("Azimuth Food Street - Section Separators")
    for item in BUS:
        draw_page(c, item["bu"], item["brands"])
    c.save()
    print("Wrote", OUT, "with", len(BUS), "separator pages")


if __name__ == "__main__":
    main()
