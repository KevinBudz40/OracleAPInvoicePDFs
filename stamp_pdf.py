"""
stamp_pdf.py  —  adds a small filled black circle to the upper-left corner
                 of page 1 of a PDF, then overwrites the file in place.

Requirements:
    pip install pypdf reportlab

Usage (called automatically by AP_Invoice_Downloader.xlsm):
    python stamp_pdf.py "C:\path\to\invoice.pdf"
"""

import sys
import io
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas

# Position of the dot centre, measured from the upper-left corner (points).
# 1 pt = 1/72 inch.  PDF y-axis runs bottom-to-top, so we flip internally.
MARGIN = 15   # pts from each edge to dot centre
RADIUS =  6   # pts


def stamp(path: str) -> None:
    reader = PdfReader(path)
    first  = reader.pages[0]
    w = float(first.mediabox.width)
    h = float(first.mediabox.height)

    # Build a single-page overlay that contains only the dot.
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(w, h))
    c.setFillColorRGB(0, 0, 0)
    # PDF origin is bottom-left, so "upper-left" → x=MARGIN, y=h-MARGIN
    c.circle(MARGIN, h - MARGIN, RADIUS, stroke=0, fill=1)
    c.save()
    buf.seek(0)

    overlay = PdfReader(buf).pages[0]
    first.merge_page(overlay)

    writer = PdfWriter()
    writer.add_page(first)
    for page in reader.pages[1:]:
        writer.add_page(page)

    with open(path, "wb") as f:
        writer.write(f)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("Usage: stamp_pdf.py <pdf_path>")
    stamp(sys.argv[1])
    print("OK")
