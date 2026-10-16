#!/usr/bin/env python3
"""Author synthetic PDF regression inputs; no downloaded or original PDF input.

Requires reportlab==4.4.3, pillow==11.3.0 and Poppler pdftoppm on PATH.
PDF metadata is deterministic and contains no local identity or document path.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
import subprocess
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle

ROOT = Path(__file__).resolve().parent
WIDTH, HEIGHT = A4
BODY = ParagraphStyle("body", fontName="Helvetica", fontSize=12, leading=18)


def document(name, title):
    pdf = canvas.Canvas(str(ROOT / name), pagesize=A4, pageCompression=1, invariant=1)
    pdf.setTitle(title)
    pdf.setAuthor("doc.haus synthetic regression fixtures")
    pdf.setSubject("Freshly authored synthetic text; no client or third-party document content")
    pdf.setCreator("generate-synthetic-pdfs.py / ReportLab")
    return pdf


def heading(pdf, text, y, size=20):
    pdf.setFont("Helvetica-Bold", size)
    pdf.drawString(54, y, text)


def paragraph(pdf, text, y):
    flow = Paragraph(text, BODY)
    _, height = flow.wrap(WIDTH - 108, HEIGHT)
    flow.drawOn(pdf, 54, y - height)
    return y - height - 16


def footer(pdf, page):
    pdf.setFont("Helvetica", 9)
    pdf.drawString(54, 36, "SYNTHETIC REGRESSION FIXTURE - NOT A REAL MATTER")
    pdf.drawRightString(WIDTH - 54, 36, str(page))


def text_layer():
    pdf = document("synthetic-text-layer.pdf", "Synthetic workshop plan")
    heading(pdf, "SYNTHETIC WORKSHOP PLAN", HEIGHT - 64)
    y = paragraph(pdf, "This invented plan describes a paper-folding workshop. Every sentence was written for parser regression testing. It identifies no actual person, organization, transaction or legal obligation.", HEIGHT - 100)
    heading(pdf, "1. Materials", y, 15)
    y = paragraph(pdf, "The organizer places twelve blue paper squares and four reusable trays on the demonstration desk. Participants return unused paper to the marked storage box after the final exercise.", y - 24)
    heading(pdf, "2.3 Schedule changes", y, 15)
    y = paragraph(pdf, "The organizer sends a written reminder before changing the workshop time. Participants keep the reminder beside the printed timetable. A new timetable replaces the earlier timetable only after the room booking is confirmed.", y - 24)
    heading(pdf, "3. Closing checklist", y, 15)
    paragraph(pdf, "Count the reusable trays, close the storage box and leave the demonstration desk clear. The final checklist is an invented extraction marker and does not describe an actual event.", y - 24)
    footer(pdf, 1)
    pdf.save()


def continuation():
    pdf = document("synthetic-continuation.pdf", "Synthetic equipment handover note")
    heading(pdf, "SYNTHETIC EQUIPMENT HANDOVER", HEIGHT - 64)
    y = paragraph(pdf, "This fictional note exercises headings and text continued on another page. A sample equipment box contains wooden counters, paper labels and an empty inventory card. All descriptions are authored test content.", HEIGHT - 100)
    heading(pdf, "1. Inspection", y, 15)
    paragraph(pdf, "Read the inventory card before moving the sample box. The counters are grouped by shape so that the next reader can check the inventory without guessing how the groups were assembled.", y - 24)
    heading(pdf, "2. Transport sequence", 164, 15)
    paragraph(pdf, "The carrier places the triangular counters into the green tray before", 140)
    footer(pdf, 1)
    pdf.showPage()
    paragraph(pdf, "placing the circular counters into the amber tray. This sentence continues the transport instruction from the preceding page and must remain after its opening words in extracted text.", HEIGHT - 74)
    heading(pdf, "3. Receipt check", HEIGHT - 170, 15)
    paragraph(pdf, "The recipient compares both trays with the inventory card and records any missing counter. The receipt check closes this synthetic note. No signatures, parties or real-world promises are represented.", HEIGHT - 194)
    footer(pdf, 2)
    pdf.save()


def table():
    pdf = document("synthetic-table.pdf", "Synthetic workshop checklist form")
    heading(pdf, "SYNTHETIC WORKSHOP CHECKLIST", HEIGHT - 64)
    paragraph(pdf, "This invented form tests two-column grid extraction. Row labels and their values should stay paired. It is a parser fixture, not a completed record or a document taken from another source.", HEIGHT - 100)
    heading(pdf, "Appendix B", HEIGHT - 200, 17)
    x, y, col, row = 54, HEIGHT - 232, (WIDTH - 108) / 2, 44
    rows = [("CHECK", "EXPECTED RESULT"), ("Tray colour", "Amber"), ("Counter shape", "Triangle"), ("Storage location", "Upper shelf"), ("Checked by: [name]", "Reviewed by: [name]")]
    pdf.setLineWidth(1)
    for index, cells in enumerate(rows):
        top = y - row * index
        for column, value in enumerate(cells):
            pdf.rect(x + column * col, top - row, col, row)
            pdf.setFont("Helvetica-Bold" if index == 0 else "Helvetica", 11)
            pdf.drawString(x + column * col + 10, top - 27, value)
    paragraph(pdf, "End of synthetic form. The table deliberately contains distinct labels and values so a flattened or reordered extraction is observable.", y - row * len(rows) - 24)
    footer(pdf, 1)
    pdf.save()


def scan():
    # Render only our authored text-layer fixture. The final PDF embeds pixels,
    # with no OCR text, invisible text or imported original PDF content.
    with TemporaryDirectory(prefix="dochaus-synthetic-scan-") as directory:
        prefix = Path(directory) / "page"
        subprocess.run(["pdftoppm", "-f", "1", "-singlefile", "-r", "180", "-gray", "-png", str(ROOT / "synthetic-text-layer.pdf"), str(prefix)], check=True)
        pdf = document("synthetic-scanned.pdf", "Synthetic workshop plan image-only scan")
        pdf.drawImage(ImageReader(str(prefix) + ".png"), 0, 0, width=WIDTH, height=HEIGHT)
        pdf.save()


if __name__ == "__main__":
    text_layer()
    continuation()
    table()
    scan()
