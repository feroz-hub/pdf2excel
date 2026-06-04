"""Generate a small multi-page test PDF for exercising the extractor.

This is a developer helper (not part of the shipped tool). It needs reportlab:
    pip install reportlab
    python samples/_make_test_pdf.py
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas


def build(path: str = "samples/test_guideline.pdf") -> None:
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    def header_footer(page_no: int) -> None:
        c.setFont("Helvetica", 8)
        c.drawString(2 * cm, height - 1.2 * cm, "Regulatory Guideline ABC-123")
        c.drawCentredString(width / 2, 1.2 * cm, f"Page {page_no}")

    def line(text, x, y, font="Helvetica", size=11):
        c.setFont(font, size)
        c.drawString(x, y, text)

    # ---- Page 1 ----
    header_footer(1)
    y = height - 3 * cm
    line("4 Scope", 2 * cm, y, "Helvetica-Bold", 13); y -= 0.9 * cm
    line("This guideline applies to all operators who process regu-", 2 * cm, y); y -= 0.6 * cm
    line("lated materials within the jurisdiction described in Annex A.", 2 * cm, y); y -= 1.2 * cm
    line("4.1 Definitions", 2 * cm, y, "Helvetica-Bold", 13); y -= 0.9 * cm
    line("For the purposes of this document the following definitions", 2 * cm, y); y -= 0.6 * cm
    # This paragraph deliberately runs off the bottom and continues on page 2.
    line("apply and shall be interpreted consistently across all", 2 * cm, y); y -= 0.6 * cm
    line("sections of the present", 2 * cm, y)

    c.showPage()

    # ---- Page 2 ----
    header_footer(2)
    y = height - 3 * cm
    # continuation (starts lowercase, previous para did not end a sentence)
    line("guideline unless explicitly stated otherwise.", 2 * cm, y); y -= 1.2 * cm
    line("4.2 General Requirements", 2 * cm, y, "Helvetica-Bold", 13); y -= 0.9 * cm
    line("Operators shall maintain records for a period of five years.", 2 * cm, y); y -= 1.2 * cm
    line("Records shall be made available to the authority on request.", 2 * cm, y)

    c.showPage()
    c.save()
    print(f"wrote {path}")


if __name__ == "__main__":
    build()
