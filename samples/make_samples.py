"""Generate the bundled sample PDFs with no third-party dependencies.

This writes two small, real PDFs into ``samples/`` so the extractors and the
router can be tried out without hunting for documents:

  * ``sample_law.pdf``  — portrait, legal-style headings (Chapter / Article ...),
                          spanning two pages with a running header/footer. Auto
                          routing picks *prose* mode.
  * ``sample_deck.pdf`` — landscape slides, including an "As-Is | To-Be" pair of
                          side-by-side tables. Auto routing picks *tables* mode.

It uses a tiny hand-rolled PDF writer (text + vector lines only) so it needs
nothing beyond the standard library.

    python samples/make_samples.py
"""

from __future__ import annotations

import os
from typing import List, Tuple


# --------------------------------------------------------------------------- #
# Minimal PDF writer (standard-library only)
# --------------------------------------------------------------------------- #

def _esc(s: str) -> str:
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def text_op(x: float, y: float, s: str, font: str = "F1", size: float = 11) -> str:
    """A content-stream snippet drawing one line of text at (x, y)."""
    return f"BT /{font} {size} Tf {x} {y} Td ({_esc(s)}) Tj ET\n"


def line_op(x1: float, y1: float, x2: float, y2: float) -> str:
    """A content-stream snippet drawing a stroked line segment."""
    return f"{x1} {y1} m {x2} {y2} l S\n"


def grid_table(
    x0: float, top: float, col_xs: List[float], row_ys: List[float],
    cells: List[List[str]], font: str = "F1", size: float = 11,
) -> str:
    """Draw a ruled grid plus its cell text.

    ``col_xs`` are the x positions of every vertical rule (len = ncols + 1).
    ``row_ys`` are the y positions of every horizontal rule, top to bottom
    (len = nrows + 1). ``cells[r][c]`` is the text for row r, column c.
    """
    parts = ["1 w\n"]
    x_left, x_right = col_xs[0], col_xs[-1]
    y_top, y_bottom = row_ys[0], row_ys[-1]
    for x in col_xs:                       # vertical rules
        parts.append(line_op(x, y_top, x, y_bottom))
    for y in row_ys:                       # horizontal rules
        parts.append(line_op(x_left, y, x_right, y))
    for r, row in enumerate(cells):        # cell text, nudged inside each cell
        for c, value in enumerate(row):
            if not value:
                continue
            tx = col_xs[c] + 4
            ty = row_ys[r] - (row_ys[r] - row_ys[r + 1]) * 0.65
            parts.append(text_op(tx, ty, value, font=font, size=size))
    return "".join(parts)


class PDF:
    """Accumulates pages and serializes a valid PDF byte string."""

    def __init__(self) -> None:
        self._pages: List[Tuple[float, float, str]] = []

    def add_page(self, width: float, height: float, content: str) -> None:
        self._pages.append((width, height, content))

    def to_bytes(self) -> bytes:
        n_pages = len(self._pages)
        # Fixed objects: 1 Catalog, 2 Pages, 3 Helvetica, 4 Helvetica-Bold.
        page_nums, content_nums = [], []
        nxt = 5
        for _ in range(n_pages):
            page_nums.append(nxt); nxt += 1
            content_nums.append(nxt); nxt += 1
        max_num = nxt - 1

        parts = {
            1: b"<< /Type /Catalog /Pages 2 0 R >>",
            3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            4: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
        }
        kids = " ".join(f"{p} 0 R" for p in page_nums)
        parts[2] = (
            f"<< /Type /Pages /Kids [{kids}] /Count {n_pages} >>".encode("latin-1")
        )
        for idx, (w, h, content) in enumerate(self._pages):
            pnum, cnum = page_nums[idx], content_nums[idx]
            parts[pnum] = (
                f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {w} {h}] "
                f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
                f"/Contents {cnum} 0 R >>"
            ).encode("latin-1")
            cb = content.encode("latin-1")
            parts[cnum] = (
                f"<< /Length {len(cb)} >>\nstream\n".encode("latin-1")
                + cb
                + b"\nendstream"
            )

        out = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
        offsets = {}
        for num in range(1, max_num + 1):
            offsets[num] = len(out)
            out += f"{num} 0 obj\n".encode("latin-1") + parts[num] + b"\nendobj\n"

        xref_pos = len(out)
        out += f"xref\n0 {max_num + 1}\n".encode("latin-1")
        out += b"0000000000 65535 f\r\n"                      # 20-byte entries
        for num in range(1, max_num + 1):
            out += f"{offsets[num]:010d} 00000 n\r\n".encode("latin-1")
        out += b"trailer\n"
        out += f"<< /Size {max_num + 1} /Root 1 0 R >>\n".encode("latin-1")
        out += f"startxref\n{xref_pos}\n%%EOF".encode("latin-1")
        return out

    def save(self, path: str) -> None:
        with open(path, "wb") as fh:
            fh.write(self.to_bytes())


# --------------------------------------------------------------------------- #
# Sample documents
# --------------------------------------------------------------------------- #

def build_law(path: str) -> None:
    """Portrait law with Chapter/Article headings spanning two pages."""
    W, H = 595, 842
    pdf = PDF()

    def header_footer(page_no: int) -> str:
        return (
            text_op(60, H - 40, "Official Journal of the Union", "F1", 8)
            + text_op(W / 2 - 20, 36, f"Page {page_no}", "F1", 8)
        )

    # ---- Page 1 ----
    c = header_footer(1)
    y = H - 90
    c += text_op(60, y, "Chapter 1 General Provisions", "F2", 14); y -= 36
    c += text_op(60, y, "Article 1 (Purpose) The purpose of this Regulation is to", "F2", 12); y -= 22
    c += text_op(60, y, "establish harmonised rules for placing products on the", "F1", 11); y -= 18
    c += text_op(60, y, "market across the territories covered by this instrument.", "F1", 11); y -= 36
    c += text_op(60, y, "Article 2 (Scope) This Regulation applies to all operators", "F2", 12); y -= 22
    c += text_op(60, y, "who manufacture, import or distribute regulated products,", "F1", 11); y -= 18
    # Deliberately leave this sentence unfinished to continue on page 2.
    c += text_op(60, y, "and to any representative acting on their behalf within the", "F1", 11)
    pdf.add_page(W, H, c)

    # ---- Page 2 ----
    c = header_footer(2)
    y = H - 90
    # Continuation (starts lowercase, prior sentence not ended) -> merges back
    # into Article 2's paragraph once the running header is stripped.
    c += text_op(60, y, "internal market of the relevant jurisdiction.", "F1", 11); y -= 36
    c += text_op(60, y, "Article 3 (Definitions) For the purposes of this Regulation", "F2", 12); y -= 22
    c += text_op(60, y, "the following definitions apply consistently throughout.", "F1", 11); y -= 36
    c += text_op(60, y, "Article 4 (Obligations) Operators shall keep records for", "F2", 12); y -= 22
    c += text_op(60, y, "a period of at least five years and make them available to", "F1", 11); y -= 18
    # Runs off the bottom unfinished, continuing on page 3.
    c += text_op(60, y, "the competent authority, and shall ensure that such records", "F1", 11)
    pdf.add_page(W, H, c)

    # ---- Page 3 ----
    c = header_footer(3)
    y = H - 90
    c += text_op(60, y, "remain accurate and complete throughout that period.", "F1", 11); y -= 36
    c += text_op(60, y, "Chapter 2 Final Provisions", "F2", 14); y -= 36
    c += text_op(60, y, "Article 5 (Entry into Force) This Regulation shall enter", "F2", 12); y -= 22
    c += text_op(60, y, "into force on the twentieth day following its publication", "F1", 11); y -= 18
    c += text_op(60, y, "in the Official Journal and must be applied in full.", "F1", 11)
    pdf.add_page(W, H, c)

    pdf.save(path)


def build_deck(path: str) -> None:
    """Landscape slides incl. an As-Is | To-Be side-by-side table pair."""
    W, H = 842, 595
    pdf = PDF()

    # ---- Slide 1: text only ----
    c = text_op(60, H - 70, "Process Transformation Programme", "F2", 22)
    c += text_op(60, H - 110, "Goals for the current fiscal year", "F2", 14)
    c += text_op(70, H - 140, "- Reduce manual effort across the back office", "F1", 12)
    c += text_op(70, H - 162, "- Improve data quality and turnaround time", "F1", 12)
    c += text_op(70, H - 184, "- Establish real-time reporting", "F1", 12)
    pdf.add_page(W, H, c)

    # ---- Slide 2: two single-column tables side by side -> stitched ----
    c = text_op(60, H - 60, "Current vs Target State", "F2", 20)
    c += text_op(60, H - 90, "The left column is today; the right is the target.", "F1", 12)
    rows = 4
    row_ys = [H - 120 - i * 34 for i in range(rows + 1)]   # 5 horizontal rules
    left = grid_table(
        60, row_ys[0], [60, 360], row_ys,
        [["As-Is"], ["Manual data entry"], ["Paper-based forms"], ["Weekly close"]],
    )
    right = grid_table(
        400, row_ys[0], [400, 720], row_ys,
        [["To-Be"], ["Automated capture"], ["Digital forms"], ["Real-time close"]],
    )
    pdf.add_page(W, H, c + left + right)

    # ---- Slide 3: a single three-column table ----
    c = text_op(60, H - 60, "Milestones", "F2", 20)
    row_ys = [H - 100 - i * 34 for i in range(5)]          # 4 rows -> 5 rules
    col_xs = [60, 260, 520, 780]
    table = grid_table(
        60, row_ys[0], col_xs, row_ys,
        [
            ["Phase", "Owner", "Target date"],
            ["Discovery", "Operations", "Q1"],
            ["Build", "Engineering", "Q2"],
            ["Rollout", "Change team", "Q3"],
        ],
    )
    pdf.add_page(W, H, c + table)

    pdf.save(path)


def main() -> None:
    here = os.path.dirname(os.path.abspath(__file__))
    law = os.path.join(here, "sample_law.pdf")
    deck = os.path.join(here, "sample_deck.pdf")
    build_law(law)
    build_deck(deck)
    print(f"wrote {law}")
    print(f"wrote {deck}")


if __name__ == "__main__":
    main()
