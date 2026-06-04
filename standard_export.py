"""Export extractor output into the "Standard Assessment" Excel format.

This writes into ``templates/Standard.xlsx`` **as a template** — it is loaded
with openpyxl and only the metadata cells and the data rows (10+) are touched.
Merged cells, the Times New Roman header styling, the data-validation dropdowns
and the named ranges (CLASSIFICATION / Requirement_Applicability /
Responsible_by / Information) are left untouched so they keep working.

The data rows are *cleared by value* (set to ``None``) rather than deleted —
deleting rows would corrupt the data validations whose ranges (E10:E1020,
I10:I1020, ...) are anchored to specific rows.

Public API:
    write_standard_assessment(items, out_path, template_path=..., standard_id=...,
                              standard_title=..., standard_edition=...,
                              document_id=..., document_name=..., document_revision=...)
    paragraphs_to_items(paras)         # prose -> items
    deck_to_items(pages_tables, slides_text)  # deck -> items

An ``item`` is a dict: {clause_id, title, text, classification}.
"""

from __future__ import annotations

import os
import re
from typing import List

# Default location of the bundled template, resolved relative to this file so
# it works regardless of the current working directory.
DEFAULT_TEMPLATE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "templates", "Standard.xlsx"
)

SHEET_NAME = "Standard Assessment"
DATA_START_ROW = 10
_LAST_COL = 17  # column Q

# Leading legal clause token, e.g. "Article 1", "Chapter 2", "Annex 3".
_HEADING_ID_RE = re.compile(
    r"^(Article|Chapter|Section|Part|Annex|Schedule)\s+\d+", re.IGNORECASE
)
# Fallback for numbered headings, e.g. "4.2 Documentation".
_NUMBERED_ID_RE = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(.*)$")
# Obligation language that promotes a paragraph to a "Requirement".
_REQUIREMENT_RE = re.compile(
    r"\b(shall not|shall|must|is required to|are required to)\b", re.IGNORECASE
)


# --------------------------------------------------------------------------- #
# Mapping layer: extractor output -> items
# --------------------------------------------------------------------------- #

def split_heading(label: str):
    """Split a heading label into (clause_id, title).

    "Article 1 (Purpose)"       -> ("Article 1", "Purpose")
    "Chapter 2 General Matters" -> ("Chapter 2", "General Matters")
    "4.2 Documentation"         -> ("4.2", "Documentation")
    Anything unrecognised        -> (label, "").
    """
    label = (label or "").strip()
    m = _HEADING_ID_RE.match(label)
    if m:
        clause_id = m.group(0).strip()
        rest = label[m.end():].strip()
        if rest.startswith("(") and rest.endswith(")"):
            rest = rest[1:-1].strip()
        return clause_id, rest
    m = _NUMBERED_ID_RE.match(label)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return label, ""


def _classify(text: str, classify_requirements: bool) -> str:
    if classify_requirements and _REQUIREMENT_RE.search(text or ""):
        return "Requirement"
    return "Information"


def paragraphs_to_items(paras, classify_requirements: bool = True) -> List[dict]:
    """Turn prose paragraphs into items.

    Heading paragraphs update the running clause id/title but emit no row of
    their own; each body paragraph becomes one item.
    """
    items: List[dict] = []
    cur_id, cur_title = "", ""
    for p in paras:
        if getattr(p, "type", "body") == "heading":
            cur_id, cur_title = split_heading(p.text)
            continue
        items.append(
            {
                "clause_id": cur_id,
                "title": cur_title,
                "text": p.text,
                "classification": _classify(p.text, classify_requirements),
            }
        )
    return items


def _slide_title(text: str, page_no: int) -> str:
    for line in (text or "").splitlines():
        if line.strip():
            return line.strip()
    return f"Slide {page_no}"


def deck_to_items(pages_tables, slides_text) -> List[dict]:
    """Turn deck output into items.

    Every table row becomes one item (cells joined by " | "), and every slide's
    full text becomes one item, so nothing is lost.
    """
    title_by_page = {}
    text_by_page = {}
    for page_no, text in slides_text:
        title_by_page[page_no] = _slide_title(text, page_no)
        text_by_page[page_no] = (text or "").strip()

    tables_by_page = {}
    for page_no, tables in pages_tables:
        tables_by_page.setdefault(page_no, []).extend(tables)

    pages = sorted(set(title_by_page) | set(tables_by_page))
    items: List[dict] = []
    for page_no in pages:
        title = title_by_page.get(page_no, f"Slide {page_no}")
        for idx, table in enumerate(tables_by_page.get(page_no, [])):
            clause = f"slide{page_no}" if idx == 0 else f"slide{page_no}_{idx + 1}"
            for row in table.rows:
                items.append(
                    {
                        "clause_id": clause,
                        "title": title,
                        "text": " | ".join(c for c in row),
                        "classification": "Information",
                    }
                )
        if page_no in text_by_page:
            items.append(
                {
                    "clause_id": f"slide{page_no}",
                    "title": title,
                    "text": text_by_page[page_no],
                    "classification": "Information",
                }
            )
    return items


# --------------------------------------------------------------------------- #
# Writer
# --------------------------------------------------------------------------- #

def write_standard_assessment(
    items,
    out_path: str,
    template_path: str = DEFAULT_TEMPLATE,
    standard_id: str = "MLSR",
    standard_title: str = "",
    standard_edition: str = "",
    document_id: str = "",
    document_name: str = "",
    document_revision: str = "",
) -> None:
    """Populate the Standard Assessment template and save it to ``out_path``."""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font

    if not os.path.isfile(template_path):
        raise FileNotFoundError(f"template not found: {template_path}")

    wb = load_workbook(template_path)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"template has no '{SHEET_NAME}' sheet")
    ws = wb[SHEET_NAME]

    # Metadata block: overwrite the external-link formulas with plain values
    # (blank when not provided). These cells are the masters of merged ranges.
    ws["B1"] = document_id or None
    ws["B2"] = document_name or None
    ws["B3"] = document_revision or None
    ws["B5"] = standard_id or None
    ws["B6"] = standard_title or None
    ws["B7"] = standard_edition or None

    # Clear existing data-row VALUES (do NOT delete rows: that breaks the
    # validations anchored to E10:E1020 / I10:I1020 / ...).
    last_row = ws.max_row
    for r in range(DATA_START_ROW, last_row + 1):
        for c in range(1, _LAST_COL + 1):
            ws.cell(row=r, column=c).value = None

    font = Font(name="Times New Roman", size=10)
    align = Alignment(vertical="top", wrap_text=True)

    prev_clause = None
    for i, item in enumerate(items, start=1):
        row = DATA_START_ROW + i - 1
        clause = (item.get("clause_id") or "").strip()
        values = {
            1: f"{standard_id}_{i}",
            # B is sparse: written only when the clause changes (like the template).
            2: clause if clause != prev_clause else None,
            3: item.get("title", ""),
            4: item.get("text", ""),
            5: item.get("classification", "Information") or "Information",
            # F..Q intentionally left empty for the analyst (dropdowns remain).
        }
        for c, v in values.items():
            cell = ws.cell(row=row, column=c)
            cell.value = v
            cell.font = font
            cell.alignment = align
        prev_clause = clause

    wb.save(out_path)
