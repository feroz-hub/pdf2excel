"""Table normalization and extraction for standard/guideline documents.

Normalizes extracted tables (merges wrapped rows, joins cells with pipes, detects headers,
and classifies rows as requirements or informative).
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional
from standard_patterns import StandardProfile

# Common TOC patterns
TOC_CELL_RE = re.compile(r"\.{4,}|·{4,}")

def is_toc_table(rows: List[List[str]]) -> bool:
    """Detect if the table is actually a Table of Contents."""
    if not rows or len(rows) < 2:
        return False
    toc_hits = 0
    total_cells = 0
    for row in rows:
        for cell in row:
            if not cell:
                continue
            total_cells += 1
            val = str(cell).strip()
            if TOC_CELL_RE.search(val) or val.lower() in ("table of contents", "contents", "page"):
                toc_hits += 1
    if total_cells > 0 and (toc_hits / total_cells) > 0.15:
        return True
    return False


def is_decorative_table(rows: List[List[str]]) -> bool:
    """Identify decorative or mostly empty tables."""
    if not rows:
        return True
    non_empty_rows = 0
    for row in rows:
        non_empty_cells = [c for c in row if str(c or "").strip()]
        if non_empty_cells:
            non_empty_rows += 1
    # If table has no rows or only 1 non-empty row, it's decorative/empty
    if non_empty_rows <= 1:
        return True
    return False


def normalize_table(
    rows: List[List[str]],
    profile: Optional[StandardProfile] = None
) -> List[Dict[str, Any]]:
    """Normalize raw table rows: merge wrapped rows, detect headers, and map to items."""
    if not rows or is_decorative_table(rows) or is_toc_table(rows):
        return []

    # Clean whitespace and None values
    cleaned_rows = []
    for row in rows:
        cleaned_rows.append([str(c or "").strip() for c in row])

    # Find the header row
    header_idx = 0
    # Search first 3 rows for table headers matching profile or common standard keywords
    for idx in range(min(3, len(cleaned_rows))):
        row = cleaned_rows[idx]
        has_header_keyword = False
        if profile and profile.table_header_patterns:
            for pat in profile.table_header_patterns:
                if any(pat.search(c) for c in row):
                    has_header_keyword = True
                    break
        if not has_header_keyword:
            # Common generic header keywords
            common_headers = ["requirement", "control", "procedure", "guidance", "description", "status", "asset", "function"]
            if any(any(h in c.lower() for h in common_headers) for c in row):
                has_header_keyword = True
        if has_header_keyword:
            header_idx = idx
            break

    headers = cleaned_rows[header_idx]
    data_rows = cleaned_rows[header_idx + 1:]

    # Merge wrapped rows:
    # If a row's key columns (like first column or ID column) are empty, but description columns have text,
    # it is likely a continuation of the previous row.
    merged_rows = []
    for row in data_rows:
        if not any(row):  # skip fully empty rows
            continue
        
        # If first column is empty and we have a previous row, merge it
        if len(row) > 1 and not row[0] and merged_rows:
            prev = merged_rows[-1]
            for col_idx in range(len(row)):
                if row[col_idx]:
                    if prev[col_idx]:
                        prev[col_idx] += "\n" + row[col_idx]
                    else:
                        prev[col_idx] = row[col_idx]
        else:
            merged_rows.append(row)

    # Convert to items
    items = []
    for r in merged_rows:
        # Join cells with " | "
        joined_text = " | ".join(c.replace("\n", " ") for c in r)
        
        # Heuristically detect clause/ID and title from columns
        # First column is often the ID
        clause_id = r[0] if len(r) > 0 else ""
        title = r[1] if len(r) > 1 else ""
        
        # If clause_id is suspiciously long, it's probably not a real ID
        if len(clause_id) > 30:
            clause_id = ""
            title = ""

        # Classification check: if it contains requirement keywords
        classification = "Information"
        req_keywords = ["shall", "must", "is required to", "are required to", "require", "ensure"]
        if profile:
            req_keywords = profile.requirement_keywords
        if any(kw in joined_text.lower() for kw in req_keywords):
            classification = "Requirement"

        items.append({
            "clause_id": clause_id,
            "title": title,
            "text": joined_text,
            "classification": classification,
            "confidence": 0.85,
            "issues": []
        })

    return items
