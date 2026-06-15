"""Validation module for standard/guideline extraction results.

Checks extracted items for quality issues (empty rows, short text, repeated headers/footers,
TOC rows, missing clause IDs/titles, duplicates, broken clause order, etc.), assigning confidence
scores and lists of warnings.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Set, Optional

# Regex for TOC dot leaders
DOT_LEADER_RE = re.compile(r"\.{4,}|·{4,}")

# Regex for dotted clause numbers like 1.2 or 4.3.1
DOTTED_RE = re.compile(r"^(\d+(?:\.\d+)+)$")

def looks_like_page_number(text: str) -> bool:
    """Check if the text is primarily a page number."""
    stripped = text.strip()
    if not stripped:
        return False
    # Matches "12", "Page 12", "12 / 30", "- 12 -", "Page 12 of 30"
    return bool(
        re.fullmatch(
            r"(page\s+)?[-–\s]*\d+([\s/]+(of\s+)?\d+)?[-–\s]*",
            stripped,
            flags=re.IGNORECASE,
        )
    )

def is_doi_url(text: str) -> bool:
    """Check if the text is primarily a DOI or footer URL."""
    stripped = text.strip().lower()
    return "doi.org" in stripped or (stripped.startswith("http") and len(stripped) < 100)


def validate_items(
    items: List[Dict[str, Any]],
    boilerplate_set: Optional[Set[str]] = None
) -> List[Dict[str, Any]]:
    """Validate a list of mapped assessment items, adding confidence and issues."""
    if not items:
        return []

    boilerplate = boilerplate_set or set()

    # Track duplicate texts across the document to flag suspicious repeats
    text_counts: Dict[str, int] = {}
    for item in items:
        txt = (item.get("text") or "").strip()
        if txt:
            text_counts[txt] = text_counts.get(txt, 0) + 1

    # Track duplicate clause IDs
    clause_counts: Dict[str, int] = {}
    for item in items:
        cid = (item.get("clause_id") or "").strip()
        if cid:
            clause_counts[cid] = clause_counts.get(cid, 0) + 1

    # Track clause sequence to detect broken ordering
    last_dotted: List[List[int]] = []

    for item in items:
        issues = list(item.get("issues", []))
        text = (item.get("text") or "").strip()
        clause_id = (item.get("clause_id") or "").strip()
        title = (item.get("title") or "").strip()
        classification = item.get("classification", "Information")
        confidence = item.get("confidence", 1.0)

        # 1. Empty text
        if not text:
            issues.append("empty_row")
            confidence = min(confidence, 0.0)

        # 2. Text too short
        if text and len(text) < 10:
            issues.append("text_too_short")
            confidence = min(confidence, 0.4)

        # 3. One-letter row
        if text and len(text) == 1:
            issues.append("one_letter_row")
            confidence = min(confidence, 0.1)

        # 4. Page number included
        if text and looks_like_page_number(text):
            issues.append("page_number_included")
            confidence = min(confidence, 0.2)

        # 5. Repeated header/footer
        if text:
            from pdf_paragraphs_to_excel import _normalize_digits
            norm = _normalize_digits(text).strip()
            if norm in boilerplate:
                issues.append("possible_header_footer")
                confidence = min(confidence, 0.3)

        # 6. DOI/footer URL
        if text and is_doi_url(text):
            issues.append("doi_footer_url")
            confidence = min(confidence, 0.2)

        # 7. TOC dotted leader row
        if text and DOT_LEADER_RE.search(text):
            issues.append("toc_dot_leader")
            confidence = min(confidence, 0.3)

        # 8. Suspicious duplicate rows
        if text and text_counts.get(text, 0) > 1 and len(text) > 30:
            issues.append("duplicate_row")
            confidence = min(confidence, 0.7)

        # 9. Suspiciously long text
        if text and len(text) > 2500:
            issues.append("suspiciously_long_text")

        # 10. Suspiciously short text for Requirement
        if text and len(text) < 25 and classification == "Requirement":
            issues.append("suspiciously_short_text")
            confidence = min(confidence, 0.6)

        # 11. Missing clause ID / title for Requirements
        if classification == "Requirement":
            if not clause_id:
                issues.append("missing_clause_id")
                confidence = min(confidence, 0.8)
            if not title:
                issues.append("missing_title")
                confidence = min(confidence, 0.8)

        # 12. Duplicate clause ID
        if clause_id and clause_counts.get(clause_id, 0) > 1 and classification == "Requirement":
            issues.append("duplicate_clause_id")
            confidence = min(confidence, 0.8)

        # 13. Broken clause order (dotted numbers decreasing)
        m = DOTTED_RE.match(clause_id)
        if m:
            try:
                parts = [int(x) for x in m.group(1).split(".")]
                if last_dotted:
                    prev_parts = last_dotted[-1]
                    if len(parts) == len(prev_parts):
                        # E.g. going from 1.2 to 1.1 on the same level
                        if parts[:-1] == prev_parts[:-1] and parts[-1] < prev_parts[-1]:
                            issues.append("broken_clause_order")
                            confidence = min(confidence, 0.8)
                last_dotted.append(parts)
            except ValueError:
                pass

        # 14. Scanned page skipped marker
        if "[scanned page" in text.lower() or "scanned_page_skipped" in issues:
            if "scanned_page_skipped" not in issues:
                issues.append("scanned_page_skipped")
            confidence = min(confidence, 0.1)

        # 15. Rotated text included
        if item.get("is_rotated"):
            issues.append("rotated_text_included")
            confidence = min(confidence, 0.7)

        # 16. Table extraction failure
        if text and "|" in text:
            cells = [c.strip() for c in text.split("|")]
            non_empty_cells = [c for c in cells if c]
            if len(non_empty_cells) <= 1:
                issues.append("table_extraction_low_confidence")
                confidence = min(confidence, 0.6)

        # General low confidence label
        if confidence < 0.5 and "text_confidence_low" not in issues:
            issues.append("text_confidence_low")

        item["issues"] = sorted(list(set(issues)))
        item["confidence"] = confidence

    return items
