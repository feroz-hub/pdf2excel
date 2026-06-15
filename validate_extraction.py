"""Validation module for pdf2excel extraction results.

Checks extracted items for various potential quality issues: empty/short rows,
boilerplate headers/footers, TOC page remains, page numbers, duplicate rows,
out-of-order clauses, missing clause IDs/titles, and low-confidence elements.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Set, Optional
from pdf_paragraphs_to_excel import _looks_like_page_number, _normalize_digits

# Regex for TOC dot leaders
DOT_LEADER_RE = re.compile(r"\.{4,}|·{4,}")

# Regex for dotted clause numbers like 1.2 or 4.3.1
DOTTED_RE = re.compile(r"^(\d+(?:\.\d+)+)$")


def validate_items(items: List[Dict[str, Any]], boilerplate_set: Optional[Set[str]] = None) -> List[Dict[str, Any]]:
    """Validate a list of mapped assessment items, adding confidence and issues."""
    if not items:
        return []
        
    boilerplate = boilerplate_set or set()
    
    # 1. Track duplicate texts across the document to flag suspicious repeats
    text_counts: Dict[str, int] = {}
    for item in items:
        txt = (item.get("text") or "").strip()
        if txt:
            text_counts[txt] = text_counts.get(txt, 0) + 1
            
    # 2. Track clause sequence to detect broken ordering
    last_dotted: List[List[int]] = []
    
    for idx, item in enumerate(items):
        issues = list(item.get("issues", []))
        text = (item.get("text") or "").strip()
        clause_id = (item.get("clause_id") or "").strip()
        title = (item.get("title") or "").strip()
        classification = item.get("classification", "Information")
        confidence = item.get("confidence", 1.0)
        
        # Issue 1: Empty text
        if not text:
            issues.append("empty_row")
            confidence = min(confidence, 0.0)
            
        # Issue 2: Very short text rows
        if text and len(text) < 15 and classification == "Requirement":
            issues.append("short_text")
            confidence = min(confidence, 0.5)
            
        # Issue 3: Page numbers accidentally included
        if text and _looks_like_page_number(text):
            issues.append("page_number_included")
            confidence = min(confidence, 0.2)
            
        # Issue 4: Repeated headers/footers still present
        if text:
            norm = _normalize_digits(text).strip()
            if norm in boilerplate:
                issues.append("possible_header_footer")
                confidence = min(confidence, 0.3)
                
        # Issue 5: TOC dotted leader rows
        if text and DOT_LEADER_RE.search(text):
            issues.append("toc_dot_leader")
            confidence = min(confidence, 0.3)
            
        # Issue 6: Suspicious duplicate rows
        if text and text_counts.get(text, 0) > 1 and len(text) > 30:
            issues.append("duplicate_row")
            confidence = min(confidence, 0.7)
            
        # Issue 7: Very long paragraphs
        if text and len(text) > 1500:
            issues.append("long_paragraph")
            # doesn't reduce confidence directly, just flags warning
            
        # Issue 8: Missing clause id / title for Requirements
        if classification == "Requirement":
            if not clause_id:
                issues.append("missing_clause_id")
                confidence = min(confidence, 0.8)
            if not title:
                issues.append("missing_title")
                confidence = min(confidence, 0.8)
                
        # Issue 9: Table extraction failure (e.g. empty cell rows or single cell rows joined)
        if text and "|" in text:
            cells = [c.strip() for c in text.split("|")]
            non_empty_cells = [c for c in cells if c]
            if len(non_empty_cells) <= 1:
                issues.append("table_row_extraction_failure")
                confidence = min(confidence, 0.6)
                
        # Issue 10: Scanned page skipped marker
        if "[scanned page" in text.lower() or "scanned page" in issues:
            issues.append("scanned_page_skipped")
            confidence = min(confidence, 0.1)
            
        # Issue 11: Broken clause order (dotted numbers decreasing)
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
                
        # Issue 12: General low confidence
        if confidence < 0.5 and "text_confidence_low" not in issues:
            issues.append("text_confidence_low")
            
        item["issues"] = sorted(list(set(issues)))
        item["confidence"] = confidence
        
    return items
