"""Validation module for standard/guideline extraction results."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set

BLOCKING_ISSUES = {
    "front_matter_skipped",
    "toc_row",
    "errata_row",
    "header_footer",
    "doi_footer",
    "doi_footer_url",
    "page_number",
    "page_number_included",
    "one_letter_fragment",
    "orphan_block_not_exported",
    "orphan_block",
    "appendix_summary_row",
    "missing_clause_id_for_control_profile",
    "empty_text",
    "raw_layout_fragment",
    "duplicate_clause_id",
}

WARNING_ISSUES = {
    "text_short",
    "text_long",
    "low_confidence",
    "missing_title",
    "table_low_confidence",
    "profile_confidence_low",
    "duplicate_row",
    "broken_clause_order",
    "toc_dot_leader",
}

DOT_LEADER_RE = re.compile(r"\.{4,}|·{4,}")
DOTTED_RE = re.compile(r"^(\d+(?:\.\d+)+)$")

HEADER_RE = re.compile(
    r"NIST\s+SP\s+800-53.*?SECURITY\s+AND\s+PRIVACY\s+CONTROLS.*?ORGANIZATIONS",
    re.I,
)
SPACED_HEADER_RE = re.compile(
    r"N\s*I\s*S\s*T\s+S\s*P\s+8\s*0\s*0\s*-\s*5\s*3\s*,\s*R\s*E\s*V\s*\.?\s*5\s+"
    r"S\s*E\s*C\s*U\s*R\s*I\s*T\s*Y\s+A\s*N\s*D\s+P\s*R\s*I\s*V\s*A\s*C\s*Y\s+"
    r"C\s*O\s*N\s*T\s*R\s*O\s*L\s*S\s+F\s*O\s*R\s+I\s*N\s*F\s*O\s*R\s*M\s*A\s*T\s*I\s*O\s*N\s+"
    r"S\s*Y\s*S\s*T\s*E\s*M\s*S\s+A\s*N\s*D\s+O\s*R\s*G\s*A\s*N\s*I\s*Z\s*A\s*T\s*I\s*O\s*N\s*S",
    re.I,
)
CHAPTER_PAGE_RE = re.compile(
    r"(CHAPTER|APPENDIX|C\s*H\s*A\s*P\s*T\s*E\s*R|A\s*P\s*P\s*E\s*N\s*D\s*I\s*X)\s+"
    r"([A-Z0-9\s]+)?\s*PAGE\s+\d+",
    re.I,
)
DOI_RE = re.compile(
    r"(This publication is available free of charge|doi\.org/10\.6028/NIST\.SP\.800-53r5)",
    re.I,
)


def looks_like_page_number(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    return bool(
        re.fullmatch(
            r"(page\s+)?[-–\s]*\d+([\s/]+(of\s+)?\d+)?[-–\s]*",
            stripped,
            flags=re.IGNORECASE,
        )
    )


def is_doi_url(text: str) -> bool:
    stripped = text.strip().lower()
    return "doi.org" in stripped or (stripped.startswith("http") and len(stripped) < 100)


def validate_items(
    items: List[Dict[str, Any]],
    boilerplate_set: Optional[Set[str]] = None,
    profile_name: str = "generic",
) -> List[Dict[str, Any]]:
    """Validate extracted items; tag blocking/warning issues and adjust confidence."""
    if not items:
        return []

    boilerplate = boilerplate_set or set()
    is_catalog = profile_name in ("nist80053", "control-catalog")

    text_counts: Dict[str, int] = {}
    for item in items:
        txt = (item.get("text") or "").strip()
        if txt:
            text_counts[txt] = text_counts.get(txt, 0) + 1

    clause_counts: Dict[str, int] = {}
    for item in items:
        cid = (item.get("clause_id") or "").strip()
        if cid:
            clause_counts[cid] = clause_counts.get(cid, 0) + 1

    last_dotted: List[List[int]] = []

    for item in items:
        issues = list(item.get("issues", []))
        text = (item.get("text") or "").strip()
        clause_id = (item.get("clause_id") or "").strip()
        title = (item.get("title") or "").strip()
        classification = item.get("classification", "Information")
        confidence = item.get("confidence", 1.0)

        if "orphan_block_not_exported" in issues:
            issues.append("orphan_block")

        if is_catalog and not clause_id and item.get("export_status") != "rejected":
            issues.append("missing_clause_id_for_control_profile")
            confidence = min(confidence, 0.1)

        if text:
            if HEADER_RE.search(text) or SPACED_HEADER_RE.search(text) or CHAPTER_PAGE_RE.search(text):
                issues.append("header_footer")
                confidence = min(confidence, 0.1)
            norm = re.sub(r"\s+", "", text)
            if norm in boilerplate or looks_like_page_number(text):
                issues.append("header_footer")
                issues.append("page_number")
                issues.append("page_number_included")
                confidence = min(confidence, 0.2)

        if "toc_row" in issues or DOT_LEADER_RE.search(text) or "TABLE OF CONTENTS" in text.upper():
            issues.append("toc_row")
            if DOT_LEADER_RE.search(text):
                issues.append("toc_dot_leader")
            confidence = min(confidence, 0.2)

        if "errata" in text.lower() or "errata" in title.lower():
            issues.append("errata_row")
            confidence = min(confidence, 0.3)

        if "TABLE C-" in text.upper() or "APPENDIX C PAGE" in text.upper():
            issues.append("appendix_summary_row")
            confidence = min(confidence, 0.3)

        if title == "NUMBER" or text == "NUMBER":
            issues.append("appendix_summary_row")

        if text and (is_doi_url(text) or DOI_RE.search(text)):
            issues.append("doi_footer")
            confidence = min(confidence, 0.1)

        if text and len(text) == 1:
            issues.append("one_letter_fragment")
            confidence = min(confidence, 0.1)

        is_withdrawn = "[withdrawn" in text.lower()
        if not text:
            issues.append("empty_text")
            confidence = min(confidence, 0.1)
        elif len(text) < 10 and not is_withdrawn:
            issues.append("text_short")
            confidence = min(confidence, 0.5)

        if text and len(text) > 2500:
            issues.append("text_long")

        if not title and clause_id and item.get("export_status") == "exported":
            issues.append("missing_title")

        if clause_id and clause_counts.get(clause_id, 0) > 1:
            issues.append("duplicate_clause_id")
            confidence = min(confidence, 0.3)

        if text and text_counts.get(text, 0) > 1:
            issues.append("duplicate_row")
            confidence = min(confidence, 0.6)

        m = DOTTED_RE.match(clause_id)
        if m:
            try:
                parts = [int(x) for x in m.group(1).split(".")]
                if last_dotted:
                    prev = last_dotted[-1]
                    if len(parts) == len(prev) and parts[:-1] == prev[:-1] and parts[-1] < prev[-1]:
                        issues.append("broken_clause_order")
                last_dotted.append(parts)
            except ValueError:
                pass

        if confidence < 0.5:
            issues.append("low_confidence")

        item["issues"] = sorted(set(issues))
        item["confidence"] = confidence

    return items


def filter_exportable_items(
    items: List[Dict[str, Any]],
    profile_name: str = "generic",
    min_confidence: float = 0.0,
    export_low_confidence: bool = False,
) -> List[Dict[str, Any]]:
    """Return only items safe to write into Standard Assessment."""
    is_catalog = profile_name in ("nist80053", "control-catalog")
    exported = []
    for it in items:
        if it.get("export_status") == "rejected":
            continue
        if not export_low_confidence and min_confidence > 0 and it.get("confidence", 1) < min_confidence:
            continue
        if any(iss in BLOCKING_ISSUES for iss in it.get("issues", [])):
            continue
        if is_catalog and not (it.get("clause_id") or "").strip():
            continue
        if not (it.get("text") or "").strip():
            continue
        exported.append(it)
    return exported
