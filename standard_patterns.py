"""Standard profiles and pattern matching for standard/guideline documents.

Re-exports profile configs from ``standard_profiles`` and provides heading/label
parsing helpers used by the layout and structure parsers.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

from standard_profiles import (
    ControlCatalogProfile,
    GenericNumberedStandardProfile,
    GenericProfile,
    IsoLikeProfile,
    LegalArticleProfile,
    Nist80053Profile,
    PciDssProfile,
    CisControlsProfile,
    StandardProfile,
    PROFILES,
    resolve_profile,
    NIST_BASE_RE,
    NIST_ENH_RE,
)

# Backward-compatible aliases
GenericNumberedProfile = GenericNumberedStandardProfile

SECTION_LABEL_RE = re.compile(
    r"^(?P<label>Control|Requirement|Discussion|Guidance|Implementation Guidance|"
    r"Testing Procedure|Related Controls|References|Notes|Objective|Purpose|"
    r"Applicability|Other information|Assessment Procedure)\s*:",
    re.IGNORECASE,
)


def match_section_label(line: str) -> Optional[str]:
    """Return section label if line starts with one (e.g. 'Discussion:')."""
    m = SECTION_LABEL_RE.match(line.strip())
    return m.group("label") if m else None


def parse_heading_with_profile(text: str, profile: StandardProfile) -> Optional[Dict[str, Any]]:
    """Try to parse a heading using profile-specific patterns."""
    text_clean = text.strip()
    for pattern in profile.heading_patterns:
        m = pattern.match(text_clean)
        if m:
            gd = m.groupdict()
            clause_id = (gd.get("id") or "").strip()
            title = (gd.get("title") or gd.get("title2") or "").strip()
            level = len(clause_id.split(".")) if "." in clause_id else 1
            return {
                "kind": "sub_clause" if level > 1 else "base_clause",
                "clause_id": clause_id,
                "title": title,
                "level": level,
                "confidence": 0.95,
                "is_new_item": True,
            }
    for pattern in profile.enhancement_patterns:
        m = pattern.match(text_clean)
        if m and _is_uppercase_title(m.group("title")):
            return {
                "kind": "enhancement",
                "clause_id": m.group("num"),
                "title": m.group("title").strip(),
                "level": 3,
                "confidence": 0.93,
                "is_enhancement_or_subitem": True,
            }
    return None


def parse_heading_generic(text: str) -> Optional[Dict[str, Any]]:
    """Try all profiles to detect a heading."""
    for profile in PROFILES.values():
        res = parse_heading_with_profile(text, profile)
        if res:
            return res
    return None


def _is_uppercase_title(title: str) -> bool:
    """NIST-style enhancements use ALL CAPS titles; reject mixed-case sub-clauses."""
    letters = [c for c in title if c.isalpha()]
    if not letters:
        return False
    upper = sum(1 for c in letters if c.isupper())
    return upper / len(letters) >= 0.85


def detect_profile_from_preflights(page_preflights: List[Dict[str, Any]]) -> str:
    """Auto-detect document profile from preflight metadata (first ~20 pages)."""
    counts: Dict[str, int] = {}
    for p in page_preflights[:20]:
        st = p.get("likely_standard_type") or p.get("profile_candidate") or "unknown"
        if st != "unknown":
            counts[st] = counts.get(st, 0) + 1
    if not counts:
        return "generic"
    best = max(counts, key=counts.get)
    from standard_profiles import PROFILE_ALIASES
    return PROFILE_ALIASES.get(best, best if best in PROFILES else "generic")
