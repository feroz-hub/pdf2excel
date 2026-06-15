"""Standard profiles and pattern matching for standard/guideline documents.

Provides regex-based heading detection, label detection, and document-level
profiles for various standard types (NIST 800-53, ISO, PCI DSS, CIS, GDPR, etc.).
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional

class StandardProfile:
    """Configuration profile for a specific standard or guideline type."""
    def __init__(
        self,
        name: str,
        heading_patterns: List[re.Pattern],
        section_markers: List[str],
        skip_page_patterns: List[re.Pattern],
        start_markers: List[re.Pattern],
        stop_markers: List[re.Pattern],
        requirement_keywords: List[str],
        table_header_patterns: List[re.Pattern],
    ):
        self.name = name
        self.heading_patterns = heading_patterns
        self.section_markers = section_markers
        self.skip_page_patterns = skip_page_patterns
        self.start_markers = start_markers
        self.stop_markers = stop_markers
        self.requirement_keywords = requirement_keywords
        self.table_header_patterns = table_header_patterns


# ---------------------------------------------------------------------------
# Pre-defined Profiles
# ---------------------------------------------------------------------------

Nist80053Profile = StandardProfile(
    name="nist80053",
    heading_patterns=[
        # AC-1 POLICY AND PROCEDURES
        re.compile(r"^(?P<id>[A-Z]{2}-\d+)\s+(?P<title>[A-Z0-9][A-Z0-9 ,/\-\u2014:&()]+)$"),
        # (1) AUTOMATED SYSTEM ACCOUNT MANAGEMENT
        re.compile(r"^\((?P<num>\d+)\)\s+(?P<title>[A-Z0-9][A-Z0-9 ,/\-\u2014:&()|]+)$"),
    ],
    section_markers=[
        "Control:", "Discussion:", "Related Controls:", "References:", "Control Enhancements:"
    ],
    skip_page_patterns=[
        re.compile(r"TABLE OF CONTENTS", re.I),
        re.compile(r"ERRATA", re.I),
    ],
    start_markers=[
        re.compile(r"CHAPTER THREE", re.I),
    ],
    stop_markers=[
        re.compile(r"^REFERENCES\b", re.I),
        re.compile(r"^APPENDIX A\b", re.I),
    ],
    requirement_keywords=[
        "shall", "must", "is required to", "are required to", "require", "required"
    ],
    table_header_patterns=[
        re.compile(r"Control ID", re.I),
        re.compile(r"Control Name", re.I),
    ]
)

IsoLikeProfile = StandardProfile(
    name="iso",
    heading_patterns=[
        # 7.5.1 Documented Information, A.8.1.1 Inventory of assets
        re.compile(r"^(?P<id>[A-Z]?\.\d+(?:\.\d+)*|\d+(?:\.\d+)+)\s+(?P<title>[A-Z0-9][A-Za-z0-9 ,/\-\u2014:&()]+)$"),
        # Top-level: 5 Leadership, 10 Improvement
        re.compile(r"^(?P<id>\d+)\s+(?P<title>[A-Z][A-Za-z0-9 ,/\-\u2014:&()]+)$"),
    ],
    section_markers=[
        "Control", "Requirement", "Implementation guidance", "Other information", "Required"
    ],
    skip_page_patterns=[
        re.compile(r"Foreword", re.I),
        re.compile(r"Introduction", re.I),
        re.compile(r"Table of Contents", re.I),
    ],
    start_markers=[
        re.compile(r"^1\s+Scope", re.I),
        re.compile(r"^4\s+Context of the organization", re.I),
    ],
    stop_markers=[
        re.compile(r"^Bibliography", re.I),
    ],
    requirement_keywords=[
        "shall", "must", "is required to", "ensure", "establish", "implement", "maintain"
    ],
    table_header_patterns=[
        re.compile(r"Control", re.I),
        re.compile(r"Control objective", re.I),
    ]
)

LegalArticleProfile = StandardProfile(
    name="legal",
    heading_patterns=[
        # Article 1, Chapter I
        re.compile(r"^(?P<id>(?:Article|Section)\s+\d+)\s*(?::|-)?\s*(?:\((?P<title>[^)]+)\)|(?P<title2>.+))?$", re.I),
        re.compile(r"^(?P<id>(?:Chapter|Part|Schedule)\s+(?:[IVXLCDM\d]+))\s*(?::|-)?\s*(?P<title>.*)$", re.I),
    ],
    section_markers=[
        "Article", "Section", "Definition", "1.", "2.", "3.", "Paragraph", "Note:", "Guidance:"
    ],
    skip_page_patterns=[
        re.compile(r"Table of contents", re.I),
    ],
    start_markers=[
        re.compile(r"CHAPTER I\b", re.I),
        re.compile(r"Article 1\b", re.I),
    ],
    stop_markers=[
        re.compile(r"ANNEX\s+[A-Z]", re.I),
    ],
    requirement_keywords=[
        "shall", "shall not", "must", "is prohibited", "are prohibited", "requires", "required"
    ],
    table_header_patterns=[]
)

PciDssProfile = StandardProfile(
    name="pci",
    heading_patterns=[
        # PCI Requirement 3.2.1
        re.compile(r"^(?P<id>(?:Req\.?|Requirement)\s+\d+(?:\.\d+)*)\s+(?P<title>[A-Z][A-Za-z0-9 ,/\-\u2014:&()]+)$", re.I),
    ],
    section_markers=[
        "PCI DSS Requirement", "Testing Procedure", "Guidance", "Applicability"
    ],
    skip_page_patterns=[],
    start_markers=[
        re.compile(r"PCI DSS Requirements", re.I),
    ],
    stop_markers=[
        re.compile(r"Appendix\s+[A-Z]", re.I),
    ],
    requirement_keywords=[
        "must", "shall", "is required", "are required", "examine", "verify", "confirm"
    ],
    table_header_patterns=[
        re.compile(r"PCI DSS Requirement", re.I),
        re.compile(r"Testing Procedure", re.I),
        re.compile(r"Guidance", re.I),
    ]
)

CisControlsProfile = StandardProfile(
    name="cis",
    heading_patterns=[
        # CIS Control 1, CIS Safeguard 1.1
        re.compile(r"^(?P<id>CIS\s+(?:Control|Safeguard)\s+\d+(?:\.\d+)*)\s+(?P<title>[A-Z][A-Za-z0-9 ,/\-\u2014:&()]+)$", re.I),
        # Control 5.1, Safeguard 5.1
        re.compile(r"^(?P<id>(?:Control|Safeguard)\s+\d+(?:\.\d+)*)\s+(?P<title>[A-Z][A-Za-z0-9 ,/\-\u2014:&()]+)$", re.I),
    ],
    section_markers=[
        "Description", "Asset Type", "Security Function", "Title"
    ],
    skip_page_patterns=[],
    start_markers=[
        re.compile(r"CIS Controls", re.I),
    ],
    stop_markers=[
        re.compile(r"Appendix\s+[A-Z]", re.I),
    ],
    requirement_keywords=[
        "ensure", "establish", "maintain", "restrict", "protect", "encrypt", "authenticate", "require", "shall"
    ],
    table_header_patterns=[
        re.compile(r"Asset Type", re.I),
        re.compile(r"Security Function", re.I),
    ]
)

ControlCatalogProfile = StandardProfile(
    name="control-catalog",
    heading_patterns=[
        # Generic Control 5.1
        re.compile(r"^(?P<id>(?:Control|Requirement)\s+\d+(?:\.\d+)*)\s*(?::|-)?\s*(?P<title>[A-Z0-9][A-Za-z0-9 ,/\-\u2014:&()]+)$", re.I),
        # Dotted control headers: AC-1 POLICY AND PROCEDURES
        re.compile(r"^(?P<id>[A-Z]{2,4}-\d+(?:\(\d+\))?)\s*(?::|-)?\s*(?P<title>[A-Z0-9][A-Za-z0-9 ,/\-\u2014:&()]+)$"),
    ],
    section_markers=[
        "Control:", "Guidance:", "Discussion:", "References:", "Notes:"
    ],
    skip_page_patterns=[],
    start_markers=[],
    stop_markers=[],
    requirement_keywords=[
        "shall", "must", "required", "ensure", "implement", "restrict"
    ],
    table_header_patterns=[]
)

GenericNumberedProfile = StandardProfile(
    name="generic",
    heading_patterns=[
        # 1.2.3 Title
        re.compile(r"^(?P<id>\d+(?:\.\d+)*)\s+(?P<title>[A-Z0-9][A-Za-z0-9 ,/\-\u2014:&()]+)$"),
        # Annex A, Appendix B
        re.compile(r"^(?P<id>(?:Annex|Appendix|Schedule)\s+[A-Z0-9](?:\.\d+)*)\s+(?P<title>[A-Z0-9][A-Za-z0-9 ,/\-\u2014:&()]+)$", re.I),
        # Article 1, Chapter I, Section 1
        re.compile(r"^(?P<id>(?:Article|Chapter|Section|Part)\s+\d+)\s*(?::|-)?\s*(?:\((?P<title>[^)]+)\)|(?P<title2>.+))?$", re.I),
    ],
    section_markers=[
        "Description", "Requirements", "Note:", "Guidance:", "Article", "Section", "Definition"
    ],
    skip_page_patterns=[],
    start_markers=[],
    stop_markers=[],
    requirement_keywords=[
        "shall", "must", "is required to", "are required to", "ensure", "implement", "maintain"
    ],
    table_header_patterns=[]
)

PROFILES = {
    "nist80053": Nist80053Profile,
    "iso": IsoLikeProfile,
    "legal": LegalArticleProfile,
    "pci": PciDssProfile,
    "cis": CisControlsProfile,
    "control-catalog": ControlCatalogProfile,
    "generic": GenericNumberedProfile
}


# ---------------------------------------------------------------------------
# Label / Section detection
# ---------------------------------------------------------------------------

SECTION_LABEL_RE = re.compile(
    r"^(?P<label>Control|Requirement|Discussion|Guidance|Implementation Guidance|Testing Procedure|Related Controls|References|Notes|Objective|Purpose|Applicability|Other information)\s*:",
    re.IGNORECASE
)

def match_section_label(line: str) -> Optional[str]:
    """Check if line starts with a standard section label (e.g. 'Discussion:')."""
    m = SECTION_LABEL_RE.match(line.strip())
    if m:
        return m.group("label")
    return None


# ---------------------------------------------------------------------------
# Heading matching helper
# ---------------------------------------------------------------------------

def parse_heading_with_profile(text: str, profile: StandardProfile) -> Optional[Dict[str, Any]]:
    """Try to parse heading with a specific profile's patterns."""
    text_clean = text.strip()
    for pattern in profile.heading_patterns:
        m = pattern.match(text_clean)
        if m:
            gd = m.groupdict()
            clause_id = gd.get("id") or ""
            title = gd.get("title") or gd.get("title2") or ""
            level = len(clause_id.split(".")) if "." in clause_id else 1
            if "num" in gd:
                clause_id = gd.get("num") or ""
                level = 3
            return {
                "kind": "sub_clause" if level > 1 else "base_clause",
                "clause_id": clause_id.strip(),
                "title": title.strip(),
                "level": level,
                "confidence": 0.95
            }
    return None

def parse_heading_generic(text: str) -> Optional[Dict[str, Any]]:
    """Try to parse heading using all profiles to find a match."""
    for profile in PROFILES.values():
        res = parse_heading_with_profile(text, profile)
        if res:
            return res
    return None


# Class Aliases
GenericNumberedStandardProfile = GenericNumberedProfile
IsoLikeProfile = IsoLikeProfile


def detect_profile_from_preflights(page_preflights: List[Dict[str, Any]]) -> str:
    """Analyze page preflight metadata to auto-detect the document's standard profile."""
    counts: Dict[str, int] = {}
    # Sample first 20 pages
    for p in page_preflights[:20]:
        st = p.get("likely_standard_type", "unknown")
        if st != "unknown":
            counts[st] = counts.get(st, 0) + 1
            
    if not counts:
        return "generic"
        
    best = max(counts, key=counts.get)
    
    mapping = {
        "nist_control_catalog": "nist80053",
        "iso_numbered_standard": "iso",
        "legal_articles": "legal",
        "pci_requirement_table": "pci",
        "cis_controls": "cis",
        "generic_numbered_guideline": "generic"
    }
    return mapping.get(best, "generic")
