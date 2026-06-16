"""Standard profile definitions for structured PDF extraction.

Each profile configures heading patterns, start/stop boundaries, and skip rules
for a class of standards/guidelines documents.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Pattern


@dataclass
class StandardProfile:
    """Configuration for parsing a standard or guideline document."""

    name: str
    heading_patterns: List[Pattern] = field(default_factory=list)
    enhancement_patterns: List[Pattern] = field(default_factory=list)
    section_markers: List[str] = field(default_factory=list)
    table_header_patterns: List[Pattern] = field(default_factory=list)
    start_markers: List[Pattern] = field(default_factory=list)
    stop_markers: List[Pattern] = field(default_factory=list)
    skip_page_patterns: List[Pattern] = field(default_factory=list)
    requirement_keywords: List[str] = field(default_factory=list)
    information_keywords: List[str] = field(default_factory=list)
    default_skip_front_matter: bool = True
    default_skip_toc: bool = True
    default_skip_references: bool = True
    default_skip_appendix: bool = True
    requires_start_boundary: bool = False

    # Back-compat aliases used elsewhere
    @property
    def section_markers_legacy(self) -> List[str]:
        return self.section_markers


# NIST / control catalog — case-sensitive IDs and ALL CAPS enhancement titles
NIST_BASE_RE = re.compile(
    r"^(?P<id>[A-Z]{2}-\d+)\s+(?P<title>[A-Z0-9][A-Z0-9 ,/\-\u2014:&()]+)$"
)
NIST_ENH_RE = re.compile(
    r"^\((?P<num>\d+)\)\s+(?P<title>[A-Z0-9][A-Z0-9 ,/\-\u2014:&()|]+)$"
)

Nist80053Profile = StandardProfile(
    name="nist80053",
    heading_patterns=[NIST_BASE_RE],
    enhancement_patterns=[NIST_ENH_RE],
    section_markers=[
        "Control:", "Discussion:", "Related Controls:", "References:",
        "Control Enhancements:",
    ],
    start_markers=[re.compile(r"^AC-1\b")],
    stop_markers=[
        re.compile(r"^REFERENCES\b", re.I),
        re.compile(r"^APPENDIX A\b", re.I),
        re.compile(r"^APPENDIX B\b", re.I),
        re.compile(r"^APPENDIX C\b", re.I),
    ],
    skip_page_patterns=[
        re.compile(r"TABLE OF CONTENTS", re.I),
        re.compile(r"ERRATA", re.I),
    ],
    requirement_keywords=["shall", "must", "is required to", "are required to"],
    default_skip_appendix=True,
    requires_start_boundary=True,
)

ControlCatalogProfile = StandardProfile(
    name="control-catalog",
    heading_patterns=[
        NIST_BASE_RE,
        re.compile(
            r"^(?P<id>(?:Control|Requirement)\s+\d+(?:\.\d+)*)\s*(?::|-)?\s*"
            r"(?P<title>[A-Z0-9][A-Za-z0-9 ,/\-\u2014:&()]+)$",
            re.I,
        ),
    ],
    enhancement_patterns=[NIST_ENH_RE],
    section_markers=["Control:", "Guidance:", "Discussion:", "References:"],
    requirement_keywords=["shall", "must", "required", "ensure", "implement"],
    default_skip_appendix=True,
    requires_start_boundary=True,
)

IsoLikeProfile = StandardProfile(
    name="iso",
    heading_patterns=[
        re.compile(r"^(?P<id>[A-Z]?\.\d+(?:\.\d+)*|\d+(?:\.\d+)+)\s+(?P<title>.+)$"),
        re.compile(r"^(?P<id>\d+)\s+(?P<title>[A-Z][A-Za-z0-9 ,/\-\u2014:&()]+)$"),
    ],
    section_markers=["Control", "Requirement", "Implementation guidance", "Other information"],
    start_markers=[
        re.compile(r"^1\s+Scope", re.I),
        re.compile(r"^4\s+Context", re.I),
    ],
    stop_markers=[re.compile(r"^Bibliography", re.I)],
    requirement_keywords=["shall", "must", "ensure", "establish", "maintain", "document"],
    default_skip_appendix=True,
)

LegalArticleProfile = StandardProfile(
    name="legal",
    heading_patterns=[
        re.compile(
            r"^(?P<id>(?:Article|Section)\s+\d+)\s*(?::|-)?\s*"
            r"(?:\((?P<title>[^)]+)\)|(?P<title2>.+))?$",
            re.I,
        ),
        re.compile(
            r"^(?P<id>(?:Chapter|Part|Schedule)\s+(?:[IVXLCDM\d]+))\s*(?::|-)?\s*(?P<title>.*)$",
            re.I,
        ),
    ],
    section_markers=["Article", "Section", "Paragraph", "Note:", "Guidance:"],
    start_markers=[re.compile(r"^Article 1\b", re.I), re.compile(r"^CHAPTER I\b", re.I)],
    stop_markers=[re.compile(r"^ANNEX\s+[A-Z]", re.I)],
    requirement_keywords=[
        "shall", "shall not", "must", "may not", "is required to",
        "prohibited", "permitted only", "responsible for",
    ],
)

PciDssProfile = StandardProfile(
    name="pci",
    heading_patterns=[
        re.compile(
            r"^(?P<id>(?:Req\.?|Requirement)\s+\d+(?:\.\d+)*)\s+(?P<title>.+)$",
            re.I,
        ),
    ],
    section_markers=["Testing Procedure", "Guidance", "Applicability", "Defined Approach"],
    table_header_patterns=[
        re.compile(r"PCI DSS Requirement", re.I),
        re.compile(r"Testing Procedure", re.I),
    ],
    start_markers=[re.compile(r"PCI DSS Requirements", re.I)],
    stop_markers=[re.compile(r"^Appendix\s+[A-Z]", re.I)],
    requirement_keywords=["must", "shall", "examine", "verify", "confirm"],
)

CisControlsProfile = StandardProfile(
    name="cis",
    heading_patterns=[
        re.compile(
            r"^(?P<id>CIS\s+(?:Control|Safeguard)\s+\d+(?:\.\d+)*)\s+(?P<title>.+)$",
            re.I,
        ),
        re.compile(
            r"^(?P<id>(?:Control|Safeguard)\s+\d+(?:\.\d+)*)\s+(?P<title>.+)$",
            re.I,
        ),
    ],
    start_markers=[re.compile(r"CIS Controls", re.I)],
    stop_markers=[re.compile(r"^Appendix\s+[A-Z]", re.I)],
    requirement_keywords=["ensure", "establish", "maintain", "restrict", "protect"],
)

GenericNumberedStandardProfile = StandardProfile(
    name="generic",
    heading_patterns=[
        re.compile(r"^(?P<id>\d+(?:\.\d+)*)\s+(?P<title>.+)$"),
        re.compile(
            r"^(?P<id>(?:Annex|Appendix|Schedule)\s+[A-Z0-9](?:\.\d+)*)\s+(?P<title>.+)$",
            re.I,
        ),
        re.compile(
            r"^(?P<id>(?:Article|Chapter|Section|Part)\s+\d+)\s*(?::|-)?\s*(?P<title>.*)$",
            re.I,
        ),
    ],
    requirement_keywords=["shall", "must", "ensure", "implement", "maintain"],
)

GenericProfile = GenericNumberedStandardProfile

PROFILES = {
    "nist80053": Nist80053Profile,
    "control-catalog": ControlCatalogProfile,
    "iso": IsoLikeProfile,
    "legal": LegalArticleProfile,
    "pci": PciDssProfile,
    "cis": CisControlsProfile,
    "generic": GenericNumberedStandardProfile,
}

PROFILE_ALIASES = {
    "nist_control_catalog": "nist80053",
    "iso_numbered_standard": "iso",
    "legal_articles": "legal",
    "pci_requirement_table": "pci",
    "cis_controls": "cis",
    "generic_numbered_guideline": "generic",
    "table-heavy": "generic",
    "unknown": "generic",
}


def resolve_profile(name: str) -> StandardProfile:
    """Return profile object for a CLI/auto profile name."""
    key = PROFILE_ALIASES.get(name, name)
    return PROFILES.get(key, GenericNumberedStandardProfile)
