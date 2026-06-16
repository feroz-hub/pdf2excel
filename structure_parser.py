"""Structure parser — converts cleaned layout blocks into finalized standard items."""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from standard_patterns import (
    NIST_BASE_RE,
    NIST_ENH_RE,
    match_section_label,
    parse_heading_generic,
    parse_heading_with_profile,
    _is_uppercase_title,
)
from standard_profiles import StandardProfile
from table_normalizer import normalize_table

OBLIGATION_RE = re.compile(
    r"\b(shall not|shall|must|is required to|are required to|requires|ensure|"
    r"implement|maintain|establish|document|monitor|review|verify|restrict|"
    r"protect|encrypt|authenticate|authorize|retain|assess)\b",
    re.IGNORECASE,
)
INFO_RE = re.compile(
    r"\b(definition|overview|purpose only|scope only|example only|references only|"
    r"glossary|acronym|withdrawn|informative note)\b",
    re.IGNORECASE,
)

HEADER_RE = re.compile(
    r"NIST\s+SP\s+800-53.*?SECURITY\s+AND\s+PRIVACY\s+CONTROLS.*?ORGANIZATIONS",
    re.IGNORECASE,
)
SPACED_HEADER_RE = re.compile(
    r"N\s*I\s*S\s*T\s+S\s*P\s+8\s*0\s*0\s*-\s*5\s*3\s*,\s*R\s*E\s*V\s*\.?\s*5\s+"
    r"S\s*E\s*C\s*U\s*R\s*I\s*T\s*Y\s+A\s*N\s*D\s+P\s*R\s*I\s*V\s*A\s*C\s*Y\s+"
    r"C\s*O\s*N\s*T\s*R\s*O\s*L\s*S\s+F\s*O\s*R\s+I\s*N\s*F\s*O\s*R\s*M\s*A\s*T\s*I\s*O\s*N\s+"
    r"S\s*Y\s*S\s*T\s*E\s*M\s*S\s+A\s*N\s*D\s+O\s*R\s*G\s*A\s*N\s*I\s*Z\s*A\s*T\s*I\s*O\s*N\s*S",
    re.IGNORECASE,
)
CHAPTER_PAGE_RE = re.compile(
    r"^(CHAPTER|APPENDIX|C\s*H\s*A\s*P\s*T\s*E\s*R|A\s*P\s*P\s*E\s*N\s*D\s*I\s*X)\s+"
    r"([A-Z0-9\s]+)?\s*PAGE\s+\d+",
    re.IGNORECASE,
)
DOI_RE = re.compile(
    r"(This publication is available free of charge|doi\.org/10\.6028/NIST\.SP\.800-53r5)",
    re.IGNORECASE,
)
NOISE_LINE_RE = re.compile(
    r"^(JOINT TASK FORCE|Authority|Abstract|Errata|Foreword|Preface|Acknowledgments?)\b",
    re.IGNORECASE,
)


def classify_text(text: str, is_catalog: bool = False) -> str:
    if "[withdrawn" in text.lower():
        return "Information"
    if is_catalog:
        return "Requirement"
    if INFO_RE.search(text):
        return "Information"
    if OBLIGATION_RE.search(text):
        return "Requirement"
    return "Information"


def clean_noise_text(text: str) -> str:
    lines = []
    for line in text.split("\n"):
        cleaned = line.strip()
        if not cleaned:
            continue
        for pat in (HEADER_RE, SPACED_HEADER_RE, CHAPTER_PAGE_RE, DOI_RE):
            cleaned = pat.sub("", cleaned).strip()
        cleaned = re.sub(r"This publication is available free of charge( from:)?", "", cleaned, flags=re.I)
        cleaned = re.sub(r"https?://(dx\.)?doi\.org/[^\s]+", "", cleaned, flags=re.I)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        if cleaned:
            lines.append(cleaned)
    return "\n".join(lines).strip()


def is_blocking_noise(text: str, block: Dict[str, Any], profile: Optional[StandardProfile]) -> Optional[str]:
    """Return blocking issue code if block should never export."""
    t = text.strip()
    if not t:
        return "empty_text"
    if block.get("block_type") in ("header", "footer"):
        return "header_footer"
    if block.get("block_type") == "toc":
        return "toc_row"
    if DOI_RE.search(t) or "doi.org/" in t.lower():
        return "doi_footer"
    if CHAPTER_PAGE_RE.search(t):
        return "header_footer"
    if NOISE_LINE_RE.match(t):
        return "front_matter_skipped"
    if len(t) == 1:
        return "one_letter_fragment"
    if re.fullmatch(r"[a-zA-Z]", t):
        return "one_letter_fragment"
    if "TABLE C-" in t.upper() or t.strip() == "NUMBER" or t.strip().upper() == "NUMBER":
        return "appendix_summary_row"
    issues = block.get("issues") or []
    if "single_char_text" in issues:
        return "one_letter_fragment"
    if "scanned_page_skipped" in issues:
        return "raw_layout_fragment"
    return None


def is_start_boundary(text: str, profile: Optional[StandardProfile], is_catalog: bool) -> bool:
    if is_catalog:
        return bool(_match_catalog_base(text))
    if profile and profile.start_markers:
        return any(m.search(text) for m in profile.start_markers)
    return False


def is_stop_boundary(text: str, profile: Optional[StandardProfile]) -> bool:
    """Match back-matter section headings only — not in-control 'References:' labels."""
    clean = text.strip()
    if len(clean) > 80:
        return False
    # In-control section labels (References:, Discussion:, etc.) are not stop boundaries
    if match_section_label(clean):
        return False
    if re.match(r"^(REFERENCES|APPENDIX [ABC])(\s+PAGE\s+\d+)?\s*$", clean, re.I):
        return True
    if profile and profile.stop_markers:
        return any(m.search(clean) for m in profile.stop_markers)
    return False


def _match_catalog_base(text: str) -> Optional[re.Match]:
    return NIST_BASE_RE.match(text.strip())


def _match_catalog_enhancement(text: str) -> Optional[re.Match]:
    m = NIST_ENH_RE.match(text.strip())
    if m and _is_uppercase_title(m.group("title")):
        return m
    return None


def parse_blocks_to_items(
    blocks: List[Dict[str, Any]],
    profile: Optional[StandardProfile] = None,
    include_front_matter: bool = False,
    include_appendix: bool = False,
    include_references: bool = False,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """State-machine parser: blocks → finalized items + rejected audit rows."""
    finalized: List[Dict[str, Any]] = []
    rejected: List[Dict[str, Any]] = []

    profile_name = profile.name if profile else "generic"
    is_catalog = profile_name in ("nist80053", "control-catalog")

    current_item: Optional[Dict[str, Any]] = None
    current_parent_id: Optional[str] = None
    current_parent_title: Optional[str] = None
    export_started = include_front_matter or not (
        (profile.requires_start_boundary if profile else False) or is_catalog
    )

    def reject(block: Dict[str, Any], text: str, reason: str, extra: Optional[List[str]] = None):
        issues = sorted(set(list(block.get("issues", [])) + [reason] + (extra or [])))
        rejected.append({
            "clause_id": "",
            "title": "",
            "text": clean_noise_text(text),
            "classification": "Information",
            "source_page": block.get("page", 1),
            "page": block.get("page", 1),
            "source_type": block.get("block_type", "unknown"),
            "confidence": 0.1,
            "issues": issues,
            "bbox": block.get("bbox"),
            "export_status": "rejected",
            "raw_text": text,
        })

    def flush_current():
        nonlocal current_item
        if current_item is None:
            return
        raw = "\n".join(current_item.pop("text_parts", [])).strip()
        cleaned = clean_noise_text(raw)
        current_item["text"] = cleaned
        if cleaned:
            current_item["classification"] = classify_text(cleaned, is_catalog)
            finalized.append(current_item)
        else:
            current_item["issues"] = sorted(set(current_item.get("issues", []) + ["empty_text"]))
            rejected.append({**current_item, "export_status": "rejected", "raw_text": raw})
        current_item = None

    def start_item(clause_id: str, title: str, block: Dict[str, Any], source_type: str):
        nonlocal current_item
        flush_current()
        current_item = {
            "clause_id": clause_id,
            "title": title,
            "text_parts": [],
            "classification": "Requirement",
            "source_page": block.get("page", 1),
            "page": block.get("page", 1),
            "source_type": source_type,
            "confidence": block.get("confidence", 0.95),
            "issues": list(block.get("issues", [])),
            "bbox": block.get("bbox"),
        }

    for block in blocks:
        text = (block.get("text") or "").strip()
        b_type = block.get("block_type", "paragraph")
        page = block.get("page", 1)

        noise = is_blocking_noise(text, block, profile)
        if noise and not export_started:
            reject(block, text, noise if noise != "front_matter_skipped" else "front_matter_skipped")
            continue

        if is_start_boundary(text, profile, is_catalog) and not export_started:
            export_started = True

        if not export_started:
            reject(block, text, "front_matter_skipped")
            continue

        if is_stop_boundary(text, profile):
            if not include_appendix and not include_references:
                flush_current()
                reject(block, text, "appendix_summary_row" if "APPENDIX" in text.upper() else "front_matter_skipped")
                break
            continue

        if b_type == "toc" or "TABLE OF CONTENTS" in text.upper():
            reject(block, text, "toc_row")
            continue

        if noise:
            reject(block, text, noise)
            continue

        # --- Catalog (NIST) headings ---
        if is_catalog:
            m_base = _match_catalog_base(text)
            if m_base:
                current_parent_id = m_base.group("id")
                current_parent_title = m_base.group("title").strip()
                start_item(current_parent_id, current_parent_title, block, "control")
                continue

            m_enh = _match_catalog_enhancement(text)
            if m_enh and current_parent_id:
                enh_num = m_enh.group("num")
                enh_title = m_enh.group("title").strip()
                title = enh_title
                if current_parent_title and current_parent_title.upper() not in title.upper():
                    title = f"{current_parent_title} | {title}"
                start_item(f"{current_parent_id}({enh_num})", title, block, "control_enhancement")
                continue
        else:
            parsed = parse_heading_with_profile(text, profile) if profile else parse_heading_generic(text)
            if parsed and parsed.get("is_new_item"):
                start_item(parsed["clause_id"], parsed.get("title") or "", block, parsed.get("kind", "base_clause"))
                continue
            if parsed and parsed.get("is_enhancement_or_subitem") and current_parent_id:
                cid = f"{current_parent_id}({parsed['clause_id']})"
                start_item(cid, parsed.get("title") or "", block, "sub_clause")
                continue

        if b_type == "table":
            if is_catalog:
                if current_item is not None:
                    current_item["text_parts"].append(text)
            else:
                flush_current()
                rows = block.get("rows")
                normalized = normalize_table(rows, profile) if rows else []
                for item in normalized:
                    item.update({
                        "page": page,
                        "source_page": page,
                        "source_type": "table",
                        "bbox": block.get("bbox"),
                        "export_status": "exported",
                    })
                    finalized.append(item)
            continue

        if current_item is not None:
            current_item["text_parts"].append(text)
            current_item["issues"] = sorted(set(current_item.get("issues", []) + list(block.get("issues", []))))
            current_item["confidence"] = min(current_item["confidence"], block.get("confidence", 1.0))
        else:
            reject(block, text, "orphan_block_not_exported")

    flush_current()

    # Attach export_status to finalized items
    for it in finalized:
        it.setdefault("export_status", "exported")

    all_rows = finalized + rejected
    return all_rows, rejected
