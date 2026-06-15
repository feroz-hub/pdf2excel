"""Structure parser module for pdf2excel.

Converts layout blocks into structured standard items based on active document profile,
applying heading contextual matching, paragraph grouping, list merging, and table normalization.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional
from standard_patterns import StandardProfile, PROFILES, parse_heading_with_profile, parse_heading_generic, match_section_label
from table_normalizer import normalize_table

# Obligation regex for classification
OBLIGATION_RE = re.compile(
    r"\b(shall not|shall|must|is required to|are required to|requires|ensure|implement|maintain|establish|document|monitor|review|verify|restrict|protect|encrypt|authenticate|authorize|retain|assess)\b",
    re.IGNORECASE
)

# Information regex
INFO_RE = re.compile(
    r"\b(definition|overview|purpose only|scope only|example only|references only|glossary|acronym|withdrawn|informative note)\b",
    re.IGNORECASE
)


def classify_text(text: str) -> str:
    """Classify text as Requirement or Information based on obligation words."""
    txt_lower = text.lower()
    if INFO_RE.search(txt_lower):
        return "Information"
    if OBLIGATION_RE.search(txt_lower):
        return "Requirement"
    return "Information"


def parse_blocks_to_items(
    blocks: List[Dict[str, Any]],
    profile: Optional[StandardProfile] = None,
    include_appendix: bool = True
) -> List[Dict[str, Any]]:
    """Convert layout blocks into structured Standard Assessment items."""
    items: List[Dict[str, Any]] = []
    
    # Track current context
    current_clause_id = ""
    current_title = ""
    
    # For control catalog accumulation
    acc_text = []
    acc_page = None
    acc_bbox = None
    acc_issues = []
    acc_confidence = 1.0
    
    # Active profile name
    profile_name = profile.name if profile else "generic"
    is_catalog = profile_name in ("nist80053", "control-catalog")

    def flush_accumulator():
        nonlocal acc_text, acc_page, acc_bbox, acc_issues, acc_confidence
        if not acc_text:
            return
        full_text = "\n".join(acc_text).strip()
        if full_text:
            classification = "Information"
            if "[withdrawn" in full_text.lower():
                classification = "Information"
            else:
                classification = classify_text(full_text)

            items.append({
                "page": acc_page or 1,
                "source_page": acc_page or 1,
                "source_type": "control" if is_catalog else "paragraph",
                "clause_id": current_clause_id,
                "title": current_title,
                "text": full_text,
                "classification": classification,
                "confidence": acc_confidence,
                "issues": sorted(list(set(acc_issues))),
                "bbox": acc_bbox
            })
        acc_text = []
        acc_issues = []
        acc_confidence = 1.0
        acc_bbox = None

    for block in blocks:
        b_type = block.get("block_type", "paragraph")
        text = block.get("text", "").strip()
        page = block.get("page", 1)
        bbox = block.get("bbox")
        issues = list(block.get("issues", []))
        confidence = block.get("confidence", 1.0)
        
        # 1. Skip TOC, errata, front matter pages/blocks if not desired
        if b_type == "toc":
            continue
            
        # Skip appendix if configured
        if not include_appendix and b_type == "heading":
            # Check if title indicates appendix
            if any(k in text.lower() for k in ("appendix", "annex", "schedule")):
                # From now on we skip until next non-appendix heading (or just skip all remaining if at the end)
                # For simplicity, we skip this heading block and keep context
                continue

        # 2. Heading block processing
        if b_type == "heading":
            # Parse heading
            parsed = None
            if profile:
                parsed = parse_heading_with_profile(text, profile)
            else:
                parsed = parse_heading_generic(text)

            if parsed:
                # If it's a catalog, flush the previous control text first
                if is_catalog:
                    flush_accumulator()

                # Update context
                clause_id = parsed["clause_id"]
                title = parsed["title"]
                
                # Check for control enhancement sub-controls
                # e.g. for (1), if current parent control is AC-2, full ID becomes AC-2(1)
                if is_catalog and clause_id.isdigit() and current_clause_id:
                    # Parse base ID: e.g. AC-2 from AC-2(3) or AC-2
                    base_m = re.match(r"^([A-Z]{2}-\d+)", current_clause_id)
                    if base_m:
                        parent_id = base_m.group(1)
                        clause_id = f"{parent_id}({clause_id})"
                
                current_clause_id = clause_id
                current_title = title if title else current_title
                
                # If not catalog, heading updates context but does not emit row directly
                # If catalog, the heading is the start of the control, we initialize accumulator
                if is_catalog:
                    acc_page = page
                    acc_bbox = bbox
                    acc_confidence = confidence
                    acc_issues = list(issues)
                    # We can add the title/heading text as start of control text if needed,
                    # or keep accumulator empty. Usually we keep it empty and accumulate body text.
            else:
                # Heading that couldn't be parsed structure-wise
                if is_catalog:
                    flush_accumulator()
                # Treat as generic heading updating title
                current_title = text
                
            continue

        # 3. Table block processing
        if b_type == "table":
            if is_catalog:
                flush_accumulator()
            
            rows = block.get("rows")
            normalized_items = []
            if rows:
                normalized_items = normalize_table(rows, profile)
                
            if normalized_items:
                for item in normalized_items:
                    # Inherit context if row didn't parse a clause ID
                    item["clause_id"] = item["clause_id"] or current_clause_id
                    item["title"] = item["title"] or current_title
                    item["page"] = page
                    item["source_page"] = page
                    item["source_type"] = "table"
                    item["bbox"] = bbox
                    # merge issues
                    item["issues"] = sorted(list(set(item["issues"] + issues)))
                    items.append(item)
            else:
                # Fallback to single cell joins
                classification = classify_text(text)
                items.append({
                    "page": page,
                    "source_page": page,
                    "source_type": "table",
                    "clause_id": current_clause_id,
                    "title": current_title,
                    "text": text,
                    "classification": classification,
                    "confidence": 0.85,
                    "issues": sorted(list(set(issues + ["table_structure_fallback"]))),
                    "bbox": bbox
                })
            continue

        # 4. Paragraph / list / note block processing
        if is_catalog:
            # For catalogs, accumulate all text under the active control heading
            # Filter boilerplate or empty text
            if text and not text.startswith("[Withdrawn:"):
                # Clean up text
                acc_text.append(text)
                if not acc_page:
                    acc_page = page
                if not acc_bbox:
                    acc_bbox = bbox
                acc_issues.extend(issues)
                acc_confidence = min(acc_confidence, confidence)
            elif text.startswith("[Withdrawn:"):
                # Withdrawn control: flush current, emit withdrawn control, and reset
                flush_accumulator()
                items.append({
                    "page": page,
                    "source_page": page,
                    "source_type": "control",
                    "clause_id": current_clause_id,
                    "title": current_title,
                    "text": text,
                    "classification": "Information",
                    "confidence": confidence,
                    "issues": sorted(list(set(issues))),
                    "bbox": bbox
                })
        else:
            # For prose, each paragraph is typically a requirement or info row
            if b_type == "list_item":
                # Check if it has requirement keywords
                is_req = OBLIGATION_RE.search(text)
                if is_req:
                    # Emit list item as independent requirement row
                    items.append({
                        "page": page,
                        "source_page": page,
                        "source_type": "list",
                        "clause_id": current_clause_id,
                        "title": current_title,
                        "text": text,
                        "classification": "Requirement",
                        "confidence": confidence,
                        "issues": sorted(list(set(issues))),
                        "bbox": bbox
                    })
                else:
                    # Merge with previous item if possible to avoid fragmentation
                    if items and items[-1]["source_page"] == page and items[-1]["clause_id"] == current_clause_id:
                        prev_item = items[-1]
                        prev_item["text"] += "\n" + text
                        prev_item["issues"] = sorted(list(set(prev_item["issues"] + issues)))
                    else:
                        items.append({
                            "page": page,
                            "source_page": page,
                            "source_type": "list",
                            "clause_id": current_clause_id,
                            "title": current_title,
                            "text": text,
                            "classification": "Information",
                            "confidence": confidence,
                            "issues": sorted(list(set(issues))),
                            "bbox": bbox
                        })
            else:
                # Regular paragraph or note
                classification = classify_text(text)
                items.append({
                    "page": page,
                    "source_page": page,
                    "source_type": "paragraph",
                    "clause_id": current_clause_id,
                    "title": current_title,
                    "text": text,
                    "classification": classification,
                    "confidence": confidence,
                    "issues": sorted(list(set(issues))),
                    "bbox": bbox
                })

    if is_catalog:
        flush_accumulator()

    # Final sweep: filter out bad items
    filtered_items = []
    for item in items:
        txt = item.get("text", "").strip()
        # never output one-letter fragments
        if len(txt) <= 1:
            continue
        filtered_items.append(item)

    return filtered_items
