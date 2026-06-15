"""Layout block extraction layer for pdf2excel.

Segments page content into unified layout blocks (headings, paragraphs, tables,
list_items, notes, footers, headers, toc, unknown) with detailed metadata and bounding boxes.
"""

from __future__ import annotations

import re
import logging
import statistics
from typing import Any, Dict, List, Optional, Tuple, Set

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

from deck_tables_to_excel import _extract_page_tables, _merge_side_by_side
from pdf_paragraphs_to_excel import _clean_text as _para_clean_text, _looks_like_page_number, _SENTENCE_END_RE, _HYPHEN_END_RE, _normalize_digits
from standard_patterns import parse_heading_generic, match_section_label

log = logging.getLogger("pdf2excel.layout_blocks")

# List start marker: bullets or (a), 1) etc.
LIST_START_RE = re.compile(r"^([•\-\*\▪\◦\♦]|\([a-z0-9]+\)|[a-z0-9]+\))\s+", re.IGNORECASE)
NOTE_START_RE = re.compile(r"^(note|remark|caution|warning|nb|n\.b\.)\b", re.IGNORECASE)
CID_RE = re.compile(r"\(cid:\d+\)")

class BlockLine:
    """A visual line of text on a page with layout metrics."""
    def __init__(
        self,
        text: str,
        top: float,
        bottom: float,
        x0: float,
        x1: float,
        font_size: float,
        font_name: str,
        is_bold: bool,
        is_rotated: bool,
        words: List[dict]
    ):
        self.text = text
        self.top = top
        self.bottom = bottom
        self.x0 = x0
        self.x1 = x1
        self.font_size = font_size
        self.font_name = font_name
        self.is_bold = is_bold
        self.is_rotated = is_rotated
        self.words = words


def clean_block_text(text: str) -> str:
    """Normalize text: remove cid, soft-hyphens, BOM, duplicate spaces."""
    text = CID_RE.sub("", text)
    text = text.replace("\u00ad", "")  # soft hyphen
    text = text.replace("\ufeff", "")  # BOM
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def extract_page_tables_with_fallbacks(page: Any) -> List[Any]:
    """Extract tables using standard and borderless fallbacks."""
    tables = _extract_page_tables(page)
    if not tables:
        try:
            from deck_tables_to_excel import _Table, _clean_cell, _is_real_table
            # Borderless check
            for tbl in page.find_tables(table_settings={
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
                "snap_y_tolerance": 5,
                "snap_x_tolerance": 5,
            }):
                raw = tbl.extract()
                rows = [[_clean_cell(c) for c in row] for row in raw]
                if _is_real_table(rows):
                    tables.append(_Table(rows=rows, bbox=tuple(tbl.bbox)))
        except Exception as exc:
            log.warning("Borderless table fallback failed: %s", exc)
    return _merge_side_by_side(tables)


def cluster_words_into_lines(words: List[dict]) -> List[BlockLine]:
    """Group pdfplumber word boxes into visual lines with style attributes."""
    if not words:
        return []

    # Sort top-to-bottom, then left-to-right
    words = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))

    lines: List[List[dict]] = []
    current = [words[0]]
    current_top = words[0]["top"]

    for w in words[1:]:
        height = max(w["bottom"] - w["top"], 1.0)
        tol = height * 0.6
        if abs(w["top"] - current_top) <= tol:
            current.append(w)
        else:
            lines.append(current)
            current = [w]
            current_top = w["top"]
    lines.append(current)

    result = []
    for group in lines:
        group = sorted(group, key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in group)
        text = clean_block_text(text)
        if not text:
            continue

        x0 = min(w["x0"] for w in group)
        x1 = max(w["x1"] for w in group)
        top = min(w["top"] for w in group)
        bottom = max(w["bottom"] for w in group)

        sizes = [w.get("size") for w in group if w.get("size") is not None]
        avg_size = sum(sizes) / len(sizes) if sizes else 10.0

        # Dominant font name
        fonts = [w.get("fontname", "") for w in group if w.get("fontname") is not None]
        font_name = max(set(fonts), key=fonts.count) if fonts else ""

        is_bold = any(k in font_name.lower() for k in ("bold", "black", "heavy", "semibold"))
        is_rotated = any(w.get("upright") is False for w in group)

        result.append(BlockLine(
            text=text, top=top, bottom=bottom, x0=x0, x1=x1,
            font_size=avg_size, font_name=font_name, is_bold=is_bold,
            is_rotated=is_rotated, words=group
        ))
    return result


def is_visual_heading(line: BlockLine, doc_median_size: float, median_gap: float, next_line_gap: float) -> bool:
    """Visual heading classifier (standard pattern match, size, boldness & spacing)."""
    text = line.text.strip()
    
    # 1. Check patterns (ISO, legal, NIST, CIS, etc.)
    p_match = parse_heading_generic(text)
    if p_match and len(text) < 150:
        return True

    # 2. Visually distinct size
    if line.font_size >= doc_median_size + 1.5 and len(text) < 120:
        return True

    # 3. Bold, short, followed by spacing
    if line.is_bold and len(text) < 80 and next_line_gap > median_gap * 1.4:
        return True

    return False


def join_block_lines(lines: List[BlockLine]) -> str:
    """Fuse lines in buffer with hyphenation resolution."""
    out = ""
    for ln in lines:
        piece = ln.text
        if not out:
            out = piece
            continue
        m = _HYPHEN_END_RE.search(out)
        if m:
            out = out[:m.start(1) + 1] + piece.lstrip()
        else:
            out = out + " " + piece
    return clean_block_text(out)


def group_lines_into_blocks(
    lines: List[BlockLine],
    page_num: int,
    page_type: str,
    gap_factor: float,
    doc_median_size: float
) -> List[dict]:
    """Cluster visual lines into paragraph, heading, list_item, note, or toc blocks."""
    if not lines:
        return []

    # Calculate line-to-line gaps
    gaps = []
    for prev, cur in zip(lines, lines[1:]):
        gap = cur.top - prev.bottom
        if gap > 0:
            gaps.append(gap)
    median_gap = statistics.median(gaps) if gaps else 0.0
    gap_threshold = median_gap * gap_factor if median_gap > 0 else 0.0

    blocks: List[dict] = []
    buffer: List[BlockLine] = []
    current_type = "paragraph"

    def flush():
        if not buffer:
            return
        text = join_block_lines(buffer)
        if text:
            # Aggregate metrics
            x0 = min(ln.x0 for ln in buffer)
            top = min(ln.top for ln in buffer)
            x1 = max(ln.x1 for ln in buffer)
            bottom = max(ln.bottom for ln in buffer)
            avg_size = sum(ln.font_size for ln in buffer) / len(buffer)
            font_name = buffer[0].font_name
            is_bold = any(ln.is_bold for ln in buffer)
            is_rotated = any(ln.is_rotated for ln in buffer)

            block_type = "toc" if page_type == "toc" else current_type
            confidence = 0.95
            issues = []

            # Safety check: one-letter fragments
            if len(text) <= 2 and not text.isalnum():
                issues.append("single_char_text")
                confidence = min(confidence, 0.2)

            blocks.append({
                "page": page_num,
                "block_type": block_type,
                "text": text,
                "bbox": [x0, top, x1, bottom],
                "font_size": avg_size,
                "font_name": font_name,
                "is_bold": is_bold,
                "is_rotated": is_rotated,
                "confidence": confidence,
                "issues": issues
            })
        buffer.clear()

    for idx, ln in enumerate(lines):
        next_gap = 0.0
        if idx + 1 < len(lines):
            next_gap = lines[idx + 1].top - ln.bottom

        # 1. Heading check
        if is_visual_heading(ln, doc_median_size, median_gap, next_gap):
            flush()
            buffer.append(ln)
            current_type = "heading"
            flush()
            current_type = "paragraph"
            continue

        # 2. Note check
        if NOTE_START_RE.match(ln.text):
            flush()
            buffer.append(ln)
            current_type = "note"
            continue

        # 3. List item check
        if LIST_START_RE.match(ln.text):
            flush()
            buffer.append(ln)
            current_type = "list_item"
            continue

        # 4. Standard grouping decision
        if buffer:
            prev = buffer[-1]
            gap = ln.top - prev.bottom
            is_break = False

            if gap_threshold and gap > gap_threshold:
                is_break = True
            elif abs(ln.font_size - prev.font_size) > 1.2:
                is_break = True
            elif ln.is_bold != prev.is_bold:
                is_break = True
            elif abs(ln.x0 - prev.x0) > 15:
                is_break = True
            elif LIST_START_RE.match(ln.text) or NOTE_START_RE.match(ln.text):
                is_break = True
            elif _SENTENCE_END_RE.search(prev.text) and ln.text[:1].isupper():
                is_break = True

            if is_break:
                flush()
                if NOTE_START_RE.match(ln.text):
                    current_type = "note"
                elif LIST_START_RE.match(ln.text):
                    current_type = "list_item"
                else:
                    current_type = "paragraph"

        buffer.append(ln)

    flush()
    return blocks


def sort_blocks_layout(blocks: List[dict], page_width: float, is_multi_column: bool) -> List[dict]:
    """Sort blocks by page reading order, column-aware."""
    if not is_multi_column:
        return sorted(blocks, key=lambda b: b["bbox"][1])

    sorted_by_y = sorted(blocks, key=lambda b: b["bbox"][1])
    center_x = page_width / 2

    bands = []
    current_band = []

    for b in sorted_by_y:
        x0, top, x1, bottom = b["bbox"]
        width = x1 - x0
        is_full_width = (width > page_width * 0.6) and (x0 < center_x < x1)

        if is_full_width:
            if current_band:
                bands.append((False, current_band))
                current_band = []
            bands.append((True, [b]))
        else:
            current_band.append(b)

    if current_band:
        bands.append((False, current_band))

    final_blocks = []
    for is_fw, band_blocks in bands:
        if is_fw:
            final_blocks.extend(band_blocks)
        else:
            left_col = []
            right_col = []
            for b in band_blocks:
                x0, top, x1, bottom = b["bbox"]
                cx = (x0 + x1) / 2
                if cx < center_x:
                    left_col.append(b)
                else:
                    right_col.append(b)
            final_blocks.extend(sorted(left_col, key=lambda b: b["bbox"][1]))
            final_blocks.extend(sorted(right_col, key=lambda b: b["bbox"][1]))

    return final_blocks


def extract_page_blocks(
    page: Any,
    preflight_meta: Dict[str, Any],
    gap_factor: float,
    doc_median_size: float,
    boilerplate_set: Set[str],
    ocr_mode: str = "off"
) -> List[Dict[str, Any]]:
    """Extract and group blocks for a single page based on preflight metadata."""
    page_num = preflight_meta["page_number"]
    page_type = preflight_meta["page_type"]

    # 1. Scanned page handling
    if preflight_meta["is_scanned"]:
        ocr_text = preflight_meta.get("ocr_text", "")
        if ocr_text:
            lines = []
            for line_str in ocr_text.splitlines():
                line_str = line_str.strip()
                if line_str:
                    lines.append(BlockLine(
                        text=line_str, top=0, bottom=0, x0=0, x1=page.width,
                        font_size=doc_median_size, font_name="", is_bold=False,
                        is_rotated=False, words=[]
                    ))
            return group_lines_into_blocks(lines, page_num, page_type, gap_factor, doc_median_size)
        else:
            return [{
                "page": page_num,
                "block_type": "scanned",
                "text": "[Scanned Page — Selectable Text Missing]",
                "bbox": [0.0, 0.0, page.width, page.height],
                "font_size": doc_median_size,
                "font_name": "",
                "is_bold": False,
                "is_rotated": False,
                "confidence": 0.1,
                "issues": ["scanned_page_skipped"]
            }]

    # 2. Extract tables
    tables = []
    if page_type in ("table", "mixed", "unknown"):
        tables = extract_page_tables_with_fallbacks(page)

    # 3. Extract prose lines
    prose_lines = []
    table_bboxes = [t.bbox for t in tables]
    
    # Check for rotated text coordinates
    words = page.extract_words(
        use_text_flow=False,
        keep_blank_chars=False,
        extra_attrs=["size", "fontname", "upright"]
    )

    prose_words = []
    for w in words:
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        
        # Rotated text filter: if upright=False and sits in margin, filter it
        if w.get("upright") is False and w["x0"] < 45:
            continue
        # Standard DOI filter
        if "doi.org" in w["text"].lower() and w["x0"] < 45:
            continue

        inside_table = False
        for bx0, btop, bx1, bbottom in table_bboxes:
            if bx0 - 2 <= cx <= bx1 + 2 and btop - 2 <= cy <= bbottom + 2:
                inside_table = True
                break
        if not inside_table:
            prose_words.append(w)

    raw_lines = cluster_words_into_lines(prose_words)

    # Filter repeating headers/footers
    for ln in raw_lines:
        norm = _normalize_digits(ln.text).strip()
        if norm in boilerplate_set or _looks_like_page_number(ln.text):
            continue
        # Skip DOI line
        if "doi.org/10.6028/NIST.SP.800-53" in ln.text:
            continue
        prose_lines.append(ln)

    # 4. Group lines into blocks
    blocks = group_lines_into_blocks(prose_lines, page_num, page_type, gap_factor, doc_median_size)

    # 5. Convert tables to blocks
    for tbl in tables:
        text_rows = []
        for row in tbl.rows:
            text_rows.append(" | ".join(str(cell or "").strip() for cell in row))
        table_text = "\n".join(text_rows)

        blocks.append({
            "page": page_num,
            "block_type": "table",
            "text": table_text,
            "bbox": list(tbl.bbox),
            "font_size": doc_median_size,
            "font_name": "",
            "is_bold": False,
            "is_rotated": False,
            "confidence": 0.9,
            "issues": [],
            "rows": tbl.rows
        })

    # 6. Sort by reading order
    is_multi_column = preflight_meta["is_multi_column"]
    return sort_blocks_layout(blocks, page.width, is_multi_column)


def extract_pdf_blocks(
    pdf_path: str,
    page_preflights: List[Dict[str, Any]],
    gap_factor: float = 1.6,
    ocr_mode: str = "off",
    skip_cover: bool = True,
    include_toc: bool = False
) -> List[Dict[str, Any]]:
    """Extract layout blocks across all pages of a PDF, detecting boilerplate."""
    if not pdfplumber:
        raise RuntimeError("pdfplumber is not installed.")

    # Calculate document-wide median font size
    sizes = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            try:
                for w in page.extract_words(extra_attrs=["size"]):
                    if w.get("size") is not None:
                        sizes.append(w["size"])
            except Exception:
                for w in page.extract_words():
                    sizes.append(10.0)
    doc_median_size = statistics.median(sizes) if sizes else 10.0

    # Boilerplate detection
    all_pages_lines = []
    preflight_by_page = {p["page_number"]: p for p in page_preflights}

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            meta = preflight_by_page.get(page_num)
            if not meta:
                continue
            if skip_cover and meta["page_type"] == "cover":
                continue
            if not include_toc and meta["page_type"] == "toc":
                continue
            if meta["is_scanned"]:
                continue

            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            tables = _extract_page_tables(page)
            table_bboxes = [t.bbox for t in tables]
            prose_words = []
            for w in words:
                cx = (w["x0"] + w["x1"]) / 2
                cy = (w["top"] + w["bottom"]) / 2
                inside = False
                for bx0, btop, bx1, bbottom in table_bboxes:
                    if bx0 - 2 <= cx <= bx1 + 2 and btop - 2 <= cy <= bbottom + 2:
                        inside = True
                        break
                if not inside:
                    prose_words.append(w)

            lines = cluster_words_into_lines(prose_words)
            # Match Line class signature for boilerplate detector
            from pdf_paragraphs_to_excel import _Line
            adapted_lines = [
                _Line(text=ln.text, top=ln.top, bottom=ln.bottom, x0=ln.x0, words=ln.words)
                for ln in lines
            ]
            all_pages_lines.append(adapted_lines)

    from pdf_paragraphs_to_excel import _detect_repeating_headers_footers
    boilerplate_set = _detect_repeating_headers_footers(all_pages_lines)

    # Perform page-by-page block extraction
    all_blocks = []
    block_id_counter = 1

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            meta = preflight_by_page.get(page_num)
            if not meta:
                continue
            if skip_cover and meta["page_type"] == "cover":
                continue
            if not include_toc and meta["page_type"] == "toc":
                continue

            page_blocks = extract_page_blocks(
                page=page,
                preflight_meta=meta,
                gap_factor=gap_factor,
                doc_median_size=doc_median_size,
                boilerplate_set=boilerplate_set,
                ocr_mode=ocr_mode
            )

            for b in page_blocks:
                b["block_id"] = block_id_counter
                block_id_counter += 1
                all_blocks.append(b)

    return all_blocks
