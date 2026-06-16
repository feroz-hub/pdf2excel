"""PDF preflight module for page-level analysis.

Analyzes each page of a PDF and returns page-level metadata including text/word counts,
orientation, tables, likely page type, layout column structure, and confidence scores.
Supports optional OCR via pytesseract when scanned pages are encountered.
"""

from __future__ import annotations

import re
import logging
from typing import Any, Dict, List

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pytesseract
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

log = logging.getLogger("pdf2excel.preflight")


def detect_multi_column(words: List[Dict[str, Any]], page_width: float) -> bool:
    """Detect if a page layout is multi-column using horizontal word gutter clustering."""
    if not words or len(words) < 20:
        return False
        
    x0s = [w["x0"] for w in words]
    x1s = [w["x1"] for w in words]
    x_min, x_max = min(x0s), max(x1s)
    width = x_max - x_min
    if width < 150:
        return False
        
    # Divide the text width into 20 bins to check for a gutter
    n_bins = 20
    bin_width = width / n_bins
    bins = [0] * n_bins
    for w in words:
        start_bin = int((w["x0"] - x_min) / bin_width)
        end_bin = int((w["x1"] - x_min) / bin_width)
        start_bin = max(0, min(start_bin, n_bins - 1))
        end_bin = max(0, min(end_bin, n_bins - 1))
        for b in range(start_bin, end_bin + 1):
            bins[b] += 1
            
    # For a multi-column page, we expect a gutter (very low word overlap) near the middle third
    sorted_bins = sorted(bins)
    avg_dense = sum(sorted_bins[-5:]) / 5 if len(sorted_bins) >= 5 else (sum(bins) / len(bins) if bins else 1.0)
    if avg_dense == 0:
        return False
        
    # Check middle third bins (typically indices 6 to 13 out of 20)
    for b in range(6, 14):
        if bins[b] < avg_dense * 0.1:  # gutter density is less than 10% of typical peak density
            left_sum = sum(bins[:b])
            right_sum = sum(bins[b+1:])
            # Verify significant text exists on both sides of the gutter
            if left_sum > len(words) * 0.2 and right_sum > len(words) * 0.2:
                return True
    return False


def ocr_page_text(page: Any) -> str:
    """Render page as image and run Tesseract OCR on it.
    
    Requires pytesseract and system tesseract binary, plus pdf2image/poppler.
    """
    if not HAS_TESSERACT:
        raise RuntimeError("pytesseract or tesseract binary is not installed.")
        
    # Test if tesseract is actually executable
    try:
        pytesseract.get_tesseract_version()
    except Exception as exc:
        raise RuntimeError(f"Tesseract binary not found or not working: {exc}") from exc
        
    try:
        # Render the page to a PIL Image at 150 DPI
        im = page.to_image(resolution=150)
        pil_img = im.original
        text = pytesseract.image_to_string(pil_img)
        return text or ""
    except Exception as exc:
        raise RuntimeError(f"Failed to render page image or run OCR: {exc}. Ensure poppler-utils is installed.") from exc


def analyze_page(page: Any, page_num: int, ocr_mode: str = "off") -> Dict[str, Any]:
    """Analyze a single page and return its preflight metadata."""
    # We extract words with extra attributes to check for rotated text
    try:
        words = page.extract_words(extra_attrs=["upright"]) or []
    except Exception:
        words = page.extract_words() or []

    word_count = len(words)
    has_rotated_text = any(w.get("upright") is False for w in words)
    
    # Text character count
    raw_text = page.extract_text() or ""
    char_count = len(raw_text)
    
    has_selectable_text = word_count > 0
    is_portrait = page.width <= page.height
    
    # Check for tables
    tables = page.find_tables()
    has_tables = len(tables) > 0
    
    # Determine if scanned / image-only
    has_visual_elements = len(page.images) > 0 or len(page.rects) > 0 or len(page.lines) > 0 or len(page.curves) > 0
    is_scanned = False
    
    if word_count < 10:
        if has_visual_elements or (char_count < 30 and (len(page.images) > 0 or len(page.rects) > 0)):
            is_scanned = True
        elif word_count == 0 and char_count == 0:
            is_scanned = True
            
    # Gutter-based multi-column check
    is_multi_column = detect_multi_column(words, page.width) if has_selectable_text else False
    
    # OCR logic
    ocr_text = ""
    warnings = []
    confidence = 1.0
    
    if is_scanned:
        if ocr_mode == "detect":
            if HAS_TESSERACT:
                try:
                    ocr_text = ocr_page_text(page)
                    char_count = len(ocr_text)
                    word_count = len(ocr_text.split())
                    has_selectable_text = word_count > 0
                    confidence = 0.7
                    log.info("Page %d: Successfully performed OCR.", page_num)
                except Exception as exc:
                    warnings.append(f"OCR failed: {exc}")
                    confidence = 0.1
            else:
                warnings.append("OCR requested but pytesseract/tesseract not installed.")
                confidence = 0.1
        else:
            warnings.append("Scanned page detected. Selectable text is missing.")
            confidence = 0.1
            
    # Classify page type
    page_type = "unknown"
    if is_scanned and not ocr_text:
        page_type = "scanned"
    else:
        text_to_check = ocr_text if ocr_text else raw_text
        lower_text = text_to_check.lower()
        
        # TOC detection
        has_toc_keywords = any(kw in lower_text for kw in ["table of contents", "contents", "index"])
        has_dots = bool(re.search(r"\.{4,}|·{4,}", text_to_check))
        
        if has_toc_keywords or has_dots:
            page_type = "toc"
        elif re.search(r"\bappendix\s+[a-c]\b", lower_text[:500]):
            page_type = "appendix"
        elif re.search(r"\breferences\b", lower_text[:500]) and page_num > 10:
            page_type = "references"
        elif re.search(r"\bglossary\b|\bacronyms\b", lower_text[:500]):
            page_type = "glossary"
        elif page_num == 1 and word_count < 300:
            page_type = "cover"
        elif has_tables:
            # Check what proportion of words is inside tables
            table_word_count = 0
            for tbl in tables:
                tb_x0, tb_top, tb_x1, tb_bottom = tbl.bbox
                for w in words:
                    cx = (w["x0"] + w["x1"]) / 2
                    cy = (w["top"] + w["bottom"]) / 2
                    if tb_x0 <= cx <= tb_x1 and tb_top <= cy <= tb_bottom:
                        table_word_count += 1
            
            non_table_words = word_count - table_word_count
            if non_table_words < 50 or (word_count > 0 and non_table_words / word_count < 0.2):
                page_type = "table"
            else:
                page_type = "mixed"
        elif word_count > 0:
            page_type = "prose"
            
    # Detect likely standard type
    likely_standard_type = "unknown"
    text_to_check = ocr_text if ocr_text else raw_text
    lower_check = text_to_check.lower()
    
    if "nist sp" in lower_check or "800-53" in lower_check or "control catalog" in lower_check:
        likely_standard_type = "nist_control_catalog"
    elif "iso/iec" in lower_check or "iso standard" in lower_check or re.search(r"\b\d+\.\d+\.\d+\b", text_to_check):
        likely_standard_type = "iso_numbered_standard"
    elif "article" in lower_check and ("chapter" in lower_check or "regulation" in lower_check):
        likely_standard_type = "legal_articles"
    elif "pci dss" in lower_check or "testing procedure" in lower_check or "assessor guidance" in lower_check:
        likely_standard_type = "pci_requirement_table"
    elif "cis control" in lower_check or "cis safeguard" in lower_check or "cis controls" in lower_check:
        likely_standard_type = "cis_controls"
    elif re.search(r"^\d+\.\d+\s+[A-Z]", text_to_check, re.MULTILINE):
        likely_standard_type = "generic_numbered_guideline"

    # Adjust confidence for tables/mixed pages since cell boundaries can shift
    if page_type in ("table", "mixed") and confidence == 1.0:
        confidence = 0.9
        
    return {
        "page": page_num,
        "page_number": page_num,
        "has_selectable_text": has_selectable_text,
        "text_char_count": char_count,
        "word_count": word_count,
        "is_scanned": is_scanned,
        "orientation": "portrait" if is_portrait else "landscape",
        "is_portrait": is_portrait,
        "has_tables": has_tables,
        "has_rotated_text": has_rotated_text,
        "page_type": page_type,
        "likely_page_type": page_type,
        "likely_standard_type": likely_standard_type,
        "profile_candidate": likely_standard_type,
        "is_multi_column": is_multi_column,
        "confidence": confidence,
        "warnings": warnings,
        "ocr_text": ocr_text,
    }
