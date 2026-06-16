"""Automatic routing between the prose and table extractors.

Looks at a PDF's first few pages and decides whether it is a long-form text
document (laws, guidelines) or a slide-deck / table-heavy file, then dispatches
to the matching extractor:

  * prose  -> pdf_paragraphs_to_excel.extract_paragraphs + write_excel
  * tables -> deck_tables_to_excel (one sheet per table + slides_text)

The output can be the project's default workbook or the "Standard Assessment"
format (``fmt="standard"``), which is populated from the bundled template.

Public API:
    detect_kind(pdf_path) -> "prose" | "tables"
    convert(pdf_path, out_path, mode="auto", fmt="default", ...) -> ConvertResult

CLI:
    python router.py input.pdf -o out.xlsx [--mode auto|prose|tables]
                                           [--format default|standard --standard-id ...]
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from typing import List, Tuple

try:
    import pdfplumber
except ImportError:  # pragma: no cover - dependency hint
    pdfplumber = None

from deck_tables_to_excel import (
    _extract_page_tables,
    _merge_side_by_side,
    extract_deck,
    write_deck_excel,
)
from pdf_paragraphs_to_excel import Paragraph, extract_paragraphs, write_excel
from standard_export import (
    DEFAULT_TEMPLATE,
    deck_to_items,
    paragraphs_to_items,
    write_standard_assessment,
    blocks_to_items,
)
import preflight
import extract_blocks
import validate_extraction
from nist80053_extractor import (
    is_nist_800_53,
    extract_nist_800_53_items,
    write_nist_review_workbook,
)
import layout_blocks
import standard_patterns
import standard_profiles
import structure_parser
import validation
import review_export

# Number of leading pages sampled when sniffing the document kind.
_SAMPLE_PAGES = 6


@dataclass
class ConvertResult:
    """Outcome of a conversion, with enough detail to build a GUI preview."""

    mode: str                                   # "prose" or "tables"
    out_path: str
    fmt: str = "default"                        # "default" or "standard"
    n_items: int = 0                            # rows written in standard format
    paragraphs: List[Paragraph] = field(default_factory=list)
    # For tables mode: (sheet_name, n_rows, n_cols) per worksheet written.
    sheets: List[Tuple[str, int, int]] = field(default_factory=list)
    enriched: bool = False                      # AI enrichment (cols E–I) ran
    n_requirements: int = 0                     # rows classified as "Requirement"
    n_information: int = 0                      # rows classified as "Information"
    items: List[dict] = field(default_factory=list)
    meta: dict = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    has_scanned_pages: bool = False
    profile: str = ""
    review_output_path: str | None = None
    quality_gate_passed: bool = True
    quality_gate_failures: List[str] = field(default_factory=list)
    issues_count: int = 0
    rejected_count: int = 0


class QualityGateError(ValueError):
    """Raised when export quality gate fails (unless suppressed by caller)."""

    def __init__(
        self,
        message: str,
        failures: List[str],
        result: ConvertResult | None = None,
    ):
        super().__init__(message)
        self.failures = failures
        self.result = result


def detect_kind(pdf_path: str) -> str:
    """Classify a PDF as ``"prose"`` or ``"tables"`` from its first pages.

    Decks tend to be landscape and/or table-heavy; documents are portrait prose.
    Returns ``"tables"`` when more than half the sampled pages are landscape OR
    at least half contain a real table; otherwise ``"prose"``.
    """
    if pdfplumber is None:
        raise RuntimeError(
            "pdfplumber is not installed. Run: pip install -r requirements.txt"
        )

    landscape = 0
    table_pages = 0
    total_chars = 0
    n = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[:_SAMPLE_PAGES]:
            n += 1
            if page.width > page.height:
                landscape += 1
            tables = _merge_side_by_side(_extract_page_tables(page))
            if tables:
                table_pages += 1
            total_chars += len(page.extract_text() or "")

    if n == 0:
        return "prose"

    landscape_fraction = landscape / n
    table_page_fraction = table_pages / n
    _mean_chars = total_chars / n  # noqa: F841 - informative, not part of rule

    if landscape_fraction > 0.5 or table_page_fraction >= 0.5:
        return "tables"
    return "prose"


def check_quality_gate(
    exported_items: List[dict],
    profile_name: str,
) -> List[str]:
    """Verify quality gate checks for NIST/control catalog standards.

    Returns a list of failure descriptions.
    """
    failures = []
    if profile_name not in ("nist80053", "control-catalog"):
        return failures

    if not exported_items:
        failures.append("No items were exported.")
        return failures

    # 1. First exported row is not AC-1
    first_id = exported_items[0].get("clause_id", "")
    if first_id != "AC-1":
        failures.append(f"First exported row is '{first_id}', expected 'AC-1'.")

    # 2. More than 2% rows have blank clause ID
    blank_ids = sum(1 for it in exported_items if not it.get("clause_id"))
    pct_blank = blank_ids / len(exported_items)
    if pct_blank > 0.02:
        failures.append(f"More than 2% of rows ({pct_blank:.1%}) have blank clause ID.")

    dup_ids = [cid for cid, n in Counter(
        it.get("clause_id") for it in exported_items if it.get("clause_id")
    ).items() if n > 1]
    if dup_ids:
        failures.append(f"Duplicate clause IDs in export: {', '.join(dup_ids[:5])}.")

    # 3. Any text contains page header/footer/DOI leaks
    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        if "NIST SP 800-53" in text.upper() and ("REV" in text.upper() or "R EV" in text.upper() or "R E V" in text.upper()):
            failures.append(f"Row {clause_id} contains page header text 'NIST SP 800-53, REV. 5'.")
            break

    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        if "This publication is available free of charge" in text:
            failures.append(f"Row {clause_id} contains DOI footer text.")
            break

    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        if "doi.org/10.6028/NIST.SP.800-53r5" in text:
            failures.append(f"Row {clause_id} contains DOI URL.")
            break

    for it in exported_items:
        title = it.get("title", "")
        clause_id = it.get("clause_id", "")
        if "Errata" in title:
            failures.append(f"Row {clause_id} title contains 'Errata'.")
            break

    for it in exported_items:
        title = it.get("title", "")
        clause_id = it.get("clause_id", "")
        if "APPENDIX C PAGE" in title.upper():
            failures.append(f"Row {clause_id} title contains 'APPENDIX C PAGE'.")
            break

    for it in exported_items:
        title = it.get("title", "")
        clause_id = it.get("clause_id", "")
        if title == "NUMBER":
            failures.append(f"Row {clause_id} has title 'NUMBER'.")
            break

    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        if text.strip() == "NUMBER":
            failures.append(f"Row {clause_id} has text exactly 'NUMBER'.")
            break

    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        is_withdrawn = "[withdrawn" in text.lower() or it.get("classification") == "Information"
        if len(text) < 10 and not is_withdrawn:
            failures.append(f"Row {clause_id} has text length less than 10 ({len(text)} chars) and is not withdrawn.")
            break

    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        if "CHAPTER THREE PAGE" in text.upper():
            failures.append(f"Row {clause_id} contains 'CHAPTER THREE PAGE'.")
            break

    for it in exported_items:
        text = it.get("text", "")
        clause_id = it.get("clause_id", "")
        if "APPENDIX C PAGE" in text.upper():
            failures.append(f"Row {clause_id} contains 'APPENDIX C PAGE'.")
            break

    return failures


def _summarize_extraction(
    items: List[dict],
    exported_items: List[dict],
    profile_name: str,
    review_output: str | None,
    gate_failures: List[str],
    warnings: List[str],
) -> dict:
    """Build GUI-friendly summary counters."""
    rejected = sum(1 for it in items if it.get("export_status") == "rejected")
    issues_count = sum(1 for it in items if it.get("issues"))
    n_req = sum(1 for it in exported_items if it.get("classification") == "Requirement")
    n_info = sum(1 for it in exported_items if it.get("classification") == "Information")
    return {
        "profile": profile_name,
        "review_output": review_output,
        "quality_gate_passed": not gate_failures,
        "quality_gate_failures": list(gate_failures),
        "issues_count": issues_count,
        "rejected_count": rejected,
        "warnings_count": len(warnings),
        "n_requirements": n_req,
        "n_information": n_info,
    }


def _run_structured_standard_pipeline(
    *,
    pdf_path: str,
    out_path: str,
    mode: str,
    profile: str,
    gap_factor: float,
    ocr_mode: str,
    skip_cover: bool,
    include_toc: bool,
    include_front_matter: bool,
    include_appendix: bool,
    include_references: bool,
    skip_references: bool,
    min_confidence: float,
    export_low_confidence: bool,
    review_output: str | None,
    force_export: bool,
    raise_on_quality_gate: bool,
    enrich_config,
    progress,
    _write_standard,
    _maybe_enrich,
    _meta,
) -> ConvertResult:
    """Structured standards pipeline for ``fmt=standard`` (NIST + generic profiles)."""
    page_preflights = []
    has_scanned = False
    warnings = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            meta = preflight.analyze_page(page, page_num, ocr_mode=ocr_mode)
            page_preflights.append(meta)
            if meta["is_scanned"]:
                has_scanned = True
            if meta["warnings"]:
                warnings.extend(meta["warnings"])

    if profile == "auto":
        profile_name = standard_patterns.detect_profile_from_preflights(page_preflights)
        if profile_name == "generic" and is_nist_800_53(pdf_path):
            profile_name = "nist80053"
    else:
        profile_name = profile

    prof_obj = standard_profiles.resolve_profile(profile_name)

    if profile_name == "nist80053" or (
        profile_name == "control-catalog" and is_nist_800_53(pdf_path)
    ):
        items = extract_nist_800_53_items(pdf_path)
        blocks = []
    else:
        blocks = layout_blocks.extract_pdf_blocks(
            pdf_path=pdf_path,
            page_preflights=page_preflights,
            gap_factor=gap_factor,
            ocr_mode=ocr_mode,
            skip_cover=skip_cover,
            include_toc=include_toc,
            include_appendix=include_appendix,
            include_references=include_references or (not skip_references),
        )
        items, _rejected = structure_parser.parse_blocks_to_items(
            blocks=blocks,
            profile=prof_obj,
            include_front_matter=include_front_matter,
            include_appendix=include_appendix,
            include_references=include_references or (not skip_references),
        )

    items = validation.validate_items(items, profile_name=profile_name)
    for it in items:
        if it.get("issues"):
            for iss in it["issues"]:
                msg = f"Page {it.get('page') or it.get('source_page')}: {iss}"
                if msg not in warnings:
                    warnings.append(msg)

    exported_items = validation.filter_exportable_items(
        items,
        profile_name=profile_name,
        min_confidence=min_confidence,
        export_low_confidence=export_low_confidence,
    )

    gate_failures = []
    if profile_name in ("nist80053", "control-catalog"):
        gate_failures = check_quality_gate(exported_items, profile_name)

    if gate_failures and not force_export:
        final_export_items = []
    else:
        final_export_items = exported_items

    if review_output:
        review_export.write_review_workbook(items, review_output, profile_name=profile_name)

    paras = [Paragraph(text=b["text"], type=b["block_type"], page=b["page"]) for b in blocks]

    final_export_items, did_enrich, n_req = _maybe_enrich(final_export_items)
    _write_standard(items, export_items=final_export_items)

    summary = _summarize_extraction(
        items, final_export_items, profile_name, review_output, gate_failures, warnings,
    )
    base_meta = _meta()
    base_meta.update(summary)

    result = ConvertResult(
        mode="structured" if mode in ("auto", "standard") else mode,
        out_path=out_path,
        fmt="standard",
        n_items=len(final_export_items),
        paragraphs=paras,
        enriched=did_enrich,
        n_requirements=n_req,
        n_information=summary["n_information"],
        items=final_export_items,
        meta=base_meta,
        warnings=warnings,
        has_scanned_pages=has_scanned,
        profile=profile_name,
        review_output_path=review_output,
        quality_gate_passed=not gate_failures,
        quality_gate_failures=gate_failures,
        issues_count=summary["issues_count"],
        rejected_count=summary["rejected_count"],
    )

    if gate_failures and not force_export:
        print("LOUD WARNING: Quality gate failed!", file=sys.stderr)
        for failure in gate_failures:
            print(f"  - {failure}", file=sys.stderr)
        msg = (
            "Quality gate failed! A blank Standard Assessment was written "
            "with issues in Extraction_Issues. Use --force-export to override."
        )
        if raise_on_quality_gate:
            raise QualityGateError(msg, gate_failures, result=result)
    return result


def convert(
    source: str,
    out_path: str,
    mode: str = "auto",
    fmt: str = "default",
    heading_style: str = "auto",
    gap_factor: float = 1.6,
    standard_id: str = "MLSR",
    standard_title: str = "",
    standard_edition: str = "",
    document_id: str = "",
    document_name: str = "",
    document_revision: str = "",
    template_path: str = DEFAULT_TEMPLATE,
    insecure: bool = False,
    ca_bundle=None,
    render: str = "auto",
    enrich_config=None,
    progress=None,
    review_output: str | None = None,
    include_toc: bool = False,
    skip_cover: bool = True,
    ocr_mode: str = "off",
    min_confidence: float = 0.0,
    show_issues: bool = False,
    profile: str = "auto",
    include_front_matter: bool = False,
    include_appendix: bool = False,
    skip_references: bool = False,
    force_export: bool = False,
    export_low_confidence: bool = False,
    include_references: bool = False,
    raise_on_quality_gate: bool = True,
) -> ConvertResult:
    """Convert ``source`` to ``out_path``, routing by ``mode`` and ``fmt``.

    When ``fmt`` is ``"standard"``, always uses the structured standards pipeline
    (never raw paragraph/block export).
    """
    if fmt not in ("default", "standard"):
        raise ValueError(f"unknown format: {fmt!r}")

    # Normalize routing keys (GUI labels / mixed case must not bypass structured flow).
    fmt = (fmt or "default").strip().lower()
    mode = (mode or "auto").strip().lower()
    profile = (profile or "auto").strip().lower()

    if fmt not in ("default", "standard"):
        raise ValueError(f"unknown format: {fmt!r}")

    if include_front_matter:
        skip_cover = False

    # Resolve a URL into either pre-extracted web paragraphs or a temp PDF path.
    web_paras = None
    tmp_pdf = None
    pdf_path = source
    if "://" in source:
        from download_and_extract import detect, fetch
        from web_extract import clean_url, extract_url

        source = clean_url(source)  # normalize before any fetch
        final_url, content_type, data = fetch(
            source, insecure=insecure, ca_bundle=ca_bundle
        )
        kind = detect(data, content_type)
        if kind == "pdf":
            fd, tmp_pdf = tempfile.mkstemp(suffix=".pdf", prefix="pdf2excel_")
            os.write(fd, data)
            os.close(fd)
            pdf_path = tmp_pdf
        elif kind in ("html", "text"):
            web_paras = extract_url(
                source, prefetched=(final_url, content_type, data),
                insecure=insecure, ca_bundle=ca_bundle, render=render,
            )
            # Carry the page's identity into the metadata block (unless supplied).
            document_id = document_id or final_url
            document_name = document_name or web_paras.title
            standard_title = standard_title or web_paras.title
        else:
            raise ValueError(f"Unsupported content type: {kind!r}")

    def _write_standard(items, export_items=None) -> None:
        write_standard_assessment(
            items,
            out_path,
            template_path=template_path,
            standard_id=standard_id,
            standard_title=standard_title,
            standard_edition=standard_edition,
            document_id=document_id,
            document_name=document_name,
            document_revision=document_revision,
            export_items=export_items,
            show_issues=show_issues,
        )


    def _maybe_enrich(items):
        """Run AI enrichment (cols E–I) when an enrich config was supplied.

        Returns ``(items, enriched_flag, n_requirements)``. ``ai_enrich`` is
        imported lazily so the router stays importable without the AI deps.
        """
        if enrich_config is None:
            return items, False, 0
        from ai_enrich import enrich as _run_enrich

        enriched_items = _run_enrich(items, enrich_config, progress)
        n_req = sum(
            1 for it in enriched_items if it.get("classification") == "Requirement"
        )
        return enriched_items, True, n_req

    def _meta() -> dict:
        """Resolved Standard Assessment metadata (for a later enrichment pass)."""
        return dict(
            standard_id=standard_id, standard_title=standard_title,
            standard_edition=standard_edition, document_id=document_id,
            document_name=document_name, document_revision=document_revision,
            template_path=template_path,
        )

    try:
        # HTML/text URL: paragraphs already extracted; treat as prose-equivalent
        # and reuse the existing writers unchanged.
        if web_paras is not None:
            paras = list(web_paras)
            if fmt == "standard":
                items = paragraphs_to_items(paras)
                items, did_enrich, n_req = _maybe_enrich(items)
                _write_standard(items)
                return ConvertResult(
                    mode="prose", out_path=out_path, fmt="standard",
                    n_items=len(items), paragraphs=paras,
                    enriched=did_enrich, n_requirements=n_req,
                    items=items, meta=_meta(),
                )
            write_excel(paras, out_path)
            return ConvertResult(mode="prose", out_path=out_path, paragraphs=paras)

        # PDF (local file or downloaded URL)

        # --- Standard Assessment: ALWAYS structured pipeline (never prose dump) ---
        if fmt == "standard":
            return _run_structured_standard_pipeline(
                pdf_path=pdf_path,
                out_path=out_path,
                mode=mode,
                profile=profile,
                gap_factor=gap_factor,
                ocr_mode=ocr_mode,
                skip_cover=skip_cover,
                include_toc=include_toc,
                include_front_matter=include_front_matter,
                include_appendix=include_appendix,
                include_references=include_references,
                skip_references=skip_references,
                min_confidence=min_confidence,
                export_low_confidence=export_low_confidence,
                review_output=review_output,
                force_export=force_export,
                raise_on_quality_gate=raise_on_quality_gate,
                enrich_config=enrich_config,
                progress=progress,
                _write_standard=_write_standard,
                _maybe_enrich=_maybe_enrich,
                _meta=_meta,
            )

        # --- NIST SP 800-53 explicit mode (default workbook format only) ---
        if mode == "nist80053":
            items = extract_nist_800_53_items(pdf_path)

            # Filter by min confidence
            if min_confidence > 0.0:
                items = [it for it in items
                         if it.get("confidence", 1.0) >= min_confidence]

            # Optional review workbook
            if review_output:
                write_nist_review_workbook(items, review_output)

            warnings = []
            for it in items:
                if it.get("issues"):
                    for iss in it["issues"]:
                        msg = f"Page {it.get('source_page')}: {iss}"
                        if msg not in warnings:
                            warnings.append(msg)

            if fmt == "standard":
                items, did_enrich, n_req = _maybe_enrich(items)
                _write_standard(items)
                return ConvertResult(
                    mode="nist80053",
                    out_path=out_path,
                    fmt="standard",
                    n_items=len(items),
                    enriched=did_enrich,
                    n_requirements=n_req,
                    items=items,
                    meta=_meta(),
                    warnings=warnings,
                )
            else:
                from openpyxl import Workbook as _Wb
                wb = _Wb()
                ws = wb.active
                ws.title = "NIST 800-53 Controls"
                ws.append(["Clause ID", "Title", "Classification",
                           "Source Page", "Text"])
                for it in items:
                    ws.append([
                        it.get("clause_id", ""),
                        it.get("title", ""),
                        it.get("classification", ""),
                        it.get("source_page", ""),
                        it.get("text", ""),
                    ])
                wb.save(out_path)
                return ConvertResult(
                    mode="nist80053",
                    out_path=out_path,
                    n_items=len(items),
                    items=items,
                    warnings=warnings,
                )

        resolved_mode = detect_kind(pdf_path) if mode == "auto" else mode
        if resolved_mode not in ("prose", "tables"):
            raise ValueError(f"unknown mode: {resolved_mode!r}")

        if resolved_mode == "prose":
            if fmt == "standard":
                raise ValueError(
                    "Internal error: fmt=standard must use structured pipeline, not prose extraction."
                )
            paras = extract_paragraphs(
                pdf_path, gap_factor=gap_factor, heading_style=heading_style
            )
            write_excel(paras, out_path)
            return ConvertResult(mode="prose", out_path=out_path, paragraphs=paras)

        # tables
        pages_tables, slides_text = extract_deck(pdf_path)

        sheets: List[Tuple[str, int, int]] = []
        for page_no, tables in pages_tables:
            for idx, table in enumerate(tables):
                name = f"slide{page_no}" if idx == 0 else f"slide{page_no}_{idx + 1}"
                sheets.append((name, len(table.rows), table.width))
        sheets.append(("slides_text", len(slides_text), 2))

        if fmt == "standard":
            raise ValueError(
                "Standard Assessment format must use the structured extraction pipeline."
            )
        write_deck_excel(pages_tables, slides_text, out_path)
        return ConvertResult(mode="tables", out_path=out_path, sheets=sheets)
    finally:
        if tmp_pdf and os.path.exists(tmp_pdf):
            os.remove(tmp_pdf)

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Auto-route a PDF file or URL to the right extractor."
    )
    parser.add_argument("source", help="Path to a PDF file, or a URL.")
    parser.add_argument("-o", "--output", help="Output .xlsx path.")
    parser.add_argument(
        "--mode",
        choices=["auto", "prose", "tables", "nist80053", "standard"],
        default="auto",
        help="Force a mode, or 'auto' to detect (default: auto).",
    )
    parser.add_argument(
        "--profile",
        choices=["auto", "generic", "nist80053", "iso", "legal", "pci", "cis", "control-catalog"],
        default="auto",
        help="Force a standard profile (default: auto)."
    )
    parser.add_argument(
        "--heading-style",
        choices=["auto", "numbered", "legal"],
        default="auto",
        help="Heading style for prose mode (default: auto).",
    )
    parser.add_argument("--include-front-matter", action="store_true", help="Include cover/front-matter.")
    parser.add_argument("--include-appendix", action="store_true", help="Include appendix pages.")
    parser.add_argument("--no-include-appendix", dest="include_appendix", action="store_false", help="Skip appendix pages (default).")
    parser.set_defaults(include_appendix=False)
    parser.add_argument("--skip-references", action="store_true", help="Skip reference pages.")
    parser.add_argument(
        "--format",
        dest="fmt",
        choices=["default", "standard"],
        default="default",
        help="Output format: default workbook or Standard Assessment template.",
    )
    parser.add_argument("--standard-id", default="MLSR", help="Standard ID (standard format).")
    parser.add_argument("--standard-title", default="", help="Standard title.")
    parser.add_argument("--standard-edition", default="", help="Standard edition.")
    parser.add_argument("--document-id", default="", help="Document ID metadata.")
    parser.add_argument("--document-name", default="", help="Document name metadata.")
    parser.add_argument("--document-revision", default="", help="Document revision metadata.")
    parser.add_argument(
        "--template",
        default=DEFAULT_TEMPLATE,
        help="Path to the Standard Assessment template .xlsx.",
    )
    parser.add_argument(
        "--insecure",
        action="store_true",
        help="Skip TLS certificate verification for URL fetches "
        "(disables MITM protection; for trusted public sources only).",
    )
    parser.add_argument(
        "--ca-bundle",
        default=None,
        help="Path to a PEM CA bundle (e.g. with the missing intermediate) "
        "used to verify URL fetches.",
    )
    parser.add_argument(
        "--render",
        choices=["auto", "always", "never"],
        default="auto",
        help="Headless-render JS pages: auto (when static looks incomplete), "
        "always, or never (default: auto).",
    )
    parser.add_argument("--review-output", default=None, help="Path to write the block-level review workbook.")
    parser.add_argument("--include-toc", action="store_true", help="Include TOC pages.")
    parser.add_argument("--no-skip-cover", dest="skip_cover", action="store_false", default=True, help="Do not skip cover pages.")
    parser.add_argument("--ocr-mode", choices=["off", "detect"], default="off", help="OCR mode.")
    parser.add_argument("--min-confidence", type=float, default=0.0, help="Minimum confidence threshold.")
    parser.add_argument("--show-issues", action="store_true", help="Show extraction issues.")
    parser.add_argument("--force-export", action="store_true", help="Force export despite quality gate or validation failures.")
    parser.add_argument("--export-low-confidence", action="store_true", help="Export low-confidence items.")
    parser.add_argument("--include-references", action="store_true", help="Include reference/bibliography pages.")

    # --- AI enrichment (fills Standard Assessment cols E–I; needs --format standard) ---
    parser.add_argument(
        "--ai-fill",
        action="store_true",
        help="After extraction, call an LLM to classify each clause and fill "
        "columns E–I (requires --format standard).",
    )
    parser.add_argument(
        "--ai-provider",
        choices=["claude", "openai", "gemini", "ollama"],
        default="claude",
        help="LLM provider for --ai-fill (default: claude; 'ollama' = local).",
    )
    parser.add_argument(
        "--ai-model",
        default="",
        help="Model id for --ai-fill (default: the provider's default model).",
    )
    parser.add_argument(
        "--ai-batch-size", type=int, default=12,
        help="Clauses per AI request (default: 12).",
    )
    parser.add_argument(
        "--ai-workers", type=int, default=4,
        help="Parallel AI requests (default: 4).",
    )
    parser.add_argument(
        "--ai-temperature", type=float, default=0.0,
        help="Sampling temperature for --ai-fill (ignored by models that reject it).",
    )
    parser.add_argument(
        "--ai-dry-run",
        action="store_true",
        help="Run the AI phase with no API calls (placeholders) — for testing.",
    )
    args = parser.parse_args(argv)

    if args.output:
        out_path = args.output
    else:
        # For a URL the basename is often unhelpful (e.g. "viewer.do"); fall back.
        base = os.path.splitext(os.path.basename(args.source))[0]
        out_path = (base or "output") + ".xlsx"

    # Build the AI-enrichment config when --ai-fill is set (standard format only).
    enrich_config = None
    if args.ai_fill:
        if args.fmt != "standard":
            print("warning: --ai-fill requires --format standard; "
                  "skipping AI enrichment.", file=sys.stderr)
        else:
            from ai_enrich import EnrichConfig

            enrich_config = EnrichConfig(
                provider=args.ai_provider, model=args.ai_model,
                temperature=args.ai_temperature, batch_size=args.ai_batch_size,
                workers=args.ai_workers, dry_run=args.ai_dry_run,
            )

    def _progress(done, total, msg):
        print(f"\r[ai] {done}/{total} {msg}".ljust(60), end="", file=sys.stderr)
        if done >= total:
            print(file=sys.stderr)

    try:
        result = convert(
            args.source, out_path,
            mode=args.mode, fmt=args.fmt, heading_style=args.heading_style,
            standard_id=args.standard_id, standard_title=args.standard_title,
            standard_edition=args.standard_edition, document_id=args.document_id,
            document_name=args.document_name, document_revision=args.document_revision,
            template_path=args.template,
            insecure=args.insecure, ca_bundle=args.ca_bundle, render=args.render,
            enrich_config=enrich_config,
            progress=_progress if enrich_config is not None else None,
            review_output=args.review_output,
            include_toc=args.include_toc,
            skip_cover=args.skip_cover,
            ocr_mode=args.ocr_mode,
            min_confidence=args.min_confidence,
            show_issues=args.show_issues,
            profile=args.profile,
            include_front_matter=args.include_front_matter,
            include_appendix=args.include_appendix,
            skip_references=args.skip_references,
            force_export=args.force_export,
            export_low_confidence=args.export_low_confidence,
            include_references=args.include_references,
        )
    except FileNotFoundError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"error: conversion failed: {exc}", file=sys.stderr)
        return 1

    if result.fmt == "standard":
        ai = " (AI-filled)" if result.enriched else ""
        reqs = f", {result.n_requirements} requirements" if result.enriched else ""
        print(f"[{result.mode}->standard{ai}] {result.n_items} rows"
              f"{reqs} -> {out_path}")
        if args.show_issues:
            for it in result.items:
                if it.get("issues"):
                    print(f"  Row {it.get('clause_id') or 'unknown'}: {', '.join(it['issues'])}")
    elif result.mode == "prose":
        headings = sum(1 for p in result.paragraphs if p.type == "heading")
        print(
            f"[prose] {len(result.paragraphs)} paragraphs "
            f"({headings} headings) -> {out_path}"
        )
    elif result.mode == "nist80053":
        reqs = sum(1 for it in result.items if it.get("classification") == "Requirement")
        infos = sum(1 for it in result.items if it.get("classification") == "Information")
        print(f"[nist80053] {result.n_items} controls/enhancements "
              f"({reqs} requirements, {infos} withdrawn/info) -> {out_path}")
        if args.show_issues:
            for it in result.items:
                if it.get("issues"):
                    print(f"  {it.get('clause_id') or 'unknown'}: {', '.join(it['issues'])}")
    elif result.mode in ("auto", "standard"):
        print(f"[{result.mode}] {result.n_items} rows -> {out_path}")
        if args.show_issues:
            for it in result.items:
                if it.get("issues"):
                    print(f"  Row {it.get('clause_id') or 'unknown'}: {', '.join(it['issues'])}")
    else:
        n_tables = max(len(result.sheets) - 1, 0)  # minus slides_text
        print(f"[tables] {n_tables} table sheets + slides_text -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
