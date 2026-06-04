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
)

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
) -> ConvertResult:
    """Convert ``source`` to ``out_path``, routing by ``mode`` and ``fmt``.

    ``source`` is a local PDF path or a URL. URLs are sniffed: PDFs go through
    the existing PDF pipeline; HTML/text pages go through
    :func:`web_extract.extract_url`. ``mode`` is ``"auto"`` (sniff), ``"prose"``
    or ``"tables"``; ``fmt`` is ``"default"`` or ``"standard"``. ``gap_factor`` /
    ``heading_style`` apply only to PDF prose; the ``standard_*`` / ``document_*``
    args only to standard format. ``insecure`` / ``ca_bundle`` control TLS
    verification, and ``render`` ("auto"/"always"/"never") the headless-render
    fallback, when ``source`` is a URL.
    """
    if fmt not in ("default", "standard"):
        raise ValueError(f"unknown format: {fmt!r}")

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

    def _write_standard(items) -> None:
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
        )

    try:
        # HTML/text URL: paragraphs already extracted; treat as prose-equivalent
        # and reuse the existing writers unchanged.
        if web_paras is not None:
            paras = list(web_paras)
            if fmt == "standard":
                items = paragraphs_to_items(paras)
                _write_standard(items)
                return ConvertResult(
                    mode="prose", out_path=out_path, fmt="standard",
                    n_items=len(items), paragraphs=paras,
                )
            write_excel(paras, out_path)
            return ConvertResult(mode="prose", out_path=out_path, paragraphs=paras)

        # PDF (local file or downloaded URL) — behaviour unchanged.
        resolved_mode = detect_kind(pdf_path) if mode == "auto" else mode
        if resolved_mode not in ("prose", "tables"):
            raise ValueError(f"unknown mode: {resolved_mode!r}")

        if resolved_mode == "prose":
            paras = extract_paragraphs(
                pdf_path, gap_factor=gap_factor, heading_style=heading_style
            )
            if fmt == "standard":
                items = paragraphs_to_items(paras)
                _write_standard(items)
                return ConvertResult(
                    mode="prose", out_path=out_path, fmt="standard",
                    n_items=len(items), paragraphs=paras,
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
            items = deck_to_items(pages_tables, slides_text)
            _write_standard(items)
            return ConvertResult(
                mode="tables", out_path=out_path, fmt="standard",
                n_items=len(items), sheets=sheets,
            )

        write_deck_excel(pages_tables, slides_text, out_path)
        return ConvertResult(mode="tables", out_path=out_path, sheets=sheets)
    finally:
        if tmp_pdf and os.path.exists(tmp_pdf):
            os.remove(tmp_pdf)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Auto-route a PDF file or URL to the right extractor."
    )
    parser.add_argument("source", help="Path to a PDF file, or a URL.")
    parser.add_argument("-o", "--output", help="Output .xlsx path.")
    parser.add_argument(
        "--mode",
        choices=["auto", "prose", "tables"],
        default="auto",
        help="Force a mode, or 'auto' to detect (default: auto).",
    )
    parser.add_argument(
        "--heading-style",
        choices=["auto", "numbered", "legal"],
        default="auto",
        help="Heading style for prose mode (default: auto).",
    )
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
    args = parser.parse_args(argv)

    if args.output:
        out_path = args.output
    else:
        # For a URL the basename is often unhelpful (e.g. "viewer.do"); fall back.
        base = os.path.splitext(os.path.basename(args.source))[0]
        out_path = (base or "output") + ".xlsx"

    try:
        result = convert(
            args.source, out_path,
            mode=args.mode, fmt=args.fmt, heading_style=args.heading_style,
            standard_id=args.standard_id, standard_title=args.standard_title,
            standard_edition=args.standard_edition, document_id=args.document_id,
            document_name=args.document_name, document_revision=args.document_revision,
            template_path=args.template,
            insecure=args.insecure, ca_bundle=args.ca_bundle, render=args.render,
        )
    except FileNotFoundError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"error: conversion failed: {exc}", file=sys.stderr)
        return 1

    if result.fmt == "standard":
        print(f"[{result.mode}->standard] {result.n_items} rows -> {out_path}")
    elif result.mode == "prose":
        headings = sum(1 for p in result.paragraphs if p.type == "heading")
        print(
            f"[prose] {len(result.paragraphs)} paragraphs "
            f"({headings} headings) -> {out_path}"
        )
    else:
        n_tables = max(len(result.sheets) - 1, 0)  # minus slides_text
        print(f"[tables] {n_tables} table sheets + slides_text -> {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
