"""Download a PDF by URL, then extract its paragraphs into Excel.

A thin convenience wrapper around :mod:`pdf_paragraphs_to_excel` for the common
case of "grab this guideline from a URL and give me a spreadsheet".

Usage:
    python download_and_extract.py https://example.org/guideline.pdf
    python download_and_extract.py <url> -o out.xlsx --keep-pdf
"""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen

from pdf_paragraphs_to_excel import extract_paragraphs, write_excel

# A descriptive, honest User-Agent so origins can identify the client.
_USER_AGENT = (
    "pdf2excel/1.0 (public-content extractor; "
    "+https://github.com/pdf2excel) Python-requests"
)
# Cap downloads so a hostile/huge URL can't exhaust memory.
_MAX_BYTES = 25 * 1024 * 1024  # 25 MB


def _filename_from_url(url: str) -> str:
    """Best-effort PDF filename from a URL path, defaulting to download.pdf."""
    path = unquote(urlparse(url).path)
    name = os.path.basename(path) or "download.pdf"
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return name


def download_pdf(url: str, dest_dir: str) -> str:
    """Download ``url`` into ``dest_dir`` and return the saved file path."""
    filename = _filename_from_url(url)
    dest = os.path.join(dest_dir, filename)

    req = Request(url, headers={"User-Agent": _USER_AGENT})
    with urlopen(req) as resp:  # noqa: S310 - user-supplied URL is expected
        content_type = (resp.headers.get("Content-Type") or "").lower()
        data = resp.read()

    if data[:5] != b"%PDF-" and "pdf" not in content_type:
        raise ValueError(
            f"URL did not return a PDF (Content-Type: {content_type!r})."
        )

    with open(dest, "wb") as fh:
        fh.write(data)
    return dest


def _tls_verify(insecure: bool, ca_bundle):
    """Resolve the requests ``verify=`` value, preferring the OS trust store.

    Injecting truststore lets the OS resolve missing intermediate certificates
    (e.g. Windows fetches them via AIA); on Python < 3.10 this is skipped and we
    fall back to certifi's bundle. A user-supplied PEM or ``insecure`` override.
    """
    if insecure:
        try:
            import urllib3

            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        except Exception:  # noqa: BLE001 - best-effort warning suppression
            pass
        return False
    if ca_bundle:
        return ca_bundle

    truststore_ok = False
    try:
        import truststore

        truststore.inject_into_ssl()
        truststore_ok = True
    except Exception:  # noqa: BLE001 - unavailable on <3.10; fall back to certifi
        pass
    if truststore_ok:
        return True  # use the (injected) OS trust store
    try:
        import certifi

        return certifi.where()
    except Exception:  # noqa: BLE001
        return True


def fetch(url: str, timeout: int = 30, insecure: bool = False, ca_bundle=None):
    """GET ``url`` with a descriptive UA, following redirects, size-capped.

    Returns ``(final_url, content_type, data)``. Raises ``PermissionError`` on
    401/403 (gated/auth-required), ``ValueError`` if the body exceeds the size
    cap, ``RuntimeError`` with a clear message on TLS verification failure, and
    propagates other network errors / non-2xx as ``requests`` exceptions.

    TLS is verified by default (OS trust store via truststore, else certifi).
    Pass ``ca_bundle`` (a PEM with the missing intermediate) or ``insecure=True``
    to override.
    """
    import requests  # lazy: keeps the module importable without the web deps

    verify = _tls_verify(insecure, ca_bundle)
    headers = {
        "User-Agent": _USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/pdf,*/*;q=0.8",
    }
    try:
        resp = requests.get(
            url, headers=headers, timeout=timeout, stream=True,
            allow_redirects=True, verify=verify,
        )
    except requests.exceptions.SSLError:
        from urllib.parse import urlparse

        host = urlparse(url).netloc or url
        raise RuntimeError(  # clear, actionable message — no raw traceback
            f"TLS certificate could not be verified for {host} — the server may "
            "not send a complete certificate chain. Fixes: install 'truststore' "
            "(Python 3.10+) and update 'certifi'; or pass --ca-bundle <pem with "
            "the intermediate>; or, for this trusted public source, re-run with "
            "--insecure to skip verification (disables MITM protection)."
        ) from None
    # Do not attempt to bypass gating / auth — report it plainly and stop.
    if resp.status_code in (401, 403):
        resp.close()
        raise PermissionError(
            f"Access denied ({resp.status_code}); the page may require "
            "authentication or be behind a paywall/anti-bot measure."
        )
    resp.raise_for_status()

    chunks, total = [], 0
    for chunk in resp.iter_content(8192):
        total += len(chunk)
        if total > _MAX_BYTES:
            resp.close()
            raise ValueError(f"response exceeds {_MAX_BYTES} byte size cap")
        chunks.append(chunk)
    content_type = (resp.headers.get("Content-Type") or "").lower()
    return resp.url, content_type, b"".join(chunks)


def detect(data: bytes, content_type: str = "") -> str:
    """Sniff content kind from magic bytes (and Content-Type as a hint).

    Returns one of ``"pdf"``, ``"html"``, ``"hwp"``, ``"text"`` or ``"unknown"``.
    """
    head = data[:1024]
    if head[:5] == b"%PDF-":
        return "pdf"
    # HWP 5.x is an OLE compound file; HWPML embeds a signature string.
    if head[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" or b"HWP Document File" in head:
        return "hwp"

    ct = content_type or ""
    if "pdf" in ct:
        return "pdf"
    if "html" in ct or "xml" in ct:
        return "html"

    lowered = head.lower()
    if b"<html" in lowered or b"<!doctype html" in lowered or b"<?xml" in lowered:
        return "html"
    if ct.startswith("text/"):
        return "text"
    # Treat mostly-printable payloads as text; otherwise unknown (e.g. binaries).
    sample = head[:512]
    if sample and sum(b in b"\t\r\n" or 32 <= b < 127 for b in sample) / len(sample) > 0.9:
        return "text"
    return "unknown"


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Download a PDF by URL and extract its paragraphs to Excel."
    )
    parser.add_argument("url", help="URL of the PDF to download.")
    parser.add_argument("-o", "--output", help="Output .xlsx path.")
    parser.add_argument(
        "-g",
        "--gap-factor",
        type=float,
        default=1.6,
        help="Paragraph-break sensitivity (default: 1.6).",
    )
    parser.add_argument(
        "--keep-pdf",
        action="store_true",
        help="Keep the downloaded PDF (saved next to the output).",
    )
    args = parser.parse_args(argv)

    work_dir = os.getcwd() if args.keep_pdf else tempfile.mkdtemp(prefix="pdf2excel_")

    try:
        print(f"Downloading {args.url} ...")
        pdf_path = download_pdf(args.url, work_dir)
    except Exception as exc:  # noqa: BLE001
        print(f"error: download failed: {exc}", file=sys.stderr)
        return 1

    base = os.path.splitext(os.path.basename(pdf_path))[0]
    out_path = args.output or f"{base}.xlsx"

    try:
        paras = extract_paragraphs(pdf_path, gap_factor=args.gap_factor)
    except Exception as exc:  # noqa: BLE001
        print(f"error: extraction failed: {exc}", file=sys.stderr)
        return 1

    write_excel(paras, out_path)
    print(f"Extracted {len(paras)} paragraphs -> {out_path}")
    if args.keep_pdf:
        print(f"Kept PDF -> {pdf_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
