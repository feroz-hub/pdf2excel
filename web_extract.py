"""Extract the main content of a public web page into paragraph objects.

The output is the SAME shape :func:`pdf_paragraphs_to_excel.extract_paragraphs`
returns — a list of :class:`Paragraph` objects (``type`` "heading"/"body",
``section``, ``text``) — so the existing mapping/writers in
:mod:`standard_export` are reused unchanged.

Responsible handling: only publicly visible content is processed, with a
descriptive User-Agent; gating (401/403) is reported, not bypassed; no
summarization or reordering — original wording and DOM order are preserved.

Public API:
    fetch_html(url, timeout=30) -> (final_url, content_type, data)
    extract_url(url, prefetched=None) -> ParagraphList   # a list[Paragraph]
                                                          # carrying .title/.final_url
"""

from __future__ import annotations

import logging
import re
import urllib.parse
from types import SimpleNamespace
from typing import List, Optional, Tuple

from pdf_paragraphs_to_excel import (
    Paragraph,
    heading_label,
    is_heading,
    resolve_heading,
)

log = logging.getLogger("pdf2excel.web")

# Below this much extracted text we treat the page as having no main content.
_MIN_CONTENT_CHARS = 200
# Warn if a heading/clause label repeats verbatim more than this many times.
_DUP_WARN_THRESHOLD = 5
# A static extraction yielding less than this looks "incomplete" -> try render.
_INCOMPLETE_CHARS = 500
_INCOMPLETE_BODY_BLOCKS = 3
# trafilatura output thinner than this falls back to a high-recall DOM walk.
_THIN_CHARS = 500
_THIN_BLOCKS = 3

_CHARSET_RE = re.compile(rb"charset=[\"']?\s*([\w\-]+)", re.IGNORECASE)
_WS_RE = re.compile(r"\s+")
_CLAUSE_NUM_RE = re.compile(r"(\d+)")
# Used to tell whether any clause structure is present (render-trigger heuristic).
_LEGALISH_RE = re.compile(r"\b(article|chapter|section|part)\s+\d+", re.IGNORECASE)
# Enumeration markers statutes place in their own HTML element, separate from the
# text they label (so each lands as a stray "marker-only" block to be merged).
_MARKER_RES = (
    re.compile(r"^\(?\d+([.\-]\d+)*\)?[.)]?$"),        # 1.  1-2.  (1)  1)
    re.compile(r"^\([a-z]\)$"),                        # (a) (b) (c)
    re.compile(r"^[①-⑳]$"),                  # circled numbers ①..⑳
    re.compile(r"^[ivxlcdm]+[.)]$", re.IGNORECASE),    # i. ii) iii.
)
# Trailing junk such as " [host]" — either literal or percent-encoded
# ("%20%5B...%5D") — that some sources append to a copied URL.
_TRAILING_JUNK_RE = re.compile(r"(\s+|%20)(\[|%5[bB]).*$")
# Matched wrapping pairs to peel from a pasted URL (e.g. "<url>", "[url]").
_WRAP_PAIRS = {"<": ">", '"': '"', "'": "'", "[": "]"}


class ParagraphList(list):
    """A ``list[Paragraph]`` that also carries page metadata.

    Subclassing ``list`` keeps the documented ``-> list[paragraph]`` contract
    while letting the caller read the discovered ``title`` / ``final_url`` for
    the Standard Assessment metadata block.
    """

    title: str = ""
    final_url: str = ""


# --------------------------------------------------------------------------- #
# URL normalization
# --------------------------------------------------------------------------- #

def clean_url(raw: str) -> str:
    """Normalize a pasted URL before fetching.

    Strips surrounding whitespace and wrapping ``<>``/quotes/``[]``, drops a
    trailing " [host]" (literal or percent-encoded) and anything after the first
    space (a valid URL has none), and validates the scheme/netloc.
    """
    s = (raw or "").strip()
    s = _TRAILING_JUNK_RE.sub("", s)            # drop " [host]" / "%20%5B...%5D"
    parts = s.split()
    s = parts[0] if parts else ""               # URLs contain no whitespace
    while len(s) >= 2 and s[0] in _WRAP_PAIRS and s[-1] == _WRAP_PAIRS[s[0]]:
        s = s[1:-1].strip()

    parsed = urllib.parse.urlparse(s)
    if parsed.scheme not in ("http", "https") or not parsed.netloc:
        raise ValueError(f"Not a valid URL: {raw!r}")
    if s != (raw or "").strip():
        log.info("normalized URL to %s", s)
    return s


# --------------------------------------------------------------------------- #
# Fetch / decode
# --------------------------------------------------------------------------- #

def fetch_html(url: str, timeout: int = 30, insecure: bool = False,
               ca_bundle: Optional[str] = None):
    """Fetch a URL (reusing :func:`download_and_extract.fetch`).

    Returns ``(final_url, content_type, data)``; raises on network error,
    non-2xx, gating (401/403), TLS verification failure or oversize responses.
    """
    from download_and_extract import fetch

    return fetch(clean_url(url), timeout=timeout, insecure=insecure,
                 ca_bundle=ca_bundle)


def fetch_rendered(url: str, timeout: int = 30, insecure: bool = False):
    """Render a JS-heavy page headlessly and return (main_html, [iframe_html, ...]).

    Uses Playwright (lazy import so it stays optional). Raises a clear,
    actionable error when the renderer isn't installed.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Page appears JavaScript-rendered. Install the renderer: "
            "pip install playwright && playwright install chromium  "
            "(or re-run with --render never to skip)."
        ) from None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            # ignore_https_errors mirrors the static fetch's `insecure` workaround.
            context = browser.new_context(ignore_https_errors=insecure)
            page = context.new_page()
            page.goto(url, wait_until="networkidle", timeout=timeout * 1000)
            html = page.content()
            frame_htmls = []
            for frame in page.frames:  # child frames cover iframe bodies
                if frame is page.main_frame:
                    continue
                try:
                    frame_htmls.append(frame.content())
                except Exception:  # noqa: BLE001 - a frame may be detached/cross-origin
                    pass
            return html, frame_htmls
        finally:
            browser.close()


def _decode(data: bytes, content_type: str) -> str:
    """Decode HTML bytes to text, honouring the charset hint where present."""
    charset = None
    m = re.search(r"charset=([\w\-]+)", content_type or "", re.IGNORECASE)
    if m:
        charset = m.group(1)
    if not charset:
        m = _CHARSET_RE.search(data[:2048])
        if m:
            charset = m.group(1).decode("ascii", "ignore")
    for enc in (charset, "utf-8", "cp949", "latin-1"):
        if not enc:
            continue
        try:
            return data.decode(enc)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", "replace")


def _html_title(html: str) -> str:
    try:
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html, "lxml")
        if soup.title and soup.title.string:
            return _clean(soup.title.string)
    except Exception:  # noqa: BLE001 - title is best-effort metadata
        pass
    return ""


def _clean(text: Optional[str]) -> str:
    """Normalise whitespace only — wording is preserved, never paraphrased."""
    if not text:
        return ""
    return _WS_RE.sub(" ", text).strip()


# --------------------------------------------------------------------------- #
# Block extraction (trafilatura, then readability + BeautifulSoup)
# --------------------------------------------------------------------------- #

def _block_text(el) -> str:
    return _clean("".join(el.itertext()))


def _walk_xml(el, out: List[Tuple[str, str]]) -> None:
    """Walk trafilatura XML in document order into (kind, text) blocks."""
    tag = el.tag
    if tag == "head":
        out.append(("heading", _block_text(el)))
    elif tag in ("p", "quote", "item"):
        out.append(("body", _block_text(el)))
    elif tag == "row":  # table row -> join its cells, preserving content/order
        cells = [_block_text(c) for c in el if getattr(c, "tag", "") == "cell"]
        out.append(("body", " | ".join(c for c in cells if c)))
    else:
        for child in el:
            _walk_xml(child, out)


def _extract_blocks_trafilatura(html: str) -> List[Tuple[str, str]]:
    """Primary extractor: trafilatura structured (XML) output."""
    try:
        import trafilatura
        from lxml import etree
    except ImportError:
        return []

    xml = None
    for kwargs in (
        dict(output_format="xml", include_tables=True, include_comments=False,
             favor_recall=True),
        dict(output_format="xml", include_tables=True, include_comments=False),
        dict(output_format="xml"),
    ):
        try:
            xml = trafilatura.extract(html, **kwargs)
        except TypeError:
            continue
        if xml:
            break
    if not xml:
        return []

    try:
        root = etree.fromstring(xml.encode("utf-8"))
    except Exception:  # noqa: BLE001 - fall back if the XML can't be parsed
        return []

    blocks: List[Tuple[str, str]] = []
    _walk_xml(root, blocks)
    return [(k, t) for k, t in blocks if t]


# Block-level tags treated as leaf content (captured whole, not recursed into).
_BLOCK_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6", "p", "li", "td"}


def _walk_dom(el, out: List[Tuple[str, str]]) -> None:
    """Walk a BeautifulSoup tree in document order into (kind, text) blocks.

    Leaf block tags (h1-h6/p/li/td) are captured whole; ``div`` contributes its
    *direct* text (then we recurse for nested blocks); other containers recurse.
    """
    from bs4 import Tag

    for child in getattr(el, "children", []):
        if not isinstance(child, Tag):
            continue
        name = child.name
        if name in _BLOCK_TAGS:
            text = _clean(child.get_text(" ", strip=True))
            if text:
                kind = "heading" if name[0] == "h" and name[1:].isdigit() else "body"
                out.append((kind, text))
        elif name == "div":
            direct = _clean(" ".join(child.find_all(string=True, recursive=False)))
            if direct:
                out.append(("body", direct))
            _walk_dom(child, out)
        else:
            _walk_dom(child, out)


def _extract_blocks_domwalk(html: str) -> List[Tuple[str, str]]:
    """High-recall fallback: a document-order DOM walk over the cleaned tree."""
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "lxml")
    for junk in soup(["script", "style", "nav", "header", "footer", "aside"]):
        junk.decompose()
    out: List[Tuple[str, str]] = []
    _walk_dom(soup.body or soup, out)
    return [(k, t) for k, t in out if t]


def _text_len(blocks: List[Tuple[str, str]]) -> int:
    return sum(len(t) for _, t in blocks)


def extract_blocks(html: str) -> List[Tuple[str, str]]:
    """Extract ordered (kind, text) blocks, highest-recall available.

    trafilatura (favor_recall) is preferred; if it returns little, fall back to
    a DOM walk and keep whichever captured more text.
    """
    blocks = _extract_blocks_trafilatura(html)
    if len(blocks) < _THIN_BLOCKS or _text_len(blocks) < _THIN_CHARS:
        dom = _extract_blocks_domwalk(html)
        if _text_len(dom) > _text_len(blocks):
            blocks = dom
    return blocks


def looks_incomplete(blocks: List[Tuple[str, str]]) -> bool:
    """Heuristic: does a static extraction look like it missed the body?"""
    body = [t for k, t in blocks if k == "body"]
    if sum(len(t) for t in body) < _INCOMPLETE_CHARS:
        return True
    if len(body) < _INCOMPLETE_BODY_BLOCKS:
        return True
    has_structure = any(k == "heading" for k, _ in blocks) or any(
        _LEGALISH_RE.search(t) for _, t in blocks
    )
    return not has_structure


def _is_marker_only(text: str) -> bool:
    """True if a block is nothing but an enumeration marker ("1.", "(a)", ...)."""
    s = text.strip()
    return bool(s) and any(rx.match(s) for rx in _MARKER_RES)


def normalize_blocks(blocks: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    """Merge orphan enumeration markers into the text they label, then drop
    consecutive duplicate blocks.

    Statute pages put each marker ("1.", "1-2.", "(a)", circled numbers, roman
    numerals) in its own element, so the raw list has a bare-marker block before
    the block holding the sub-paragraph text. We prepend the accumulated
    marker(s) to that following block and inherit its type (a marker never
    becomes a row of its own and never flips heading vs body); a marker with no
    following text block is kept as-is. Then consecutive blocks with the same
    type and whitespace-normalized text are de-duplicated (e.g. repeated
    CHAPTER/Article headings), keeping the first occurrence. Wording and order
    are otherwise preserved.
    """
    marker_count = sum(1 for _, t in blocks if _is_marker_only(t))

    # Phase 1 — fold marker-only blocks into the next non-marker block.
    merged: List[Tuple[str, str]] = []
    pending: List[Tuple[str, str]] = []
    for kind, text in blocks:
        if _is_marker_only(text):
            pending.append((kind, text))
            continue
        if pending:
            prefix = " ".join(p.strip() for _, p in pending)
            text = f"{prefix} {text.strip()}".strip()
            pending = []
        merged.append((kind, text))
    merged.extend(pending)  # trailing markers with no target: keep as-is
    merged_markers = marker_count - len(pending)

    # Phase 2 — drop consecutive duplicates (same type + normalized text).
    deduped: List[Tuple[str, str]] = []
    prev_key: Optional[Tuple[str, str]] = None
    for kind, text in merged:
        key = (kind, _WS_RE.sub(" ", text).strip())
        if key == prev_key:
            continue
        deduped.append((kind, text))
        prev_key = key
    removed_dupes = len(merged) - len(deduped)

    if marker_count or removed_dupes:
        log.info("merged %d orphan markers; removed %d duplicate blocks",
                 merged_markers, removed_dupes)
    return deduped


# --------------------------------------------------------------------------- #
# Block -> Paragraph, with heading detection over text
# --------------------------------------------------------------------------- #

def _blocks_to_paragraphs(blocks: List[Tuple[str, str]]) -> List[Paragraph]:
    # Reuse the existing heading detector; "auto" picks legal vs numbered from
    # the text so inline "Article N (Title)" lines are recognised as headings.
    spec = resolve_heading("auto", [[SimpleNamespace(text=t) for _, t in blocks]])

    paragraphs: List[Paragraph] = []
    section = ""
    pid = 0

    def emit_heading(text: str) -> None:
        nonlocal section, pid
        label, trailing = heading_label(text)
        section = label
        pid += 1
        paragraphs.append(
            Paragraph(para_id=pid, page=None, type="heading", section=section, text=label)
        )
        if trailing:  # sentence trailing "Article 1 (Purpose) ..." starts the body
            pid += 1
            paragraphs.append(
                Paragraph(para_id=pid, page=None, type="body", section=section, text=trailing)
            )

    for kind, text in blocks:
        text = _clean(text)
        if not text:
            continue
        if kind == "heading" or is_heading(text, spec):
            emit_heading(text)
            continue
        pid += 1
        paragraphs.append(
            Paragraph(para_id=pid, page=None, type="body", section=section, text=text)
        )
    return paragraphs


# --------------------------------------------------------------------------- #
# Validation (warn; hard-fail only on empty)
# --------------------------------------------------------------------------- #

def _validate(paragraphs: List[Paragraph]) -> None:
    total = sum(len(p.text) for p in paragraphs)
    if not paragraphs or total < _MIN_CONTENT_CHARS:
        raise ValueError(
            "No main content extracted; page may be empty, gated, or "
            "JavaScript-rendered."
        )

    headings = [p.text for p in paragraphs if p.type == "heading"]

    # Duplicate-section warning.
    seen = {}
    for h in headings:
        seen[h] = seen.get(h, 0) + 1
    for label, count in seen.items():
        if count > _DUP_WARN_THRESHOLD:
            log.warning("heading repeats %d times verbatim: %r", count, label[:60])

    # Ordering warning: numeric clause ids should not run backwards (resets per
    # chapter are normal, so we only flag a strict decrease).
    nums = []
    for h in headings:
        m = _CLAUSE_NUM_RE.search(h)
        if m:
            nums.append(int(m.group(1)))
    for prev, cur in zip(nums, nums[1:]):
        if cur < prev:
            log.warning(
                "clause numbering is non-monotonic (%d follows %d); "
                "check document order.", cur, prev,
            )
            break


# --------------------------------------------------------------------------- #
# Public entry point
# --------------------------------------------------------------------------- #

def extract_url(url: str, prefetched=None, insecure: bool = False,
                ca_bundle: Optional[str] = None,
                render: str = "auto") -> ParagraphList:
    """Extract a public page's main content into a :class:`ParagraphList`.

    ``prefetched`` may be a ``(final_url, content_type, data)`` tuple (e.g. the
    bytes the router already fetched while sniffing the type) to avoid a second
    request. ``insecure`` / ``ca_bundle`` control TLS verification on fetch.
    ``render`` is ``"auto"`` (render only when the static result looks
    incomplete), ``"always"`` or ``"never"``.
    """
    if render not in ("auto", "always", "never"):
        raise ValueError(f"unknown render mode: {render!r}")

    if prefetched is not None:
        final_url, content_type, data = prefetched
    else:
        final_url, content_type, data = fetch_html(
            clean_url(url), insecure=insecure, ca_bundle=ca_bundle
        )

    html = _decode(data, content_type)
    title = _html_title(html) or final_url

    blocks = extract_blocks(html)
    log.info("static: %d blocks / %d chars", len(blocks), _text_len(blocks))

    # JS-heavy pages (body injected by script / inside an iframe) need a render.
    if render == "always" or (render == "auto" and looks_incomplete(blocks)):
        rendered_html, frame_htmls = fetch_rendered(
            final_url or url, insecure=insecure
        )
        rblocks = extract_blocks(rendered_html)
        for frame_html in frame_htmls:  # append iframe bodies, in order
            rblocks += extract_blocks(frame_html)
        log.info("rendered: %d blocks / %d chars", len(rblocks), _text_len(rblocks))
        blocks = rblocks
        title = _html_title(rendered_html) or title

    # Merge orphan enumeration markers + drop duplicate blocks (both paths).
    blocks = normalize_blocks(blocks)

    paragraphs = _blocks_to_paragraphs(blocks)
    _validate(paragraphs)

    result = ParagraphList(paragraphs)
    result.title = title
    result.final_url = final_url
    return result
