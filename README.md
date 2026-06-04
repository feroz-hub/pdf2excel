# pdf2excel

Turn a PDF into a tidy Excel workbook — for **two very different kinds of PDF**,
with automatic routing between them:

- **Prose mode** — long-form text documents (laws, regulations, guidelines).
  One row per paragraph, with page number, paragraph/heading type, and the
  nearest section heading.
- **Tables mode** — slide decks and table-heavy PDFs. One worksheet per table
  (side-by-side boxes are stitched together), plus a `slides_text` overview.

The input can be a **local PDF, a PDF URL, or any public web page (HTML)** —
the router sniffs each source and picks the right path automatically; you can
also force a mode.

Either source can be exported in **two output formats**:

- **Default** — the native workbook described under each mode below.
- **Standard Assessment** (`--format standard`) — populates the bundled
  `templates/Standard.xlsx` so the result drops straight into an existing
  assessment workflow (see below).

> Works on PDFs that contain real text (selectable in a viewer). It does **not**
> OCR scanned/image-only PDFs.

## Install

Requires **Python 3.9+**.

```bash
pip install -r requirements.txt
```

Dependencies: `pdfplumber` + `openpyxl` for the core PDF→Excel path; the URL
feature additionally needs `trafilatura`, `beautifulsoup4`, `lxml`, `requests`,
`certifi`, `truststore` and (for JS-heavy pages) `playwright` — all lazily
imported, so non-URL use never touches them. The headless renderer needs a
one-time browser download:

```bash
playwright install chromium      # only needed for JavaScript-rendered pages
```

## Quick start

```bash
# Let the router decide (recommended):
python router.py samples/sample_law.pdf      # -> detected: prose
python router.py samples/sample_deck.pdf     # -> detected: tables

# A public web page (HTML) straight into the Standard Assessment format:
python router.py "https://example.org/statute" --format standard \
    --standard-id MLSR230 -o statute.xlsx

# Desktop GUI:
python pdf_to_excel_gui.py
```

## The router (auto mode)

`router.py` inspects the first few pages and classifies the document:

- **landscape fraction** — share of pages wider than they are tall.
- **table-page fraction** — share of pages with at least one real table.

It returns **tables** if more than half the sampled pages are landscape *or* at
least half contain a real table; otherwise **prose**. Decks are landscape and/or
table-heavy; documents are portrait prose.

```bash
python router.py input.pdf -o out.xlsx --mode auto      # auto (default)
python router.py input.pdf -o out.xlsx --mode prose     # force prose
python router.py input.pdf -o out.xlsx --mode tables    # force tables
python router.py input.pdf --mode prose --heading-style legal
```

As a library:

```python
from router import detect_kind, convert

print(detect_kind("input.pdf"))          # "prose" | "tables"
result = convert("input.pdf", "out.xlsx", mode="auto")
print(result.mode, result.out_path)
```

## Prose mode — `pdf_paragraphs_to_excel.py`

Reconstructs paragraphs from the raw word boxes that
[`pdfplumber`](https://github.com/jsvine/pdfplumber) reports, rather than
trusting the PDF's internal text flow:

1. **Lines** — words are clustered by vertical position into visual lines.
2. **Paragraphs** — lines are grouped by detecting vertical gaps larger than the
   page's typical line spacing (`median gap × gap_factor`).
3. **Headers / footers** — text appearing at the top/bottom of most pages is
   detected by its *digit-normalized* form (so `Page 1` and `Page 2` collapse to
   `Page #`) and removed, along with bare page numbers.
4. **Headings** — see heading styles below; each row is tagged `heading` or
   `body` and carries the nearest heading in its `section` column.
5. **Cleanup** — words hyphenated across a line break are re-joined, paragraphs
   continuing across a page boundary are merged, and `(cid:NNN)` glyph artifacts
   are stripped.

Output: a single `paragraphs` sheet —

| para_id | page | type | section | text |
|---------|------|------|---------|------|

in **Arial**, with a styled/frozen header row, wrapped text, an auto-filter and
sensible column widths.

### Heading styles (`--heading-style`)

| style | matches | example |
|-------|---------|---------|
| `numbered` | dotted clause numbers | `4.2 Documentation` |
| `legal` | Chapter / Article / Section / Part / Annex / Schedule | `Article 7 (Records)` |
| `auto` (default) | picks `legal` when ≥3 distinct legal headings are present, else `numbered` | |

For a parenthetical legal heading, the heading **row** holds only the clean
label (e.g. `Article 1 (Purpose)`) and the trailing sentence begins the body
paragraph. The `legal` pattern uses a lookahead so in-text citations like
`Article 2, Paragraph 1 …` are *not* treated as headings.

```bash
python pdf_paragraphs_to_excel.py guideline.pdf -o out.xlsx --gap-factor 1.8
python pdf_paragraphs_to_excel.py law.pdf --heading-style legal
```

```python
from pdf_paragraphs_to_excel import extract_paragraphs, write_excel
paras = extract_paragraphs("law.pdf", gap_factor=1.6, heading_style="auto")
write_excel(paras, "out.xlsx")
```

## Tables mode — `deck_tables_to_excel.py`

For each page it pulls every *real* table (≥2 rows and ≥2 non-empty cells —
which discards the full-page border boxes pdfplumber reports as single cells),
cleans each cell (`(cid:NNN)` removal + whitespace collapse), and **stitches
side-by-side tables**: tables that are horizontally adjacent and vertically
aligned are concatenated column-wise, so an "As-Is | To-Be" pair becomes one
table.

Output: one sheet per resulting table named `slide{page}` (with `_2`, `_3`
suffixes when a page has several), plus a `slides_text` sheet (one row per slide
from `page.extract_text()`). Arial, bold white-on-navy header row, frozen
header, wrapped text, autosized columns (capped at width 60).

```bash
python deck_tables_to_excel.py deck.pdf -o out.xlsx
```

## Standard Assessment format — `standard_export.py`

`--format standard` writes into `templates/Standard.xlsx` **as a template**. The
file is loaded with openpyxl and only the metadata cells and data rows are
touched, so the merged cells, Times New Roman header styling, data-validation
dropdowns and named ranges (`CLASSIFICATION` / `Requirement_Applicability` /
`Responsible_by` / `Information`) are preserved. Data rows are *cleared by value*
(never deleted) — deleting rows would corrupt validations anchored to fixed row
ranges (e.g. `E10:E1020`, `I10:I1020`).

**Metadata** (rows 1–7) is set from the flags: `B1` Document ID, `B2` Document
Name, `B3` Document Revision, `B5` Standard ID, `B6` Standard Title, `B7`
Standard Edition. The external-link formulas in those cells are replaced with
plain values (blank when not provided).

**Column mapping** (data from row 10):

| col | header | filled by | content |
|-----|--------|-----------|---------|
| A | S.No. | auto | `{standard_id}_{i}` |
| B | Standard Clause/Section ID | auto | clause id, written sparsely (only when it changes) |
| C | Clause/Section Title | auto | heading title (prose) or slide title (tables) |
| D | Standard's Text | auto | paragraph text (prose) or row cells joined by ` \| ` (tables) |
| E | Classification | auto | `Information`, or `Requirement` when the text matches obligation language (`shall`, `must`, …) |
| F–Q | Requirement … Recommended Measure | analyst | left empty; dropdowns remain available |

Items are built from the extractor output as follows:

- *Prose*: headings update the running clause id + title but emit no row of their
  own (`Article 1 (Purpose)` → id `Article 1`, title `Purpose`); each body
  paragraph becomes one row.
- *Tables/deck*: every table row becomes one row (cells joined by ` | `) and every
  slide's full text becomes one row, so nothing is lost.

```bash
python router.py law.pdf --format standard --standard-id MLSR221 \
    --standard-title "My Standard" --standard-edition "2024" -o out.xlsx
python router.py deck.pdf --format standard --standard-id DECK01 -o out.xlsx
```

Flags: `--format {default,standard}`, `--standard-id`, `--standard-title`,
`--standard-edition`, `--document-id`, `--document-name`, `--document-revision`,
`--template PATH`.

> The cascading NIST dropdowns (columns J–N) were bound to specific rows in the
> original sample; for a new standard the analyst re-applies them. The
> full-range `E`/`I` dropdowns survive because rows are cleared, not deleted.

## Web pages (URLs) — `web_extract.py`

Give the router a URL instead of a file path and it sniffs the content by magic
bytes (`download_and_extract.detect`): a PDF goes through the normal PDF
pipeline; an HTML page goes through `web_extract.extract_url`, which pulls the
**main content** with [trafilatura](https://trafilatura.readthedocs.io/)
(structured XML output, falling back to readability-lxml + BeautifulSoup). The
cleaned content is walked **in document order** into the same `Paragraph`
objects the PDF path produces — `<h1>`–`<h6>` become headings, `<p>`/`<li>`
become body — and the existing `resolve_heading` / `is_heading` / `heading_label`
detectors run over the text so inline `Article N (Title)` lines are recognised
too. From there the existing writers are reused unchanged. For the Standard
Assessment format the page `<title>` becomes the document name (and standard
title if not given) and the URL becomes the document id.

```bash
python router.py "https://example.org/statute" --format standard \
    --standard-id MLSR230 -o statute.xlsx
```

This is **generic** — no per-site rules, no language/layout assumptions, and no
summarization: original wording and order are preserved. Pasted URLs are
**normalized** before fetching (`clean_url`): surrounding `<>`/quotes/`[]` and a
trailing " [host]" (e.g. a copied `…/key=4 [elaw.klri.re.kr]`, literal or
percent-encoded) are stripped, and the scheme/host are validated.

### JavaScript-rendered pages (`--render`)

Some pages (e.g. the KLRI statute viewer) deliver only a masthead statically and
inject the article body via JavaScript and/or an `<iframe>`, so a plain fetch
sees just the title and amendment history. The extractor first does the fast
static fetch; if the result **looks incomplete** (little body text, no
Article/Chapter/Section structure, or fewer than three paragraphs) it falls back
to a **headless Playwright render** and re-extracts from the rendered page plus
every child frame, in order.

```bash
python router.py "<url>" --render auto    # render only if static looks thin (default)
python router.py "<url>" --render always  # always render
python router.py "<url>" --render never   # never render (fast static only)
```

The diagnostics log shows which path ran (`static: N blocks / M chars`, and
`rendered: …` when it fires), so it's clear whether the body was captured. The
renderer is **optional and lazily imported** — a static page launches no browser.
Enable it once with:

```bash
pip install playwright && playwright install chromium
```

If a page needs rendering but Playwright isn't installed, a clear error explains
the fix (or re-run with `--render never`).

### TLS certificate verification

TLS is **verified by default**. Some servers (e.g. `elaw.klri.re.kr`) don't send
a complete certificate chain, so the default bundle can't build a trust path and
you'll see `CERTIFICATE_VERIFY_FAILED: unable to get local issuer certificate`.
Three remedies, most-secure first:

1. **OS trust store + certifi** — `truststore` (installed via `requirements.txt`,
   Python 3.10+) makes verification use the operating system's trust store,
   which resolves missing intermediates (Windows fetches them via AIA); `certifi`
   keeps the fallback CA bundle current. This usually just works after install.
2. **`--ca-bundle <pem>`** — supply a PEM that includes the missing intermediate;
   it is passed straight to `requests` `verify=`. Secure and explicit.
3. **`--insecure`** — skip verification entirely (suppresses the urllib3 warning).
   This **disables MITM protection** — use only for a trusted public source.

```bash
python router.py "<url>" --ca-bundle chain.pem -o out.xlsx
python router.py "<url>" --insecure -o out.xlsx        # trusted sources only
```

### Responsible data handling

- Only publicly visible content is processed, with a descriptive `User-Agent`.
- Gating is respected, not bypassed: a `401`/`403` (or a paywall/anti-bot wall)
  stops with a clear error rather than an empty file.
- Response size is capped; redirects are followed; charset is auto-detected.
- No personal/private data is collected.

**Validation** (warnings are logged; only emptiness is fatal):

- < ~200 chars or zero blocks → `ValueError` ("No main content extracted; page
  may be empty, gated, or JavaScript-rendered").
- A heading/clause label repeating verbatim many times → warning.
- Numeric clause IDs running backwards → ordering warning.

## Desktop GUI — `pdf_to_excel_gui.py`

A thin Tkinter shell over `router.convert` (no extraction logic of its own):

- **PDF file or URL** input — a local path or a URL (PDF or HTML page).
- **Render JavaScript** dropdown (Auto / Always / Never) and an **Allow insecure
  TLS** checkbox for URL fetches.
- **Mode** dropdown: Auto / Prose / Tables.
- **Output format** dropdown: Default / Standard Assessment, plus a **Standard
  ID** field used when the standard format is selected.
- **Gap factor** slider (prose only).
- Runs on a background thread; previews the result in a table — paragraph rows
  for prose, or a per-sheet `sheet | rows | cols` summary for tables. The Excel
  file the router produced is saved to the chosen output path.

```bash
python pdf_to_excel_gui.py
```

## Download by URL, then extract (prose) — `download_and_extract.py`

```bash
python download_and_extract.py https://example.org/guideline.pdf
python download_and_extract.py <url> -o out.xlsx --keep-pdf
```

## Tuning

`--gap-factor` / the GUI slider controls paragraph-break sensitivity in prose
mode. Default `1.6`. If paragraphs split too eagerly, raise it (e.g. `2.0`); if
separate paragraphs merge, lower it (e.g. `1.3`).

## Samples

`samples/` ships two ready-to-try PDFs:

- `sample_law.pdf` — portrait, legal headings; auto-routes to **prose**, and the
  `section` column tracks the `Chapter`/`Article` headings.
- `sample_deck.pdf` — landscape slides with an As-Is | To-Be table pair;
  auto-routes to **tables** and produces one stitched table sheet plus
  `slides_text`.

Regenerate them anytime (no extra dependencies):

```bash
python samples/make_samples.py
```

## Project layout

```text
pdf2excel/
├── README.md
├── requirements.txt              # pdfplumber, openpyxl
├── .gitignore
├── pdf_paragraphs_to_excel.py    # prose engine + CLI
├── deck_tables_to_excel.py       # slide-deck / table extractor + CLI
├── web_extract.py                # URL -> main content -> Paragraph objects
├── standard_export.py            # Standard Assessment writer + mapping layer
├── router.py                     # auto-routing (file/URL) + format + CLI
├── download_and_extract.py       # fetch()/detect() + download-by-URL CLI
├── pdf_to_excel_gui.py           # Tkinter GUI over the router
├── templates/
│   └── Standard.xlsx             # bundled Standard Assessment template
└── samples/
    ├── make_samples.py           # dependency-free sample generator
    ├── sample_law.pdf
    └── sample_deck.pdf
```
