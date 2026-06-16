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

Optionally, an **AI enrichment** pass then fills the assessment's analyst columns
**E–I** — classifying each clause as *Requirement* or *Information* and, for
requirements, drafting the organisation-perspective requirement — using **Claude,
OpenAI, or Gemini**. See *AI enrichment* below.

> Works on PDFs that contain real text (selectable in a viewer). It does **not**
> OCR scanned/image-only PDFs.

## Install

Requires **Python 3.9+**.

It is recommended (and often required on modern Linux distributions) to install the dependencies inside a virtual environment:

```bash
# Create a virtual environment
python3 -m venv venv

# Activate it
source venv/bin/activate  # On Windows use: venv\Scripts\activate

# Install requirements
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

## The router and hybrid pipeline (auto mode)

`router.py` routes the document based on the selected mode:
- **`auto` (default)** — auto-detects the document type and selects the best extraction strategy. **NIST SP 800-53 Rev. 5** PDFs are automatically detected and routed to the dedicated NIST extractor. Other PDFs go through the Hybrid Page-by-Page Extraction Pipeline.
- **`prose`** — forces the paragraph-based prose extraction mode for the whole document.
- **`tables`** — forces the slide-deck / table extraction mode for the whole document.
- **`nist80053`** — forces the dedicated NIST SP 800-53 Rev. 5 extractor (one row per base control or control enhancement).

### NIST SP 800-53 Rev. 5 — `nist80053_extractor.py`

A dedicated extractor for the NIST SP 800-53 Rev. 5 control catalog. The correct
extraction unit is **one base control or one control enhancement = one Standard
Assessment row**, not paragraphs or visual line fragments.

**Auto-detection**: When `mode="auto"`, the router checks the first few pages
for `"NIST SP 800-53"` and `"Revision 5"` / `"Rev. 5"` markers. If found, the
NIST extractor runs automatically.

**What it does**:
- Identifies the Chapter Three control catalog boundaries (skips cover, TOC,
  chapters 1–2, errata, references, appendices).
- Filters out rotated side-margin text (`upright=False`, DOI boilerplate at x0 < 45).
- Removes page headers, footers, and underline separator lines.
- Detects base control headings (e.g. `AC-1 POLICY AND PROCEDURES`) via bold
  Calibri font + regex.
- Detects control enhancements (e.g. `(1) AUTOMATED SYSTEM ACCOUNT MANAGEMENT`)
  and assigns compound IDs like `AC-2(1)`.
- Classifies withdrawn controls (`[Withdrawn: ...]`) as `Information`; all active
  controls and enhancements as `Requirement`.
- Runs validation checks: duplicate IDs, missing fields, DOI/TOC contamination.

**Expected output for NIST 800-53 Rev. 5**:
- ~1,100 rows (946 requirements + ~176 withdrawn/information items)
- Zero one-letter or broken-fragment rows
- Zero DOI or TOC rows
- Every row has a valid `XX-NN` or `XX-NN(M)` clause ID

```bash
# Auto-detect NIST 800-53 (recommended):
python router.py NIST.SP.800-53r5.pdf \
  --format standard \
  --standard-id NIST80053R5 \
  -o NIST.SP.800-53r5_fixed.xlsx

# Explicit NIST mode with full metadata and review output:
python router.py NIST.SP.800-53r5.pdf \
  --mode nist80053 \
  --format standard \
  --standard-id NIST80053R5 \
  --standard-title "NIST SP 800-53 Rev. 5 Security and Privacy Controls" \
  --document-name "NIST SP 800-53 Revision 5" \
  --review-output nist_review.xlsx \
  -o NIST.SP.800-53r5_fixed.xlsx
```

### Hybrid Extraction Options
When using `mode="auto"`, you can fine-tune extraction with these new CLI options:
- `--review-output PATH` — path to write an intermediate block-level review Excel sheet.
- `--include-toc` — include table of contents pages (skipped by default).
- `--no-skip-cover` — do not skip cover pages (the first page is skipped by default if word count < 300).
- `--ocr-mode {off,detect}` — set to `detect` to enable Tesseract OCR on scanned/image-only pages.
- `--min-confidence THRESHOLD` — filter out rows with confidence below the threshold (e.g. `0.5`).
- `--show-issues` — print warning codes for rows with post-extraction validation issues in CLI.

### Pipeline Architecture
The hybrid pipeline consists of:
1. **Preflight (`preflight.py`)**: Checks page geometry/orientation, table layout presence, character counts, multi-column gutter spacing, and scanned-page status.
2. **Block Extraction (`extract_blocks.py`)**: Visual sorting of layout columns, title/heading identification, and table extraction. Can write a color-coded review workbook.
3. **Validation (`validate_extraction.py`)**: Runs post-extraction rule validation on output items, checking for duplicates, naked page numbers, empty fields, TOC dot leaders, or broken clause numbering order, assigning confidence scores and issues list to each item.

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

## AI enrichment (columns E–I) — `ai_enrich.py` / `ai_providers.py`

A second, optional phase. After extraction has produced the Standard Assessment
rows (columns A–E), an LLM reads each clause and fills the analyst columns:

| col | header | filled by the model |
|-----|--------|---------------------|
| E | Classification | `Requirement` or `Information` (overrides the keyword heuristic; the keyword rule is the fallback if a row can't be parsed) |
| F | Requirement | the organisation-perspective requirement in *Standard Format* ("The organization shall …"), Direct or Derived from the text |
| G | Detailed Description | `REQ-NNN \| Verification: <method> \| <Direct/Derived> — <trace>` followed by a plain-language description of what complying entails |
| H | Change in Requirement | the concrete change/action the organisation must implement to comply |
| I | Requirement Classification | one of a configurable controlled vocabulary (default `Product` / `Process` / `Other`) that feeds the template's NIST cascade |

For a clause classified **Information**, columns F–I are left **blank** (toggle
with `fill_only_requirements`). Requirement IDs are renumbered into one
continuous `REQ-001, REQ-002, …` sequence across the whole document.

The prompt for E+F is the supplied compliance-analyst prompt; the instructions
for G, H and I are editable defaults (see `DEFAULT_PROMPTS` in `ai_enrich.py`, or
the prompt editors in the GUI). One **consolidated structured-JSON call per
batch** of clauses keeps E–I mutually consistent and costs far less than a call
per column; output is parsed defensively (a JSON-repair retry, then a per-row
keyword fallback), and batches run concurrently.

### Providers, models & keys

| provider | flag value | default model | API key (env var or saved config) |
|----------|------------|---------------|-----------------------------------|
| Claude (Anthropic) | `claude` | `claude-opus-4-8` | `ANTHROPIC_API_KEY` |
| OpenAI | `openai` | `gpt-4o` | `OPENAI_API_KEY` |
| Gemini (Google) | `gemini` | `gemini-2.0-flash` | `GOOGLE_API_KEY` |
| Ollama (local) | `ollama` | `gemma4` | none — local server (`OLLAMA_HOST`, default `localhost:11434`) |

Keys resolve **saved config → environment variable**. The GUI saves them (and
your model/prompt/option choices) to `~/.pdf2excel.json` (gitignored, `0600`).
SDKs are **lazily imported** — install only the provider(s) you use
(`pip install anthropic` / `openai` / `google-generativeai` / `ollama`). For
Claude the large instruction prompt is sent with prompt caching, JSON is enforced
via structured outputs, and `temperature` is dropped for Opus 4.7/4.8 (which
reject it). **Ollama** runs models locally — **no key, no cost, no rate limits**:
start the Ollama server, `ollama pull <model>`, choose *Ollama (local)* and click
**List models** to pick one you've pulled (optionally set a host in the field).
It's the ideal way to test the whole flow for free.

> **Local-model resources.** Local models are memory-heavy: if Ollama is killed
> mid-run (exit code 137 / out of memory), switch to a smaller model
> (e.g. `gemma3`, `llama3.2`, `qwen2.5:3b`), keep **Workers = 1** and a small
> **Batch size**, and validate with **Rows = First N** before enriching a large
> document. Big documents (thousands of clauses) on a laptop are slow — prefer a
> hosted provider, or a small local model, for the full run.

Model availability changes over time (e.g. Google retired the Gemini 1.5
models). The model field is editable, and the GUI's **List models** button
queries the provider's API for the exact models your key can use — click it and
pick one if a default model returns a *not found / not supported* error.

### CLI

```bash
# Dry-run (no API calls, no cost) — exercises the whole pipeline with placeholders:
python router.py samples/sample_law.pdf --format standard --ai-fill --ai-dry-run -o out.xlsx

# Real run with Claude (needs ANTHROPIC_API_KEY):
python router.py law.pdf --format standard --standard-id MLSR221 \
    --ai-fill --ai-provider claude --ai-model claude-opus-4-8 -o out.xlsx

# OpenAI, smaller batches:
python router.py law.pdf --format standard --ai-fill \
    --ai-provider openai --ai-model gpt-4o-mini --ai-batch-size 8 -o out.xlsx
```

Flags: `--ai-fill` (requires `--format standard`), `--ai-provider`, `--ai-model`,
`--ai-batch-size`, `--ai-workers`, `--ai-temperature`, `--ai-dry-run`.

As a library:

```python
from router import convert
from ai_enrich import EnrichConfig

cfg = EnrichConfig(provider="claude", model="claude-opus-4-8", batch_size=12)
result = convert("law.pdf", "out.xlsx", fmt="standard",
                 standard_id="MLSR221", enrich_config=cfg)
print(result.n_requirements, "requirements /", result.n_items, "rows")
```

### Rate limits & quota (429)

API calls are rate-limited. On a `429` the run **retries with backoff** (honouring
the server's `retry-after`) up to **Retries** times (default 5), so transient
per-minute limits clear on their own. If the quota genuinely can't be satisfied
(e.g. a free-tier `limit: 0` for the chosen model), the run stops with an
actionable message. Remedies:

- **Lower throughput** — set **Workers** to 1 and reduce **Batch size** to stay
  under per-minute limits.
- **Pick a model with quota** — click **List models**; free-tier quota varies by
  model/region (Gemini's free tier may grant `0` for some models).
- **Enable billing** on the provider account, or **switch provider**.
- **Dry-run** needs no quota — use it to validate the pipeline first.

### Responsible use

Only the extracted clause text (plus the standard's metadata) is sent to the
chosen provider — no other document or system data. Keys are stored locally and
gitignored. The model is instructed to stay faithful to the source and **not**
invent obligations, thresholds, or scope the text does not support; a `dry-run`
mode and a pre-run **cost estimate** let you validate the flow before spending.

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

A thin Tkinter shell over `router.convert` + `ai_enrich.enrich` (no extraction logic
in the GUI). Three tabs:

1. **Input & Extract** — PDF/URL, output path, **Mode** (Auto / Prose / Tables /
   Standard / Structured / NIST 800-53), **Format** (Default / Standard Assessment),
   **Profile** (Auto, Generic, NIST 800-53, Control Catalog, ISO-like, Legal,
   PCI, CIS), standard & document metadata, structured extraction options (review
   workbook, Extraction_Issues sheet, force export, low-confidence rows, include
   front matter/TOC/appendix/references), OCR mode, minimum confidence, then
   **Run Extraction**.
2. **AI Configuration** — provider, model, API key, batch/workers/temperature,
   editable prompts, cost estimate. AI runs only after a successful extraction.
3. **Run & Results** — exported-items preview, log summary, AI enrichment grid,
   export to Excel.

### Recommended workflow for standards PDFs

1. Open the GUI: `python pdf_to_excel_gui.py`
2. Select your PDF (e.g. `NIST.SP.800-53r5.pdf`)
3. **Format** = Standard Assessment · **Profile** = Auto
4. Leave **Generate review workbook** and **Show Extraction_Issues sheet** checked
5. Click **Run Extraction**
6. Open the review workbook (`*_review.xlsx`) to inspect rejected blocks
7. If the quality gate passes (or you intentionally force export), run **Generate with AI** on tab 3

### GUI verification checklist (NIST)

| Check | Expected |
|-------|----------|
| Row 10 in output | `AC-1` |
| Front matter | not exported |
| AC-2(1) | one complete row with Discussion inside text |
| Review workbook | generated beside output |
| GUI summary | issue/warning counts visible in log |

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
├── requirements.txt              # core: pdfplumber, openpyxl (+ optional web/AI)
├── .gitignore
├── pdf_paragraphs_to_excel.py    # prose engine + CLI
├── deck_tables_to_excel.py       # slide-deck / table extractor + CLI
├── preflight.py                  # page preflight and column/scanned detection
├── extract_blocks.py             # visual column-aware block extraction & sorting
├── validate_extraction.py        # post-extraction rules and confidence validation
├── web_extract.py                # URL -> main content -> Paragraph objects
├── standard_export.py            # Standard Assessment writer + mapping layer
├── ai_providers.py               # LLM provider abstraction (Claude/OpenAI/Gemini)
├── ai_enrich.py                  # AI enrichment: items -> cols E–I (+ prompts)
├── config.py                     # local settings/keys persistence (~/.pdf2excel.json)
├── router.py                     # hybrid-routing (file/URL) + format + AI + CLI
├── download_and_extract.py       # fetch()/detect() + download-by-URL CLI
├── pdf_to_excel_gui.py           # 3-tab Tkinter GUI (extract → AI → export)
├── templates/
│   └── Standard.xlsx             # bundled Standard Assessment template
└── samples/
    ├── make_samples.py           # dependency-free sample generator
    ├── sample_law.pdf
    └── sample_deck.pdf
```
