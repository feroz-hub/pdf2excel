# Implementation Analysis — pdf2excel Standard Extraction

## 1. Router flow (`router.py`)

**Decision order today:**

| Input | Path |
|-------|------|
| URL (HTML/text) | `web_extract` → `paragraphs_to_items` |
| `--mode nist80053` | `nist80053_extractor.extract_nist_800_53_items` |
| `--mode auto\|standard` **or** `--format standard` | Structured pipeline (preflight → blocks → parser → validation) |
| `--mode prose` + `--format default` | `pdf_paragraphs_to_excel` |
| `--mode tables` + `--format default` | `deck_tables_to_excel` |

**Previous failure:** `--format standard` with `--mode prose` (or legacy auto→prose) wrote one Excel row per PDF paragraph via `paragraphs_to_items`, dumping cover/TOC/DOI/fragments.

**Fix:** `--format standard` always routes through the structured pipeline. NIST/control-catalog profiles use the dedicated line-level NIST extractor.

## 2. Paragraph extraction (`pdf_paragraphs_to_excel.py`)

Clusters pdfplumber words into lines/paragraphs using vertical gaps and heading heuristics. Each paragraph becomes a row in default format. **Not suitable for Standard Assessment** — no clause boundaries, no front-matter skip, continuations become separate rows.

## 3. Table extraction (`deck_tables_to_excel.py`)

Extracts per-page tables for slide/table PDFs. `deck_to_items` joins cells with ` | `. Used for `--mode tables` default format; table normalization in the structured pipeline reuses its table helpers.

## 4. Standard export (`standard_export.py`)

Loads `templates/Standard.xlsx`, writes metadata B1–B7, clears row values from 10+ (never deletes rows), fills A–E from finalized items. Column B is sparse (repeated clause IDs omitted). `export_items` allows writing a filtered subset while keeping full `items` for `Extraction_Issues`.

## 5. Where raw fragments leaked

1. **Wrong path:** `paragraphs_to_items` / `blocks_to_items` mapped every block → row.
2. **Block pipeline:** Side-margin rotated DOI text, one-letter fragments, headers/footers before AC-1.
3. **Parser bug:** `(1) Changing one or more…` matched enhancement regex (`IGNORECASE`) → duplicate/wrong clause IDs.
4. **Export filter gaps:** Duplicate clause IDs exported (sparse column B looked like “blank ID”).
5. **Appendix included by default** (`include_appendix=True`).

## 6. Why complex standards fail

- PDF layout ≠ logical structure (multi-column, rotated margin text, repeated headers).
- NIST controls use `(N) ALL CAPS TITLE` enhancements; base control text uses `(1) mixed case` sub-clauses — regex must distinguish them.
- Continuation labels (`Discussion:`, `Related Controls:`) must append, not become rows.
- TOC/errata/appendix tables look like content without boundary detection.

## 7. Changes implemented

```
PDF → preflight (page types, profile hints)
    → layout_blocks (unified blocks, noise marked)
    → profile detection (standard_profiles / auto)
    → structure_parser state machine OR nist80053 line parser
    → validation (blocking vs warning issues)
    → review workbook (rejected blocks)
    → Standard.xlsx (export_items only)
    → optional AI enrichment
```

**Rules enforced:**

- Never export raw layout blocks to Standard Assessment.
- Front matter / TOC / errata / references / appendix skipped by default.
- NIST starts at `AC-1`, stops before REFERENCES/APPENDIX A–C.
- Enhancements require ALL CAPS titles (no `IGNORECASE`).
- Blocking issues → review workbook / Extraction_Issues only.
- Quality gate fails loudly unless `--force-export`.

## GUI integration (`pdf_to_excel_gui.py`)

The GUI is a thin wrapper over `router.convert()`. Tab 1 collects all structured
pipeline options (profile, review workbook, issue flags, include sections, OCR,
min confidence, metadata) and calls `convert(..., raise_on_quality_gate=False)`.
On quality gate failure, `QualityGateError.result` carries partial stats; the GUI
shows the review path and disables AI until extraction is clean or force export
is enabled.

### Manual GUI regression (NIST SP 800-53 Rev. 5)

Settings on tab **1 · Input & Extract**:

| Field | Value |
|-------|-------|
| Input | `NIST.SP.800-53r5.pdf` |
| Format | Standard Assessment |
| Mode | Auto |
| Profile | Auto or NIST 800-53 |
| Generate review workbook | checked |
| Show issues sheet | checked |
| Force export | unchecked |
| Export low-confidence rows | unchecked |
| Include front matter / TOC / appendix / references | unchecked |
| Output | `NIST_fixed_gui.xlsx` |

After extraction, the log panel should show `Format: Standard Assessment -> standard`,
`Mode: Auto -> auto`, and ~1,100 items exported. Verify:

```bash
python verify_nist_output.py NIST_fixed_gui.xlsx
```

Expected: row 10 = `AC-1`, no `JOINT TASK FORCE`, no duplicate clause IDs.
