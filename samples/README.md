# samples/

Two ready-to-try PDFs are bundled here:

| file | shape | auto-routes to |
|------|-------|----------------|
| `sample_law.pdf` | portrait, `Chapter`/`Article` headings, 2 pages | **prose** |
| `sample_deck.pdf` | landscape slides, an As-Is vs To-Be table pair | **tables** |

Try them:

```bash
python ../router.py sample_law.pdf      # prose; section tracks Article/Chapter
python ../router.py sample_deck.pdf     # tables; one sheet per table + slides_text
```

## Regenerating the samples

`make_samples.py` rebuilds both PDFs using a tiny hand-rolled PDF writer, so it
needs **no third-party dependencies**:

```bash
python make_samples.py
```

`_make_test_pdf.py` is an older helper that builds a guideline PDF with
`reportlab` (`pip install reportlab`) — optional.

Drop your own PDFs here too; generated `.xlsx` files in this folder are
gitignored.
