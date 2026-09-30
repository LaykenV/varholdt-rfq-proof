# Synthetic RFQ fixtures

Every file in this directory is invented test data. No customer supplied these
documents, catalogues, or prices. The Arabic text appears in the generated
workbook headings, not the source PDFs.

Run `python3 scripts/generate_rfq_fixtures.py` to reproduce all JSON and PDF
fixtures. The generator uses only Python's standard library. Its PDFs contain
ordinary selectable text and a fixed pipe-separated layout. They do not test
arbitrary vendor layouts, OCR, handwriting, or Arabic PDF extraction.

| Fixture | Expected behavior |
| --- | --- |
| `baseline.pdf` | Two exact matches, AED 25.00 and AED 0.02, subtotal AED 25.02 |
| `unseen.pdf` | Eight rows, two priced, six unpriced for review, subtotal AED 23.00 |
| `multipage.pdf` | Three rows across two pages, subtotal AED 37.52 |
| `count-mismatch.pdf` | Declared row count differs from extracted items, run fails |
| `unrecognized-line.pdf` | Unexpected text creates a count conflict, run fails |
| `parse-exception.pdf` | Unexpected text survives as an unpriced review row because the count reconciles |
| `duplicate-rfq-row.pdf` | Both rows with the same row ID stay unpriced |
| `no-text.pdf` | Valid PDF without extractable text, run fails |
| `precision-boundary.pdf` | With precision price fixture, rounded subtotal is AED 999995000000000000.00 |

The baseline deliberately includes a three-decimal unit price. The unseen
same-layout PDF covers an unknown code, description-only matching, unit
mismatch, missing approved price, conflicting description, ambiguous
description, and formula-like text. Description matches suggest candidates
but never transfer a price without review. The formula-like description must
remain plain text in Excel.

`duplicate-catalogue.json`, `duplicate-prices.json`, and
`unapproved-prices.json` must fail before producing a workbook. The changed
price fixture has a new version and changes the baseline subtotal to AED 26.02.
Reusing the original version with changed bytes must fail.

The precision fixture uses extreme values within the documented input bounds
to catch intermediate Decimal rounding. It is not a plausible customer quote
or a claim about transaction size.

Independent acceptance checks run with:

```bash
.venv-rfq/bin/python -m unittest discover -s tests -p test_rfq_acceptance.py -v
```

The checks inspect actual saved workbooks, original source snapshots, SHA-256
provenance, row counts, null prices, failed-run state, and formula cell types.
The unseen fixture was authored independently against the agreed input
contract. It is synthetic, not an unseen real customer document.
