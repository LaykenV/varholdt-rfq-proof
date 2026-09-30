# Independent RFQ proof review

September 29, 2026. Pass for the specified synthetic text-PDF layout and local
draft-workbook behavior. Nine independent acceptance tests passed against the
final helper after correcting an arithmetic boundary defect.

This proves a working PDF extraction, catalogue check, separate price lookup,
and Excel-file path on invented data. It does not prove compatibility with a
customer's PDF, production readiness, or that the owner has reviewed the file
in Excel.

## Independent fixtures and checks

The reviewer authored the fixtures against the agreed input contract rather
than deriving expected results from the helper. The generator uses Python's
standard library and creates valid PDFs with selectable text. Baseline and
unseen PDFs share one explicit layout. The unseen file is another invented
document, not an unseen customer document.

| Case | Checked result |
| --- | --- |
| Baseline | Two priced rows, AED 25.00 and AED 0.02, subtotal AED 25.02 |
| Unseen same-layout PDF | Eight rows preserved, two priced, six unpriced for review, subtotal AED 23.00 |
| Description-only and ambiguous descriptions | Review suggestions appear without an automatic price |
| Unknown code, conflicting description, unit mismatch, missing price | Each row stays visible with reasons and null prices |
| Duplicate catalogue or price code | Run fails before a workbook appears |
| Unapproved price list | Run fails before a workbook appears |
| Duplicate RFQ row ID | Both input rows remain visible and unpriced |
| Two-page PDF | All three rows and their source page/line references survive, subtotal AED 37.52 |
| Unexpected extracted text | It survives as an unpriced parse exception when the count reconciles; a count conflict fails the run |
| Empty text layer | Run fails; the prototype does not claim OCR |
| Formula-like description | `=1+1` remains an Excel string, not a formula |
| New price version | Baseline changes to AED 26.02 and points to the new version and hash |
| Changed data under the same price version | Run fails instead of silently changing the approved snapshot |
| Failure after a successful run | Failed current state has no workbook pointer; the old archived workbook remains unchanged |

The tests reopen the actual XLSX files. Every input item appears in the draft
sheet. Every review item also appears in the separate exceptions sheet, with
empty price cells. Source text, catalogue and price versions, hashes, and
decisions match the JSON result. Every nonblank extracted source line appears
in the source audit. All workbook cells are checked for formula type.

Input PDF, catalogue, and approved-price bytes remain unchanged and are copied
into each successful run. Their SHA-256 hashes match the result and manifest.
Input counts equal priced rows plus review rows. Unpriced rows are not treated
as zero-cost quoted items. The displayed subtotal covers only priced rows.

## Arithmetic finding corrected

The original helper used Decimal's default intermediate precision. Two allowed
values could multiply to a number just below a half-cent, round up during the
multiplication, and then round up again when converted to cents.

The regression multiplies `999999999.999999` by `999995000.000001`. The exact
product is `999995000000000000.004999999999`, whose correct two-decimal result
is `999995000000000000.00`. The first implementation returned a final cent of
`01`. The builder increased local precision for multiplication, addition, and
quantization. The independent regression now passes, and ordinary half-cent
rounding still passes. These extreme values test the documented bounds; they
do not represent a plausible customer quote.

## n8n review

The reviewer inspected `scripts/verify_rfq_n8n.py`. It imports and executes an
actual workflow in an isolated local n8n installation. The workflow fetches the
PDF through HTTP, posts its binary data to the local helper, and receives the
XLSX as binary output. The verifier compares the posted PDF bytes with the
source and the returned workbook bytes with the helper output.

The final executions use the reviewed helper after the arithmetic fix:

- `artifacts/private/rfq-final-baseline` has two priced rows and an AED 25.02
  subtotal.
- `artifacts/private/rfq-final-unseen-serial` has two priced rows and six review
  rows, with an AED 23.00 subtotal.
- `artifacts/private/rfq-final-rejection-serial` records the expected failed
  parser state, no current workbook pointer, and no XLSX anywhere in the run.

The reviewer independently reopened these final returned workbooks. Their bytes
equal the helper's workbook bytes, their hashes match the verification records,
their exception sheets contain the expected rows, and none contains a formula
cell. All recorded helper, verifier, and requirements hashes match the current
files. Each execution reports n8n 2.41.3 and the expected GET then POST sequence.

Concurrent verification attempts conflicted on n8n's task-broker port before
workflow execution. Those attempts are not passes. The final unseen and
rejection checks ran serially. Use serial local verification unless the n8n
runtime's ports are isolated separately.

This demonstrates local n8n orchestration. It does not demonstrate a customer's
n8n instance, authenticated upload service, deployed webhook, production
credentials, or concurrent-job handling.

## Limits and reproduction

The workbook has English and Arabic headings. The review checked their
presence, sheet structure, and cell values. It did not include visual inspection
in desktop Excel or linguistic review by an Arabic speaker. Amount cells and
subtotals are exact text values. Editing a quantity in Excel does not recalculate
the workbook.

Only exact product codes with consistent descriptions, units, quantities, and
approved prices receive draft amounts. The code's description normalization and
suggestions are deterministic. They are not evidence of AI matching accuracy.
Human approval remains necessary. Taxes, discounts, shipping, stock, and delivery
promises are outside this proof.

Run the independent checks with:

```bash
python3 scripts/generate_rfq_fixtures.py
.venv-rfq/bin/python -m unittest discover -s tests -p test_rfq_acceptance.py -v
```

Reviewed helper SHA-256 is
`deaece12ce5f530e6f802ee22750da4624b43162ff791c61f623bcede625fa36`.
The nine-test acceptance file SHA-256 is
`6bab1378c93441673565333e6a643882af713c393ee7280a85b35214609dfe51`.
The reproducible fixture generator SHA-256 is
`56d8a5562830b3685c633d4f4aca203cfdf43c65461a5e0b32deb832d053e6a7`.
