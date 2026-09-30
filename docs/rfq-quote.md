# RFQ quotation proof

This local prototype reads one explicitly synthetic text-PDF layout, matches
its rows to a catalogue, looks up a separate approved AED price list, and writes
an Excel quotation draft. It does not issue quotes, send messages, or connect
production accounts. The n8n wrapper is separate; running this command alone is
not evidence of an n8n execution.

## Install and run

```bash
uv venv .venv-rfq
uv pip install --python .venv-rfq/bin/python -r requirements-rfq.txt
.venv-rfq/bin/python src/rfq_quote.py fixtures/rfq-demo/baseline.pdf fixtures/rfq-demo/catalogue.json fixtures/rfq-demo/prices.json artifacts/private/rfq-output --run-id baseline-001
.venv-rfq/bin/python -m unittest discover -s tests -v
```

The environment here lacks system `ensurepip`; `uv` creates the project-local
environment without changing system Python. Dependencies are pinned to pypdf
6.19.0, openpyxl 3.1.5, and et_xmlfile 2.0.0. Versions were checked against their
package releases on September 29, 2026. Installation needs package downloads;
the RFQ command itself has no network calls or paid model/API dependency.

## Agreed synthetic input contract

The PDF contains selectable text. Its first header is `SYNTHETIC RFQ V1`, one
count line says `ITEMS: N`, and the table header is
`ROW | CODE | DESCRIPTION | QTY | UNIT`. Each item occupies one extracted text
line with those five pipe-separated fields. A blank code is allowed and requires
review. A final `END RFQ` line closes the document. Layout and table headers may
repeat on later pages. Pipes and wrapped descriptions inside an item are outside
this layout.

Every nonblank extracted text line enters the source audit. Unrecognized rows
enter the item list as unpriced parse exceptions. The declared count must equal
the parsed and exception row count; a mismatch fails the run. Page and text-line
numbers refer to pypdf's extraction, not a claim about pixel coordinates.

Limits are ten pages, 200 item rows, a 10 MB PDF, 5 MB decoded content per page,
and 5,000 catalogue or price-list entries. These are bounded prototype inputs,
not a hostile-upload service. Scans, encryption, arbitrary PDF layouts, OCR,
rotated or wrapped tables, and Arabic source-PDF extraction are not supported.
English/Arabic workbook headings are included; Arabic catalogue translation and
linguistic review are separate work.

Catalogue JSON:

```json
{"version":"demo-cat-v1","items":[{"code":"P001","description":"Synthetic part","unit":"EA"}]}
```

Separate approved price JSON:

```json
{"version":"demo-prices-v1","currency":"AED","approved":true,"prices":[{"code":"P001","unit":"EA","unit_price":"0.005"}]}
```

An operator supplies approved data. The `approved` flag is not a signature and
cannot establish authorization by itself. Codes must be unique within each
source. Catalogue and price versions are explicit. Price snapshots are preserved
in `OUTPUT/price-versions/`; reusing a version with different bytes fails. This
immutability check is local to that output directory, not a global registry.

## Matching and arithmetic

Only an exact case-sensitive product code can receive a price automatically.
Its description must equal the catalogue description after case and whitespace
normalization. Its unit must exactly match both catalogue and approved price
units. Quantity must be positive and the approved price must exist. An empty
description, conflict, missing price, invalid quantity, or duplicate row ID makes
the entire row unpriced and requires review. Both copies of a duplicate row ID
remain visible. Repeated product codes on distinct row IDs are separate requests.

Unknown and similar codes and description-only matches remain unpriced regardless
of their similarity. Up to three catalogue codes can appear as review suggestions.
Suggestions never transfer prices and no substitutions are approved. There is no
human-approval persistence or override workflow in this proof.

Amounts use Decimal with a local precision of 60 digits for multiplication,
rounding and accumulation. This preserves all permitted operand digits before
the final rounding and does not change the caller's global context.
Quantities and unit prices accept at most nine integer
digits and six decimal places. Prices can be zero. Line totals round to two AED
decimals using `ROUND_HALF_UP`: `0.145 × 1.00` becomes `0.15`, and
`0.005 × 1` becomes `0.01`. The subtotal adds the rounded priced line totals.
Review rows are excluded; it is not the total cost of fulfilling the whole RFQ.
Tax, discounts, freight, currency conversion, stock and delivery promises are
outside scope.

## Output and failure handling

`run(PDF, CATALOGUE, PRICES, OUTPUT, RUN_ID)` and the CLI produce a unique
`OUTPUT/run-RUN_ID/` directory containing the unchanged input PDF, catalogue,
approved prices, `result.json`, `quote.xlsx`, and `manifest.json`. The manifest
hashes the inputs and outputs. Result rows record the price and catalogue version
and hash, source page and text line, original extracted text, decision reasons,
and explicit null prices for review rows. Input, exact-draft, and review counts
reconcile.

The workbook has a draft sheet with every input item, a separate exceptions
sheet containing every review row, a summary, and a source text audit. The
exceptions sheet repeats rows from the draft, not additional input rows.
Headings include English and Arabic. Cells are explicit Excel text,
including exact decimal amounts, so spreadsheet software cannot reinterpret a
leading `=` as a formula. The subtotal is precomputed and displayed as exact
text. This also means amount cells are not a numeric spreadsheet model for
editing or recalculating quotes. Human approval is required for every output.

The command writes `OUTPUT/last-run.json` as running, succeeded, or failed. A
successful state includes paths for this run's workbook, result and manifest.
Failure contains no workbook pointer, and a partial new run directory is removed.
Existing run directories cannot be reused or overwritten. A previous successful
run remains archived. An integration must check the command's zero exit code,
the expected run ID and succeeded state before using its named workbook. Never
pick a workbook by a directory wildcard or treat any old file as the latest run.
Use one writer per output directory; concurrent writers need isolated roots.

## Evidence and remaining work

The buyer's actual requirement is the complete PDF to n8n to Excel route,
including a separate price version, AED, bilingual headings, an unseen PDF of
the same layout, and responsibility for review and fixes. This parser implements
the local PDF and workbook steps. It does not by itself prove n8n orchestration,
client-layout compatibility, a paid engagement, a personally delivered client
project, or the owner's review. The buyer has not commissioned this work.

The extraction method follows the official pypdf text-extraction API. Its docs
explain that PDF tables do not carry a reliable semantic structure and that pypdf
does not perform OCR. That is why this proof fixes one known text layout.
[pypdf extraction documentation](https://pypdf.readthedocs.io/en/stable/user/extract-text.html)

Workbook creation and saving use openpyxl's documented API. Dependency licenses
remain applicable when handing over code: pypdf is BSD-3-Clause and openpyxl and
et_xmlfile use MIT. Owner approval of the client-specific source license is still
needed; these library licenses are not a contract with the buyer.
[openpyxl tutorial](https://openpyxl.readthedocs.io/en/3.1/tutorial.html),
[pypdf package](https://pypi.org/project/pypdf/),
[openpyxl package](https://pypi.org/project/openpyxl/),
[et_xmlfile package](https://pypi.org/project/et_xmlfile/)
