# Run the complete RFQ demonstration

This demonstration runs an actual n8n workflow on this machine. Its HTTP nodes
fetch a synthetic PDF, pass those bytes to a local Python parser, and receive
an Excel draft. The verification script compares the returned workbook bytes
with the parser's output. It does not activate a public webhook or run a
background service after the test.

Read [the parser contract](rfq-quote.md) first. Only the documented synthetic
PDF layout is implemented. A customer's PDF still needs sample review and
adaptation. These fixtures do not establish a customer's acceptance.

## Local installation

Use the Python environment from the parser instructions. Install Node and n8n
inside this project's ignored tooling directory:

```bash
mkdir -p tmp/rfq-n8n
npm install --prefix tmp/rfq-n8n node@24.21.0 --no-audit --no-fund
PATH="$PWD/tmp/rfq-n8n/node_modules/node/bin:$PATH" npm install --prefix tmp/rfq-n8n n8n@2.41.3 --no-audit --no-fund
```

The tested n8n release requires Node 24 or later. The isolated installation
leaves the machine's default Node unchanged. Package installation uses the
network. The synthetic workflow calls only the local loopback service.

The workflow uses the built-in Manual Trigger and HTTP Request nodes. It does
not enable Execute Command, use a community node, require a model API, or reuse
an existing n8n account. The generated n8n database stays inside the output
directory and contains no customer credentials.

## Execute and verify

Run these commands from the repository root. Each output path must be new:

```bash
.venv-rfq/bin/python scripts/verify_rfq_n8n.py fixtures/rfq-demo/baseline.pdf fixtures/rfq-demo/catalogue.json fixtures/rfq-demo/prices.json artifacts/private/rfq-proof-baseline
.venv-rfq/bin/python scripts/verify_rfq_n8n.py fixtures/rfq-demo/unseen.pdf fixtures/rfq-demo/catalogue.json fixtures/rfq-demo/prices.json artifacts/private/rfq-proof-unseen
.venv-rfq/bin/python scripts/verify_rfq_n8n.py fixtures/rfq-demo/count-mismatch.pdf fixtures/rfq-demo/catalogue.json fixtures/rfq-demo/prices.json artifacts/private/rfq-proof-rejection --expect-rejection
```

Run these checks one at a time. n8n's task broker uses local port 5679 even
though each check has a separate database and parser port. Concurrent checks
can fail before executing the workflow because that broker port is already
in use. A broker startup failure is a failed test, never a passing rejection.

The script starts an ephemeral HTTP service bound to `127.0.0.1`. It serves
only the selected synthetic input. The POST route accepts only the same PDF
bytes, invokes the parser with fixed reference paths, and returns the workbook.
This is a test adapter, not a production upload API. It rejects arbitrary input
changes. It has no customer authentication, tenancy, queue, or hosted delivery.

Next, the script creates an inactive workflow, imports it into a fresh local
n8n database, and runs it through n8n's CLI. Each success requires all three
workflow nodes to execute, the expected HTTP request sequence, unchanged PDF
bytes, and a returned workbook with the same SHA256 as the parser output.
An expected rejection requires an actual failed n8n execution, failed parser
state, and no workbook. An import failure cannot count as an expected rejection.

On success, inspect `verification.json`, `n8n-returned.xlsx`, and
`quote-output/run-n8n-proof/result.json`. The generated `workflow.json` is the
exact executed workflow, including its ephemeral local port. Logs and the
local n8n database are private diagnostic artifacts. They are not needed in a
buyer attachment. The script shuts down its service before it exits.

[The editable workflow template](../workflows/rfq-local-demo.json) uses port
8876 as a placeholder. The verifier generates a working copy with the test
server's assigned port. Importing the template alone does not start a parser
service. A customer installation needs agreed service hosting and access.

## Claims and handoff

The Python source and workflow are original demonstration material. The proposed
client rights and support terms require owner approval. Dependency licenses
remain separate. We do not bundle the n8n runtime in the proof archive or
represent it as our code. Consult n8n's license for the buyer's intended hosting
model before committing to a deployment.

The workbook has English and Arabic headings. Automated checks read its actual
cells, types, row counts, hashes, and amounts. They do not establish Arabic
translation quality or compatibility with every spreadsheet application.

Official references checked September 29, 2026:

- [n8n Server CLI](https://docs.n8n.io/deploy/host-n8n/configure-n8n/use-the-command-line) documents workflow import and execution by ID.
- [HTTP Request node](https://docs.n8n.io/integrations/builtin/core-nodes/n8n-nodes-base.httprequest/) documents binary request bodies and file responses.
- [n8n license](https://github.com/n8n-io/n8n/blob/master/LICENSE.md) governs the separate runtime.
