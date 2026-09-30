"""Run the synthetic RFQ through a real, isolated local n8n installation."""

import argparse
import base64
from datetime import datetime, timezone
import hashlib
from http.server import BaseHTTPRequestHandler, HTTPServer
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import uuid

ROOT = Path(__file__).resolve().parents[1]


def workflow(port):
    nodes = [
        {"id": "start", "name": "Manual start", "type": "n8n-nodes-base.manualTrigger",
         "typeVersion": 1, "position": [0, 0], "parameters": {}},
        {"id": "input", "name": "Read synthetic PDF", "type": "n8n-nodes-base.httpRequest",
         "typeVersion": 4.2, "position": [240, 0], "parameters": {
             "url": f"http://127.0.0.1:{port}/input.pdf",
             "options": {"response": {"response": {"responseFormat": "file", "outputPropertyName": "pdf"}}}}},
        {"id": "quote", "name": "Create draft Excel", "type": "n8n-nodes-base.httpRequest",
         "typeVersion": 4.2, "position": [480, 0], "parameters": {
             "method": "POST", "url": f"http://127.0.0.1:{port}/quote",
             "sendHeaders": True,
             "headerParameters": {"parameters": [{"name": "Content-Type", "value": "application/pdf"}]},
             "sendBody": True, "contentType": "binaryData", "inputDataFieldName": "pdf",
             "options": {"timeout": 60000, "response": {"response": {
                 "responseFormat": "file", "outputPropertyName": "quote"}}}}},
    ]
    return {"id": "rfqSyntheticProof", "name": "Synthetic RFQ PDF to draft Excel", "active": False,
            "nodes": nodes, "connections": {
                "Manual start": {"main": [[{"node": "Read synthetic PDF", "type": "main", "index": 0}]]},
                "Read synthetic PDF": {"main": [[{"node": "Create draft Excel", "type": "main", "index": 0}]]}},
            "settings": {"executionOrder": "v1"}}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf", type=Path)
    parser.add_argument("catalogue", type=Path)
    parser.add_argument("prices", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--n8n-root", type=Path, default=ROOT / "tmp/rfq-n8n")
    parser.add_argument("--expect-rejection", action="store_true")
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=False)
    out = args.output.resolve()
    pdf_bytes = args.pdf.read_bytes()
    source_hashes = {name: hashlib.sha256((ROOT / name).read_bytes()).hexdigest()
                     for name in ("src/rfq_quote.py", "scripts/verify_rfq_n8n.py", "requirements-rfq.txt")}
    if len(pdf_bytes) > 10_000_000:
        raise ValueError("PDF exceeds demo limit")
    requests = []
    result = {}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def send_bytes(self, status, body, mime):
            self.send_response(status)
            self.send_header("Content-Type", mime)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            if self.path != "/input.pdf":
                self.send_error(404)
                return
            requests.append("GET /input.pdf")
            self.send_bytes(200, pdf_bytes, "application/pdf")

        def do_POST(self):
            if self.path != "/quote":
                self.send_error(404)
                return
            try:
                size = int(self.headers.get("Content-Length", "0"))
                if not 0 < size <= 10_000_000:
                    raise ValueError("Invalid length")
                self.connection.settimeout(10)
                body = self.rfile.read(size)
                if body != pdf_bytes:
                    raise ValueError("n8n did not preserve the PDF bytes")
                requests.append("POST /quote")
                received = out / "received.pdf"
                received.write_bytes(body)
                command = [sys.executable, str(ROOT / "src/rfq_quote.py"), str(received),
                           str(args.catalogue.resolve()), str(args.prices.resolve()),
                           str(out / "quote-output"), "--run-id", "n8n-proof"]
                completed = subprocess.run(command, capture_output=True, text=True, timeout=45)
                (out / "parser.log").write_text(completed.stdout + completed.stderr)
                for name, digest in source_hashes.items():
                    if hashlib.sha256((ROOT / name).read_bytes()).hexdigest() != digest:
                        raise RuntimeError("Source changed during verification")
                if completed.returncode:
                    raise ValueError("Parser rejected input; see parser.log")
                book = out / "quote-output/run-n8n-proof/quote.xlsx"
                workbook = book.read_bytes()
                result["workbook_sha256"] = hashlib.sha256(workbook).hexdigest()
                result["workbook_bytes"] = len(workbook)
                self.send_bytes(200, workbook, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as exc:
                result["error"] = str(exc)
                self.send_bytes(422, b"RFQ validation failed. No quote returned.", "text/plain")

    server = HTTPServer(("127.0.0.1", 0), Handler)
    server.timeout = 15
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    config = workflow(server.server_port)
    workflow_file = out / "workflow.json"
    workflow_file.write_text(json.dumps(config, indent=2) + "\n")
    install = args.n8n_root.resolve()
    node = install / "node_modules/node/bin/node"
    n8n = install / "node_modules/n8n/bin/n8n"
    env = {key: os.environ[key] for key in ("HOME", "LANG", "TZ", "TMPDIR") if key in os.environ}
    env.update(PATH=str(node.parent) + os.pathsep + os.environ.get("PATH", ""),
               N8N_USER_FOLDER=str(out / "n8n-user"), N8N_DIAGNOSTICS_ENABLED="false",
               N8N_VERSION_NOTIFICATIONS_ENABLED="false", N8N_TEMPLATES_ENABLED="false",
               N8N_DEFAULT_BINARY_DATA_MODE="default", N8N_RUNNERS_ENABLED="false",
               N8N_ENCRYPTION_KEY=uuid.uuid4().hex)
    try:
        for label, command in [
            ("import", ["import:workflow", "--input", str(workflow_file)]),
            ("execute", ["execute", "--id", config["id"], "--rawOutput"])]:
            run = subprocess.run([str(node), str(n8n), *command], env=env,
                                 capture_output=True, text=True, timeout=180, cwd=ROOT)
            (out / f"{label}.log").write_text(run.stdout + run.stderr)
            if run.returncode:
                if label == "execute" and args.expect_rejection and result.get("error"):
                    state_file = out / "quote-output/last-run.json"
                    state = json.loads(state_file.read_text())
                    if state["status"] != "failed" or "workbook" in state:
                        raise RuntimeError("Failed parser left a current workbook pointer")
                    if list((out / "quote-output").rglob("*.xlsx")):
                        raise RuntimeError("Rejected PDF produced a workbook")
                    if requests != ["GET /input.pdf", "POST /quote"]:
                        raise RuntimeError("Rejection did not traverse the expected n8n path")
                    evidence = {"status": "expected_rejection_passed", "synthetic": True,
                                "verified_at": datetime.now(timezone.utc).isoformat(),
                                "source_sha256": source_hashes,
                                "n8n_version": json.loads((install / "node_modules/n8n/package.json").read_text())["version"],
                                "pdf_sha256": hashlib.sha256(pdf_bytes).hexdigest(),
                                "request_sequence": requests, "workbook_returned": False,
                                "parser_status": state["status"], "parser_error": state["error"]}
                    (out / "verification.json").write_text(json.dumps(evidence, indent=2) + "\n")
                    print(json.dumps(evidence, indent=2))
                    return
                raise RuntimeError(f"n8n {label} failed; inspect {out / (label + '.log')}")
        if args.expect_rejection:
            raise RuntimeError("n8n accepted a PDF that should have been rejected")
        # n8n may prefix diagnostics before the JSON execution result.
        decoder = json.JSONDecoder()
        execution = None
        for index, char in enumerate(run.stdout):
            if char != "{":
                continue
            try:
                value, _ = decoder.raw_decode(run.stdout[index:])
                if isinstance(value, dict) and "data" in value and "resultData" in value["data"]:
                    execution = value
                    break
            except (ValueError, TypeError):
                continue
        if execution is None:
            raise RuntimeError("No n8n execution result found")
        steps = execution["data"]["resultData"]["runData"]
        quote = steps["Create draft Excel"][-1]["data"]["main"][0][0]["binary"]["quote"]
        returned = base64.b64decode(quote["data"], validate=True)
        digest = hashlib.sha256(returned).hexdigest()
        if digest != result.get("workbook_sha256"):
            raise RuntimeError("n8n workbook bytes differ from parser output")
        if requests != ["GET /input.pdf", "POST /quote"]:
            raise RuntimeError("Unexpected HTTP request sequence")
        (out / "n8n-returned.xlsx").write_bytes(returned)
        evidence = {"status": "passed", "synthetic": True,
                    "verified_at": datetime.now(timezone.utc).isoformat(),
                    "source_sha256": source_hashes,
                    "n8n_version": json.loads((install / "node_modules/n8n/package.json").read_text())["version"],
                    "pdf_sha256": hashlib.sha256(pdf_bytes).hexdigest(), **result,
                    "request_sequence": requests, "executed_nodes": list(steps),
                    "n8n_returned_workbook_sha256": digest}
        (out / "verification.json").write_text(json.dumps(evidence, indent=2) + "\n")
        print(json.dumps(evidence, indent=2))
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
