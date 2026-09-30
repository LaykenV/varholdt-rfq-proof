"""Independent black-box acceptance checks against synthetic PDF and XLSX files."""

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

try:
    import openpyxl
    from pypdf import PdfReader
except ModuleNotFoundError:
    openpyxl = None
    PdfReader = None

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = ROOT / "fixtures" / "rfq-demo"


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


@unittest.skipUnless(openpyxl is not None and PdfReader is not None,
                     "run with .venv-rfq/bin/python after installing requirements-rfq.txt")
class RFQAcceptanceTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.output = Path(self.temporary.name) / "outputs"

    def tearDown(self):
        self.temporary.cleanup()

    def invoke(self, pdf="baseline.pdf", catalogue="catalogue.json", prices="prices.json", run_id="test", succeeds=True):
        sources = [FIXTURES / pdf, FIXTURES / catalogue, FIXTURES / prices]
        before = {str(p): digest(p) for p in sources}
        process = subprocess.run(
            [sys.executable, str(ROOT / "src" / "rfq_quote.py"), *map(str, sources),
             str(self.output), "--run-id", run_id],
            cwd=ROOT, text=True, capture_output=True, timeout=30,
        )
        self.assertEqual(process.returncode, 0 if succeeds else 1, process.stderr + process.stdout)
        self.assertEqual(before, {str(p): digest(p) for p in sources}, "an input changed")
        state = json.loads((self.output / "last-run.json").read_text())
        self.assertEqual(state["run_id"], run_id)
        self.assertEqual(state["status"], "succeeded" if succeeds else "failed")
        if not succeeds:
            self.assertNotIn("workbook", state)
            return state
        run_dir = self.output / f"run-{run_id}"
        self.assertEqual(Path(state["workbook"]), run_dir / "quote.xlsx")
        result = json.loads((run_dir / "result.json").read_text())
        manifest = json.loads((run_dir / "manifest.json").read_text())
        self.assertTrue(result["synthetic"])
        self.assertTrue(result["draft_only"])
        for name, expected_hash in manifest["files"].items():
            self.assertEqual(digest(run_dir / name), expected_hash)
        for source, copied in zip(sources, ["input.pdf", "catalogue.json", "approved-prices.json"]):
            self.assertEqual(source.read_bytes(), (run_dir / copied).read_bytes())
        self.assertEqual(result["pdf_sha256"], digest(sources[0]))
        self.assertEqual(result["catalogue_sha256"], digest(sources[1]))
        self.assertEqual(result["price_sha256"], digest(sources[2]))
        self.assertEqual(result["input_rows"], len(result["items"]))
        self.assertEqual(result["input_rows"], result["exact_rows"] + result["review_rows"])
        for row in result["items"]:
            self.assertEqual(row["price_sha256"], result["price_sha256"])
            self.assertEqual(row["catalogue_sha256"], result["catalogue_sha256"])
            self.assertEqual(row["price_version"], result["price_version"])
            self.assertEqual(row["catalogue_version"], result["catalogue_version"])
            self.assertEqual(row["currency"], "AED")
        return run_dir, result

    def assert_workbook_matches(self, run_dir, result):
        workbook = openpyxl.load_workbook(run_dir / "quote.xlsx", data_only=False)
        try:
            sheet = workbook["Draft مسودة"]
            self.assertEqual(sheet.max_row, len(result["items"]) + 2)
            self.assertIn("SYNTHETIC", sheet["A1"].value)
            for position, item in enumerate(result["items"], 3):
                for column, key in [(1, "row_id"), (2, "code"), (3, "description"), (4, "quantity"),
                                    (5, "unit"), (6, "status"), (9, "unit_price"), (10, "line_total"),
                                    (11, "page"), (12, "text_line"), (13, "price_version"),
                                    (14, "price_sha256"), (15, "catalogue_version"),
                                    (16, "catalogue_sha256"), (17, "raw")]:
                    expected = None if item[key] in (None, "") else str(item[key])
                    self.assertEqual(sheet.cell(position, column).value, expected, f"item {position}, {key}")
            for sheet in workbook:
                for row in sheet:
                    for cell in row:
                        self.assertNotEqual(cell.data_type, "f", f"formula in {sheet.title}:{cell.coordinate}")
            audit = workbook["Source text المصدر"]
            self.assertEqual(audit.max_row, len(result["extracted_text_audit"]) + 1)
            for position, entry in enumerate(result["extracted_text_audit"], 2):
                self.assertEqual([audit.cell(position, column).value for column in range(1, 5)],
                                 [str(entry[key]) for key in ["page", "text_line", "kind", "raw"]])
            exceptions = workbook["Exceptions مراجعة"]
            expected_exceptions = [row for row in result["items"] if row["status"] == "REVIEW"]
            self.assertEqual(exceptions.max_row, len(expected_exceptions) + 2)
            for position, item in enumerate(expected_exceptions, 3):
                self.assertEqual(exceptions.cell(position, 1).value, item["row_id"] or None)
                self.assertIsNone(exceptions.cell(position, 9).value)
                self.assertIsNone(exceptions.cell(position, 10).value)
        finally:
            workbook.close()

    def test_baseline_arithmetic_provenance_and_workbook(self):
        run_dir, result = self.invoke()
        self.assertEqual((result["input_rows"], result["exact_rows"], result["review_rows"]), (2, 2, 0))
        self.assertEqual([r["line_total"] for r in result["items"]], ["25.00", "0.02"])
        self.assertEqual(result["items"][1]["unit_price"], "0.005")
        self.assertEqual(result["priced_subtotal"], "25.02")
        self.assertEqual(result["price_version"], "synthetic-prices-v1")
        self.assert_workbook_matches(run_dir, result)

    def test_unseen_rows_are_preserved_without_guessing_or_formulas(self):
        run_dir, result = self.invoke(pdf="unseen.pdf")
        self.assertEqual((result["input_rows"], result["exact_rows"], result["review_rows"]), (8, 2, 6))
        self.assertEqual(result["priced_subtotal"], "23.00")
        items = {r["row_id"]: r for r in result["items"]}
        self.assertEqual(set(items), {str(i) for i in range(1, 9)})
        reasons = {
            "2": "unknown_or_similar_code_requires_review",
            "3": "description_only_requires_review",
            "4": "unit_differs_from_catalogue",
            "5": "missing_approved_price",
            "6": "description_conflicts_with_exact_code",
            "7": "description_only_requires_review",
        }
        for row_id, reason in reasons.items():
            self.assertEqual(items[row_id]["status"], "REVIEW")
            self.assertIn(reason, items[row_id]["review_reasons"])
            self.assertIsNone(items[row_id]["unit_price"])
            self.assertIsNone(items[row_id]["line_total"])
        self.assertIn("SYN-B2", items["3"]["suggested_codes"])
        self.assertTrue({"SYN-D4", "SYN-E5"} <= set(items["7"]["suggested_codes"]))
        self.assertEqual(items["8"]["description"], "=1+1")
        self.assertEqual(items["8"]["line_total"], "8.00")
        self.assert_workbook_matches(run_dir, result)

    def test_allowed_numeric_boundary_does_not_double_round(self):
        run_dir, result = self.invoke(pdf="precision-boundary.pdf", prices="precision-prices.json")
        # Exact product is 999995000000000000.004999999999, below the half-cent.
        self.assertEqual(result["items"][0]["line_total"], "999995000000000000.00")
        self.assertEqual(result["priced_subtotal"], "999995000000000000.00")
        self.assert_workbook_matches(run_dir, result)

    def test_multipage_source_row_accounting(self):
        run_dir, result = self.invoke(pdf="multipage.pdf")
        self.assertEqual(result["priced_subtotal"], "37.52")
        self.assertEqual([r["page"] for r in result["items"]], [1, 2, 2])
        extracted = [page.extract_text() for page in PdfReader(FIXTURES / "multipage.pdf").pages]
        expected_raw = [line for text in extracted for line in text.splitlines() if line.strip()]
        self.assertEqual([r["raw"] for r in result["extracted_text_audit"]], expected_raw)
        self.assert_workbook_matches(run_dir, result)

    def test_unrecognized_text_is_visible_or_fails_the_batch(self):
        run_dir, result = self.invoke(pdf="parse-exception.pdf", run_id="preserved")
        self.assertEqual((result["input_rows"], result["exact_rows"], result["review_rows"]), (3, 2, 1))
        exception = result["items"][-1]
        self.assertIn("unrecognized_layout_row", exception["review_reasons"])
        self.assertEqual(exception["raw"], "UNRECOGNIZED SYNTHETIC EXTRA ITEM 500 AED")
        self.assertIsNone(exception["line_total"])
        self.assert_workbook_matches(run_dir, result)
        for pdf in ["unrecognized-line.pdf", "count-mismatch.pdf", "no-text.pdf"]:
            with self.subTest(pdf=pdf):
                self.invoke(pdf=pdf, run_id=pdf.replace(".pdf", ""), succeeds=False)

    def test_duplicate_catalogue_prices_or_unapproved_prices_fail(self):
        for key, kwargs in [
            ("catalogue", {"catalogue": "duplicate-catalogue.json"}),
            ("prices", {"prices": "duplicate-prices.json"}),
            ("unapproved", {"prices": "unapproved-prices.json"}),
        ]:
            with self.subTest(key=key):
                self.invoke(run_id=key, succeeds=False, **kwargs)
                self.assertFalse((self.output / f"run-{key}" / "quote.xlsx").exists())

    def test_duplicate_rfq_rows_are_both_unpriced(self):
        run_dir, result = self.invoke(pdf="duplicate-rfq-row.pdf")
        self.assertEqual((result["input_rows"], result["exact_rows"], result["review_rows"]), (2, 0, 2))
        self.assertEqual(result["priced_subtotal"], "0.00")
        self.assertTrue(all("duplicate_or_conflicting_row_id" in r["review_reasons"] for r in result["items"]))
        self.assertTrue(all(r["line_total"] is None for r in result["items"]))
        self.assert_workbook_matches(run_dir, result)

    def test_price_revision_changes_amounts_and_same_version_reuse_fails(self):
        first_dir, first = self.invoke(run_id="v1")
        second_dir, second = self.invoke(prices="changed-prices.json", run_id="v2")
        self.assertEqual(first["priced_subtotal"], "25.02")
        self.assertEqual(second["priced_subtotal"], "26.02")
        self.assertNotEqual(first["price_sha256"], second["price_sha256"])
        self.assertEqual(second["price_version"], "synthetic-prices-v2")
        self.assert_workbook_matches(second_dir, second)
        altered = json.loads((FIXTURES / "changed-prices.json").read_text())
        altered["version"] = "synthetic-prices-v1"
        local = Path(self.temporary.name) / "altered-prices.json"
        local.write_text(json.dumps(altered))
        state = self.invoke(prices=str(local), run_id="collision", succeeds=False)
        self.assertIn("version was reused", state["error"])
        self.assertEqual((first_dir / "approved-prices.json").read_bytes(), (FIXTURES / "prices.json").read_bytes())

    def test_failed_run_cannot_present_prior_workbook_as_current(self):
        run_dir, _ = self.invoke(run_id="good")
        before = digest(run_dir / "quote.xlsx")
        self.invoke(pdf="count-mismatch.pdf", run_id="bad", succeeds=False)
        self.assertEqual(digest(run_dir / "quote.xlsx"), before)
        self.assertFalse((self.output / "run-bad").exists())
        self.invoke(run_id="good", succeeds=False)
        self.assertEqual(digest(run_dir / "quote.xlsx"), before)


if __name__ == "__main__":
    unittest.main()
