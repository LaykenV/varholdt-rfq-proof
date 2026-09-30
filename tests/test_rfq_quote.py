import copy
from decimal import getcontext
from io import BytesIO
import json
from pathlib import Path
import sys
import tempfile
import unittest

from openpyxl import load_workbook
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from rfq_quote import RFQError, indexes, parse_pdf, quote_items, run


def make_pdf(lines):
    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    font = DictionaryObject({NameObject("/Type"): NameObject("/Font"), NameObject("/Subtype"): NameObject("/Type1"),
                             NameObject("/BaseFont"): NameObject("/Helvetica")})
    page[NameObject("/Resources")] = DictionaryObject({NameObject("/Font"): DictionaryObject({NameObject("/F1"): font})})
    escaped = [line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)") for line in lines]
    data = "BT /F1 10 Tf 40 740 Td 14 TL " + " T* ".join(f"({line}) Tj" for line in escaped) + " ET"
    stream = DecodedStreamObject()
    stream.set_data(data.encode("ascii"))
    page[NameObject("/Contents")] = stream
    output = BytesIO()
    writer.write(output)
    return output.getvalue()


def pdf_rows(rows, declared=None):
    return make_pdf(["SYNTHETIC RFQ V1", f"ITEMS: {len(rows) if declared is None else declared}",
                     "ROW | CODE | DESCRIPTION | QTY | UNIT", *rows, "END RFQ"])


class RFQTests(unittest.TestCase):
    def setUp(self):
        self.catalogue = {"version": "synthetic-cat-v1", "items": [{"code": "P001", "description": "Synthetic part", "unit": "EA"}]}
        self.prices = {"version": "synthetic-price-v1", "approved": True, "currency": "AED",
                       "prices": [{"code": "P001", "unit": "EA", "unit_price": "0.005"}]}

    def test_pdf_fields_and_source_rows_preserved_with_half_up_rounding(self):
        items, audit = parse_pdf(pdf_rows(["1 | P001 | Synthetic part | 1.00 | EA"]))
        self.assertEqual((items[0]["page"], items[0]["text_line"]), (1, 4))
        self.assertEqual(len(audit), 5)
        quoted, total = quote_items(items, self.catalogue, self.prices)
        self.assertEqual(quoted[0]["line_total"], "0.01")
        self.assertEqual(quoted[0]["unit_price"], "0.005")
        self.assertEqual(total, "0.01")
        self.assertEqual(quoted[0]["raw"], items[0]["raw"])
        self.prices["prices"][0]["unit_price"] = "0.145"
        self.assertEqual(quote_items(items, self.catalogue, self.prices)[1], "0.15")

    def test_uncertain_missing_conflicting_and_invalid_rows_never_priced(self):
        lines = ["1 | | Synthetic part | 1 | EA", "2 | P00I | Synthetic part | 1 | EA",
                 "3 | P001 | Different part | 1 | EA", "4 | P001 | Synthetic part | 1 | BOX",
                 "5 | P001 | Synthetic part | -1 | EA", "6 | P001 | Synthetic part | NaN | EA",
                 "7 | P001 | Synthetic part | 99999999999999999999999 | EA"]
        items, _ = parse_pdf(pdf_rows(lines))
        result, total = quote_items(items, self.catalogue, self.prices)
        self.assertEqual(len(result), len(lines))
        self.assertEqual(total, "0.00")
        self.assertTrue(all(row["status"] == "REVIEW" and row["unit_price"] is None and row["line_total"] is None for row in result))
        self.assertEqual(result[0]["suggested_codes"], ["P001"])

    def test_allowed_range_does_not_round_intermediate_product(self):
        items, _ = parse_pdf(pdf_rows(["1 | P001 | Synthetic part | 999999999.999999 | EA"]))
        self.prices["prices"][0]["unit_price"] = "999995000.000001"
        before = getcontext().prec
        quoted, total = quote_items(items, self.catalogue, self.prices)
        self.assertEqual(quoted[0]["line_total"], "999995000000000000.00")
        self.assertEqual(total, "999995000000000000.00")
        self.assertEqual(getcontext().prec, before)

    def test_duplicate_row_ids_and_missing_prices_are_exceptions(self):
        items, _ = parse_pdf(pdf_rows(["1 | P001 | Synthetic part | 1 | EA", "1 | P001 | Synthetic part | 2 | EA"]))
        result, _ = quote_items(items, self.catalogue, self.prices)
        self.assertTrue(all("duplicate_or_conflicting_row_id" in r["review_reasons"] and r["line_total"] is None for r in result))
        self.prices["prices"] = []
        items, _ = parse_pdf(pdf_rows(["1 | P001 | Synthetic part | 1 | EA"]))
        self.assertIn("missing_approved_price", quote_items(items, self.catalogue, self.prices)[0][0]["review_reasons"])

    def test_invalid_metadata_rejected(self):
        for changed in [{"currency": "USD"}, {"approved": False}, {"version": "../escape"},
                        {"prices": self.prices["prices"] * 2}]:
            with self.subTest(changed=changed), self.assertRaises(RFQError):
                indexes(self.catalogue, dict(self.prices, **changed))
        self.catalogue["items"] *= 2
        with self.assertRaisesRegex(RFQError, "duplicate catalogue"):
            indexes(self.catalogue, self.prices)

    def test_unrecognized_rows_are_retained_and_count_mismatch_fails(self):
        items, audit = parse_pdf(pdf_rows(["unparseable original row"]))
        self.assertEqual(items[0]["raw"], "unparseable original row")
        self.assertEqual(items[0]["parse_issue"], "unrecognized_layout_row")
        self.assertEqual(audit[3]["kind"], "parse_exception")
        with self.assertRaisesRegex(RFQError, "count"):
            parse_pdf(pdf_rows(["1 | P001 | Synthetic part | 1 | EA"], declared=2))
        with self.assertRaises(RFQError):
            parse_pdf(b"not a PDF")

    def test_workbook_treats_untrusted_text_as_strings_and_audits_inputs(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            description = '=HYPERLINK("https://example.invalid","synthetic")'
            cat = copy.deepcopy(self.catalogue)
            cat["items"][0]["description"] = description
            (root / "rfq.pdf").write_bytes(pdf_rows([f"1 | P001 | {description} | 1 | EA",
                                                     "2 | UNKNOWN | Synthetic unknown | 2 | EA",
                                                     "3 | | Synthetic part | 1 | EA"]))
            (root / "cat.json").write_text(json.dumps(cat))
            (root / "prices.json").write_text(json.dumps(self.prices))
            state = run(root / "rfq.pdf", root / "cat.json", root / "prices.json", root / "out", "first")
            book = load_workbook(state["workbook"], data_only=False)
            self.assertEqual(book["Draft مسودة"]["C3"].value, description)
            self.assertEqual(book["Draft مسودة"]["C3"].data_type, "s")
            self.assertFalse(any(cell.data_type == "f" for sheet in book for row in sheet for cell in row))
            result = json.loads(Path(state["result"]).read_text())
            self.assertEqual(result["input_rows"], result["exact_rows"] + result["review_rows"])
            self.assertEqual(result["items"][0]["price_sha256"], result["price_sha256"])
            self.assertIn("الوصف", book["Draft مسودة"]["C2"].value)
            exceptions = book["Exceptions مراجعة"]
            self.assertEqual(exceptions.max_row - 2, result["review_rows"])
            self.assertEqual([exceptions.cell(i, 1).value for i in range(3, exceptions.max_row + 1)],
                             [item["row_id"] for item in result["items"] if item["status"] == "REVIEW"])
            self.assertTrue(all(exceptions.cell(i, 9).value is None and exceptions.cell(i, 10).value is None
                                for i in range(3, exceptions.max_row + 1)))
            book.close()

    def test_failure_clears_latest_success_pointer_and_preserves_old_run(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "rfq.pdf").write_bytes(pdf_rows(["1 | P001 | Synthetic part | 1 | EA"]))
            (root / "cat.json").write_text(json.dumps(self.catalogue))
            price_path = root / "prices.json"
            price_path.write_text(json.dumps(self.prices))
            args = [root / "rfq.pdf", root / "cat.json", price_path, root / "out"]
            good = run(*args, "good")
            original = Path(good["workbook"]).read_bytes()
            self.prices["prices"][0]["unit_price"] = "10"
            price_path.write_text(json.dumps(self.prices))
            with self.assertRaisesRegex(RFQError, "reused"):
                run(*args, "failed")
            latest = json.loads((root / "out" / "last-run.json").read_text())
            self.assertEqual(latest["status"], "failed")
            self.assertNotIn("workbook", latest)
            self.assertFalse((root / "out" / "run-failed").exists())
            self.assertEqual(Path(good["workbook"]).read_bytes(), original)
            with self.assertRaisesRegex(RFQError, "already exists"):
                run(*args, "good")


if __name__ == "__main__":
    unittest.main()
