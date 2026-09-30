"""Generate deterministic synthetic text PDFs without third-party dependencies."""

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "fixtures" / "rfq-demo"


def pdf_bytes(pages):
    """Write a small valid PDF with an ordinary extractable Helvetica text layer."""
    objects = []
    page_refs = [4 + 2 * index for index in range(len(pages))]
    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    kids = " ".join(f"{number} 0 R" for number in page_refs)
    objects.append(f"<< /Type /Pages /Kids [{kids}] /Count {len(pages)} >>".encode())
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    for index, lines in enumerate(pages):
        page_number = page_refs[index]
        objects.append((f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
                        f"/Resources << /Font << /F1 3 0 R >> >> /Contents {page_number + 1} 0 R >>").encode())
        operations = ["BT", "/F1 9 Tf", "12 TL", "36 752 Td"]
        for line in lines:
            escaped = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            operations.extend([f"({escaped}) Tj", "T*"])
        operations.append("ET")
        stream = "\n".join(operations).encode("ascii") + b"\n"
        objects.append(b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"endstream")
    output = bytearray(b"%PDF-1.4\n% SYNTHETIC TEST FIXTURE\n")
    offsets = [0]
    for number, body in enumerate(objects, 1):
        offsets.append(len(output))
        output.extend(f"{number} 0 obj\n".encode() + body + b"\nendobj\n")
    startxref = len(output)
    output.extend(f"xref\n0 {len(objects) + 1}\n0000000000 65535 f \n".encode())
    for offset in offsets[1:]:
        output.extend(f"{offset:010d} 00000 n \n".encode())
    output.extend((f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\n"
                   f"startxref\n{startxref}\n%%EOF\n").encode())
    return bytes(output)


def write_json(name, value):
    (OUTPUT / name).write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def build():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    catalogue = {
        "version": "synthetic-catalogue-v1",
        "items": [
            {"code": "SYN-A1", "description": "SYNTHETIC steel bracket", "unit": "EA"},
            {"code": "SYN-B2", "description": "SYNTHETIC cable clip", "unit": "EA"},
            {"code": "SYN-C3", "description": "SYNTHETIC gasket", "unit": "EA"},
            {"code": "SYN-D4", "description": "SYNTHETIC generic bolt", "unit": "EA"},
            {"code": "SYN-E5", "description": "SYNTHETIC generic bolt", "unit": "EA"},
            {"code": "SYN-F6", "description": "=1+1", "unit": "EA"},
        ],
    }
    prices = {
        "version": "synthetic-prices-v1", "currency": "AED", "approved": True,
        "prices": [
            {"code": "SYN-A1", "unit": "EA", "unit_price": "12.50"},
            {"code": "SYN-B2", "unit": "EA", "unit_price": "0.005"},
            {"code": "SYN-D4", "unit": "EA", "unit_price": "1.25"},
            {"code": "SYN-E5", "unit": "EA", "unit_price": "1.30"},
            {"code": "SYN-F6", "unit": "EA", "unit_price": "4.00"},
        ],
    }
    write_json("catalogue.json", catalogue)
    write_json("prices.json", prices)
    duplicate = dict(catalogue, items=catalogue["items"] + [dict(catalogue["items"][0])])
    write_json("duplicate-catalogue.json", duplicate)
    duplicate_prices = dict(prices, prices=prices["prices"] + [dict(prices["prices"][0], unit_price="99.00")])
    write_json("duplicate-prices.json", duplicate_prices)
    write_json("unapproved-prices.json", dict(prices, approved=False))
    write_json("changed-prices.json", dict(prices, version="synthetic-prices-v2", prices=[dict(row, unit_price="13.00") if row["code"] == "SYN-A1" else row for row in prices["prices"]]))
    write_json("precision-prices.json", dict(prices, version="synthetic-precision-v1", prices=[dict(row, unit_price="999995000.000001") if row["code"] == "SYN-A1" else row for row in prices["prices"]]))
    header = ["SYNTHETIC RFQ V1", "ITEMS: 2", "ROW | CODE | DESCRIPTION | QTY | UNIT"]
    baseline_rows = [
        "1 | SYN-A1 | SYNTHETIC steel bracket | 2 | EA",
        "2 | SYN-B2 | SYNTHETIC cable clip | 3 | EA",
    ]
    docs = {
        "baseline.pdf": [header + baseline_rows + ["END RFQ"]],
        "unseen.pdf": [[
            "SYNTHETIC RFQ V1", "ITEMS: 8", "ROW | CODE | DESCRIPTION | QTY | UNIT",
            "1 | SYN-A1 | SYNTHETIC steel bracket | 1.2 | EA",
            "2 | SYN-UNKNOWN | SYNTHETIC unknown item | 2 | EA",
            "3 |  | SYNTHETIC cable clip | 1 | EA",
            "4 | SYN-A1 | SYNTHETIC steel bracket | 2 | M",
            "5 | SYN-C3 | SYNTHETIC gasket | 2 | EA",
            "6 | SYN-A1 | SYNTHETIC conflicting widget | 1 | EA",
            "7 |  | SYNTHETIC generic bolt | 3 | EA",
            "8 | SYN-F6 | =1+1 | 2 | EA", "END RFQ",
        ]],
        "multipage.pdf": [
            ["SYNTHETIC RFQ V1", "ITEMS: 3", "ROW | CODE | DESCRIPTION | QTY | UNIT", baseline_rows[0]],
            ["ROW | CODE | DESCRIPTION | QTY | UNIT", baseline_rows[1], "3 | SYN-A1 | SYNTHETIC steel bracket | 1 | EA", "END RFQ"],
        ],
        "count-mismatch.pdf": [["SYNTHETIC RFQ V1", "ITEMS: 3", "ROW | CODE | DESCRIPTION | QTY | UNIT"] + baseline_rows + ["END RFQ"]],
        "unrecognized-line.pdf": [header + baseline_rows + ["UNRECOGNIZED SYNTHETIC EXTRA ITEM 500 AED", "END RFQ"]],
        "parse-exception.pdf": [["SYNTHETIC RFQ V1", "ITEMS: 3", "ROW | CODE | DESCRIPTION | QTY | UNIT"] + baseline_rows + ["UNRECOGNIZED SYNTHETIC EXTRA ITEM 500 AED", "END RFQ"]],
        "duplicate-rfq-row.pdf": [header + [baseline_rows[0], "1 | SYN-A1 | SYNTHETIC steel bracket | 3 | EA", "END RFQ"]],
        "precision-boundary.pdf": [["SYNTHETIC RFQ V1", "ITEMS: 1", "ROW | CODE | DESCRIPTION | QTY | UNIT", "1 | SYN-A1 | SYNTHETIC steel bracket | 999999999.999999 | EA", "END RFQ"]],
        "no-text.pdf": [[]],
    }
    for name, pages in docs.items():
        (OUTPUT / name).write_bytes(pdf_bytes(pages))
    write_json("expected.json", {
        "synthetic": True,
        "currency": "AED",
        "rounding": "ROUND_HALF_UP per line to two decimal places; sum rounded lines",
        "baseline": {"rows": 2, "priced_rows": 2, "line_totals": {"1": "25.00", "2": "0.02"}, "total": "25.02"},
        "unseen": {"rows": 8, "priced_row_numbers": [1, 8], "line_totals": {"1": "15.00", "8": "8.00"}, "total": "23.00", "unpriced_reasons": {"2": "unknown code", "3": "description-only requires review", "4": "unit mismatch", "5": "missing approved price", "6": "description conflicts with code", "7": "ambiguous description"}},
        "multipage": {"rows": 3, "total": "37.52"},
        "changed_price_baseline": {"total": "26.02", "price_version": "synthetic-prices-v2"},
        "not_demonstrated": ["real vendor PDF compatibility", "OCR", "AI matching accuracy", "customer savings", "production quote approval"],
    })


if __name__ == "__main__":
    build()
