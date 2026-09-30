"""Synthetic-layout text PDF to a review-only AED quotation workbook."""

import argparse
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal, ROUND_HALF_UP, localcontext
from difflib import SequenceMatcher
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import re
import shutil
import tempfile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pypdf import PdfReader


class RFQError(ValueError):
    pass


def require(condition, message):
    if not condition:
        raise RFQError(message)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def decimal_value(value, *, positive=False):
    require(isinstance(value, str) and re.fullmatch(r"\d{1,9}(?:\.\d{1,6})?", value),
            "number must have at most nine integer digits and six decimals")
    result = Decimal(value)
    require(not positive or result > 0, "quantity must be positive")
    return result


def cents(value):
    with localcontext() as context:
        context.prec = 60
        return str(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def norm(value):
    return " ".join(value.split()).casefold()


def parse_pdf(pdf_bytes):
    require(len(pdf_bytes) <= 10_000_000, "PDF exceeds 10 MB prototype limit")
    try:
        reader = PdfReader(BytesIO(pdf_bytes), strict=True)
        require(not reader.is_encrypted, "encrypted PDFs are outside this prototype")
        require(1 <= len(reader.pages) <= 10, "PDF needs one to ten pages")
        texts = []
        for page in reader.pages:
            content = page.get_contents()
            require(content is not None and len(content.get_data()) <= 5_000_000,
                    "missing or oversized PDF page stream")
            text = page.extract_text()
            require(text and text.strip(), "image-only or empty page needs another extraction scope")
            texts.append(text)
    except RFQError:
        raise
    except Exception as exc:
        raise RFQError(f"PDF extraction failed: {type(exc).__name__}") from exc
    items, audit, declared = [], [], []
    saw_title = saw_table = False
    ended = False
    for page_number, text in enumerate(texts, 1):
        for line_number, raw in enumerate(text.splitlines(), 1):
            stripped = raw.strip()
            if not stripped:
                continue
            entry = {"page": page_number, "text_line": line_number, "raw": raw}
            if stripped == "SYNTHETIC RFQ V1" and not ended:
                saw_title = True
                entry["kind"] = "layout_header"
            elif re.fullmatch(r"ITEMS: \d+", stripped) and not ended:
                declared.append(int(stripped.split(":")[1]))
                entry["kind"] = "declared_count"
            elif [x.strip() for x in stripped.split("|")] == ["ROW", "CODE", "DESCRIPTION", "QTY", "UNIT"] and not ended:
                saw_table = True
                entry["kind"] = "table_header"
            elif stripped == "END RFQ" and not ended:
                ended = True
                entry["kind"] = "footer"
            else:
                fields = [x.strip() for x in raw.split("|")]
                issue = None if len(fields) == 5 and saw_title and saw_table and not ended else "unrecognized_layout_row"
                if len(fields) != 5:
                    fields = ["", "", "", "", ""]
                item = dict(zip(("row_id", "code", "description", "quantity", "unit"), fields))
                item.update(page=page_number, text_line=line_number, raw=raw, parse_issue=issue)
                items.append(item)
                entry["kind"] = "item" if issue is None else "parse_exception"
            audit.append(entry)
    require(saw_title and saw_table and ended, "PDF does not match the declared synthetic layout")
    require(len(declared) == 1 and declared[0] == len(items), "declared item count does not reconcile with extracted rows")
    require(1 <= len(items) <= 200, "PDF needs one to 200 item rows")
    return items, audit


def indexes(catalogue, prices):
    for data in (catalogue, prices):
        require(isinstance(data, dict) and isinstance(data.get("version"), str)
                and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,79}", data["version"]), "invalid data version")
    require(prices.get("currency") == "AED" and prices.get("approved") is True,
            "separate price list must explicitly be approved and use AED")
    catalogue_rows, price_rows = catalogue.get("items"), prices.get("prices")
    require(isinstance(catalogue_rows, list) and 1 <= len(catalogue_rows) <= 5000, "invalid catalogue size")
    require(isinstance(price_rows, list) and len(price_rows) <= 5000, "invalid price-list size")
    cats, rates = {}, {}
    for row in catalogue_rows:
        require(isinstance(row, dict) and all(isinstance(row.get(k), str) and row[k].strip()
                for k in ("code", "description", "unit")), "invalid catalogue row")
        require(row["code"] not in cats, "duplicate catalogue code")
        cats[row["code"]] = row
    for row in price_rows:
        require(isinstance(row, dict) and all(isinstance(row.get(k), str) and row[k].strip()
                for k in ("code", "unit", "unit_price")), "invalid approved price row")
        require(row["code"] not in rates, "duplicate approved price code")
        decimal_value(row["unit_price"])
        rates[row["code"]] = row
    return cats, rates


def quote_items(items, catalogue, prices):
    cats, rates = indexes(catalogue, prices)
    counts = Counter(item["row_id"] for item in items)
    quoted = []
    total = Decimal("0")
    for item in items:
        row = dict(item)
        issues = [item["parse_issue"]] if item["parse_issue"] else []
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", item["row_id"]):
            issues.append("missing_or_invalid_row_id")
        elif counts[item["row_id"]] > 1:
            issues.append("duplicate_or_conflicting_row_id")
        quantity = None
        try:
            quantity = decimal_value(item["quantity"], positive=True)
        except RFQError:
            issues.append("invalid_quantity")
        if not item["description"]:
            issues.append("missing_description")
        if not item["unit"]:
            issues.append("missing_unit")
        code = item["code"]
        suggestions = []
        if code not in cats:
            issues.append("description_only_requires_review" if not code else "unknown_or_similar_code_requires_review")
            if item["description"]:
                scored = sorted(((SequenceMatcher(None, norm(item["description"]), norm(c["description"])).ratio(), key)
                                 for key, c in cats.items()), key=lambda x: (-x[0], x[1]))
                suggestions = [key for score, key in scored[:3] if score >= 0.5]
        else:
            if norm(item["description"]) != norm(cats[code]["description"]):
                issues.append("description_conflicts_with_exact_code")
            if item["unit"] != cats[code]["unit"]:
                issues.append("unit_differs_from_catalogue")
            if code not in rates:
                issues.append("missing_approved_price")
            elif item["unit"] != rates[code]["unit"]:
                issues.append("unit_differs_from_approved_price")
        row.update(status="REVIEW" if issues else "EXACT_CODE_DRAFT", review_reasons=issues,
                   suggested_codes=suggestions, currency="AED", unit_price=None, line_total=None,
                   price_version=prices["version"], catalogue_version=catalogue["version"])
        if not issues:
            row["unit_price"] = rates[code]["unit_price"]
            with localcontext() as context:
                context.prec = 60
                row["line_total"] = cents(quantity * decimal_value(row["unit_price"]))
                total += Decimal(row["line_total"])
        quoted.append(row)
    return quoted, cents(total)


HEADINGS = [
    ("row_id", "Row / الصف"), ("code", "Code / الرمز"), ("description", "Description / الوصف"),
    ("quantity", "Quantity / الكمية"), ("unit", "Unit / الوحدة"), ("status", "Status / الحالة"),
    ("review_reasons", "Review reasons / أسباب المراجعة"), ("suggested_codes", "Candidates / خيارات للمراجعة"),
    ("unit_price", "Unit price AED / سعر الوحدة"), ("line_total", "Line total AED / إجمالي السطر"),
    ("page", "PDF page / صفحة المصدر"), ("text_line", "Text row / سطر المصدر"),
    ("price_version", "Price version / إصدار الأسعار"), ("price_sha256", "Price SHA256 / بصمة الأسعار"),
    ("catalogue_version", "Catalogue version / إصدار الكتالوج"),
    ("catalogue_sha256", "Catalogue SHA256 / بصمة الكتالوج"), ("raw", "Original text / النص الأصلي"),
]


def text_cell(sheet, row_number, column, value):
    cell = sheet.cell(row_number, column)
    cell.value = "" if value is None else str(value)
    # Explicit string typing prevents leading '=' from becoming an Excel formula.
    cell.data_type = "s"
    cell.alignment = Alignment(vertical="top", wrap_text=True)
    return cell


def workbook(result, path):
    book = Workbook()
    book.remove(book.active)
    for title, items in (("Draft مسودة", result["items"]),
                         ("Exceptions مراجعة", [r for r in result["items"] if r["status"] == "REVIEW"])):
        sheet = book.create_sheet(title)
        text_cell(sheet, 1, 1, "SYNTHETIC DEMO. DRAFT ONLY. HUMAN APPROVAL REQUIRED. / مسودة للمراجعة")
        for column, (_, heading) in enumerate(HEADINGS, 1):
            cell = text_cell(sheet, 2, column, heading)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17445B")
        for row_number, item in enumerate(items, 3):
            for column, (field, _) in enumerate(HEADINGS, 1):
                value = item[field]
                if isinstance(value, list):
                    value = "; ".join(value)
                text_cell(sheet, row_number, column, value)
        sheet.freeze_panes = "D3"
        sheet.auto_filter.ref = f"A2:Q{len(items) + 2}"
        for column in range(1, len(HEADINGS) + 1):
            sheet.column_dimensions[sheet.cell(2, column).column_letter].width = 24 if column not in (3, 7, 17) else 45
    summary = book.create_sheet("Summary ملخص")
    for position, (key, value) in enumerate([
        ("Status / الحالة", "SYNTHETIC DRAFT. Not a customer quote."),
        ("Currency / العملة", "AED"), ("Input rows / صفوف الإدخال", len(result["items"])),
        ("Exact draft rows / الصفوف المطابقة", result["exact_rows"]),
        ("Review rows / صفوف المراجعة", result["review_rows"]),
        ("Priced subtotal AED / المجموع الجزئي", result["priced_subtotal"]),
        ("Subtotal scope / نطاق المجموع", "Only exact draft rows. Unpriced review rows are excluded. Tax and delivery not included."),
        ("Price version / إصدار الأسعار", result["price_version"]),
        ("Price SHA256 / بصمة الأسعار", result["price_sha256"]),
        ("PDF SHA256 / بصمة الطلب", result["pdf_sha256"]),
    ], 1):
        text_cell(summary, position, 1, key)
        text_cell(summary, position, 2, value)
    summary.column_dimensions["A"].width = 38
    summary.column_dimensions["B"].width = 85
    audit = book.create_sheet("Source text المصدر")
    for column, key in enumerate(("page", "text_line", "kind", "raw"), 1):
        text_cell(audit, 1, column, key)
    for row_number, entry in enumerate(result["extracted_text_audit"], 2):
        for column, key in enumerate(("page", "text_line", "kind", "raw"), 1):
            text_cell(audit, row_number, column, entry[key])
    audit.column_dimensions["D"].width = 100
    book.save(path)


def write_json(path, value):
    data = (json.dumps(value, ensure_ascii=False, indent=2) + "\n").encode()
    handle, temporary = tempfile.mkstemp(prefix=".state-", dir=path.parent)
    try:
        with os.fdopen(handle, "wb") as stream:
            stream.write(data)
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


def run(pdf_path, catalogue_path, prices_path, output_dir, run_id):
    require(isinstance(run_id, str) and re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,79}", run_id), "invalid run ID")
    root = Path(output_dir).resolve()
    inputs = [Path(p).resolve() for p in (pdf_path, catalogue_path, prices_path)]
    state_path = root / "last-run.json"
    require(all(p != state_path for p in inputs), "state file would overwrite an input")
    root.mkdir(parents=True, exist_ok=True)
    state = {"run_id": run_id, "status": "running", "started_at": datetime.now(timezone.utc).isoformat()}
    write_json(state_path, state)
    run_path = root / f"run-{run_id}"
    created = False
    try:
        require(not run_path.exists(), "run ID already exists; choose a new ID")
        pdf_bytes, cat_bytes, price_bytes = [p.read_bytes() for p in inputs]
        catalogue, prices = json.loads(cat_bytes), json.loads(price_bytes)
        indexes(catalogue, prices)
        items, audit = parse_pdf(pdf_bytes)
        quoted, total = quote_items(items, catalogue, prices)
        version_path = root / "price-versions"
        version_path.mkdir(exist_ok=True)
        immutable = version_path / (prices["version"] + ".json")
        if immutable.exists():
            require(immutable.read_bytes() == price_bytes, "approved price version was reused with different bytes")
        else:
            with immutable.open("xb") as stream:
                stream.write(price_bytes)
        run_path.mkdir()
        created = True
        for name, data in (("input.pdf", pdf_bytes), ("catalogue.json", cat_bytes), ("approved-prices.json", price_bytes)):
            (run_path / name).write_bytes(data)
        for item in quoted:
            item.update(price_sha256=sha(price_bytes), catalogue_sha256=sha(cat_bytes))
        result = {"schema_version": 1, "synthetic": True, "draft_only": True, "run_id": run_id,
                  "currency": "AED", "items": quoted, "input_rows": len(items),
                  "exact_rows": sum(row["status"] == "EXACT_CODE_DRAFT" for row in quoted),
                  "review_rows": sum(row["status"] == "REVIEW" for row in quoted),
                  "priced_subtotal": total, "pdf_sha256": sha(pdf_bytes), "catalogue_sha256": sha(cat_bytes),
                  "price_sha256": sha(price_bytes), "price_version": prices["version"],
                  "catalogue_version": catalogue["version"], "extracted_text_audit": audit}
        write_json(run_path / "result.json", result)
        workbook(result, run_path / "quote.xlsx")
        manifest = {"run_id": run_id, "status": "succeeded", "draft_only": True, "synthetic": True,
                    "files": {p.name: sha(p.read_bytes()) for p in sorted(run_path.iterdir()) if p.is_file()},
                    "price_version": prices["version"], "price_sha256": sha(price_bytes)}
        write_json(run_path / "manifest.json", manifest)
        state.update(status="succeeded", workbook=str(run_path / "quote.xlsx"), manifest=str(run_path / "manifest.json"),
                     result=str(run_path / "result.json"))
        write_json(state_path, state)
        return state
    except Exception as exc:
        if created:
            shutil.rmtree(run_path)
        state.update(status="failed", error=f"{type(exc).__name__}: {exc}")
        write_json(state_path, state)
        raise RFQError(state["error"]) from exc


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ("pdf", "catalogue", "prices", "output"):
        parser.add_argument(name)
    parser.add_argument("--run-id", required=True)
    args = parser.parse_args()
    try:
        print(json.dumps(run(args.pdf, args.catalogue, args.prices, args.output, args.run_id)))
    except RFQError as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}))
        raise SystemExit(1)
