"""
Extract "Employee Gross To Net" report data from PDF(s) and build, per source file,
two separate output workbooks:

    <filename>_extracted.xlsx  - every column found in the PDF (EE ID, Employee
                                  Name, SSN, Payment Date, Total Hrs/Units,
                                  Total Earnings, Total Deductions, Total Taxes,
                                  Net Pay)
    <filename>_template.xlsx   - the same rows mapped into the BDE Import
                                  Template layout (Doc ID, First/Middle/Last
                                  Name, Entity Type_Employee = TRUE, SSN),
                                  header rows copied from the template so
                                  column positions match

Rotation-aware: PyMuPDF returns word positions in correct reading order
regardless of a page's /Rotate flag (0/90/180/270), so no manual re-rendering
is needed for digitally generated PDFs.

Usage:
    python extract_and_build_bde.py <pdf_file_or_folder> [bde_template.xlsx] [-o output_dir]

Requires:
    pip install pymupdf openpyxl
"""

import re
import argparse
from pathlib import Path

import fitz  # PyMuPDF
import openpyxl
from openpyxl.utils import get_column_letter

SSN_PATTERN = re.compile(r"\d{3}-\d{2}-\d{4}|[*Xx]{3}-[*Xx]{2}-\d{4}")
EEID_PATTERN = re.compile(r"^[A-Za-z0-9-]{2,15}$")

# (column label, token sequence that identifies its header cell)
HEADER_DEFS = [
    ("EE ID", ["EE", "ID"]),
    ("Employee Name", ["Employee", "Name"]),
    ("SSN", ["SSN"]),
    ("Payment Date", ["Payment", "Date"]),
    ("Total Hrs/Units", ["Total", "Hrs/Units"]),
    ("Total Earnings", ["Total", "Earnings"]),
    ("Total Deductions", ["Total", "Deductions"]),
    ("Total Taxes", ["Total", "Taxes"]),
    ("Net Pay", ["Net", "Pay"]),
]

TEMPLATE_SHEET_NAME = "Import Template"
TEMPLATE_COL = {
    "Doc ID": 1,
    "First Name": 2,
    "Middle Name": 3,
    "Last Name": 4,
    "Entity Type_Employee": 7,
    "Social Security Number (SSN)": 23,
}


def builtin_header_rows(ncols=None):
    """Two header rows built from TEMPLATE_COL, used when no template workbook is supplied."""
    width = max(max(TEMPLATE_COL.values()), ncols or 0)
    row = [""] * width
    for name, col in TEMPLATE_COL.items():
        row[col - 1] = name
    return [list(row), list(row)]


def group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def find_token_sequence_start(line, tokens):
    words = [t for _, _, t in line]
    n = len(tokens)
    for i in range(len(words) - n + 1):
        if words[i:i + n] == tokens:
            return line[i][0]
    for i, w in enumerate(words):
        if w == tokens[0]:
            return line[i][0]
    return None


def find_header_bounds(lines):
    for line in lines:
        texts = [t for _, _, t in line]
        if "EE" in texts and "SSN" in texts:
            positions = {}
            for label, tokens in HEADER_DEFS:
                pos = find_token_sequence_start(line, tokens)
                if pos is not None:
                    positions[label] = pos
            if "EE ID" in positions and "SSN" in positions:
                return positions
    return None


def extract_rows(lines, bounds, page_num):
    ordered = sorted(bounds.items(), key=lambda kv: kv[1])
    col_names = [c for c, _ in ordered]
    col_starts = [x for _, x in ordered]

    records = []
    header_seen = False
    for line in lines:
        texts = [t for _, _, t in line]
        if "EE" in texts and "SSN" in texts:
            header_seen = True
            continue
        if not header_seen:
            continue

        buckets = {name: [] for name in col_names}
        for x0, x1, text in line:
            col = col_names[0]
            for name, start in zip(col_names, col_starts):
                if x0 + 1 >= start:
                    col = name
            buckets[col].append(text)

        row = {name: " ".join(buckets[name]).strip() for name in col_names}
        if EEID_PATTERN.match(row.get("EE ID", "")) or SSN_PATTERN.search(row.get("SSN", "")):
            row["Page"] = page_num
            records.append(row)
    return records


def process_pdf(path: Path):
    doc = fitz.open(path)
    all_records = []
    for i, page in enumerate(doc, start=1):
        rotation = page.rotation
        if rotation != 0:
            print(f"[{path.name}] page {i}: detected rotation {rotation} deg, normalizing for extraction")
        words = page.get_text("words")
        lines = group_words_into_lines(words)
        bounds = find_header_bounds(lines)
        if not bounds:
            print(f"[{path.name}] page {i}: header row not found, skipping")
            continue
        all_records.extend(extract_rows(lines, bounds, i))
    doc.close()
    return all_records


def split_last_first_middle(name: str):
    name = (name or "").strip()
    if not name:
        return "", "", ""
    if "," in name:
        last, rest = name.split(",", 1)
        parts = rest.strip().split()
        first = parts[0] if parts else ""
        middle = " ".join(parts[1:]) if len(parts) > 1 else ""
        return last.strip(), first, middle
    print(f"  [warn] name '{name}' has no comma ('Last, First' expected) -- placed entirely in Last Name")
    return name, "", ""


def build_extracted_workbook(records, report_cols):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted"
    headers = ["Page"] + report_cols
    ws.append(headers)
    for rec in records:
        ws.append([rec.get("Page", "")] + [rec.get(c, "") for c in report_cols])
    for i, header in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(header) + 2)
    return wb


def build_template_workbook(records, template_header_rows, template_col_widths):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = TEMPLATE_SHEET_NAME
    for row_vals in template_header_rows:
        ws.append(row_vals)
    for i, width in enumerate(template_col_widths, start=1):
        if width:
            ws.column_dimensions[get_column_letter(i)].width = width

    row_ptr = len(template_header_rows) + 1
    for rec in records:
        last, first, middle = split_last_first_middle(rec.get("Employee Name", ""))
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Doc ID"], value=rec.get("Doc ID", ""))
        ws.cell(row=row_ptr, column=TEMPLATE_COL["First Name"], value=first)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Middle Name"], value=middle)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Last Name"], value=last)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Entity Type_Employee"], value=True)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Social Security Number (SSN)"], value=rec.get("SSN", ""))
        row_ptr += 1
    return wb


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("template", nargs="?", default=None,
                        help="Optional path to BDE Template_Community Health.xlsx; "
                             "if omitted, built-in headers are used")
    parser.add_argument("-o", "--output-dir", default=".", help="Directory to write the workbooks into")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.template:
        tmpl_wb = openpyxl.load_workbook(args.template)
        tmpl_ws = tmpl_wb[TEMPLATE_SHEET_NAME]
        template_header_rows = [[c.value for c in tmpl_ws[1]], [c.value for c in tmpl_ws[2]]]
        template_col_widths = [
            tmpl_ws.column_dimensions[get_column_letter(i)].width
            for i in range(1, tmpl_ws.max_column + 1)
        ]
    else:
        template_header_rows = builtin_header_rows()
        template_col_widths = []

    report_cols = [label for label, _ in HEADER_DEFS]

    for pdf in pdf_files:
        print(f"Processing {pdf.name} ...")
        records = process_pdf(pdf)
        for r in records:
            r["Doc ID"] = pdf.name
        if not records:
            print(f"  no matching rows extracted from {pdf.name}")
            continue

        extracted_path = output_dir / f"{pdf.stem}_extracted.xlsx"
        template_path = output_dir / f"{pdf.stem}_template.xlsx"

        build_extracted_workbook(records, report_cols).save(extracted_path)
        build_template_workbook(records, template_header_rows, template_col_widths).save(template_path)

        print(f"  -> {extracted_path}")
        print(f"  -> {template_path}")


if __name__ == "__main__":
    main()
