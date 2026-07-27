"""
Extract EE ID, Employee Name, and SSN from "Employee Gross To Net" PDF
report(s), writing per source PDF:

    <pdf_stem>_extracted.csv    - EE ID, Employee Name, SSN
    <pdf_stem>_template.xlsx    - rows appended to a copy of the supplied
                                  BDE Import Template workbook (Doc ID,
                                  First/Middle/Last Name, Entity
                                  Type_Employee = TRUE, SSN)

Both files are written next to the source PDF. The template's existing
headers, styles, and column layout are left untouched -- data rows start at
row 3 (below the two header rows), and no other file is copied into it.

The BDE template workbook is auto-detected: any *.xlsx file in the input
folder with "template" in its name (case-insensitive, excluding Excel lock
files and this script's own "*_template.xlsx" output) is used -- no need to
pass its path separately. Place it in the same folder as the PDFs.

Standalone script (no dependency on employee_gross_to_pay.py) so it can be
copied next to the PDFs and run from any folder. Uses the same rotation-aware
PyMuPDF header/row detection approach.

Usage:
    python emp_gross_pay.py <pdf_file_or_folder>

Requires:
    pip install pymupdf openpyxl tqdm
"""

import argparse
import builtins
import contextlib
import csv
import re
from pathlib import Path

import pymupdf as fitz
import openpyxl
from tqdm import tqdm

SSN_PATTERN = re.compile(r"\d{3}-\d{2}-\d{4}|[*Xx]{3}-[*Xx]{2}-\d{4}")
EEID_PATTERN = re.compile(r"^[A-Za-z0-9-]{2,15}$")

# (column label, token sequence that identifies its header cell)
HEADER_DEFS = [
    ("EE ID", ["EE", "ID"]),
    ("Employee Name", ["Employee", "Name"]),
    ("SSN", ["SSN"]),
]

OUTPUT_COLUMNS = ["EE ID", "Employee Name", "SSN"]

TEMPLATE_SHEET_NAME = "Import Template"
TEMPLATE_COL = {
    "Doc ID": 1,
    "First Name": 2,
    "Middle Name": 3,
    "Last Name": 4,
    "Entity Type_Employee": 7,
    "Social Security Number (SSN)": 23,
}


@contextlib.contextmanager
def print_via_tqdm():
    """Route plain print() calls through tqdm.write so they don't clobber
    the active progress bar line."""
    original_print = builtins.print
    builtins.print = tqdm.write
    try:
        yield
    finally:
        builtins.print = original_print


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


def write_csv(records, out_path):
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(OUTPUT_COLUMNS)
        for rec in records:
            row = [rec.get(col, "") for col in OUTPUT_COLUMNS]
            row[OUTPUT_COLUMNS.index("SSN")] = format_ssn(rec.get("SSN", ""))
            writer.writerow(row)


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


def format_ssn(raw: str) -> str:
    """Return only the SSN itself, formatted as ###-##-####, discarding any
    other characters/text captured alongside it in the same column."""
    raw = (raw or "").strip()
    match = SSN_PATTERN.search(raw)
    if match:
        return match.group(0)
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 9:
        return f"{digits[0:3]}-{digits[3:5]}-{digits[5:9]}"
    return digits


def build_template_workbook(template_path, records):
    wb = openpyxl.load_workbook(template_path)
    ws = wb[TEMPLATE_SHEET_NAME]

    row_ptr = 3
    for rec in records:
        last, first, middle = split_last_first_middle(rec.get("Employee Name", ""))
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Doc ID"], value=rec.get("Doc ID", ""))
        ws.cell(row=row_ptr, column=TEMPLATE_COL["First Name"], value=first)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Middle Name"], value=middle)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Last Name"], value=last)
        ws.cell(row=row_ptr, column=TEMPLATE_COL["Entity Type_Employee"], value=True)
        ssn_cell = ws.cell(row=row_ptr, column=TEMPLATE_COL["Social Security Number (SSN)"],
                            value=format_ssn(rec.get("SSN", "")))
        ssn_cell.number_format = "@"  # keep as text so Excel doesn't strip dashes/leading zeros
        row_ptr += 1
    return wb


def find_template(folder: Path) -> Path:
    candidates = [
        p for p in folder.glob("*.xlsx")
        if not p.name.startswith("~$")
        and "template" in p.name.lower()
        and not p.name.lower().endswith("_template.xlsx")
    ]
    if not candidates:
        raise FileNotFoundError(
            f"No BDE template file found in {folder} "
            f"(expected an .xlsx with 'template' in its name, e.g. 'BDE Template_Community Health.xlsx')"
        )
    if len(candidates) > 1:
        raise FileNotFoundError(
            f"Multiple possible template files found in {folder}: {[c.name for c in candidates]}. "
            f"Keep only one."
        )
    return candidates[0]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    search_dir = input_path.parent if input_path.is_file() else input_path
    try:
        template_path = find_template(search_dir)
    except FileNotFoundError as e:
        print(e)
        return
    print(f"Using template: {template_path.name}")

    check_wb = openpyxl.load_workbook(template_path, read_only=True)
    if TEMPLATE_SHEET_NAME not in check_wb.sheetnames:
        print(f"Sheet '{TEMPLATE_SHEET_NAME}' not found in template: {template_path}")
        return
    check_wb.close()

    with print_via_tqdm():
        for pdf in tqdm(pdf_files, desc="Extracting", unit="file"):
            records = process_pdf(pdf)
            if not records:
                tqdm.write(f"  no matching rows extracted from {pdf.name}")
                continue
            for r in records:
                r["Doc ID"] = pdf.name

            csv_path = pdf.parent / f"{pdf.stem}_extracted.csv"
            write_csv(records, csv_path)

            template_out_path = pdf.parent / f"{pdf.stem}_template.xlsx"
            build_template_workbook(template_path, records).save(template_out_path)

            tqdm.write(f"  {pdf.name} -> {csv_path.name}, {template_out_path.name}")


if __name__ == "__main__":
    main()
