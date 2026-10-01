# =============================================================================
# extract_adp.py — Structured field extraction for ADP-generated payroll
# PDFs (earnings statements / paystubs): employee ID, masked SSN, pay
# period, gross/net pay, and the common tax withholding lines.
#
# HOW THIS WORKS: see form_extract_common.py's header — each field is found
# by searching the page's text for its printed label, then reading the
# next value of the expected shape (money/ID/date) after it. Assumes one
# employee's statement per page, which covers ADP's standard single-employee
# earnings statement/paystub PDF export.
#
# IF YOUR EXPORT IS A MULTI-EMPLOYEE PAYROLL REGISTER (several employees
# listed per page in a table) instead of one-employee-per-page statements,
# this label-per-page approach won't separate rows correctly — that needs a
# per-line record scanner like extract_student_records.py's instead of this
# per-page one. Flag it and that variant can be added.
#
# LABELS BELOW ARE PLACEHOLDERS: ADP's exact wording/layout varies by client
# configuration. Run this against a real sample first — with --debug, which
# prints each page's raw extracted text — and adjust FIELDS below for any
# value that comes back empty.
#
# USAGE:
#   python extract_adp.py statement.pdf
#   python extract_adp.py .\statements\ --output adp_data.csv
#   python extract_adp.py .\statements\ --ocr always     # scanned copies
#   python extract_adp.py combined_text.csv --input-format csv
#   python extract_adp.py statement.pdf --debug          # inspect raw page text
#
# DEPENDENCIES: pip install pdfplumber pandas (+ pdf2image pytesseract for OCR)
# =============================================================================

import argparse
import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from form_extract_common import (  # noqa: E402
    MONEY_RE, SSN_RE, FREE_TEXT_RE,
    extract_field, iter_pages, default_output_csv, add_common_cli_args,
)

DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")
ID_RE = re.compile(r"[A-Za-z0-9\-]+")

FIELDS = {
    "Employee Name":     {"labels": [r"Employee [Nn]ame", r"Name\s*:"], "value_re": FREE_TEXT_RE},
    "Employee ID":       {"labels": [r"Employee (?:ID|Number|File\s*#)", r"File Number"], "value_re": ID_RE},
    "SSN (masked)":      {"labels": [r"SSN", r"Social Security"], "value_re": SSN_RE},
    "Company/Client":    {"labels": [r"Company Code", r"Company Name", r"Client"], "value_re": FREE_TEXT_RE},
    "Pay Date":          {"labels": [r"Pay Date"], "value_re": DATE_RE},
    "Period Begin":      {"labels": [r"Period Begin(?:ning)?"], "value_re": DATE_RE},
    "Period End":        {"labels": [r"Period End(?:ing)?"], "value_re": DATE_RE},
    "Gross Pay":         {"labels": [r"Gross Pay"], "value_re": MONEY_RE},
    "Net Pay":           {"labels": [r"Net Pay"], "value_re": MONEY_RE},
    "Federal Tax WH":    {"labels": [r"Federal (?:Income Tax|Withholding|W/H)"], "value_re": MONEY_RE},
    "State Tax WH":      {"labels": [r"State (?:Income Tax|Withholding|W/H)"], "value_re": MONEY_RE},
    "Social Security WH":{"labels": [r"Social Security(?: Tax| EE)?\s*(?:Withheld)?"], "value_re": MONEY_RE},
    "Medicare WH":       {"labels": [r"Medicare(?: Tax| EE)?\s*(?:Withheld)?"], "value_re": MONEY_RE},
    "Regular Hours":     {"labels": [r"Regular Hours"], "value_re": re.compile(r"[\d,]+\.\d{2}")},
    "Overtime Hours":    {"labels": [r"Overtime Hours", r"O/?T Hours"], "value_re": re.compile(r"[\d,]+\.\d{2}")},
}


def extract_record(text: str) -> dict:
    return {name: extract_field(text, [re.compile(p) for p in spec["labels"]], spec["value_re"])
            for name, spec in FIELDS.items()}


def main():
    parser = argparse.ArgumentParser(
        description="Extract employee/pay/withholding fields from ADP payroll statement PDFs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    add_common_cli_args(parser)
    parser.add_argument("--debug", action="store_true", help="Print each page's raw extracted text before parsing.")
    args = parser.parse_args()
    input_path = " ".join(args.path)

    output_csv = args.output or default_output_csv(input_path, "_adp_data.csv")

    all_records = []
    for file_name, page_num, text in iter_pages(
        input_path, args.input_format, args.ocr, args.lang, args.dpi, args.poppler_path, args.tesseract_cmd
    ):
        if args.debug:
            print(f"\n----- {file_name} p{page_num} raw text -----\n{text}\n")
        record = extract_record(text)
        record["File Name"] = file_name
        record["Page Number"] = page_num
        all_records.append(record)
        print(f"  {file_name} p{page_num}: {record['Employee Name'] or '(not found)'}  "
              f"Net Pay={record['Net Pay'] or '(not found)'}")

    if not all_records:
        print("\nNo pages processed. CSV not written.")
        sys.exit(0)

    cols = ["File Name", "Page Number"] + list(FIELDS.keys())
    df = pd.DataFrame(all_records)[cols]
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    print(f"\nDone. {len(all_records)} record(s) saved to: {output_csv}")
    print("Reminder: this output is payroll PII — keep it in an authorized/approved location.")


if __name__ == "__main__":
    main()
