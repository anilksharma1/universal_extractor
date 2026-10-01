# =============================================================================
# extract_w2.py — Structured field extraction for IRS Form W-2 (Wage and
# Tax Statement) PDFs: SSN, EIN, employer/employee name, and the numbered
# wage/tax boxes.
#
# HOW THIS WORKS: see form_extract_common.py's header — each field is found
# by searching the page's text for its printed box label, then reading the
# next value of the expected shape (money/SSN/EIN) after it. One record per
# page (each W-2 copy is a single page).
#
# LABELS BELOW ARE PLACEHOLDERS: box wording is standardized by the IRS, but
# spacing/line-wrapping varies by payroll provider (ADP, Paychex, in-house).
# Run this against a real sample first — with --debug, which prints each
# page's raw extracted text so you can see exactly how labels/values line
# up before trusting the output — and adjust FIELDS below if a value comes
# back empty for your provider's layout.
#
# USAGE:
#   python extract_w2.py w2.pdf
#   python extract_w2.py .\w2s\ --output w2_data.csv
#   python extract_w2.py .\w2s\ --ocr always          # scanned/faxed copies
#   python extract_w2.py combined_text.csv --input-format csv
#   python extract_w2.py w2.pdf --debug               # inspect raw page text
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
    MONEY_RE, SSN_RE, EIN_RE, FREE_TEXT_RE,
    extract_field, iter_pages, default_output_csv, add_common_cli_args,
)

FIELDS = {
    "Tax Year":                {"labels": [r"Form\s+W-?2", r"Wage and Tax Statement"], "value_re": re.compile(r"20\d{2}")},
    "Employee SSN":            {"labels": [r"[Ee]mployee'?s social security number", r"\ba\b\s*Employee"], "value_re": SSN_RE},
    "Employer EIN":            {"labels": [r"[Ee]mployer identification number", r"\bb\b\s*Employer"], "value_re": EIN_RE},
    "Employer Name/Address":   {"labels": [r"[Ee]mployer'?s name, address"], "value_re": FREE_TEXT_RE},
    "Control Number":          {"labels": [r"[Cc]ontrol number"], "value_re": re.compile(r"[A-Za-z0-9\-]+")},
    "Employee Name/Address":   {"labels": [r"[Ee]mployee'?s (?:first name|name and address)"], "value_re": FREE_TEXT_RE},
    "Box 1 Wages":             {"labels": [r"\b1\b\s*Wages,?\s*tips"], "value_re": MONEY_RE},
    "Box 2 Federal Tax WH":    {"labels": [r"\b2\b\s*Federal income tax withheld"], "value_re": MONEY_RE},
    "Box 3 SS Wages":          {"labels": [r"\b3\b\s*Social security wages"], "value_re": MONEY_RE},
    "Box 4 SS Tax WH":         {"labels": [r"\b4\b\s*Social security tax withheld"], "value_re": MONEY_RE},
    "Box 5 Medicare Wages":    {"labels": [r"\b5\b\s*Medicare wages"], "value_re": MONEY_RE},
    "Box 6 Medicare Tax WH":   {"labels": [r"\b6\b\s*Medicare tax withheld"], "value_re": MONEY_RE},
    "Box 15 State":            {"labels": [r"\b15\b\s*State\b"], "value_re": re.compile(r"\b[A-Z]{2}\b")},
    "Box 16 State Wages":      {"labels": [r"\b16\b\s*State wages"], "value_re": MONEY_RE},
    "Box 17 State Tax WH":     {"labels": [r"\b17\b\s*State income tax"], "value_re": MONEY_RE},
}


def extract_record(text: str) -> dict:
    return {name: extract_field(text, [re.compile(p) for p in spec["labels"]], spec["value_re"])
            for name, spec in FIELDS.items()}


def main():
    parser = argparse.ArgumentParser(
        description="Extract SSN/EIN/wage/tax box fields from IRS Form W-2 PDFs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    add_common_cli_args(parser)
    parser.add_argument("--debug", action="store_true", help="Print each page's raw extracted text before parsing.")
    args = parser.parse_args()
    input_path = " ".join(args.path)

    output_csv = args.output or default_output_csv(input_path, "_w2_data.csv")

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
        print(f"  {file_name} p{page_num}: SSN={record['Employee SSN'] or '(not found)'}  "
              f"Box1={record['Box 1 Wages'] or '(not found)'}")

    if not all_records:
        print("\nNo pages processed. CSV not written.")
        sys.exit(0)

    cols = ["File Name", "Page Number"] + list(FIELDS.keys())
    df = pd.DataFrame(all_records)[cols]
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    print(f"\nDone. {len(all_records)} record(s) saved to: {output_csv}")
    print("Reminder: this output is tax/SSN PII — keep it in an authorized/approved location.")


if __name__ == "__main__":
    main()
