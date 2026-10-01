# =============================================================================
# extract_1099.py — Structured field extraction for IRS Form 1099 (NEC/MISC)
# PDFs: payer/recipient TIN, name, and the amount/withholding boxes.
#
# HOW THIS WORKS: see form_extract_common.py's header — each field is found
# by searching the page's text for its printed box label, then reading the
# next value of the expected shape (money/TIN) after it. One record per
# page (each 1099 copy is a single page).
#
# COVERS 1099-NEC (Box 1 Nonemployee compensation) by default, the most
# common 1099 for contractor payments. 1099-MISC uses different box numbers
# for rents/royalties/other income — if you're processing 1099-MISC, add
# those box labels to FIELDS below (same pattern as the existing entries).
#
# LABELS BELOW ARE PLACEHOLDERS: box wording is standardized by the IRS, but
# spacing/line-wrapping varies by issuer. Run this against a real sample
# first — with --debug, which prints each page's raw extracted text — and
# adjust FIELDS below for any value that comes back empty.
#
# USAGE:
#   python extract_1099.py form1099.pdf
#   python extract_1099.py .\1099s\ --output 1099_data.csv
#   python extract_1099.py .\1099s\ --ocr always        # scanned copies
#   python extract_1099.py combined_text.csv --input-format csv
#   python extract_1099.py form1099.pdf --debug         # inspect raw page text
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

TIN_RE = re.compile(rf"(?:{EIN_RE.pattern}|{SSN_RE.pattern})")

FIELDS = {
    "Tax Year":                 {"labels": [r"Form\s+1099-?(?:NEC|MISC)"], "value_re": re.compile(r"20\d{2}")},
    "Payer TIN":                {"labels": [r"PAYER'?S TIN"], "value_re": TIN_RE},
    "Payer Name/Address":       {"labels": [r"PAYER'?S name, street address"], "value_re": FREE_TEXT_RE},
    "Recipient TIN":            {"labels": [r"RECIPIENT'?S TIN"], "value_re": TIN_RE},
    "Recipient Name/Address":   {"labels": [r"RECIPIENT'?S name"], "value_re": FREE_TEXT_RE},
    "Account Number":           {"labels": [r"[Aa]ccount number"], "value_re": re.compile(r"[A-Za-z0-9\-]+")},
    "Box 1 Nonemployee Comp":   {"labels": [r"\b1\b\s*Nonemployee compensation"], "value_re": MONEY_RE},
    "Box 4 Federal Tax WH":     {"labels": [r"\b4\b\s*Federal income tax withheld"], "value_re": MONEY_RE},
    "Box 5 State Tax WH":       {"labels": [r"\b5\b\s*State tax withheld"], "value_re": MONEY_RE},
    "Box 6 State/Payer State No": {"labels": [r"\b6\b\s*State/Payer'?s state no"], "value_re": re.compile(r"[A-Za-z0-9\-]+")},
    "Box 7 State Income":       {"labels": [r"\b7\b\s*State income"], "value_re": MONEY_RE},
}


def extract_record(text: str) -> dict:
    return {name: extract_field(text, [re.compile(p) for p in spec["labels"]], spec["value_re"])
            for name, spec in FIELDS.items()}


def main():
    parser = argparse.ArgumentParser(
        description="Extract payer/recipient TIN and amount/withholding box fields from IRS Form 1099 PDFs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    add_common_cli_args(parser)
    parser.add_argument("--debug", action="store_true", help="Print each page's raw extracted text before parsing.")
    args = parser.parse_args()
    input_path = " ".join(args.path)

    output_csv = args.output or default_output_csv(input_path, "_1099_data.csv")

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
        print(f"  {file_name} p{page_num}: Recipient TIN={record['Recipient TIN'] or '(not found)'}  "
              f"Box1={record['Box 1 Nonemployee Comp'] or '(not found)'}")

    if not all_records:
        print("\nNo pages processed. CSV not written.")
        sys.exit(0)

    cols = ["File Name", "Page Number"] + list(FIELDS.keys())
    df = pd.DataFrame(all_records)[cols]
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")
    print(f"\nDone. {len(all_records)} record(s) saved to: {output_csv}")
    print("Reminder: this output is tax/TIN PII — keep it in an authorized/approved location.")


if __name__ == "__main__":
    main()
