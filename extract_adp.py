# =============================================================================
# extract_adp.py — Structured field extraction for ADP-generated payroll
# PDFs (earnings statements / paystubs): employee ID, masked SSN, pay
# period, gross/net pay, and the common tax withholding lines.
#
# HOW THIS WORKS: see the EMBEDDED COMMON ENGINE section below — each field is found
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
# DEPENDENCIES: pip install pdfplumber pandas tqdm (+ pdf2image pytesseract for OCR)
# =============================================================================

import argparse
import contextlib
import os
import re
import sys

import pandas as pd
import pdfplumber
from tqdm import tqdm


# ---------------------------------------------------------------------------
# EMBEDDED COMMON ENGINE (self-contained -- no other repo file is needed)
# Label/value scanning + PDF text extraction (native text via pdfplumber,
# PDFium fallback, OCR fallback for scanned pages) shared by the form
# extractors. Each field is found by searching the page text for its printed
# label, then reading the next value of the expected shape within a short
# window after it.
# ---------------------------------------------------------------------------

# Set these to match your machine so OCR runs with no path flags at all.
# Leave as None (or pass --poppler-path / --tesseract-cmd) if Poppler's bin
# folder and tesseract.exe are already on PATH.
DEFAULT_POPPLER_PATH = r"C:\Poppler\Release-26.02.0-0\poppler-26.02.0\Library\bin"
DEFAULT_TESSERACT_CMD = r"C:\Tessaract\tesseract.exe"

# Reusable value-shape patterns
MONEY_RE = re.compile(r"\(?-?\$?\s?\d[\d,]*(?:\.\d{2})?\)?")
SSN_RE = re.compile(r"(?:\d{3}-\d{2}-\d{4}|[Xx\*]{3}-[Xx\*]{2}-\d{4}|\d{9})")
EIN_RE = re.compile(r"\d{2}-\d{7}")
FREE_TEXT_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9 ,.'&\-]*")


def resolve_pdf_files(path: str) -> list:
    """Return list of absolute PDF paths from a file or folder."""
    if os.path.isfile(path):
        if not path.lower().endswith(".pdf"):
            print(f"Error: '{path}' is not a PDF file.")
            sys.exit(1)
        return [os.path.abspath(path)]
    elif os.path.isdir(path):
        files = [
            os.path.abspath(os.path.join(path, f))
            for f in os.listdir(path)
            if f.lower().endswith(".pdf")
        ]
        if not files:
            print(f"Error: No PDF files found in '{path}'.")
            sys.exit(1)
        return sorted(files)
    else:
        print(f"Error: Path does not exist: '{path}'")
        sys.exit(1)


def extract_field(text: str, label_alternatives: list, value_re: re.Pattern, max_gap: int = 120) -> str:
    """Search `text` for the first of several label spellings, then return
    the first `value_re` match found within `max_gap` characters after it.
    Returns "" if the label (or a matching value after it) isn't found."""
    for label_re in label_alternatives:
        label_match = label_re.search(text)
        if not label_match:
            continue
        window = text[label_match.end(): label_match.end() + max_gap]
        value_match = value_re.search(window)
        if value_match:
            return value_match.group(0).strip()
    return ""


def is_garbled(text: str, threshold: float = 0.3, single_letter_ratio: float = 0.4) -> bool:
    """
    Heuristic: detect text pdfplumber decoded incorrectly, via two failure modes.

    1. Font with no ToUnicode map -- decodes to Private Use Area codepoints
       (or other unmapped control chars), rendering as tofu boxes.
    2. Interleaved multi-column layout -- nearly every whitespace-split token
       ends up a single stray letter, a pattern real extracted text never
       produces. OCR reads the rendered page image, so it isn't fooled.
    """
    if not text.strip():
        return False
    bad = sum(
        1 for ch in text
        if 0xE000 <= ord(ch) <= 0xF8FF  # Private Use Area
        or (ord(ch) < 0x20 and ch not in "\t\n\r")
    )
    if (bad / len(text)) > threshold:
        return True

    for line in text.split("\n"):
        tokens = [t for t in line.split() if t.isalpha()]
        if len(tokens) < 6:
            continue
        singles = sum(1 for t in tokens if len(t) == 1 and t.lower() not in ("a", "i"))
        if singles / len(tokens) > single_letter_ratio:
            return True
    return False


def ocr_page_image(
    pdf_path: str, page_num: int, lang: str, dpi: int, poppler_path: str = None, tesseract_cmd: str = None
) -> str:
    """Render a single PDF page to an image and return OCR text."""
    from pdf2image import convert_from_path
    import pytesseract

    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    images = convert_from_path(
        pdf_path, dpi=dpi, first_page=page_num, last_page=page_num, poppler_path=poppler_path
    )
    if not images:
        return ""
    return pytesseract.image_to_string(images[0], lang=lang)


def open_pdfium(pdf_path: str):
    """Open a PDF with PDFium (pypdfium2) as a fallback text engine; None if unavailable."""
    try:
        import pypdfium2
        return pypdfium2.PdfDocument(pdf_path)
    except Exception:
        return None


def pdfium_page_text(doc, page_index: int) -> str:
    """Return one page's text layer via PDFium, or '' on any failure."""
    try:
        page = doc[page_index]
        try:
            textpage = page.get_textpage()
            try:
                text = textpage.get_text_range() or ""
            finally:
                textpage.close()
        finally:
            page.close()
    except Exception:
        return ""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def extract_pdf_text_lines(
    pdf_path: str,
    file_index: int,
    total_files: int,
    ocr_mode: str,
    lang: str,
    dpi: int,
    poppler_path: str = None,
    tesseract_cmd: str = None,
) -> list:
    """
    Extract non-empty text lines from every page of one PDF.

    ocr_mode:
      'never'  -- pdfplumber, then PDFium for pages it can't read; no OCR
      'auto'   -- pdfplumber, then PDFium; OCR fallback for pages with no text
      'always' -- OCR every page

    Returns a list of {"File Name", "Page Number", "Extracted Text", "Method"}.
    """
    records = []
    pdf_file = os.path.basename(pdf_path)

    try:
        with contextlib.ExitStack() as stack:
            try:
                pdf = stack.enter_context(pdfplumber.open(pdf_path))
                doc_pages = len(pdf.pages)
            except Exception as e:
                pdf = None
                doc_pages = 0
                plumber_error = e

            # PDFium is opened lazily -- only when pdfplumber fails.
            pdfium_state = {"doc": None, "tried": False}

            def get_pdfium():
                if not pdfium_state["tried"]:
                    pdfium_state["tried"] = True
                    pdfium_state["doc"] = open_pdfium(pdf_path)
                    if pdfium_state["doc"] is not None:
                        stack.callback(pdfium_state["doc"].close)
                return pdfium_state["doc"]

            if pdf is None:
                if get_pdfium() is None:
                    raise plumber_error
                doc_pages = len(pdfium_state["doc"])
                tqdm.write(f"  Note: pdfplumber couldn't open '{pdf_file}' ({plumber_error}) -- using PDFium.")

            page_iter = tqdm(
                range(doc_pages),
                desc=f"  {pdf_file[:32]}",
                unit="pg",
                total=doc_pages,
                bar_format=(
                    f"  [{file_index}/{total_files}] {{desc}} | "
                    "{n_fmt}/{total_fmt} pages [{elapsed}<{remaining}, {rate_fmt}]"
                ),
            )

            for page_index in page_iter:
                page_num = page_index + 1
                try:
                    text = ""
                    method = "native"

                    if ocr_mode != "always":
                        if pdf is not None:
                            try:
                                text = pdf.pages[page_index].extract_text() or ""
                            except Exception:
                                text = ""
                        if not text.strip() or is_garbled(text):
                            doc = get_pdfium()
                            alt = pdfium_page_text(doc, page_index) if doc is not None else ""
                            if alt.strip() and not is_garbled(alt):
                                text = alt
                                method = "pdfium"

                    needs_ocr = not text.strip() or (ocr_mode == "auto" and is_garbled(text))
                    if needs_ocr and ocr_mode != "never":
                        ocr_text = ocr_page_image(pdf_path, page_num, lang, dpi, poppler_path, tesseract_cmd)
                        if ocr_text.strip():
                            text = ocr_text
                            method = "ocr"

                    for line in text.split("\n"):
                        stripped = line.strip()
                        if stripped:
                            records.append(
                                {
                                    "File Name": pdf_file,
                                    "Page Number": page_num,
                                    "Extracted Text": stripped,
                                    "Method": method,
                                }
                            )
                except Exception as e:
                    tqdm.write(f"    Warning: page {page_num} skipped -- {e}")
    except Exception as e:
        tqdm.write(f"  Failed to open '{pdf_file}': {e}")

    return records


def iter_pages_from_pdfs(pdf_files: list, ocr_mode: str = "auto", lang: str = "eng", dpi: int = 300,
                          poppler_path: str = None, tesseract_cmd: str = None):
    """Yield (file_name, page_num, page_text) for every non-empty page."""
    for file_index, pdf_path in enumerate(pdf_files, start=1):
        records = extract_pdf_text_lines(
            pdf_path, file_index, len(pdf_files), ocr_mode, lang, dpi, poppler_path, tesseract_cmd,
        )
        pages = {}
        for rec in records:
            key = (rec["File Name"], rec["Page Number"])
            pages.setdefault(key, []).append(rec["Extracted Text"])
        for (file_name, page_num), lines in sorted(pages.items(), key=lambda kv: kv[0][1]):
            yield file_name, page_num, "\n".join(lines)


def iter_pages_from_csv(csv_path: str):
    """Yield (file_name, page_num, page_text) from a CSV with columns
    File Name, Page Number, Extracted Text (one row per text line)."""
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    required = {"File Name", "Page Number", "Extracted Text"}
    if not required.issubset(df.columns):
        sys.exit(f"Error: CSV must have columns {sorted(required)}.")

    for (file_name, page_num), group in df.groupby(["File Name", "Page Number"], sort=False):
        yield file_name, int(page_num), "\n".join(group["Extracted Text"].tolist())


def iter_pages(input_path: str, input_format: str = "auto", ocr_mode: str = "auto", lang: str = "eng",
               dpi: int = 300, poppler_path: str = None, tesseract_cmd: str = None):
    """Dispatch to the PDF or CSV page iterator based on `input_format`
    ('auto' guesses from the file extension)."""
    fmt = input_format
    if fmt == "auto":
        fmt = "csv" if os.path.isfile(input_path) and input_path.lower().endswith(".csv") else "pdf"

    if fmt == "csv":
        if not os.path.isfile(input_path):
            sys.exit(f"Error: '{input_path}' is not a file.")
        yield from iter_pages_from_csv(input_path)
    else:
        pdf_files = resolve_pdf_files(input_path)
        yield from iter_pages_from_pdfs(pdf_files, ocr_mode, lang, dpi, poppler_path, tesseract_cmd)


def default_output_csv(input_path: str, suffix: str) -> str:
    if os.path.isfile(input_path):
        return os.path.splitext(os.path.basename(input_path))[0] + suffix
    return os.path.basename(os.path.normpath(input_path)) + suffix


def add_common_cli_args(parser):
    """Shared argparse options for the form extractor CLIs."""
    parser.add_argument("path", nargs="+", help="Path to a PDF file, a folder of PDFs, or a text CSV.")
    parser.add_argument("--output", "-o", default=None, help="Output CSV file name.")
    parser.add_argument(
        "--input-format", choices=["auto", "pdf", "csv"], default="auto",
        help="Force input type instead of guessing from the file extension (default: auto).",
    )
    parser.add_argument("--ocr", choices=["never", "auto", "always"], default="auto",
                         help="OCR mode for PDF input (default: auto).")
    parser.add_argument("--lang", default="eng", help="Tesseract language code (default: eng).")
    parser.add_argument("--dpi", type=int, default=300, help="OCR render DPI (default: 300).")
    parser.add_argument("--poppler-path", default=DEFAULT_POPPLER_PATH, help="Folder with pdftoppm.exe, if not on PATH.")
    parser.add_argument("--tesseract-cmd", default=DEFAULT_TESSERACT_CMD, help="Full path to tesseract.exe, if not on PATH.")
    return parser

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
