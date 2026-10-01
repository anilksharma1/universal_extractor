# =============================================================================
# form_extract_common.py — Shared label/value scanning engine for the
# form-specific extractors (extract_w2.py, extract_adp.py, extract_1099.py).
#
# WHY THIS EXISTS:
#   W2, 1099, and ADP paystub PDFs are printed as label-then-value pairs
#   ("Box 1  Wages, tips, other compensation   54,321.00"), not as fixed
#   comb-boxes like the WA ESD form extract_5208a.py handles, and (unlike
#   extract_student_records.py's rosters) each is normally one record per
#   page. So instead of per-format pixel coordinates, each field here is
#   found by searching the page's native text for its printed label, then
#   reading the next value of the expected shape (money / SSN / EIN / plain
#   text) within a short window after that label. This is more resilient to
#   minor layout differences across employers/payroll providers than a
#   fixed box position, but label wording placeholders below still need
#   confirming against a real sample of each form before results are
#   trusted — see the header of each extractor script.
#
# INPUT: either raw PDFs (native text via pdfplumber, OCR fallback for
# scanned pages, reusing text_from_pdf.py's extract_pdf_range so OCR logic
# isn't duplicated) or a CSV already produced by text_from_pdf.py.
#
# PII NOTE: W2/1099/ADP output is tax and payroll PII (SSN/EIN, wages,
# withholding). Keep it in an authorized/approved location, not a personal
# or shared-outside-team folder. These scripts send no file contents
# anywhere — they only read local files and write a local CSV.
# =============================================================================

import os
import re
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from text_from_pdf import extract_pdf_range, DEFAULT_POPPLER_PATH, DEFAULT_TESSERACT_CMD  # noqa: E402

# ---------------------------------------------------------------------------
# Reusable value-shape patterns
# ---------------------------------------------------------------------------

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


def iter_pages_from_pdfs(pdf_files: list, ocr_mode: str = "auto", lang: str = "eng", dpi: int = 300,
                          poppler_path: str = None, tesseract_cmd: str = None, progress_queue=None):
    """Yield (file_name, page_num, page_text) for every non-empty page,
    reusing text_from_pdf.py's native-text + OCR-fallback extraction so
    that logic lives in exactly one place."""
    for file_index, pdf_path in enumerate(pdf_files, start=1):
        records, _stats = extract_pdf_range(
            pdf_path, file_index, len(pdf_files), ocr_mode, lang, dpi,
            poppler_path, tesseract_cmd, show_progress=(progress_queue is None),
            progress_queue=progress_queue,
        )
        pages = {}
        for rec in records:
            key = (rec["File Name"], rec["Page Number"])
            pages.setdefault(key, []).append(rec["Extracted Text"])
        for (file_name, page_num), lines in sorted(pages.items(), key=lambda kv: kv[0][1]):
            yield file_name, page_num, "\n".join(lines)


def iter_pages_from_csv(csv_path: str):
    """Yield (file_name, page_num, page_text) from a text_from_pdf.py CSV."""
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    required = {"File Name", "Page Number", "Extracted Text"}
    if not required.issubset(df.columns):
        sys.exit(f"Error: CSV must have columns {sorted(required)} (from text_from_pdf.py).")

    for (file_name, page_num), group in df.groupby(["File Name", "Page Number"], sort=False):
        yield file_name, int(page_num), "\n".join(group["Extracted Text"].tolist())


def iter_pages(input_path: str, input_format: str = "auto", ocr_mode: str = "auto", lang: str = "eng",
               dpi: int = 300, poppler_path: str = None, tesseract_cmd: str = None, progress_queue=None):
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
        yield from iter_pages_from_pdfs(pdf_files, ocr_mode, lang, dpi, poppler_path, tesseract_cmd, progress_queue)


def default_output_csv(input_path: str, suffix: str) -> str:
    if os.path.isfile(input_path):
        return os.path.splitext(os.path.basename(input_path))[0] + suffix
    return os.path.basename(os.path.normpath(input_path)) + suffix


def add_common_cli_args(parser):
    """Shared argparse options for the form extractor CLIs."""
    parser.add_argument("path", nargs="+", help="Path to a PDF file, a folder of PDFs, or a text_from_pdf.py CSV.")
    parser.add_argument("--output", "-o", default=None, help="Output CSV file name.")
    parser.add_argument(
        "--input-format", choices=["auto", "pdf", "csv"], default="auto",
        help="Force input type instead of guessing from the file extension (default: auto).",
    )
    parser.add_argument("--ocr", choices=["never", "auto", "always"], default="auto",
                         help="OCR mode for PDF input, same meaning as in text_from_pdf.py (default: auto).")
    parser.add_argument("--lang", default="eng", help="Tesseract language code (default: eng).")
    parser.add_argument("--dpi", type=int, default=300, help="OCR render DPI (default: 300).")
    parser.add_argument("--poppler-path", default=DEFAULT_POPPLER_PATH, help="Folder with pdftoppm.exe, if not on PATH.")
    parser.add_argument("--tesseract-cmd", default=DEFAULT_TESSERACT_CMD, help="Full path to tesseract.exe, if not on PATH.")
    return parser
