# w2_extractor.py
# Extracts Employee Name, SSN, and Address from all pages of
# W-2 PDF files (ADP layout and standard IRS layouts).
# Uses coordinate-based word extraction for accurate field detection.
# Returns one row per employee per page. If extraction is wrong, run with --debug first.
#
# Outputs:
#   <output>.xlsx          — two sheets: "Extracted Data" + "Standardized Data"
#   <output>_processing.xlsx — per-document summary (entity count, SSN count, etc.)
#
# Requires: pdfplumber, pandas, openpyxl, tqdm
#   pip install pdfplumber pandas openpyxl tqdm
#
# ---------------------------------------------------------------------------
# USAGE
# ---------------------------------------------------------------------------
#
# Debug mode — print raw word positions pdfplumber reads from the first page.
#
#   python w2_extractor.py "C:\YourFolder\sample_w2.pdf" --debug
#
# Single file extraction:
#
#   python w2_extractor.py "C:\YourFolder\sample_w2.pdf"
#   python w2_extractor.py "C:\YourFolder\sample_w2.pdf" "C:\Output\results.xlsx"
#
# Batch folder extraction (all PDFs in a directory):
#
#   python w2_extractor.py "C:\YourFolder\PDFs\" "C:\Output\results.xlsx"
#
# Limit pages per file:
#
#   python w2_extractor.py "C:\YourFolder\sample_w2.pdf" --pages 10
#
# ---------------------------------------------------------------------------
# DO NOT commit real W-2 documents or any file containing PII as test data.
# Use synthetic/anonymized samples only.
# ---------------------------------------------------------------------------

import re
import sys
import os
from pathlib import Path

import pdfplumber
import pandas as pd
from tqdm import tqdm


# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

# SSN: XXX-XX-XXXX (not EIN format XX-XXXXXXX)
_SSN_PATTERN = re.compile(r"^\d{3}-\d{2}-\d{4}$")
_SSN_IN_LINE = re.compile(r"\b(\d{3}-\d{2}-\d{4})\b")

# City, ST ZIP pattern
_CITY_STATE_ZIP_RE = re.compile(r"^.+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\s*$")

# US ZIP / ZIP+4
_ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")


# ---------------------------------------------------------------------------
# Layout constants (tuned for ADP W-2 layout)
# The 'e/f' block label sits at top < 280 and x0 < 60
# ---------------------------------------------------------------------------

_EF_LABEL_MAX_Y = 280
_EF_LABEL_MAX_X = 60
_SSN_OFFSET_MIN = 30   # points below e/f label where SSN appears
_SSN_OFFSET_MAX = 80
_BOTTOM_ROW_Y_MIN = 450  # SSNs below this are employer copies — ignore


# ---------------------------------------------------------------------------
# Coordinate-based extraction (primary path — matches extract_w2_gui.py logic)
# ---------------------------------------------------------------------------

def _find_ef_blocks(words: list) -> list:
    """Return position dicts for each 'e/f' label found in the top region of the page."""
    blocks = []
    for w in words:
        if w["top"] > _EF_LABEL_MAX_Y:
            continue
        if w["x0"] > _EF_LABEL_MAX_X:
            continue
        text_lower = w["text"].lower()
        if text_lower == "e/f" or text_lower.startswith("e/f"):
            blocks.append({"top": w["top"], "x0": w["x0"]})
    return blocks


def _extract_employee_from_block(words: list, label_top: float, label_x0: float) -> dict | None:
    """
    Given the position of an 'e/f' label, extract name/street/city-state-zip
    from the 3 lines below it, and the SSN from the area further below.
    """
    # Words in the ~35pt window below the label, within the left column
    block_words = [
        w for w in words
        if label_top + 2 < w["top"] < label_top + 35
        and w["x0"] < 200
        and w["x0"] > label_x0 - 5
    ]
    block_words.sort(key=lambda w: (w["top"], w["x0"]))

    # Group into visual lines by proximity of 'top' coordinate
    lines = []
    LINE_TOL = 4
    for w in block_words:
        if not lines:
            lines.append([w])
            continue
        if abs(w["top"] - lines[-1][0]["top"]) <= LINE_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])

    text_lines = []
    for line in lines:
        line.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in line).strip()
        if text and len(text) > 1:
            text_lines.append(text)

    if len(text_lines) < 2:
        return None

    # Find SSN in the y-offset band below the block label
    ssn_candidates = []
    for w in words:
        if label_top + _SSN_OFFSET_MIN < w["top"] < label_top + _SSN_OFFSET_MAX:
            if w["x0"] < 200:
                m = _SSN_IN_LINE.search(w["text"])
                if m and _SSN_PATTERN.match(m.group(1)):
                    ssn_candidates.append(m.group(1))

    # Fallback: search full text of that region (handles fused tokens)
    if not ssn_candidates:
        region_words = [
            w for w in words
            if label_top + _SSN_OFFSET_MIN < w["top"] < label_top + _SSN_OFFSET_MAX
            and w["x0"] < 250
        ]
        line_text = " ".join(w["text"] for w in sorted(region_words, key=lambda w: w["x0"]))
        for m in _SSN_IN_LINE.finditer(line_text):
            if _SSN_PATTERN.match(m.group(1)):
                ssn_candidates.append(m.group(1))

    ssn = ssn_candidates[0] if ssn_candidates else ""

    # Identify city/state/zip line, work backward for name and street
    name = street = csz = ""
    csz_idx = -1
    for i, ln in enumerate(text_lines):
        if _CITY_STATE_ZIP_RE.match(ln):
            csz_idx = i
            csz = ln
            break

    if csz_idx == -1:
        # Positional fallback when no city/state/zip pattern matched
        if len(text_lines) >= 1:
            name = text_lines[0]
        if len(text_lines) >= 2:
            street = text_lines[1]
        if len(text_lines) >= 3:
            csz = text_lines[2]
    elif csz_idx == 1:
        name = text_lines[0]
    elif csz_idx >= 2:
        name = text_lines[0]
        street = " ".join(text_lines[1:csz_idx]).strip()

    return {
        "name": name.strip(),
        "street": street.strip(),
        "city_state_zip": csz.strip(),
        "ssn": ssn,
    }


# ---------------------------------------------------------------------------
# Text-based fallback (for PDFs where coordinate extraction finds nothing)
# ---------------------------------------------------------------------------

_SSN_RE_LOOSE = re.compile(r"\b(\d{3}[-\s]\d{2}[-\s]\d{4})\b")
_ZIP_RE_TEXT = re.compile(r"\b\d{5}(?:-\d{4})?\b")
_NAME_RE = re.compile(r"\b([A-Z][a-zA-Z\-']+(?:\s+[A-Z][a-zA-Z\-']*){1,4})\b")

_NON_NAME_WORDS = {
    "EARNINGS", "SUMMARY", "WAGES", "TIPS", "COMPENSATION", "FEDERAL", "STATE",
    "LOCAL", "INCOME", "TAX", "WITHHELD", "SECURITY", "MEDICARE", "DEPENDENT",
    "CARE", "BENEFITS", "NONQUALIFIED", "PLANS", "ALLOCATED", "ADVANCE", "EIC",
    "PAYMENT", "SALARY", "BONUS", "TOTAL", "EMPLOYEE", "EMPLOYER", "GROSS",
    "NET", "PAY", "PERIOD", "YTD", "YEAR", "DATE", "AMOUNT", "COPY", "VOID",
    "CORRECTED", "DEPARTMENT", "TREASURY", "REVENUE", "SERVICE", "INTERNAL",
    "STATEMENT", "FORM", "WAGE", "AND", "TAX",
}

_EF_LABELS_TEXT = [
    "e/f employee's name, address, and zip code",
    "e/f employee's name, address, zip code",
    "employee's name, address, and zip code",
    "employee's name, address, zip code",
    "employee's name, address",
]

_SSN_LABELS_TEXT = [
    "employee's social security number",
    "social security number",
    "ssn",
    "social security no",
]


def _is_form_label(text: str) -> bool:
    return bool({w.upper() for w in text.split()} & _NON_NAME_WORDS)


def _extract_name_from_text(block: str) -> str:
    candidates = [c for c in _NAME_RE.findall(block) if not _is_form_label(c)]
    if not candidates:
        return ""
    return max(candidates, key=lambda s: len(s.split()))


def _text_fallback(page_text: str) -> dict:
    """Text-based extraction used when coordinate method yields no results."""
    lines = [ln for ln in page_text.splitlines() if ln.strip()]

    # SSN
    ssn = ""
    for line in lines:
        norm = line.lower().strip()
        for lbl in _SSN_LABELS_TEXT:
            if lbl in norm:
                idx = norm.index(lbl)
                remainder = line[idx + len(lbl):].strip().lstrip(":- ").strip()
                m = _SSN_RE_LOOSE.search(remainder or page_text)
                if m:
                    ssn = m.group(1)
                break
    if not ssn:
        m = _SSN_RE_LOOSE.search(page_text)
        ssn = m.group(1) if m else ""

    # Name + address via e/f label
    name = street = csz = ""
    for i, line in enumerate(lines):
        norm = line.lower().strip()
        if any(lbl in norm for lbl in _EF_LABELS_TEXT):
            subsequent = [ln.strip() for ln in lines[i + 1:] if ln.strip()]
            if subsequent and not _is_form_label(subsequent[0]):
                name = subsequent[0]
                for j in range(1, len(subsequent)):
                    if _ZIP_RE_TEXT.search(subsequent[j]):
                        csz = subsequent[j]
                        if j >= 2 and not _is_form_label(subsequent[j - 1]):
                            street = subsequent[j - 1]
                        break
            break

    return {"EmployeeName": name, "SSN": ssn, "Address": ", ".join(p for p in [street, csz] if p)}


# ---------------------------------------------------------------------------
# Core page extraction — coordinate-based primary, text fallback
# ---------------------------------------------------------------------------

def _extract_page_words(words: list, page_text: str) -> list[dict]:
    """
    Extract all employee records from a page.
    Returns a list of dicts (one per employee found).
    Uses coordinate-based extraction; falls back to text if nothing found.
    """
    blocks = _find_ef_blocks(words)
    records = []
    seen_ssns: set = set()
    seen_keys: set = set()

    for block in blocks:
        result = _extract_employee_from_block(words, block["top"], block["x0"])
        if not result:
            continue
        if result["ssn"] and result["ssn"] in seen_ssns:
            continue
        key = (result["name"], result["street"], result["city_state_zip"])
        if key == ("", "", ""):
            continue
        if not result["ssn"] and key in seen_keys:
            continue
        if result["ssn"]:
            seen_ssns.add(result["ssn"])
        seen_keys.add(key)

        address_parts = [p for p in [result["street"], result["city_state_zip"]] if p]
        records.append({
            "EmployeeName": result["name"],
            "SSN": result["ssn"],
            "Address": ", ".join(address_parts),
        })

    # If coordinate extraction found nothing, use text fallback
    if not records and page_text.strip():
        fb = _text_fallback(page_text)
        if fb["EmployeeName"] or fb["SSN"]:
            records.append(fb)

    return records


# ---------------------------------------------------------------------------
# Name and address normalization
# ---------------------------------------------------------------------------

_NAME_SUFFIXES = {"JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "ESQ", "ESQ."}

# "CITY STATE ZIP"  e.g. "PLEASANTON CA 94566-4477"
_CSZ_RE = re.compile(
    r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$"
)
# "STATE ZIP"  e.g. "WI 53923"  (city is a separate comma-segment)
_STATE_ZIP_RE = re.compile(r"^(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$")


def _split_name(full_name: str) -> dict:
    """
    Split a full name string into First, Middle, Last, Suffix.

    Handles ALL-CAPS W-2 names like "JOHN MICHAEL DOE JR" or "SMITH JANE A".
    W-2 names are sometimes stored Last-First, but we treat the raw order as-is
    (first token = first name) since ADP stores them First-Last.
    """
    tokens = full_name.strip().split()
    suffix = ""
    if tokens and tokens[-1].upper() in _NAME_SUFFIXES:
        suffix = tokens[-1]
        tokens = tokens[:-1]

    if not tokens:
        return {"FirstName": "", "MiddleName": "", "LastName": "", "Suffix": suffix}
    if len(tokens) == 1:
        return {"FirstName": tokens[0], "MiddleName": "", "LastName": "", "Suffix": suffix}
    if len(tokens) == 2:
        return {"FirstName": tokens[0], "MiddleName": "", "LastName": tokens[1], "Suffix": suffix}

    # 3+ tokens: first / middle(s) / last
    return {
        "FirstName":  tokens[0],
        "MiddleName": " ".join(tokens[1:-1]),
        "LastName":   tokens[-1],
        "Suffix":     suffix,
    }


def _split_address(address: str) -> dict:
    """
    Split a combined address string into StreetAddress, City, State, ZipCode.

    Handles two formats:
        "519 TRADITION PKWY 4200, PLEASANTON CA 94566-4477"  -> 2-part (city+state+zip fused)
        "801 W. COMMERCE ST UNIT 12, CAMBRIA, WI 53923"      -> 3-part (city and state+zip separate)
    """
    parts = [p.strip() for p in address.split(",") if p.strip()]
    street = city = state = zip_code = ""

    if not parts:
        return {"StreetAddress": street, "City": city, "State": state, "ZipCode": zip_code}

    # Format A: last segment is "STATE ZIP"  e.g. "WI 53923"
    m = _STATE_ZIP_RE.match(parts[-1])
    if m:
        state    = m.group("state")
        zip_code = m.group("zip")
        city     = parts[-2] if len(parts) >= 2 else ""
        street   = ", ".join(parts[:-2])
        return {"StreetAddress": street, "City": city, "State": state, "ZipCode": zip_code}

    # Format B: last segment is "CITY STATE ZIP"  e.g. "PLEASANTON CA 94566-4477"
    m = _CSZ_RE.match(parts[-1])
    if m:
        city     = m.group("city").strip()
        state    = m.group("state").strip()
        zip_code = m.group("zip").strip()
        street   = ", ".join(parts[:-1])
        return {"StreetAddress": street, "City": city, "State": state, "ZipCode": zip_code}

    # Fallback: return the whole string as street
    return {"StreetAddress": address, "City": city, "State": state, "ZipCode": zip_code}


# ---------------------------------------------------------------------------
# Standardized output builders
# ---------------------------------------------------------------------------

# Full column list matching the template (blank columns filled with "")
_STD_COLUMNS = [
    "Document Id", "Document",
    "First Name", "Middle Name", "Last Name", "Suffix",
    "Entity Type",
    "Street Address", "City", "State", "Zip Code",
    "Phone Number", "Email",
    "Int'l Street Address", "Int'l City", "Int'l Province/Region",
    "Int'l Zip Code", "Country (if applicable)",
    "Social Security Number", "Individual Tax ID",
    "Date of Birth", "Is Deceased", "Is Minor",
    "Driver's License Number", "Driver's License State",
    "Passport Number", "Other Government ID",
    "Financial Account Number", "Health Insurance ID",
    "Date of Service", "Medical Record Number", "Medical History",
    "Diagnosis/Condition", "Hospital/Facility",
    "Patient Account Number", "Biometric ID", "Vehicle Identification Number",
]


def _build_standardized_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    """Convert the raw extracted DataFrame into the standardized template layout."""
    rows = []
    for _, row in raw_df.iterrows():
        doc_id   = Path(row["File"]).stem          # filename without extension
        doc_name = row["File"]

        name_parts = _split_name(row.get("EmployeeName", "") or "")
        addr_parts = _split_address(row.get("Address", "") or "")

        std_row = {col: "" for col in _STD_COLUMNS}
        std_row["Document Id"]            = doc_id
        std_row["Document"]               = doc_name
        std_row["First Name"]             = name_parts["FirstName"]
        std_row["Middle Name"]            = name_parts["MiddleName"]
        std_row["Last Name"]              = name_parts["LastName"]
        std_row["Suffix"]                 = name_parts["Suffix"]
        std_row["Entity Type"]            = "Employee"
        std_row["Street Address"]         = addr_parts["StreetAddress"]
        std_row["City"]                   = addr_parts["City"]
        std_row["State"]                  = addr_parts["State"]
        std_row["Zip Code"]               = addr_parts["ZipCode"]
        std_row["Social Security Number"] = row.get("SSN", "")
        rows.append(std_row)

    return pd.DataFrame(rows, columns=_STD_COLUMNS)


def _build_processing_summary(frames_info: list) -> pd.DataFrame:
    """
    Build a per-document processing summary.

    frames_info — list of dicts:
        file, pages_processed, records, ssn_count, status, error
    """
    summary_rows = []
    for info in frames_info:
        doc_id = Path(info["file"]).stem
        summary_rows.append({
            "Document Name":    info["file"],
            "Document Id":      doc_id,
            "Pages Processed":  info["pages_processed"],
            "Entity Count":     info["records"],
            "SSN Count":        info["ssn_count"],
            "Names Found":      info["names_found"],
            "Addresses Found":  info["addresses_found"],
            "Status":           info["status"],
            "Error":            info.get("error", ""),
        })
    return pd.DataFrame(summary_rows)


# ---------------------------------------------------------------------------
# Excel writer — two sheets in one file + separate processing summary
# ---------------------------------------------------------------------------

def _save_excel(raw_df: pd.DataFrame, output_path: str,
                summary_df: pd.DataFrame | None = None) -> None:
    """
    Write output_path.xlsx with two sheets:
        "Extracted Data"   — raw fields (File, Page, EmployeeName, SSN, Address)
        "Standardized Data" — template-formatted fields

    Also writes <stem>_processing.xlsx alongside it with the document summary.
    """
    out = Path(output_path).with_suffix(".xlsx")
    std_df = _build_standardized_df(raw_df)

    with pd.ExcelWriter(str(out), engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Extracted Data", index=False)
        std_df.to_excel(writer, sheet_name="Standardized Data", index=False)

    print(f"  Excel saved: {out.name}  "
          f"({len(raw_df)} record(s), 2 sheets)")

    if summary_df is not None and not summary_df.empty:
        proc_path = out.parent / f"{out.stem}_processing.xlsx"
        with pd.ExcelWriter(str(proc_path), engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Processing Summary", index=False)
        print(f"  Processing summary: {proc_path.name}")


def extract_w2(file_path: str, max_pages: int = 0) -> tuple[pd.DataFrame, dict]:
    """
    Extract Employee Name, SSN, and Address from a W-2 or Earnings Summary PDF.

    Parameters
    ----------
    file_path : str
        Path to the PDF file.
    max_pages : int
        Maximum number of pages to process. 0 (default) means all pages.

    Returns
    -------
    (DataFrame, info_dict)
        DataFrame columns: File, Page, EmployeeName, SSN, Address.
        info_dict: file, pages_processed, records, ssn_count, names_found,
                   addresses_found, status, error.
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"PDF not found: {file_path}")

    print(f"Opening {path.name} ...", flush=True)
    rows = []
    with pdfplumber.open(str(path)) as pdf:
        if not pdf.pages:
            raise ValueError(f"PDF has no pages: {file_path}")

        total = len(pdf.pages)
        limit = min(total, max_pages) if max_pages > 0 else total
        pages_to_process = pdf.pages[:limit]

        if max_pages > 0 and max_pages < total:
            print(f"  {total} page(s) found. Processing first {limit} page(s)...", flush=True)
        else:
            print(f"  {total} page(s) found. Extracting...", flush=True)

        for page_num, page in enumerate(
            tqdm(pages_to_process, desc=f"  {path.name}", unit="pg", leave=True,
                 dynamic_ncols=True, miniters=1),
            start=1,
        ):
            try:
                words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
            except Exception:
                words = []
            page_text = page.extract_text() or ""

            records = _extract_page_words(words, page_text)
            for record in records:
                record["File"] = path.name
                record["Page"] = page_num
                rows.append(record)

    if not rows:
        df = pd.DataFrame(columns=["File", "Page", "EmployeeName", "SSN", "Address"])
    else:
        df = pd.DataFrame(rows)[["File", "Page", "EmployeeName", "SSN", "Address"]]

    info = {
        "file":             path.name,
        "pages_processed":  limit,
        "records":          len(df),
        "ssn_count":        int((df["SSN"] != "").sum()) if not df.empty else 0,
        "names_found":      int((df["EmployeeName"] != "").sum()) if not df.empty else 0,
        "addresses_found":  int((df["Address"] != "").sum()) if not df.empty else 0,
        "status":           "OK",
        "error":            "",
    }
    return df, info


# ---------------------------------------------------------------------------
# Batch extraction
# ---------------------------------------------------------------------------

def extract_w2_batch(input_dir: str, dest_dir: str, max_pages: int = 0) -> None:
    """
    Process every PDF in *input_dir*. Each PDF gets its own output file:
        <dest_dir>/<pdf_stem>.xlsx          — Extracted Data + Standardized Data sheets
        <dest_dir>/<pdf_stem>_processing.xlsx — per-document summary

    Parameters
    ----------
    input_dir : str
        Directory containing W-2 PDF files.
    dest_dir : str
        Destination directory for output files.
    max_pages : int
        Maximum pages to process per PDF. 0 (default) means all pages.
    """
    pdf_files = list(Path(input_dir).glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {input_dir}")
        return

    Path(dest_dir).mkdir(parents=True, exist_ok=True)
    total_records = 0

    for idx, pdf_path in enumerate(sorted(pdf_files), start=1):
        print(f"\n[{idx}/{len(pdf_files)}] {pdf_path.name}", flush=True)
        out_xlsx = Path(dest_dir) / f"{pdf_path.stem}.xlsx"
        try:
            df, info = extract_w2(str(pdf_path), max_pages=max_pages)
            summary_df = _build_processing_summary([info])
            _save_excel(df, str(out_xlsx), summary_df)
            total_records += len(df)
            print(f"  [OK]  {len(df)} records -> {out_xlsx.name}", flush=True)
        except Exception as exc:
            print(f"  [ERR] {exc}", flush=True)
            err_info = {
                "file": pdf_path.name, "pages_processed": 0, "records": 0,
                "ssn_count": 0, "names_found": 0, "addresses_found": 0,
                "status": "ERROR", "error": str(exc),
            }
            _save_excel(
                pd.DataFrame(columns=["File", "Page", "EmployeeName", "SSN", "Address"]),
                str(out_xlsx),
                _build_processing_summary([err_info]),
            )

    print(f"\nBatch done. {total_records} total record(s) across {len(pdf_files)} file(s).")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def _debug_lines(file_path: str) -> None:
    """Print word positions from the first page to help tune layout constants."""
    with pdfplumber.open(file_path) as pdf:
        words = pdf.pages[0].extract_words(use_text_flow=False, keep_blank_chars=False)
    print(f"\n--- Word positions from first page ({len(words)} words) ---")
    print(f"  {'x0':>6}  {'top':>6}  text")
    print(f"  {'-'*6}  {'-'*6}  ----")
    for w in words[:120]:
        print(f"  {w['x0']:>6.1f}  {w['top']:>6.1f}  {w['text']}")
    if len(words) > 120:
        print(f"  ... ({len(words) - 120} more words)")
    print("---")


if __name__ == "__main__":
    # Usage:
    #   python w2_extractor.py <pdf_file>  [dest_dir]  [--pages N]
    #   python w2_extractor.py <pdf_folder> [dest_dir] [--pages N]
    #   python w2_extractor.py <pdf_file> --debug
    #
    # Single file — output saved as <dest_dir>/<pdf_stem>.xlsx  (default: same folder as PDF)
    #   python w2_extractor.py sample.pdf
    #   python w2_extractor.py sample.pdf "C:\Output"
    #   python w2_extractor.py sample.pdf --pages 10
    #
    # Folder — each PDF gets its own <pdf_stem>.xlsx in dest_dir
    #   python w2_extractor.py "C:\PDFs\"
    #   python w2_extractor.py "C:\PDFs\" "C:\Output" --pages 20
    #
    # DO NOT pass real W-2 files containing live PII — use anonymised samples only.

    import argparse

    parser = argparse.ArgumentParser(
        prog="w2_extractor",
        description="Extract Employee Name, SSN, and Address from W-2 PDFs.",
    )
    parser.add_argument("target", help="PDF file or folder of PDFs to process")
    parser.add_argument("dest", nargs="?", default=None,
                        help="Destination folder for output files (default: same folder as input)")
    parser.add_argument("--pages", type=int, default=0, metavar="N",
                        help="Only process the first N pages per file (default: all pages)")
    parser.add_argument("--debug", action="store_true",
                        help="Print word positions from first page and exit")
    args = parser.parse_args()

    if args.debug:
        _debug_lines(args.target)
        sys.exit(0)

    if args.pages < 0:
        print("Error: --pages must be a positive integer.")
        sys.exit(1)

    if os.path.isdir(args.target):
        dest = args.dest or args.target
        extract_w2_batch(args.target, dest, max_pages=args.pages)
    else:
        pdf_path = Path(args.target)
        dest_dir = Path(args.dest) if args.dest else pdf_path.parent
        out_xlsx = dest_dir / f"{pdf_path.stem}.xlsx"
        result, info = extract_w2(args.target, max_pages=args.pages)
        summary_df = _build_processing_summary([info])
        print(result.to_string(index=False))
        print(f"\nExtracted {len(result)} record(s) total")
        _save_excel(result, str(out_xlsx), summary_df)
