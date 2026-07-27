# w2_wages_extractor.py
# Extracts Employee Name, SSN, Address, AND wage/tax box amounts from W-2 PDF files.
# Based on w2_extractor.py — adds Box 1–6 amount extraction.
# In batch mode, each PDF is written to its own output Excel file as it is processed.
#
# Extracted W-2 boxes:
#   Box 1  — Wages, Tips, Other Compensation
#   Box 2  — Federal Income Tax Withheld
#   Box 3  — Social Security Wages
#   Box 4  — Social Security Tax Withheld
#   Box 5  — Medicare Wages and Tips
#   Box 6  — Medicare Tax Withheld
#
# Outputs (batch mode — one file per PDF, written as each is processed):
#   <dest>/<pdf_stem>_wages.xlsx          — "Extracted Data" + "Standardized Data" sheets
#   <dest>/<pdf_stem>_wages_processing.xlsx — per-document summary
#
# Outputs (single file mode):
#   <dest>/<pdf_stem>.xlsx
#   <dest>/<pdf_stem>_processing.xlsx
#
# Requires: pdfplumber, pandas, openpyxl, tqdm
#   pip install pdfplumber pandas openpyxl tqdm
#
# ---------------------------------------------------------------------------
# USAGE
# ---------------------------------------------------------------------------
#
# Debug mode — print raw word positions from first page:
#
#   python w2_wages_extractor.py "C:\YourFolder\sample_w2.pdf" --debug
#
# Single file:
#
#   python w2_wages_extractor.py "C:\YourFolder\sample_w2.pdf"
#   python w2_wages_extractor.py "C:\YourFolder\sample_w2.pdf" "C:\Output\results.xlsx"
#
# Batch folder (one Excel file written per PDF, as processed):
#
#   python w2_wages_extractor.py "C:\YourFolder\PDFs\" "C:\Output\"
#   python w2_wages_extractor.py "C:\YourFolder\PDFs\" "C:\Output\" --pages 10
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

_SSN_PATTERN = re.compile(r"^\d{3}-\d{2}-\d{4}$")
_SSN_IN_LINE = re.compile(r"\b(\d{3}-\d{2}-\d{4})\b")
_CITY_STATE_ZIP_RE = re.compile(r"^.+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\s*$")
_ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")

# Dollar amount: optional $, digits with commas, optional decimal
_AMOUNT_RE = re.compile(r"\$?\s*(\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)")


# ---------------------------------------------------------------------------
# Layout constants (tuned for ADP W-2 layout)
# ---------------------------------------------------------------------------

_EF_LABEL_MAX_Y = 280
_EF_LABEL_MAX_X = 60
_SSN_OFFSET_MIN = 30
_SSN_OFFSET_MAX = 80
_BOTTOM_ROW_Y_MIN = 450


# ---------------------------------------------------------------------------
# Wage box labels — the text that appears ON the W-2 form near each box
# ---------------------------------------------------------------------------

# Each entry: (box_number, list of label substrings to match, case-insensitive)
_WAGE_BOX_DEFS = [
    ("box1_wages",    ["wages, tips, other comp", "wages, tips, other", "1 wages"]),
    ("box2_fed_tax",  ["federal income tax withheld", "fed income tax", "2 federal"]),
    ("box3_ss_wages", ["social security wages", "3 social security wages", "ss wages"]),
    ("box4_ss_tax",   ["social security tax withheld", "4 social security tax", "ss tax withheld"]),
    ("box5_med_wages",["medicare wages and tips", "5 medicare wages", "medicare wages"]),
    ("box6_med_tax",  ["medicare tax withheld", "6 medicare tax", "medicare tax"]),
]


# ---------------------------------------------------------------------------
# Coordinate-based employee block extraction (unchanged from w2_extractor.py)
# ---------------------------------------------------------------------------

def _find_ef_blocks(words: list) -> list:
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
    # Window is generous (covers a name line plus several address lines, e.g.
    # a street line AND a unit/apt line before the city/state/zip line) since
    # multi-line addresses would otherwise get truncated before reaching the
    # city/state/zip line. Line collection below stops naturally once the
    # city/state/zip line is found, so this doesn't risk pulling in unrelated
    # content (e.g. the SSN row) in the common case.
    block_words = [
        w for w in words
        if label_top + 2 < w["top"] < label_top + 90
        and w["x0"] < 200
        and w["x0"] > label_x0 - 5
    ]
    block_words.sort(key=lambda w: (w["top"], w["x0"]))

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

    _MAX_BLOCK_LINES = 6
    text_lines = []
    for line in lines:
        line.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in line).strip()
        if not text or len(text) <= 1:
            continue
        if _SSN_IN_LINE.search(text):
            # Reached the SSN row — stop before pulling it into the address.
            break
        text_lines.append(text)
        if _CITY_STATE_ZIP_RE.match(text) or len(text_lines) >= _MAX_BLOCK_LINES:
            # City/state/zip line found — the name/address block ends here.
            break

    if len(text_lines) < 2:
        return None

    ssn_candidates = []
    for w in words:
        if label_top + _SSN_OFFSET_MIN < w["top"] < label_top + _SSN_OFFSET_MAX:
            if w["x0"] < 200:
                m = _SSN_IN_LINE.search(w["text"])
                if m and _SSN_PATTERN.match(m.group(1)):
                    ssn_candidates.append(m.group(1))

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

    name = street = csz = ""
    csz_idx = -1
    for i, ln in enumerate(text_lines):
        if _CITY_STATE_ZIP_RE.match(ln):
            csz_idx = i
            csz = ln
            break

    if csz_idx == -1:
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
# Wage amount extraction — text-based, scans the full page text
# ---------------------------------------------------------------------------

def _parse_amount(text: str) -> str:
    """Return the first dollar amount found in text, or empty string."""
    m = _AMOUNT_RE.search(text)
    if m:
        raw = m.group(1).replace(",", "")
        try:
            return f"{float(raw):.2f}"
        except ValueError:
            return ""
    return ""


def _extract_wages_from_text(page_text: str) -> dict:
    """
    Scan page text for W-2 wage/tax boxes.
    Returns a dict keyed by the box field names in _WAGE_BOX_DEFS.
    Strategy: find the label line, then look at the same line or the next line
    for the dollar amount.
    """
    lines = page_text.splitlines()
    results = {key: "" for key, _ in _WAGE_BOX_DEFS}

    for box_key, label_list in _WAGE_BOX_DEFS:
        for i, line in enumerate(lines):
            line_lower = line.lower()
            matched = any(lbl in line_lower for lbl in label_list)
            if not matched:
                continue

            # Try to find amount on the same line (after the label)
            amount = _parse_amount(line)
            if not amount and i + 1 < len(lines):
                # Try the next line
                amount = _parse_amount(lines[i + 1])
            if amount:
                results[box_key] = amount
                break

    return results


def _extract_wages_from_words(words: list) -> dict:
    """
    Coordinate-aware wage extraction.
    For each wage box label found, look for a nearby amount word to the right or below.
    Falls back gracefully — used to supplement _extract_wages_from_text.
    """
    results = {key: "" for key, _ in _WAGE_BOX_DEFS}

    # Build a joined string per Y-band to find label positions
    LINE_TOL = 6

    def words_near(target_top: float, x_min: float = 0, x_max: float = 9999,
                   y_slack: float = 20) -> list:
        return [
            w for w in words
            if abs(w["top"] - target_top) <= y_slack
            and x_min <= w["x0"] <= x_max
        ]

    for box_key, label_list in _WAGE_BOX_DEFS:
        if results[box_key]:
            continue
        for w in words:
            text_lower = w["text"].lower()
            if any(lbl in text_lower for lbl in label_list):
                # Amount typically appears to the right of the label on the same line
                nearby = words_near(w["top"], x_min=w["x1"], y_slack=LINE_TOL)
                nearby.sort(key=lambda ww: ww["x0"])
                for cand in nearby:
                    amt = _parse_amount(cand["text"])
                    if amt:
                        results[box_key] = amt
                        break
                break

    return results


# ---------------------------------------------------------------------------
# Text-based fallback for employee identity (unchanged from w2_extractor.py)
# ---------------------------------------------------------------------------

_SSN_RE_LOOSE = re.compile(r"\b(\d{3}[-\s]\d{2}[-\s]\d{4})\b")
_ZIP_RE_TEXT  = re.compile(r"\b\d{5}(?:-\d{4})?\b")
_NAME_RE      = re.compile(r"\b([A-Z][a-zA-Z\-']+(?:\s+[A-Z][a-zA-Z\-']*){1,4})\b")

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
    lines = [ln for ln in page_text.splitlines() if ln.strip()]

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
# Core page extraction
# ---------------------------------------------------------------------------

def _extract_page(words: list, page_text: str) -> list[dict]:
    """
    Extract all employee records + wage amounts from a page.
    Returns list of dicts, one per employee found.
    """
    # Employee identity (coordinate-based primary)
    blocks = _find_ef_blocks(words)
    id_records = []
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
        id_records.append({
            "EmployeeName": result["name"],
            "SSN": result["ssn"],
            "Address": ", ".join(address_parts),
        })

    if not id_records and page_text.strip():
        fb = _text_fallback(page_text)
        if fb["EmployeeName"] or fb["SSN"]:
            id_records.append(fb)

    if not id_records:
        return []

    # Wage amounts — extract once per page, apply to all records on the page
    wages_coord = _extract_wages_from_words(words)
    wages_text  = _extract_wages_from_text(page_text)

    # Merge: coordinate result wins if present, otherwise text result
    wages = {}
    for box_key, _ in _WAGE_BOX_DEFS:
        wages[box_key] = wages_coord.get(box_key) or wages_text.get(box_key) or ""

    records = []
    for rec in id_records:
        records.append({**rec, **wages})

    return records


# ---------------------------------------------------------------------------
# Name and address normalization
# ---------------------------------------------------------------------------

_NAME_SUFFIXES = {"JR", "JR.", "SR", "SR.", "II", "III", "IV", "V", "ESQ", "ESQ."}

_CSZ_RE = re.compile(
    r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$"
)
_STATE_ZIP_RE = re.compile(r"^(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$")


def _split_name(full_name: str) -> dict:
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
    return {
        "FirstName":  tokens[0],
        "MiddleName": " ".join(tokens[1:-1]),
        "LastName":   tokens[-1],
        "Suffix":     suffix,
    }


def _split_address(address: str) -> dict:
    parts = [p.strip() for p in address.split(",") if p.strip()]
    street = city = state = zip_code = ""

    if not parts:
        return {"StreetAddress": street, "City": city, "State": state, "ZipCode": zip_code}

    m = _STATE_ZIP_RE.match(parts[-1])
    if m:
        state    = m.group("state")
        zip_code = m.group("zip")
        city     = parts[-2] if len(parts) >= 2 else ""
        street   = ", ".join(parts[:-2])
        return {"StreetAddress": street, "City": city, "State": state, "ZipCode": zip_code}

    m = _CSZ_RE.match(parts[-1])
    if m:
        city     = m.group("city").strip()
        state    = m.group("state").strip()
        zip_code = m.group("zip").strip()
        street   = ", ".join(parts[:-1])
        return {"StreetAddress": street, "City": city, "State": state, "ZipCode": zip_code}

    return {"StreetAddress": address, "City": city, "State": state, "ZipCode": zip_code}


# ---------------------------------------------------------------------------
# Standardized output columns
# ---------------------------------------------------------------------------

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
    # Wage fields appended after the standard template columns
    "Box 1 Wages Tips Other Comp",
    "Box 2 Federal Tax Withheld",
    "Box 3 SS Wages",
    "Box 4 SS Tax Withheld",
    "Box 5 Medicare Wages",
    "Box 6 Medicare Tax Withheld",
]

_BOX_KEY_TO_STD = {
    "box1_wages":     "Box 1 Wages Tips Other Comp",
    "box2_fed_tax":   "Box 2 Federal Tax Withheld",
    "box3_ss_wages":  "Box 3 SS Wages",
    "box4_ss_tax":    "Box 4 SS Tax Withheld",
    "box5_med_wages": "Box 5 Medicare Wages",
    "box6_med_tax":   "Box 6 Medicare Tax Withheld",
}


def _build_standardized_df(raw_df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in raw_df.iterrows():
        doc_id   = Path(row["File"]).stem
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

        for box_key, col_name in _BOX_KEY_TO_STD.items():
            std_row[col_name] = row.get(box_key, "")

        rows.append(std_row)

    return pd.DataFrame(rows, columns=_STD_COLUMNS)


def _build_processing_summary(frames_info: list) -> pd.DataFrame:
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
            "Wages Found":      info["wages_found"],
            "Status":           info["status"],
            "Error":            info.get("error", ""),
        })
    return pd.DataFrame(summary_rows)


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

def _save_excel(raw_df: pd.DataFrame, output_path: str,
                summary_df: pd.DataFrame | None = None) -> None:
    """
    Write output_path.xlsx with two sheets:
        "Extracted Data"    — raw fields including wage boxes
        "Standardized Data" — template-formatted fields
    Also writes <stem>_processing.xlsx alongside it.
    """
    out = Path(output_path).with_suffix(".xlsx")
    std_df = _build_standardized_df(raw_df)

    with pd.ExcelWriter(str(out), engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Extracted Data", index=False)
        std_df.to_excel(writer, sheet_name="Standardized Data", index=False)

    print(f"  Excel saved: {out.name}  ({len(raw_df)} record(s), 2 sheets)")

    if summary_df is not None and not summary_df.empty:
        proc_path = out.parent / f"{out.stem}_processing.xlsx"
        with pd.ExcelWriter(str(proc_path), engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Processing Summary", index=False)
        print(f"  Processing summary: {proc_path.name}")


# ---------------------------------------------------------------------------
# Per-file extraction
# ---------------------------------------------------------------------------

_RAW_COLS = [
    "File", "Page", "EmployeeName", "SSN", "Address",
    "box1_wages", "box2_fed_tax", "box3_ss_wages",
    "box4_ss_tax", "box5_med_wages", "box6_med_tax",
]


def extract_w2(file_path: str, max_pages: int = 0) -> tuple[pd.DataFrame, dict]:
    """
    Extract employee identity + wage box amounts from a W-2 PDF.

    Returns
    -------
    (DataFrame, info_dict)
        DataFrame columns: File, Page, EmployeeName, SSN, Address,
                           box1_wages ... box6_med_tax.
        info_dict: file, pages_processed, records, ssn_count, names_found,
                   addresses_found, wages_found, status, error.
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

            records = _extract_page(words, page_text)
            for rec in records:
                rec["File"] = path.name
                rec["Page"] = page_num
                rows.append(rec)

    if not rows:
        df = pd.DataFrame(columns=_RAW_COLS)
    else:
        df = pd.DataFrame(rows)
        for col in _RAW_COLS:
            if col not in df.columns:
                df[col] = ""
        df = df[_RAW_COLS]

    wages_found = int(
        df["box1_wages"].ne("").sum()
        if not df.empty else 0
    )

    info = {
        "file":             path.name,
        "pages_processed":  limit,
        "records":          len(df),
        "ssn_count":        int((df["SSN"] != "").sum()) if not df.empty else 0,
        "names_found":      int((df["EmployeeName"] != "").sum()) if not df.empty else 0,
        "addresses_found":  int((df["Address"] != "").sum()) if not df.empty else 0,
        "wages_found":      wages_found,
        "status":           "OK",
        "error":            "",
    }
    return df, info


# ---------------------------------------------------------------------------
# Batch extraction — one Excel file written per PDF, as each is processed
# ---------------------------------------------------------------------------

def extract_w2_batch(input_dir: str, output_dir: str, max_pages: int = 0) -> None:
    """
    Process every PDF in *input_dir*, writing a separate Excel file for each
    one as soon as it finishes (no combined output).

    Parameters
    ----------
    input_dir : str
        Directory containing W-2 PDF files.
    output_dir : str
        Directory to write the per-PDF output .xlsx files into.
    max_pages : int
        Maximum pages to process per PDF. 0 = all pages.
    """
    pdf_files = list(Path(input_dir).glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in: {input_dir}")
        return

    out_dir = Path(output_dir)
    if out_dir.suffix:
        out_dir = out_dir.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    total_records = 0

    for idx, pdf_path in enumerate(sorted(pdf_files), start=1):
        print(f"\n[{idx}/{len(pdf_files)}] {pdf_path.name}", flush=True)
        try:
            df, info = extract_w2(str(pdf_path), max_pages=max_pages)
            print(f"  [OK]  {len(df)} record(s)  wages found: {info['wages_found']}", flush=True)
            summary_df = _build_processing_summary([info])
            out_xlsx = out_dir / f"{pdf_path.stem}_wages.xlsx"
            _save_excel(df, str(out_xlsx), summary_df)
            total_records += len(df)
        except Exception as exc:
            print(f"  [ERR] {exc}", flush=True)

    print(f"\nBatch done. {total_records} total record(s) across {len(pdf_files)} file(s). "
          f"Individual Excel files written to {out_dir}")


# ---------------------------------------------------------------------------
# Debug helper
# ---------------------------------------------------------------------------

def _debug_lines(file_path: str) -> None:
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


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        prog="w2_wages_extractor",
        description="Extract Employee Name, SSN, Address, and Wage Amounts from W-2 PDFs.",
    )
    parser.add_argument("target", help="PDF file or folder of PDFs to process")
    parser.add_argument("dest", nargs="?", default=None,
                        help="Output file path or folder (batch: single combined .xlsx)")
    parser.add_argument("--pages", type=int, default=0, metavar="N",
                        help="Only process the first N pages per file (default: all)")
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
        if dest_dir.suffix == ".xlsx":
            out_xlsx = dest_dir
        else:
            out_xlsx = dest_dir / f"{pdf_path.stem}_wages.xlsx"
        result, info = extract_w2(args.target, max_pages=args.pages)
        summary_df = _build_processing_summary([info])
        print(result.to_string(index=False))
        print(f"\nExtracted {len(result)} record(s)  wages found: {info['wages_found']}")
        _save_excel(result, str(out_xlsx), summary_df)
