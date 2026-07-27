r"""
bucket_11.py — Extract Employee Payroll Changes from ADP "Employee Payroll Changes" PDFs.

Document structure
------------------
The PDF is a 3-column table:

    CHANGED FIELD  |  CHANGED FROM  |  CHANGED TO

Employee header rows span all three columns and look like:
    [code]  Last, First [MI]  Associate ID: <id>  Position ID: <id>
    Home Department: <dept>  Home Cost Number: <n>

This script detects column boundaries from the "CHANGED FIELD / CHANGED FROM /
CHANGED TO" header row, then assigns each subsequent line's words to the
appropriate column by x-coordinate.  Employee header rows are identified by
the presence of "Associate ID:" and parsed for name + IDs.

Output — one row per change entry per employee:
    DOCID | Last Name | First Name | Middle Name |
    Associate ID | Position ID | Home Department | Home Cost Number |
    Changed Field | Changed From | Changed To

USAGE
    python bucket_11.py "C:\path\to\file.pdf"
    python bucket_11.py "C:\path\to\folder" --recursive
    python bucket_11.py "C:\path\to\file.pdf" --debug
    python bucket_11.py "C:\path\to\file.pdf" --diagnose [--diagnose-pages N]
    python bucket_11.py --selftest
"""

from __future__ import annotations

import argparse
import difflib
import os
import re
import sys
import time
from pathlib import Path

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import openpyxl
    from openpyxl.styles import Font
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from tqdm import tqdm as _tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False


# ---------------------------------------------------------------------------
# Output schema
# ---------------------------------------------------------------------------

OUTPUT_HEADERS = [
    "DOCID",
    "Name",
    "Associate ID",
    "Position ID",
    "Home Department",
    "Home Cost Number",
    "Changed Field",
    "Changed From",
    "Changed To",
]

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_LINE_TOLERANCE = 3.0   # px: words within this vertical gap share a line
_NAME_SUFFIXES   = {"jr", "sr", "ii", "iii", "iv", "v"}

# Page-width fractions used when the table header can't be found
_DEFAULT_COL2_FRAC = 0.38
_DEFAULT_COL3_FRAC = 0.67

_NO_RECORDS = (
    "No employee records found — no 'Associate ID:' header rows detected in "
    "this PDF.  Run --diagnose to inspect the raw text."
)

# Words that are clearly page/table chrome, not data
_CHROME_WORDS = {
    "changed", "field", "from", "to", "employee", "payroll", "changes",
    "page", "of",
}

# ---------------------------------------------------------------------------
# Bucket-11 → Template column mapping
# ---------------------------------------------------------------------------
# Each entry: bucket11_header → (value_col, cat_col, cat_subitem)
#   value_col   : template column that receives the actual changed value (None = no direct col)
#   cat_col     : category column to merge a sub-item into via ";" (None = skip)
#   cat_subitem : sub-item label appended to cat_col; None = don't touch cat_col
_BUCKET11_MAP: dict[str, tuple] = {
    "Additional Earnings Amount":       (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Address - City":                   ("City",                          "Contact Information",              "Home address"),
    "Address - Country":                ("Country of Residence",          "Contact Information",              "Home address"),
    "Address - Line 1":                 ("Residential Address",           "Contact Information",              "Home address"),
    "Address - Line 2":                 ("Residential Address",           "Contact Information",              "Home address"),
    "Address - Line 2 No":              ("Residential Address",           "Contact Information",              "Home address"),
    "Address - Line 3":                 ("Residential Address",           "Contact Information",              "Home address"),
    "Address - State":                  ("State of Residence (if US)",    "Contact Information",              "Home address"),
    "Address - State No":               ("State of Residence (if US)",    "Contact Information",              "Home address"),
    "Address - Zip / Postal Code":      ("Zip Code",                      "Contact Information",              "Home address"),
    "Address - Zip / Postal Code No":   ("Zip Code",                      "Contact Information",              "Home address"),
    "Basis of Pay":                     (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Basis of Pay No":                  (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Defer Social Security tax":        (None,                            "Government-Issued Identification", "Taxpayer Identification Number (TIN)"),
    "Dependents No":                    (None,                            "Family Information",               None),
    "Direct Deposit - Account Number":  (None,                            "Financial Account Information",    "Financial account number"),
    "Employee Name - First":            ("First Name",                    None,                               None),
    "Employee Name - First No":         ("First Name",                    None,                               None),
    "Employee Name - Last":             ("Last Name",                     None,                               None),
    "Employee Name - Last No":          ("Last Name",                     None,                               None),
    "Employee Name - Middle":           ("Middle Name",                   None,                               None),
    "Employee Name - Preferred":        ("PI Notes",                      None,                               None),
    "Employee Name - Salutation":       ("Suffix",                        None,                               None),
    "Ethnicity/Race":                   (None,                            "Demographic Information",          "Race/ Ethnicity"),
    "Ethnicity/Race No":                (None,                            "Demographic Information",          "Race/ Ethnicity"),
    "Gender for Insurance Coverage No": (None,                            "Health Related Information",       "Health Insurance Information"),
    "Home Phone":                       ("Phone Number",                  "Contact Information",              "Personal phone number (home)"),
    "Job Title No":                     (None,                            "Work-Related Information",         "Employment Application Information"),
    "Lien Dependent Medical Insurance": (None,                            "Health Related Information",       "Health Insurance Information"),
    "Other Income":                     (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Payee":                            (None,                            "Financial Account Information",    "Financial account number"),
    "Payroll Name - First":             ("PI Notes",                      None,                               None),
    "Payroll Name - First No":          ("PI Notes",                      None,                               None),
    "Payroll Name - Last":              ("PI Notes",                      None,                               None),
    "Payroll Name - Last No":           ("PI Notes",                      None,                               None),
    "Personal E-mail":                  ("Email Address - Personal",      "Contact Information",              "Personal email address"),
    "Personal E-mail No":               ("Email Address - Personal",      "Contact Information",              "Personal email address"),
    "Personal Mobile":                  ("Phone Number",                  "Contact Information",              "Personal phone number (mobile)"),
    "Personal Mobile No":               ("Phone Number",                  "Contact Information",              "Personal phone number (mobile)"),
    "Rate 1":                           (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Rate 1 No":                        (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Rate 2":                           (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Rate 6":                           (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Rate 8":                           (None,                            "Work-Related Information",         "Salary or Compensation Information"),
    "Social Security Number":           ("Social Security Number",        "Government-Issued Identification", "Social Security Number (SSN)"),
    "Status":                           (None,                            "Work-Related Information",         "Employment Application Information"),
    "Tax ID Type":                      (None,                            "Government-Issued Identification", "Taxpayer Identification Number (TIN)"),
    "Termination Date":                 (None,                            "Work-Related Information",         "Employment Application Information"),
    "Termination Reason":               (None,                            "Work-Related Information",         "Disciplinary Record or Report"),
}

# Alias table: normalised alias → exact Bucket-11 header name.
# Checked before fuzzy matching so abbreviations and common variants always resolve.
_ALIASES: dict[str, str] = {
    # Social Security
    "ssn":                          "Social Security Number",
    "ss number":                    "Social Security Number",
    "social security":              "Social Security Number",
    "soc sec":                      "Social Security Number",
    "ss no":                        "Social Security Number",
    # Name fields
    "first name":                   "Employee Name - First",
    "last name":                    "Employee Name - Last",
    "middle name":                  "Employee Name - Middle",
    "preferred name":               "Employee Name - Preferred",
    "salutation":                   "Employee Name - Salutation",
    "payroll first name":           "Payroll Name - First",
    "payroll last name":            "Payroll Name - Last",
    # Contact
    "email":                        "Personal E-mail",
    "e mail":                       "Personal E-mail",
    "personal email":               "Personal E-mail",
    "work email":                   "Personal E-mail",
    "mobile":                       "Personal Mobile",
    "cell":                         "Personal Mobile",
    "cell phone":                   "Personal Mobile",
    "phone":                        "Home Phone",
    "home phone":                   "Home Phone",
    "telephone":                    "Home Phone",
    # Address
    "city":                         "Address - City",
    "state":                        "Address - State",
    "zip":                          "Address - Zip / Postal Code",
    "zip code":                     "Address - Zip / Postal Code",
    "postal code":                  "Address - Zip / Postal Code",
    "postcode":                     "Address - Zip / Postal Code",
    "country":                      "Address - Country",
    "address":                      "Address - Line 1",
    "address line 1":               "Address - Line 1",
    "address line 2":               "Address - Line 2",
    "address line 3":               "Address - Line 3",
    "street":                       "Address - Line 1",
    # Pay / rates
    "pay rate":                     "Rate 1",
    "hourly rate":                  "Rate 1",
    "rate":                         "Rate 1",
    "wage":                         "Rate 1",
    "wages":                        "Rate 1",
    "hourly wage":                  "Rate 1",
    "salary rate":                  "Rate 1",
    "base pay":                     "Basis of Pay",
    "salary":                       "Basis of Pay",
    "annual salary":                "Basis of Pay",
    "base salary":                  "Basis of Pay",
    "gross pay":                    "Basis of Pay",
    "pay basis":                    "Basis of Pay",
    "earnings":                     "Additional Earnings Amount",
    "additional pay":               "Additional Earnings Amount",
    "other pay":                    "Other Income",
    "income":                       "Other Income",
    "other income":                 "Other Income",
    "gross income":                 "Other Income",
    "annual income":                "Other Income",
    "net income":                   "Other Income",
    # Demographics
    "gender":                       "Gender for Insurance Coverage No",
    "sex":                          "Gender for Insurance Coverage No",
    "race":                         "Ethnicity/Race",
    "ethnicity":                    "Ethnicity/Race",
    "race ethnicity":               "Ethnicity/Race",
    # Employment
    "job title":                    "Job Title No",
    "title":                        "Job Title No",
    "status":                       "Status",
    "employment status":            "Status",
    "term date":                    "Termination Date",
    "termination":                  "Termination Date",
    "separation date":              "Termination Date",
    "term reason":                  "Termination Reason",
    "separation reason":            "Termination Reason",
    "reason for termination":       "Termination Reason",
    # Financial
    "direct deposit":               "Direct Deposit - Account Number",
    "bank account":                 "Direct Deposit - Account Number",
    "account number":               "Direct Deposit - Account Number",
    "dd account":                   "Direct Deposit - Account Number",
    "payee":                        "Payee",
    # Tax / ID
    "tax id":                       "Tax ID Type",
    "tin":                          "Tax ID Type",
    "taxpayer id":                  "Tax ID Type",
    "defer ss":                     "Defer Social Security tax",
    "defer soc sec":                "Defer Social Security tax",
    # Other
    "dependents":                   "Dependents No",
    "dependent count":              "Dependents No",
    "medical insurance":            "Lien Dependent Medical Insurance",
    "lien":                         "Lien Dependent Medical Insurance",
}


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _normalise(text: str) -> str:
    return " ".join(str(text or "").split()).lower()


def _clean(text: str) -> str:
    return " ".join(str(text or "").split()).strip()


# ---------------------------------------------------------------------------
# Name parsing
# ---------------------------------------------------------------------------

def parse_name(cell: str) -> dict:
    """Split "Last, First Middle [Suffix]" into components."""
    blank = {"Last Name": "", "First Name": "", "Middle Name": "", "Suffix": ""}
    s = _clean(cell)
    if not s or "," not in s:
        return blank

    last_part, _, rest = s.partition(",")
    last_toks = last_part.split()
    rest_toks = rest.split()

    suffix = ""
    if rest_toks and rest_toks[-1].strip(".").lower() in _NAME_SUFFIXES:
        suffix = rest_toks[-1]
        rest_toks = rest_toks[:-1]
    elif last_toks and last_toks[-1].strip(".").lower() in _NAME_SUFFIXES:
        suffix = last_toks[-1]
        last_toks = last_toks[:-1]

    return {
        "Last Name":   " ".join(last_toks),
        "First Name":  rest_toks[0] if rest_toks else "",
        "Middle Name": " ".join(rest_toks[1:]) if len(rest_toks) > 1 else "",
        "Suffix":      suffix,
    }


# ---------------------------------------------------------------------------
# PDF word-group helpers
# ---------------------------------------------------------------------------

def _group_lines(words: list[dict]) -> list[list[dict]]:
    """Cluster pdfplumber word dicts into visual lines by top-coordinate."""
    if not words:
        return []
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines: list[list[dict]] = []
    current: list[dict] = [ordered[0]]
    current_top = ordered[0]["top"]
    for w in ordered[1:]:
        if abs(w["top"] - current_top) <= _LINE_TOLERANCE:
            current.append(w)
        else:
            lines.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_top = w["top"]
    lines.append(sorted(current, key=lambda x: x["x0"]))
    return lines


def _line_text(line: list[dict]) -> str:
    """Join all words in a line left-to-right."""
    return " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))


# ---------------------------------------------------------------------------
# Column-boundary detection
# ---------------------------------------------------------------------------

def _detect_boundaries(lines: list[list[dict]], page_width: float) -> tuple[float, float]:
    """Return (col2_x, col3_x) from the "CHANGED FIELD / FROM / TO" header row.

    Falls back to fixed page-width fractions if the header row is not found.
    """
    for line in lines:
        norm = _normalise(_line_text(line))
        if "changed" in norm and "from" in norm and "to" in norm:
            # Map each word to its x0
            words = sorted(line, key=lambda w: w["x0"])
            from_x = to_x = None
            for i, w in enumerate(words):
                t = w["text"].lower()
                # "CHANGED FROM" / "CHANGED TO": use x0 of "CHANGED" as the
                # column start — this is where the column header begins.
                if t == "changed" and i + 1 < len(words):
                    nxt = words[i + 1]["text"].lower()
                    if nxt == "from":
                        from_x = w["x0"]
                    elif nxt == "to":
                        to_x = w["x0"]
                # Standalone "FROM" / "TO" — only apply if not already set
                elif t == "from" and from_x is None:
                    from_x = w["x0"]
                elif t == "to" and to_x is None:
                    to_x = w["x0"]
            if from_x and to_x and from_x < to_x:
                return from_x, to_x
    # Fallback
    return page_width * _DEFAULT_COL2_FRAC, page_width * _DEFAULT_COL3_FRAC


# ---------------------------------------------------------------------------
# Line classification
# ---------------------------------------------------------------------------

def _is_table_header(line: list[dict]) -> bool:
    """True if this line is the CHANGED FIELD / FROM / TO column-header row."""
    norm = _normalise(_line_text(line))
    return "changed" in norm and "from" in norm and ("field" in norm or "to" in norm)


def _is_employee_header(line: list[dict]) -> bool:
    """True if this line is an employee block header.

    Matches any of the known labeled fields that appear on employee header rows.
    'Associate ID' is the primary signal but some PDFs omit it; the others
    (Position ID, Home Department, Home Cost Number) serve as fallbacks.
    """
    norm = _normalise(_line_text(line))
    return (
        "associate id" in norm
        or "position id" in norm
        or "home department" in norm
        or "home cost number" in norm
    )


def _is_page_chrome(line: list[dict]) -> bool:
    """True for page-level decorations to skip (page numbers, document title)."""
    words = [w["text"].lower() for w in line]
    norm = " ".join(words)
    # Page X of Y
    if re.match(r"page\s+\d+\s+of\s+\d+", norm):
        return True
    # Document title line
    if norm.strip() in {"employee payroll changes", "payroll changes"}:
        return True
    return False


def _is_name_line(line: list[dict], col2_x: float) -> bool:
    """True if line looks like a name-only employee header (Last, First pattern in col1).

    Some PDFs split the employee header across lines: the name sits on one line
    and 'Associate ID: / Position ID:' appear on the next.  This detects the
    name-only line so the accumulator can join them before parsing.
    """
    if _is_table_header(line) or _is_page_chrome(line) or _is_employee_header(line):
        return False
    # Only consider words that fall in the first column (left of col2_x)
    col1_text = " ".join(
        w["text"] for w in sorted(line, key=lambda w: w["x0"])
        if w["x0"] < col2_x
    )
    return bool(re.search(r"[A-Za-z][A-Za-z'\-\.]*\s*,\s*[A-Za-z]", col1_text))


# ---------------------------------------------------------------------------
# Employee header parsing
# ---------------------------------------------------------------------------

def parse_employee_header(line_text: str) -> dict:
    """Parse an employee header row into name + ID components.

    Input example:
        "7TG  Smith, Jonathan  Associate ID: 12345  Position ID: 67890
         Home Department: B04103-LM Bellingham Warehouse  Home Cost Number: 1"

    Returns dict with keys: Last Name, First Name, Middle Name,
    Associate ID, Position ID, Home Department, Home Cost Number.
    """
    result = {
        "Last Name": "", "First Name": "", "Middle Name": "",
        "Associate ID": "", "Position ID": "",
        "Home Department": "", "Home Cost Number": "",
    }

    text = _clean(line_text)

    # ── Extract labeled fields with regex ────────────────────────────────────
    assoc_m = re.search(r"Associate\s+ID\s*:\s*(\S+)", text, re.I)
    pos_m   = re.search(r"Position\s+ID\s*:\s*(\S+)", text, re.I)
    dept_m  = re.search(
        r"Home\s+Department\s*:\s*(.+?)(?=\s{2,}Home\s+Cost|\s*Home\s+Cost|$)",
        text, re.I)
    cost_m  = re.search(r"Home\s+Cost\s+Number\s*:\s*(\S*)", text, re.I)

    if assoc_m:
        result["Associate ID"]    = assoc_m.group(1).strip()
    if pos_m:
        result["Position ID"]     = pos_m.group(1).strip()
    if dept_m:
        result["Home Department"] = _clean(dept_m.group(1))
    if cost_m:
        result["Home Cost Number"] = cost_m.group(1).strip()

    # ── Extract name: text before the first labeled field ────────────────────
    # Split on whichever header label appears first (Associate ID, Position ID,
    # Home Department, or Home Cost Number).  This handles PDFs that omit
    # "Associate ID" entirely.
    _HEADER_LABELS = re.compile(
        r"Associate\s+ID|Position\s+ID|Home\s+Department|Home\s+Cost",
        re.I,
    )
    split_m = _HEADER_LABELS.search(text)
    before = text[:split_m.start()].strip() if split_m else text

    # Split on the first comma: "CODE  Last [Last2] [,] First [Middle]"
    # Spaces before/after the comma are allowed ("Aas sda , First").
    # Fall back to full line if 'before' has no comma.
    search_text = before if "," in before else text
    if "," in search_text:
        comma_pos = search_text.index(",")
        raw_last = search_text[:comma_pos].strip()
        raw_rest = search_text[comma_pos + 1:].strip()

        last_tokens = raw_last.split()
        # Drop leading code tokens: all-uppercase/digit, ≤8 chars (e.g. "7TG")
        while last_tokens and re.match(r"^[A-Z0-9]{1,8}$", last_tokens[0]):
            last_tokens.pop(0)

        if last_tokens:
            rest_tokens = raw_rest.split()
            if rest_tokens and rest_tokens[-1].strip(".").lower() in _NAME_SUFFIXES:
                rest_tokens.pop()  # discard suffix
            result["Last Name"]   = " ".join(last_tokens)
            result["First Name"]  = rest_tokens[0] if rest_tokens else ""
            result["Middle Name"] = " ".join(rest_tokens[1:])

    return result


# ---------------------------------------------------------------------------
# 3-column line splitter
# ---------------------------------------------------------------------------

def _split_three_cols(
    line: list[dict], col2_x: float, col3_x: float
) -> tuple[str, str, str]:
    """Assign each word to col1, col2, or col3 by x-coordinate.

    Returns (changed_field, changed_from, changed_to).
    """
    col1: list[str] = []
    col2: list[str] = []
    col3: list[str] = []

    for w in sorted(line, key=lambda w: w["x0"]):
        # Use the word's left edge to determine its column
        if w["x0"] < col2_x:
            col1.append(w["text"])
        elif w["x0"] < col3_x:
            col2.append(w["text"])
        else:
            col3.append(w["text"])

    return " ".join(col1).strip(), " ".join(col2).strip(), " ".join(col3).strip()


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------

def extract_changes(
    pdf_path: Path, debug: bool = False
) -> tuple[list[dict], str, int]:
    """Return (change_rows, warning, page_count).

    Scans the PDF for the 3-column table structure.  Employee header rows
    (containing 'Associate ID:') start a new employee block; subsequent
    non-header rows are change entries belonging to that employee.
    """
    change_rows: list[dict] = []
    page_count  = 0
    debug_count = 0

    # Carry column boundaries across pages; detect once from the first
    # table-header row found anywhere in the document.
    col2_x: float | None = None
    col3_x: float | None = None

    cur_emp: dict | None = None   # current employee header fields
    acc_emp_lines: list[str] = []  # lines being accumulated for the current header

    def _emit(field: str, frm: str, to: str) -> None:
        if cur_emp is None:
            return
        if not field:
            return
        row = {
            "Last Name":      cur_emp.get("Last Name", ""),
            "First Name":     cur_emp.get("First Name", ""),
            "Middle Name":    cur_emp.get("Middle Name", ""),
            "Associate ID":   cur_emp.get("Associate ID", ""),
            "Position ID":    cur_emp.get("Position ID", ""),
            "Home Department":   cur_emp.get("Home Department", ""),
            "Home Cost Number":  cur_emp.get("Home Cost Number", ""),
            "Changed Field":  field,
            "Changed From":   frm,
            "Changed To":     to,
        }
        change_rows.append(row)
        nonlocal debug_count
        if debug and debug_count < 6:
            print(f"  [{cur_emp.get('Last Name')} {cur_emp.get('First Name')}] "
                  f"{field!r} | {frm!r} -> {to!r}")
            debug_count += 1

    def _flush_emp_acc() -> None:
        """Parse accumulated header lines and set cur_emp."""
        nonlocal cur_emp, debug_count
        if acc_emp_lines:
            cur_emp = parse_employee_header("  ".join(acc_emp_lines))
            acc_emp_lines.clear()
            if debug and debug_count < 20:
                print(f"  EMP: {cur_emp.get('Last Name')}, "
                      f"{cur_emp.get('First Name')}  "
                      f"assocID={cur_emp.get('Associate ID')!r}")
                debug_count += 1

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_count = len(pdf.pages)

        for page_no, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            pw    = float(page.width)
            lines = _group_lines(words)

            # Detect column boundaries from this page's header row if not yet known
            if col2_x is None:
                col2_x, col3_x = _detect_boundaries(lines, pw)

            assert col3_x is not None

            for line in lines:
                if not line:
                    continue
                if _is_page_chrome(line):
                    continue
                if _is_table_header(line):
                    continue

                if _is_employee_header(line):
                    # Line has ID labels — always part of an employee header block
                    acc_emp_lines.append(_line_text(line))
                    continue

                if _is_name_line(line, col2_x):
                    # Name-only line (Last, First) — start of a (possibly multi-line) header.
                    # If we were already accumulating, flush the previous employee first.
                    if acc_emp_lines:
                        _flush_emp_acc()
                    acc_emp_lines.append(_line_text(line))
                    continue

                # Regular change row — flush any pending header, then split into 3 columns
                _flush_emp_acc()
                field, frm, to_ = _split_three_cols(line, col2_x, col3_x)
                _emit(field, frm, to_)

            _flush_emp_acc()  # flush header at end of each page

    warning = "" if change_rows else _NO_RECORDS
    return change_rows, warning, page_count


# ---------------------------------------------------------------------------
# Diagnostics
# ---------------------------------------------------------------------------

def _diagnose(pdf_path: Path, max_pages: int = 3) -> None:
    """Dump raw PDF structure to stdout for debugging.

    Shows: detected column boundaries, every line's raw text, and which
    lines are classified as employee headers or change rows.
    """
    print(f"\n=== DIAGNOSE: {pdf_path.name} ===\n")
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            if page_no > max_pages:
                print(f"  (capped at {max_pages} pages — use --diagnose-pages N)")
                break
            words = page.extract_words()
            pw    = float(page.width)
            lines = _group_lines(words)
            c2, c3 = _detect_boundaries(lines, pw)
            print(f"--- PAGE {page_no}  col2_x={c2:.1f}  col3_x={c3:.1f} "
                  f"(page_width={pw:.1f}) ---")
            for li, line in enumerate(lines, start=1):
                raw = _line_text(line)
                tag = ""
                if _is_page_chrome(line):
                    tag = "  [CHROME]"
                elif _is_table_header(line):
                    tag = "  [TABLE HEADER]"
                elif _is_employee_header(line):
                    emp = parse_employee_header(raw)
                    tag = (f"  [EMP-ID: {emp['Last Name']}, {emp['First Name']}"
                           f"  assocID={emp['Associate ID']!r}"
                           f"  posID={emp['Position ID']!r}]")
                elif _is_name_line(line, c2):
                    tag = "  [EMP-NAME]"
                else:
                    field, frm, to_ = _split_three_cols(line, c2, c3)
                    if field:
                        tag = f"  [CHG: {field!r} | {frm!r} -> {to_!r}]"
                print(f"  L{li:03d}: {raw}{tag}")
            print()


# ---------------------------------------------------------------------------
# Output helpers
# ---------------------------------------------------------------------------

def _csv_safe(value):
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        try:
            float(value.replace(",", "").strip())
        except ValueError:
            return "'" + value
    return value


def _format_name(record: dict) -> str:
    last   = record.get("Last Name", "").strip()
    first  = record.get("First Name", "").strip()
    middle = record.get("Middle Name", "").strip()
    parts  = " ".join(filter(None, [first, middle]))
    return f"{last}, {parts}" if last and parts else last or parts


def _build_row(record: dict, docid: str) -> list:
    row = []
    for col in OUTPUT_HEADERS:
        if col == "DOCID":
            row.append(docid)
        elif col == "Name":
            row.append(_format_name(record))
        else:
            row.append(record.get(col, ""))
    return [_csv_safe(v) for v in row]


def _atomic_replace(tmp_path: Path, out_path: Path) -> bool:
    for attempt in range(5):
        try:
            os.replace(tmp_path, out_path)
            return True
        except PermissionError:
            if attempt < 4:
                time.sleep(1)
            else:
                print(f"  WARNING: could not write {out_path.name} after 5 attempts "
                      "(file locked) — skipping")
                try:
                    tmp_path.unlink(missing_ok=True)
                except OSError:
                    pass
                return False
    return False


def write_workbook(
    out_path: Path, change_rows: list[dict], docid: str, warning: str
) -> None:
    """Write extracted rows to {stem}_Changes.xlsx."""
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = "Payroll_Changes"
        ws.append(OUTPUT_HEADERS)
        for record in change_rows:
            ws.append(_build_row(record, docid))
        if warning:
            ws_w = wb.create_sheet("Warnings")
            ws_w.append(["Warning"])
            ws_w.cell(row=1, column=1).font = Font(bold=True)
            ws_w.append([warning])
            ws_w.column_dimensions["A"].width = 120
        wb.save(tmp_path)
        wb.close()
        _atomic_replace(tmp_path, out_path)
    except Exception:
        wb.close()
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


REPORT_HEADERS = ["File Name", "Change Rows", "Pages", "Records Found", "Warnings"]


def write_processing_report(report_path: Path, rows: list[dict]) -> None:
    tmp_path = report_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    try:
        ws = wb.active
        ws.title = "Processing_Report"
        ws.append(REPORT_HEADERS)
        for r in rows:
            ws.append([_csv_safe(r["file"]), r["records"], r["pages"],
                       r["records_found"], r["warning"]])
        if any(r["warning"] for r in rows):
            ws.append([])
            ws.append(["--- Warnings Detail ---"])
            for r in rows:
                if r["warning"]:
                    ws.append([_csv_safe(r["file"]), r["warning"]])
        wb.save(tmp_path)
        wb.close()
        _atomic_replace(tmp_path, report_path)
    except Exception:
        wb.close()
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


# ---------------------------------------------------------------------------
# Template-output helpers (fuzzy match + pivot)
# ---------------------------------------------------------------------------

def _load_bucket11_headers(path: Path) -> list[str]:
    """Load the Bucket-11 responsive header list from an Excel file."""
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active
    headers = [
        str(row[0]).strip()
        for i, row in enumerate(ws.iter_rows(values_only=True), 1)
        if i > 1 and row[0]
    ]
    wb.close()
    return headers


def _load_template(path: Path) -> tuple[list[str], list]:
    """Return (headers, tags_row) from rows 1 and 2 of the template.

    Row 1 = column headers.
    Row 2 = pre-filled dropdown tag values (semicolon-separated where multiple).
    """
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active
    rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
    headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    tags    = list(rows[1]) if len(rows) > 1 else []
    # Pad tags list to the same width as headers
    while len(tags) < len(headers):
        tags.append(None)
    wb.close()
    return headers, tags


def _norm_col(name: str) -> str:
    """Normalise a column name for fuzzy lookup (collapse whitespace/dashes)."""
    return re.sub(r"[\s–—\-]+", " ", name or "").lower().strip()


def _build_col_lookup(headers: list[str]) -> dict[str, str]:
    """Return {normalised_name: actual_name} for template column resolution."""
    return {_norm_col(h): h for h in headers if h}


def _fuzzy_match_header(field: str, candidates: list[str], cutoff: float = 0.65) -> str | None:
    """Return the closest Bucket-11 header for a Changed Field string, or None.

    Checks the _ALIASES table first (exact normalised match), then falls back
    to difflib fuzzy matching against the full candidate list.
    """
    field_norm = _norm_col(field)
    # 1. Alias table — handles abbreviations and common variants
    alias_hit = _ALIASES.get(field_norm)
    if alias_hit:
        return alias_hit
    # 2. Fuzzy match against the Bucket-11 header list
    cand_norms = [_norm_col(c) for c in candidates]
    matches = difflib.get_close_matches(field_norm, cand_norms, n=1, cutoff=cutoff)
    if not matches:
        return None
    return candidates[cand_norms.index(matches[0])]


_BLANK_VALUES = {"", "no value", "n/a", "none", "null", "-", "–", "—"}


def _select_value(changed_from: str, changed_to: str) -> str:
    """Use Changed To unless it is blank/no-value, then fall back to Changed From."""
    to_clean = changed_to.strip()
    return changed_from.strip() if to_clean.lower() in _BLANK_VALUES else to_clean


def _pivot_to_template_rows(
    change_rows: list[dict],
    docid: str,
    template_headers: list[str],
    bucket11_headers: list[str],
) -> tuple[list[dict], int, int]:
    """Convert per-change rows to one row per employee in template column layout.

    Changed Field values are fuzzy-matched against the Bucket-11 header list,
    then mapped to the appropriate template column via _BUCKET11_MAP.
    Category columns (e.g. Contact Information) accumulate unique sub-items
    separated by ";".

    Returns (rows, matched_count, unmatched_count).
    """
    col_lookup = _build_col_lookup(template_headers)

    def _find_col(name: str | None) -> str | None:
        if not name:
            return None
        return col_lookup.get(_norm_col(name))

    # emp_key → template row dict
    emp_rows: dict[tuple, dict] = {}
    # emp_key → {cat_col: [sub_items...]}
    cat_acc: dict[tuple, dict[str, list[str]]] = {}

    matched   = 0
    unmatched = 0

    for rec in change_rows:
        emp_key = (rec.get("Associate ID", ""), rec.get("Last Name", ""), rec.get("First Name", ""))

        if emp_key not in emp_rows:
            row: dict = {h: "" for h in template_headers}
            row[_find_col("DOCID")        or "DOCID"]       = rec.get("DOCID", docid)
            row[_find_col("Last Name")    or "Last Name"]   = rec.get("Last Name", "")
            row[_find_col("First Name")   or "First Name"]  = rec.get("First Name", "")
            row[_find_col("Middle Name")  or "Middle Name"] = rec.get("Middle Name", "")
            row[_find_col("Data Subject Type") or "Data Subject Type"] = "Employee"
            eic = _find_col("Employee Identification Number")
            if eic:
                row[eic] = rec.get("Associate ID", "")
            emp_rows[emp_key] = row
            cat_acc[emp_key]  = {}

        row  = emp_rows[emp_key]
        cats = cat_acc[emp_key]

        field = rec.get("Changed Field", "")
        value = _select_value(rec.get("Changed From", ""), rec.get("Changed To", ""))

        b11_hdr = _fuzzy_match_header(field, bucket11_headers)
        if b11_hdr is None:
            unmatched += 1
            continue
        mapping = _BUCKET11_MAP.get(b11_hdr)
        if mapping is None:
            unmatched += 1
            continue

        matched += 1
        value_col_name, cat_col_name, cat_subitem = mapping

        # ── Write to direct value column ────────────────────────────────────
        vcol = _find_col(value_col_name)
        if vcol and value:
            existing = row.get(vcol, "")
            if not existing:
                row[vcol] = value
            elif value not in existing:
                row[vcol] = existing + "; " + value

        # ── Accumulate category sub-item ────────────────────────────────────
        ccol = _find_col(cat_col_name)
        if ccol and cat_subitem:
            items = cats.setdefault(ccol, [])
            if cat_subitem not in items:
                items.append(cat_subitem)

    # Merge accumulated category sub-items into each row
    for emp_key, row in emp_rows.items():
        for ccol, items in cat_acc[emp_key].items():
            if items:
                existing = row.get(ccol, "")
                for item in items:
                    if item not in existing:
                        existing = (existing + ";" + item) if existing else item
                row[ccol] = existing

    return list(emp_rows.values()), matched, unmatched


def write_per_pdf_output(
    out_path: Path,
    change_rows: list[dict],
    warning: str,
    template_headers: list[str] | None = None,
    template_tags: list | None = None,
    bucket11_headers: list[str] | None = None,
) -> dict:
    """Write one workbook per PDF with two sheets.

    Sheet 1 — "Native extracted": every raw change row from this PDF.
    Sheet 2 — "Import Template":  one row per employee mapped to the template
               format.  Row 1 = headers, Row 2 = pre-filled dropdown tags
               (copied from Latest Template.xlsx row 2), Row 3+ = data.
               Only written when template_headers and bucket11_headers are supplied.

    Returns a stats dict: {template_employees, matched_fields, unmatched_fields}.
    """
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    stats = {"template_employees": 0, "matched_fields": 0, "unmatched_fields": 0}
    try:
        # ── Sheet 1: Native extracted ─────────────────────────────────────────
        ws1 = wb.active
        ws1.title = "Native extracted"
        ws1.append(OUTPUT_HEADERS)
        _name_idx = OUTPUT_HEADERS.index("Name")
        _last_name = ""
        for rec in change_rows:
            row = _build_row(rec, rec.get("DOCID", ""))
            if not row[_name_idx]:
                row[_name_idx] = _last_name
            else:
                _last_name = row[_name_idx]
            ws1.append(row)

        # ── Sheet 2: Import Template ──────────────────────────────────────────
        if template_headers and bucket11_headers:
            tmpl_rows, matched, unmatched = _pivot_to_template_rows(
                change_rows, "", template_headers, bucket11_headers
            )
            stats["template_employees"] = len(tmpl_rows)
            stats["matched_fields"]     = matched
            stats["unmatched_fields"]   = unmatched
            ws2 = wb.create_sheet("Import Template")
            ws2.append(template_headers)
            if template_tags:
                ws2.append(template_tags)
            for row_dict in tmpl_rows:
                ws2.append([_csv_safe(row_dict.get(h, "")) for h in template_headers])

        # ── Warnings sheet ────────────────────────────────────────────────────
        if warning:
            ws_w = wb.create_sheet("Warnings")
            ws_w.append(["Warning"])
            ws_w.cell(row=1, column=1).font = Font(bold=True)
            ws_w.append([warning])
            ws_w.column_dimensions["A"].width = 120

        wb.save(tmp_path)
        wb.close()
        _atomic_replace(tmp_path, out_path)
    except Exception:
        wb.close()
        try:
            tmp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise
    return stats


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------

def _selftest() -> int:
    ok = True

    # ── parse_name ────────────────────────────────────────────────────────────
    name_cases = [
        ("Smith, Jonathan",     ("Smith",   "Jonathan", "",   "")),
        ("Torres, Efrain",      ("Torres",  "Efrain",   "",   "")),
        ("Vaughan, Antione D.", ("Vaughan", "Antione",  "D.", "")),
        ("Valencia, Patrick G.","Valencia", "Patrick",  "G.", ""),
        ("Strunk, William",     ("Strunk",  "William",  "",   "")),
        ("",                    ("",        "",         "",   "")),
    ]
    for args in name_cases:
        raw  = args[0]
        exp  = args[1] if isinstance(args[1], tuple) else args[1:]
        got  = parse_name(raw)
        tup  = (got["Last Name"], got["First Name"],
                got["Middle Name"], got["Suffix"])
        status = "OK " if tup == exp else "FAIL"
        if tup != exp:
            ok = False
        print(f"  [{status}] {raw!r:30} -> {tup}")

    # ── parse_employee_header ─────────────────────────────────────────────────
    hdr = ("7TG  Smith, Jonathan  Associate ID: A12345  Position ID: P67890  "
           "Home Department: B04103-LM Bellingham Warehouse  Home Cost Number: 99")
    emp = parse_employee_header(hdr)
    emp_ok = (emp["Last Name"] == "Smith"
              and emp["First Name"] == "Jonathan"
              and emp["Associate ID"] == "A12345"
              and emp["Position ID"] == "P67890"
              and "Bellingham" in emp["Home Department"])
    print(f"  [{'OK ' if emp_ok else 'FAIL'}] parse_employee_header -> "
          f"{emp['Last Name']}, {emp['First Name']}  "
          f"assoc={emp['Associate ID']}  pos={emp['Position ID']}  "
          f"dept={emp['Home Department']!r}")
    if not emp_ok:
        ok = False

    # ── _split_three_cols ─────────────────────────────────────────────────────
    def _w(text, x0, x1):
        return {"text": text, "x0": x0, "x1": x1, "top": 0.0, "bottom": 10.0}

    col_line = [
        _w("Deduction", 10, 80), _w("Code", 83, 120), _w("(W)", 123, 145),
        _w("Active", 240, 290),
        _w("Inactive", 450, 510),
    ]
    f, frm, to_ = _split_three_cols(col_line, col2_x=230.0, col3_x=440.0)
    split_ok = (f == "Deduction Code (W)"
                and frm == "Active"
                and to_ == "Inactive")
    print(f"  [{'OK ' if split_ok else 'FAIL'}] _split_three_cols -> "
          f"field={f!r}  from={frm!r}  to={to_!r}")
    if not split_ok:
        ok = False

    # ── _is_employee_header ───────────────────────────────────────────────────
    hdr_line = [_w("7TG", 10, 30), _w("Smith,", 33, 70), _w("Jonathan", 73, 120),
                _w("Associate", 130, 185), _w("ID:", 188, 205), _w("X999", 208, 240)]
    non_hdr  = [_w("Deduction", 10, 80), _w("Code", 83, 120)]
    hdr_ok   = _is_employee_header(hdr_line) and not _is_employee_header(non_hdr)
    print(f"  [{'OK ' if hdr_ok else 'FAIL'}] _is_employee_header")
    if not hdr_ok:
        ok = False

    # ── _detect_boundaries ───────────────────────────────────────────────────
    tbl_header_line = [
        _w("CHANGED", 10, 80), _w("FIELD", 83, 120),
        _w("CHANGED", 240, 300), _w("FROM", 303, 340),
        _w("CHANGED", 450, 510), _w("TO", 513, 530),
    ]
    c2, c3 = _detect_boundaries([tbl_header_line], page_width=612.0)
    bounds_ok = (230 < c2 < 260 and 440 < c3 < 460)
    print(f"  [{'OK ' if bounds_ok else 'FAIL'}] _detect_boundaries -> "
          f"col2_x={c2:.1f}  col3_x={c3:.1f}  (expected ~240, ~450)")
    if not bounds_ok:
        ok = False

    print("Self-test:", "PASSED" if ok else "FAILED")
    return 0 if ok else 1


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

def process_file(
    pdf_path: Path,
    debug: bool,
    log=print,
) -> tuple[dict, list[dict]]:
    """Extract changes from one PDF.

    Returns (report_dict, change_rows).  change_rows have a 'DOCID' key added.
    No output files are written here; the caller accumulates rows and writes
    the combined workbook after all PDFs are processed.
    """
    log(f"Processing: {pdf_path.name}")
    change_rows, warning, pages = extract_changes(pdf_path, debug=debug)
    if warning:
        log(f"  WARNING: {warning}")
    log(f"  Extracted {len(change_rows)} change row(s).")
    # Tag every row with its source file so the combined sheet shows DOCID correctly
    for rec in change_rows:
        rec["DOCID"] = pdf_path.name
    return (
        {
            "file": pdf_path.name, "records": len(change_rows),
            "pages": pages, "records_found": "Yes" if change_rows else "No",
            "warning": warning,
        },
        change_rows,
    )


def _iter_pdfs(target: Path, recursive: bool):
    if target.is_file():
        yield target
        return
    pattern = "**/*.pdf" if recursive else "*.pdf"
    yield from sorted(target.glob(pattern))


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract Employee Payroll Changes from ADP native-text PDFs."
    )
    parser.add_argument("target", nargs="?",
                        help="PDF file or folder of PDFs")
    parser.add_argument("--recursive", action="store_true",
                        help="Recurse into subfolders when target is a folder")
    parser.add_argument("--debug", action="store_true",
                        help="Print sample detections; write no output files")
    parser.add_argument("--diagnose", action="store_true",
                        help="Dump raw PDF text and column detection to stdout")
    parser.add_argument("--diagnose-pages", type=int, default=3, metavar="N",
                        help="Pages to show in --diagnose mode (default 3)")
    parser.add_argument("--selftest", action="store_true",
                        help="Run built-in parser tests and exit")
    parser.add_argument("--template-file", metavar="XLSX",
                        help="Latest Template xlsx; when provided, also write "
                             "{stem}_Template_Output.xlsx per PDF")
    parser.add_argument("--bucket-list", metavar="XLSX",
                        help="Bucket-11 header list xlsx "
                             "(default: Bucket_11_List.xlsx beside this script)")
    args = parser.parse_args()

    if args.selftest:
        return _selftest()

    if not args.target:
        parser.error("a PDF file or folder is required (or use --selftest)")
    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required.  pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required.  pip install openpyxl")

    target = Path(args.target)
    if not target.exists():
        sys.exit(f"ERROR: path not found — {target}")

    pdfs = list(_iter_pdfs(target, args.recursive))
    if not pdfs:
        sys.exit(f"ERROR: no PDF files found at {target}")

    if args.diagnose:
        for pdf_path in pdfs[:5]:
            _diagnose(pdf_path, max_pages=args.diagnose_pages)
        return 0

    # ── Load template + bucket-list (auto-detect from script folder if not given) ──
    _script_dir = Path(__file__).parent
    tmpl_path = (
        Path(args.template_file) if args.template_file
        else _script_dir / "Latest Template.xlsx"
    )
    bl_path = (
        Path(args.bucket_list) if args.bucket_list
        else _script_dir / "Bucket_11_List.xlsx"
    )

    template_headers: list | None = None
    template_tags:    list | None = None
    bucket11_headers: list | None = None

    if tmpl_path.exists() and bl_path.exists():
        template_headers, template_tags = _load_template(tmpl_path)
        bucket11_headers = _load_bucket11_headers(bl_path)
        print(f"Template mode: {len(template_headers)} template cols, "
              f"{len(bucket11_headers)} Bucket-11 headers loaded.")
    else:
        if not tmpl_path.exists():
            print(f"WARNING: template file not found — {tmpl_path} (Import Template sheet skipped)")
        if not bl_path.exists():
            print(f"WARNING: bucket-list not found — {bl_path} (Import Template sheet skipped)")

    pbar = (
        _tqdm(total=len(pdfs), unit="file", desc="Processing",
              ncols=80, colour="cyan")
        if HAS_TQDM and not args.debug else None
    )

    def _log(msg: str) -> None:
        if pbar:
            pbar.write(msg)
        else:
            print(msg)

    report_rows:  list[dict] = []
    all_warnings: list[str]  = []
    output_folder = target if target.is_dir() else target.parent

    for pdf_path in pdfs:
        try:
            rpt, rows = process_file(pdf_path, args.debug, _log)
            report_rows.append(rpt)

            if not args.debug:
                out_path = output_folder / f"{pdf_path.stem}_extracted.xlsx"
                stats = write_per_pdf_output(
                    out_path,
                    rows,
                    rpt["warning"],
                    template_headers=template_headers,
                    template_tags=template_tags,
                    bucket11_headers=bucket11_headers,
                )
                sheets = "Native extracted"
                if template_headers and bucket11_headers:
                    emp_n   = stats["template_employees"]
                    matched = stats["matched_fields"]
                    unmatched = stats["unmatched_fields"]
                    sheets += f" + Import Template ({emp_n} employee(s), {matched} field(s) mapped"
                    if unmatched:
                        sheets += f", {unmatched} unmatched"
                    sheets += ")"
                elif not template_headers:
                    sheets += "  [Import Template skipped — template file not loaded]"
                _log(f"  Wrote: {out_path.name}  [{sheets}]  ({len(rows)} change row(s))")

            if rpt["warning"]:
                all_warnings.append(f"{pdf_path.name}: {rpt['warning']}")
        except Exception as exc:
            _log(f"  ERROR processing {pdf_path.name}: {exc}")
            report_rows.append({
                "file": pdf_path.name, "records": 0, "pages": 0,
                "records_found": "No",
                "warning": f"Processing failed: {exc}",
            })
            all_warnings.append(f"{pdf_path.name}: Processing failed: {exc}")
        if pbar:
            pbar.update(1)

    if pbar:
        pbar.close()

    if not args.debug and report_rows:
        # ── Processing report ─────────────────────────────────────────────────
        report_path = output_folder / "Processing_Report_Changes.xlsx"
        write_processing_report(report_path, report_rows)
        print(f"Wrote: {report_path.name} ({len(report_rows)} file(s))")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
