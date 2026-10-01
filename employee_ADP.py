r"""
employee_ADP.py — Extract Employee Payroll Changes from ADP "Employee Payroll Changes" PDFs.

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
    python employee_ADP.py "C:\path\to\file.pdf"
    python employee_ADP.py "C:\path\to\folder" --recursive
    python employee_ADP.py "C:\path\to\file.pdf" --debug
    python employee_ADP.py "C:\path\to\file.pdf" --diagnose [--diagnose-pages N]
    python employee_ADP.py --selftest
"""

from __future__ import annotations

import argparse
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


def write_per_pdf_output(
    out_path: Path,
    change_rows: list[dict],
    warning: str,
) -> None:
    """Write one workbook per PDF.

    Sheet 1 — "Native extracted": every raw change row from this PDF.
    Sheet 2 — "Warnings" (only when there is a warning).
    """
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
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
                write_per_pdf_output(out_path, rows, rpt["warning"])
                _log(f"  Wrote: {out_path.name}  ({len(rows)} change row(s))")

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
