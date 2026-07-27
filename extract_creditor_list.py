r"""
extract_creditor_list.py — Extract creditor names and addresses from
bankruptcy mailing-matrix PDFs into Excel.

Each page is laid out as a three-column grid of creditor entries.  The script
detects column boundaries from word x-coordinates, groups words into lines,
then splits lines into name + address blocks using vertical gap detection
(no text heuristics — entry boundaries are the whitespace gaps between blocks).

Output — Excel workbook with three sheets:

  "extracted"              — Full Name | Address Information | Page Number
  "parsed names addresses" — First Name | Last Name | Middle Name | Suffix |
                             Street Address | City | State | ZIP Code | Page Number
  "standard data"          — Same columns as "parsed names addresses" but
                             filtered to individual people only; rows where
                             First Name is blank or is inc / inc. / llc / llc.
                             (companies) are excluded.

USAGE
    python extract_creditor_list.py "C:\path\to\file.pdf"
    python extract_creditor_list.py "C:\path\to\folder"
    python extract_creditor_list.py "C:\path\to\folder" --recursive
    python extract_creditor_list.py                          # prompts for path

DEPENDENCIES
    pip install pdfplumber openpyxl
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


_LINE_TOL   = 3.0   # pt — vertical tolerance for grouping words into one line
_GAP_FACTOR = 0.6   # gap > avg_line_height * _GAP_FACTOR  →  new entry

# First-name values that indicate a company entry, not a person
_COMPANY_FIRST_NAMES = {"inc", "inc.", "llc", "llc."}

_NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v", "esq", "md", "phd"}

_ENTITY_TOKENS = {"inc", "llc", "corp", "ltd", "co", "ag", "lp", "llp",
                  "plc", "pc", "pa", "na", "nv", "sa", "gmbh"}


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _clean(text: str) -> str:
    return " ".join(str(text or "").split()).strip()


def _csv_safe(v) -> str:
    s = str(v or "")
    return ("'" + s) if s[:1] in ("=", "+", "-", "@") else s


def _is_person(first_name: str) -> bool:
    """Return True when the First Name field represents an individual, not a company."""
    fn = first_name.strip().lower()
    return bool(fn) and fn not in _COMPANY_FIRST_NAMES


# ---------------------------------------------------------------------------
# Name parser
# ---------------------------------------------------------------------------

def parse_name(cell: str) -> dict:
    """Split "Last, First Middle [Suffix]" or a company name into components.

    Person   — "SMITH, JOHN A JR"  → Last / First / Middle / Suffix fields.
    Company  — "ART RESOURCE INC"  → Last Name = full name; others blank.
    """
    blank = {"Last Name": "", "First Name": "", "Middle Name": "", "Suffix": ""}
    s = _clean(cell)
    if not s:
        return blank

    tokens = s.split()

    if "," not in s:
        if tokens[-1].strip(".,").lower() in _ENTITY_TOKENS:
            return {**blank, "Last Name": s}
        return {**blank, "Last Name": s}

    last_part, _, rest = s.partition(",")
    last_toks = last_part.split()
    rest_toks = rest.split()

    # Comma present but last_part ends with an entity token → company
    if last_toks and last_toks[-1].strip(".,").lower() in _ENTITY_TOKENS:
        return {**blank, "Last Name": s.rstrip(",")}

    suffix = ""
    if rest_toks and rest_toks[-1].strip(".").lower() in _NAME_SUFFIXES:
        suffix    = rest_toks[-1]
        rest_toks = rest_toks[:-1]
    elif last_toks and last_toks[-1].strip(".").lower() in _NAME_SUFFIXES:
        suffix    = last_toks[-1]
        last_toks = last_toks[:-1]

    return {
        "Last Name":   " ".join(last_toks),
        "First Name":  rest_toks[0] if rest_toks else "",
        "Middle Name": " ".join(rest_toks[1:]) if len(rest_toks) > 1 else "",
        "Suffix":      suffix,
    }


# Country names treated as a trailing line to skip when parsing city/state/zip
_COUNTRY_NAMES = {
    "canada", "mexico", "australia", "united kingdom", "uk", "england",
    "germany", "france", "italy", "spain", "japan", "china", "india",
}

# US ZIP:      5 digits, optional +4          e.g. 06070, 33004-2243
# Canada post: letter-digit-letter space digit-letter-digit  e.g. V7W 1X3
_RE_US_ZIP = re.compile(r'\b([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$')
_RE_CA_POST = re.compile(r'\b([A-Z]{2})\s+([A-Z]\d[A-Z]\s*\d[A-Z]\d)\s*$')


# ---------------------------------------------------------------------------
# Address parser
# ---------------------------------------------------------------------------

def parse_address(lines: list[str]) -> dict:
    """Parse 1–3 address lines into street, city, state, zipcode.

    Handles both US addresses ("CITY, ST 12345") and Canadian addresses
    ("CITY, ON L1P1L5") where "CANADA" may appear as a trailing country line.
    The country line is stripped before parsing so it does not interfere.
    """
    result = {"street": "", "city": "", "state": "", "zipcode": ""}
    clean_lines = [_clean(ln) for ln in lines if _clean(ln)]
    if not clean_lines:
        return result

    # Drop a trailing country name line (e.g. "CANADA")
    if clean_lines[-1].lower() in _COUNTRY_NAMES:
        clean_lines = clean_lines[:-1]
    if not clean_lines:
        return result

    last = clean_lines[-1]

    # Try US ZIP first, then Canadian postal code
    m = _RE_US_ZIP.search(last) or _RE_CA_POST.search(last)
    if m:
        result["state"]   = m.group(1)
        result["zipcode"] = m.group(2).replace(" ", "")   # normalise CA postal spacing
        result["city"]    = last[:m.start()].rstrip(", ").strip()
        result["street"]  = ", ".join(clean_lines[:-1])
    else:
        result["street"] = ", ".join(clean_lines)

    return result


# ---------------------------------------------------------------------------
# PDF word-level layout helpers
# ---------------------------------------------------------------------------

def _group_lines(words: list[dict]) -> list[list[dict]]:
    """Cluster pdfplumber words into visual lines by their 'top' coordinate."""
    if not words:
        return []
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines: list[list[dict]] = [[ordered[0]]]
    cur_top = ordered[0]["top"]
    for w in ordered[1:]:
        if abs(w["top"] - cur_top) <= _LINE_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])
            cur_top = w["top"]
    for ln in lines:
        ln.sort(key=lambda w: w["x0"])
    return lines


def _remove_page_headers(lines: list[list[dict]], page_width: float) -> list[list[dict]]:
    """Drop lines whose x-span exceeds 55% of the page width.

    Court document page headers (case number, filing date, page counter) run
    across the full page.  Creditor entries are confined to roughly one column
    (~33% of the page width) so they are never removed by this filter.
    """
    threshold = page_width * 0.55
    return [
        ln for ln in lines
        if ln and (max(w["x1"] for w in ln) - min(w["x0"] for w in ln)) <= threshold
    ]


def _assign_columns(
    lines: list[list[dict]], n_cols: int = 3
) -> list[list[list[dict]]]:
    """Assign each word to one of n_cols columns by x-centre.

    Column boundaries are derived from the overall page x-span divided
    evenly into thirds.  Returns n_cols ordered lists of lines.
    """
    all_x0 = [w["x0"] for ln in lines for w in ln]
    all_x1 = [w["x1"] for ln in lines for w in ln]
    if not all_x0:
        return [[] for _ in range(n_cols)]

    page_left  = min(all_x0)
    page_right = max(all_x1)
    col_width  = (page_right - page_left) / n_cols

    buckets: list[dict[float, list[dict]]] = [defaultdict(list) for _ in range(n_cols)]
    for ln in lines:
        for w in ln:
            cx      = (w["x0"] + w["x1"]) / 2
            col_idx = min(int((cx - page_left) / col_width), n_cols - 1)
            buckets[col_idx][w["top"]].append(w)

    result = []
    for ci in range(n_cols):
        col_lines = []
        for top_y in sorted(buckets[ci]):
            col_lines.append(sorted(buckets[ci][top_y], key=lambda w: w["x0"]))
        result.append(col_lines)
    return result


# ---------------------------------------------------------------------------
# Entry detection  — gap-based, no text heuristics
# ---------------------------------------------------------------------------

def _parse_entries_by_gap(col_lines: list[list[dict]], page_num: int) -> list[dict]:
    """Group column lines into creditor entries using vertical gap detection.

    Lines within one entry are separated by small gaps (normal line leading).
    Entries are separated by a larger whitespace gap — the same visual space
    visible between each creditor block.

    Threshold: gap between consecutive line bottom→top > avg_line_height * 0.6
    → new entry starts.  The first line of every group is the name;
    subsequent lines are the address.  page_num is stamped on every entry.
    """
    if not col_lines:
        return []

    tops    = [min(w["top"]    for w in ln) for ln in col_lines]
    bottoms = [max(w["bottom"] for w in ln) for ln in col_lines]
    heights = [b - t for t, b in zip(tops, bottoms)]
    avg_height = sum(heights) / len(heights) if heights else 12.0
    threshold  = avg_height * _GAP_FACTOR

    # Split lines into entry groups at large vertical gaps
    groups: list[list[list[dict]]] = [[col_lines[0]]]
    for i in range(len(col_lines) - 1):
        gap = tops[i + 1] - bottoms[i]
        if gap > threshold:
            groups.append([col_lines[i + 1]])
        else:
            groups[-1].append(col_lines[i + 1])

    entries: list[dict] = []
    for group in groups:
        text_lines = [_clean(" ".join(w["text"] for w in ln)) for ln in group]
        text_lines = [l for l in text_lines if l]
        if not text_lines:
            continue
        name       = text_lines[0]
        addr_lines = text_lines[1:]
        addr       = parse_address(addr_lines)
        full_addr  = ", ".join(l for l in addr_lines if l)
        entries.append({
            "full_name":  name,
            "address":    full_addr,
            "page_num":   page_num,
            **addr,
        })

    return entries


# ---------------------------------------------------------------------------
# Per-PDF extraction
# ---------------------------------------------------------------------------

def extract_pdf(pdf_path: Path) -> list[dict]:
    """Return all creditor entries from every page of a PDF.

    Each entry dict contains:
      full_name, address, street, city, state, zipcode, page_num
    """
    all_entries: list[dict] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if not words:
                continue
            lines   = _group_lines(words)
            lines   = _remove_page_headers(lines, float(page.width))
            columns = _assign_columns(lines, n_cols=3)
            for col_lines in columns:
                all_entries.extend(_parse_entries_by_gap(col_lines, page_num))
    return all_entries


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

_HDR_FONT  = None
_HDR_FILL  = None
_BODY_FONT = None


def _init_styles() -> None:
    global _HDR_FONT, _HDR_FILL, _BODY_FONT
    _HDR_FONT  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    _HDR_FILL  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    _BODY_FONT = Font(name="Calibri", size=11)


def _sh(cell) -> None:
    """Apply header style to a cell."""
    cell.font      = _HDR_FONT
    cell.fill      = _HDR_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _sb(cell) -> None:
    """Apply body style to a cell."""
    cell.font      = _BODY_FONT
    cell.alignment = Alignment(vertical="center")


def _autofit(ws, widths: list[int]) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_sheet1(wb, entries: list[dict]) -> None:
    """Sheet 1 — extracted: Full Name | Address Information | Page Number."""
    ws = wb.active
    ws.title = "extracted"
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(["Full Name", "Address Information", "Page Number"], 1):
        _sh(ws.cell(row=1, column=col, value=h))

    for r, e in enumerate(entries, 2):
        _sb(ws.cell(row=r, column=1, value=_csv_safe(e["full_name"])))
        _sb(ws.cell(row=r, column=2, value=_csv_safe(e["address"])))
        _sb(ws.cell(row=r, column=3, value=e["page_num"]))

    _autofit(ws, [42, 62, 14])


_PARSED_HEADERS = [
    "First Name", "Last Name", "Middle Name", "Suffix",
    "Street Address", "City", "State", "ZIP Code", "Page Number",
]


def _parsed_row(e: dict) -> list:
    nm = parse_name(e["full_name"])
    return [
        _csv_safe(nm["First Name"]),
        _csv_safe(nm["Last Name"]),
        _csv_safe(nm["Middle Name"]),
        _csv_safe(nm["Suffix"]),
        _csv_safe(e["street"]),
        _csv_safe(e["city"]),
        _csv_safe(e["state"]),
        _csv_safe(e["zipcode"]),
        e["page_num"],
    ]


def _write_sheet2(wb, entries: list[dict]) -> None:
    """Sheet 2 — parsed names addresses: all entries with parsed name fields."""
    ws = wb.create_sheet("parsed names addresses")
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(_PARSED_HEADERS, 1):
        _sh(ws.cell(row=1, column=col, value=h))

    for r, e in enumerate(entries, 2):
        for col, v in enumerate(_parsed_row(e), 1):
            _sb(ws.cell(row=r, column=col, value=v))

    _autofit(ws, [22, 35, 16, 10, 45, 22, 10, 14, 14])


def _write_sheet3(wb, entries: list[dict]) -> None:
    """Sheet 3 — standard data: individuals only (companies filtered out).

    Excluded when First Name is blank or is one of: inc, inc., llc, llc.
    """
    ws = wb.create_sheet("standard data")
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(_PARSED_HEADERS, 1):
        _sh(ws.cell(row=1, column=col, value=h))

    r = 2
    for e in entries:
        row = _parsed_row(e)
        first_name = row[0]          # column index 0 = First Name
        if not _is_person(first_name):
            continue
        for col, v in enumerate(row, 1):
            _sb(ws.cell(row=r, column=col, value=v))
        r += 1

    _autofit(ws, [22, 35, 16, 10, 45, 22, 10, 14, 14])


def write_workbook(out_path: Path, entries: list[dict]) -> None:
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    _init_styles()
    wb = openpyxl.Workbook()

    _write_sheet1(wb, entries)
    _write_sheet2(wb, entries)
    _write_sheet3(wb, entries)

    wb.save(tmp_path)
    wb.close()
    _atomic_replace(tmp_path, out_path)


# ---------------------------------------------------------------------------
# Atomic file replace
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# File iteration
# ---------------------------------------------------------------------------

def _iter_pdfs(target: Path, recursive: bool):
    if target.is_file():
        yield target
        return
    pattern = "**/*.pdf" if recursive else "*.pdf"
    yield from sorted(target.glob(pattern))


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract creditor mailing-matrix entries from PDF into Excel."
    )
    parser.add_argument(
        "target", nargs="?",
        help="PDF file or folder of PDFs (prompted if omitted)",
    )
    parser.add_argument(
        "--recursive", action="store_true",
        help="Recurse into subfolders when target is a folder",
    )
    parser.add_argument(
        "--output", default=None,
        help="Custom output .xlsx path (default: next to input file)",
    )
    args = parser.parse_args()

    target_str = args.target
    if not target_str:
        target_str = input("Enter path to PDF file or folder: ").strip().strip("\"'")
    if not target_str:
        sys.exit("ERROR: no path provided.")

    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required.  Run:  pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required.    Run:  pip install openpyxl")

    target = Path(target_str)
    if not target.exists():
        sys.exit(f"ERROR: path not found — {target}")

    pdfs = list(_iter_pdfs(target, args.recursive))
    if not pdfs:
        sys.exit(f"ERROR: no PDF files found at {target}")

    print(f"Found {len(pdfs)} PDF file(s).  Extracting...")

    all_entries: list[dict] = []
    for pdf_path in pdfs:
        print(f"  Processing: {pdf_path.name}")
        try:
            entries = extract_pdf(pdf_path)
            all_entries.extend(entries)
            print(f"    {len(entries)} entr(y/ies) extracted.")
        except Exception as exc:
            print(f"    ERROR: {exc}")

    if args.output:
        out_path = Path(args.output)
    else:
        stem     = target.stem if target.is_file() else target.name
        folder   = target.parent if target.is_file() else target
        out_path = folder / f"{stem}_extracted.xlsx"

    if all_entries:
        person_count = sum(
            1 for e in all_entries
            if _is_person(parse_name(e["full_name"])["First Name"])
        )
        write_workbook(out_path, all_entries)
        print(f"\nWrote: {out_path}")
        print(f"  Total entries  : {len(all_entries)}")
        print(f"  Standard data  : {person_count} individual(s) (companies filtered out)")
    else:
        print("\nNo entries extracted — no output written.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
