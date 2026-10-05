r"""
Universal PDF Extract -- GOVERNMENT / BENEFIT FORMS ONLY: W-2, 401k, 1095-C, 1099.

    python "<this file>" <pdf_file_or_folder> [-o output_folder] [--recursive] [--format w2|1095c|1099|401k]

Each PDF is auto-detected as one of the four supported forms and written to its own
<filename>_extracted.xlsx (next to the PDF, or in -o). Anything else is NOT extracted: it is listed in
the summary as "Not a supported form". There is no rows/tables fallback and no AI step in this script.

Values are kept exactly as the document prints them. Only SSN (###-##-####, masked as XXX-XX-####) and
Date of Birth (MM/DD/YYYY) are formatted. "Last, First M" is split across the name columns without
changing the text. Every cell is stored as Excel Text, so Excel never strips leading zeros or reformats a
number, date or SSN. A column is written only when it has a value in that file.

RUN SUMMARY (always written)
    "<yymmdd> AS govt forms extraction summary.xlsx" (next to the input, or in -o): one row per
    document with file name, document type, pages, time taken, records, extracted columns, status and
    output file, plus a "Not Extracted" sheet listing every document that produced nothing and why.
    It holds column names and counts only, never extracted values.

DEPENDENCIES
    pip install pymupdf pdfplumber openpyxl tqdm
    pip install pypdf          (optional: reads fillable-form fields in 401k PDFs)

PII NOTICE
    The output contains SSNs, names and addresses. Save it only in the appropriate Global Insider
    folder -- never on the desktop or in a personal location. Do not paste its contents into Claude
    conversations.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Font
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pymupdf as fitz
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    def tqdm(iterable=None, *a, **k):
        return iterable if iterable is not None else _NullBar()

    class _NullBar:
        def write(self, msg):
            print(msg)
        def set_postfix_str(self, *a, **k):
            pass
        def update(self, *a, **k):
            pass
        def close(self):
            pass

# ===========================================================================
# Shared utilities — used across every subcommand
# ===========================================================================

def _clean(text) -> str:
    return " ".join(str(text or "").split()).strip()


def _iter_pdfs(target: Path, recursive: bool = False):
    if target.is_file():
        yield target
        return
    pattern = "**/*.pdf" if recursive else "*.pdf"
    yield from sorted(target.glob(pattern))


def mask_shape(s: str) -> str:
    """Digit/letter shape mask for safe-to-share debug output (PHI scripts)."""
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def normalize_text(s: str) -> str:
    """Fold OCR's curly quotes/dashes and odd whitespace to plain ASCII."""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate({0x2018: "'", 0x2019: "'", 0x00B4: "'", 0x0060: "'",
                      0x201C: '"', 0x201D: '"', 0x00A0: " ", 0x00AD: "-",
                      0x2013: "-", 0x2014: "-"})
    return re.sub(r"[ \t]+", " ", s).strip()


# ===========================================================================
# w2 — Employee Name/Address/SSN (+ optional wage boxes) from W-2 PDFs
#
# Identity engine ported from "260710 te ak 22222w2.py" (pymupdf/fitz,
# caption-regex driven, OCR-tolerant, multi-employee-per-page grid support) --
# the more robust of the two W-2 lineages found in this project, and a
# documented superset of w2_ssn_name_address_csv.py (dropped).
#
# Optional --wages flag adds Box 1-6 wage/tax extraction, ported from
# w2_wages_extractor.py's text/word proximity search (which is a superset of
# w2_extractor.py, also dropped) -- adapted here to fitz word tuples instead
# of pdfplumber dicts so the whole w2 subcommand runs on one PDF library.
# ===========================================================================

W2_MIN_TEXT_CHARS_PER_PAGE = 20


W2_NAME_SUFFIXES = {"JR", "SR", "II", "III", "IV", "V"}

W2_SSN_CAPTION_RE = re.compile(r"Employee.?s\s+(?:social security number|SSN|SSA number)\b", re.IGNORECASE)
W2_NAME_ONLY_CAPTION_RE = re.compile(r"Employee.?s\s+first name and initial\b", re.IGNORECASE)
W2_COMBINED_NAME_ADDR_CAPTION_RE = re.compile(r"Employee.?s\s+name,\s*address", re.IGNORECASE)
W2_ADDRESS_CAPTION_RE = re.compile(r"\bf\s+Employee.?s address\b", re.IGNORECASE)
W2_STOP_LABEL_RE = re.compile(
    r"^(?:\d{1,2}\s+State\b|Employer.s state ID|Local income tax|Locality name|"
    r"[a-f]\s+Employer.s\b|[a-f]\s+Employee.s\s+(?:SSN|SSA|social security)|Employer.s (?:name|FED|federal)|"
    r"Form\s*W-?2\b|Wage\s*(?:&|and)\s*Tax Statement|Copy\s+[A-Z0-9]|"
    r"For (?:Official|Privacy)|Department of the Treasury|20\d{2}$)",
    re.IGNORECASE,
)
W2_BOX_NUMERIC_LABEL_RE = re.compile(r"^1[1-4][a-d]?$")
W2_SSN_VALUE_RE = re.compile(r"(?<![A-Za-z0-9*])(\d{3}|[Xx*]{3})[-\s]?(\d{2}|[Xx*]{2})[-\s]?(\d{4})(?!\d)")
W2_CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b")
W2_EIN_VALUE_RE = re.compile(r"\b\d{2}-\d{7}\b")
W2_EMPLOYER_BLOCK_MAX_GAP_LINES = 5

W2_OCR_DIGIT_FIX = str.maketrans({
    "O": "0", "o": "0", "Q": "0", "D": "0",
    "I": "1", "l": "1", "i": "1", "|": "1",
    "Z": "2", "z": "2",
    "S": "5", "s": "5",
    "G": "6", "b": "6",
    "T": "7",
    "B": "8",
    "g": "9", "q": "9",
})


def w2_normalize_ssn(match):
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"


def group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, text = w[0], w[1], w[2], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def w2_find_ssn(plain_lines, page_num, column_label):
    for i, line in enumerate(plain_lines):
        if W2_SSN_CAPTION_RE.search(line):
            lo, hi = max(0, i - 2), min(len(plain_lines), i + 3)
            window_lines = plain_lines[lo:hi]
            window_text = " ".join(window_lines)
            m = W2_SSN_VALUE_RE.search(window_text)
            if m:
                return w2_normalize_ssn(m)
            m = W2_SSN_VALUE_RE.search(window_text.translate(W2_OCR_DIGIT_FIX))
            if m:
                print(f"  page {page_num} ({column_label}): SSN recovered via OCR digit-lookalike correction")
                return w2_normalize_ssn(m)
            shape = " | ".join(l if W2_SSN_CAPTION_RE.search(l) else mask_shape(l) for l in window_lines)
            print(f"  page {page_num} ({column_label}): SSN caption found but no SSN-shaped value nearby "
                  f"-- nearby line shapes (safe to share): {shape}")
            return ""
    print(f"  page {page_num} ({column_label}): SSN caption not found")
    return ""


def w2_find_right_column_boundary(word_line):
    # A street number that starts the row ("12 Fake St") is not a box-12 label.
    for i, (x0, x1, text) in enumerate(word_line):
        if i > 0 and W2_BOX_NUMERIC_LABEL_RE.match(text.strip(".:,")):
            return x0
    return None


W2_BOX_CAPTION_WORD_RE = re.compile(
    r"^(?:Wages|Federal|Social|Medicare|Allocated|Dependent|Nonqualified|Statutory|Retirement|Third|See|"
    r"Employer|Employer's|Control|State|Local|Locality|Advance|Verification|Deferred)\b", re.IGNORECASE)


def w2_find_right_column_boundary_strict(word_line):
    """x of the first numbered box caption ("6 Medicare tax withheld", "12a See instructions") to the
    right of the row's first word, or None. Used where boxes 5/6/12 sit beside the name/address box."""
    for i in range(0, len(word_line) - 1):
        x0, _, text = word_line[i]
        if re.fullmatch(r"\d{1,2}[a-d]?", text.strip(".:,")) and W2_BOX_CAPTION_WORD_RE.match(word_line[i + 1][2]):
            return x0
    return None


W2_PANEL_GAP = 40  # a horizontal gap this wide inside one row separates side-by-side panels (e.g. a copy of the name block)


def w2_row_text_left_of(word_line, max_x):
    words = word_line if max_x is None else [w for w in word_line if w[0] < max_x]
    kept = []
    for w in words:
        if kept and w[0] - kept[-1][1] > W2_PANEL_GAP:
            break
        kept.append(w)
    return normalize_text(" ".join(t for _, _, t in kept))


def w2_find_e_box_dimensions(caption_word_line):
    last_x0 = suff_x0 = suff_x1 = None
    for x0, x1, text in caption_word_line:
        low = text.strip(".:,").lower()
        if low == "last" and last_x0 is None:
            last_x0 = x0
        elif low == "suff" and suff_x0 is None:
            suff_x0 = x0
            suff_x1 = x1
    static_boundary = (suff_x1 + 15) if suff_x1 is not None else None
    return last_x0, suff_x0, static_boundary


def w2_split_name_row(word_line, last_x0, suff_x0, max_x):
    words = word_line if max_x is None else [w for w in word_line if w[0] < max_x]
    join = lambda ws: normalize_text(" ".join(t for _, _, t in ws))
    if last_x0 is None:
        return join(words), "", ""
    first_words = [w for w in words if w[0] < last_x0]
    if suff_x0 is not None:
        last_words = [w for w in words if last_x0 <= w[0] < suff_x0]
        suff_words = [w for w in words if w[0] >= suff_x0]
    else:
        last_words = [w for w in words if w[0] >= last_x0]
        suff_words = []
    return join(first_words), join(last_words), join(suff_words)


def w2_split_name_by_whitespace(name):
    tokens = name.split()
    suffix = ""
    if tokens and tokens[-1].strip(".").upper() in W2_NAME_SUFFIXES:
        suffix = tokens[-1]
        tokens = tokens[:-1]
    if not tokens:
        return "", "", suffix
    if len(tokens) == 1:
        return tokens[0], "", suffix
    return " ".join(tokens[:-1]), tokens[-1], suffix


def w2_empty_name_address_fields():
    return {"First Name": "", "Last Name": "", "Suffix": "",
            "Street Address": "", "City": "", "State": "", "Zip Code": ""}


def w2_split_address_block(block, start_idx, page_num, column_label):
    fields = {"Street Address": "", "City": "", "State": "", "Zip Code": ""}
    addr_lines = block[start_idx:]
    city_idx = next((idx for idx, l in enumerate(addr_lines) if W2_CITY_STATE_ZIP_RE.match(l)), None)
    if city_idx is not None:
        m = W2_CITY_STATE_ZIP_RE.match(addr_lines[city_idx])
        fields["City"] = m.group("city").rstrip(",")
        fields["State"] = m.group("state")
        fields["Zip Code"] = m.group("zip")
        fields["Street Address"] = " ".join(addr_lines[:city_idx])
    else:
        fields["Street Address"] = " ".join(addr_lines)
        if addr_lines:
            shape = " | ".join(mask_shape(l) for l in addr_lines)
            print(f"  page {page_num} ({column_label}): could not find a city/state/zip-shaped line "
                  f"in the address block -- line shapes (safe to share): {shape}")
    return fields


def w2_find_name_address(lines, page_num, column_label):
    plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]

    name_idx = next((i for i, l in enumerate(plain_lines) if W2_NAME_ONLY_CAPTION_RE.search(l)), None)
    if name_idx is not None:
        return _w2_find_name_address_separate_boxes(lines, plain_lines, name_idx, page_num, column_label)

    # Boxes 5/6/12 sit in a column to the right of the name/address box: clip every row at that column's x
    # (taken from the numbered captions), so values printed without a caption on their row are cut too.
    left_x = min((w[0] for ln in lines for w in ln), default=0)
    cap_idx = next((i for i, l in enumerate(plain_lines) if W2_COMBINED_NAME_ADDR_CAPTION_RE.search(l)), None)
    block_end = len(lines)
    if cap_idx is not None:  # only captions beside the name block count; boxes below its stop label don't
        block_end = next((i for i in range(cap_idx + 1, len(plain_lines))
                          if W2_STOP_LABEL_RE.search(plain_lines[i])), len(lines))
    right_xs = [b for b in (w2_find_right_column_boundary_strict(ln) for ln in lines[:block_end])
                if b is not None and b >= left_x + 100]  # a box caption at the left margin is not a right column
    right_x = min(right_xs) if right_xs else None
    # Rows that only hold text from a side-by-side panel (e.g. a "Box 1 of W-2" header printed to the right of the
    # name block) start far right of the caption's own left edge; blank them so they aren't read as name/address.
    cap_x0 = lines[cap_idx][0][0] if cap_idx is not None and lines[cap_idx] else left_x
    clipped = [normalize_text(w2_row_text_left_of(ln, right_x)) if ln and ln[0][0] <= cap_x0 + W2_PANEL_GAP else ""
               for ln in lines]
    for i, line in enumerate(plain_lines):
        if W2_COMBINED_NAME_ADDR_CAPTION_RE.search(line):
            return _w2_find_name_address_combined_box(clipped, i, page_num, column_label)

    print(f"  page {page_num} ({column_label}): name/address caption not found")
    return w2_empty_name_address_fields()


def _w2_find_name_address_combined_box(plain_lines, caption_idx, page_num, column_label):
    raw_block = plain_lines[caption_idx + 1:min(caption_idx + 9, len(plain_lines))]
    block = []
    for candidate in raw_block:
        candidate = candidate.strip()
        if not candidate:
            continue
        if W2_STOP_LABEL_RE.search(candidate) or W2_ADDRESS_CAPTION_RE.search(candidate):
            break
        block.append(candidate)
    if not block:
        shape = " | ".join(mask_shape(l) for l in raw_block)
        print(f"  page {page_num} ({column_label}): name/address caption found but block below it is "
              f"empty (safe to share): {shape}")
        return w2_empty_name_address_fields()

    first, last, suffix = w2_split_name_by_whitespace(block[0])
    fields = w2_split_address_block(block, 1, page_num, column_label)
    fields.update({"First Name": first, "Last Name": last, "Suffix": suffix})
    return fields


def _w2_find_name_address_separate_boxes(lines, plain_lines, name_idx, page_num, column_label):
    last_x0, suff_x0, static_boundary = w2_find_e_box_dimensions(lines[name_idx])

    first = last = suffix = ""
    plain_block = []
    name_row_used = False
    for j in range(name_idx + 1, min(name_idx + 40, len(plain_lines))):
        if W2_STOP_LABEL_RE.search(plain_lines[j]):
            break
        if W2_ADDRESS_CAPTION_RE.search(plain_lines[j]):
            continue
        row_boundary = w2_find_right_column_boundary(lines[j]) or static_boundary
        text = w2_row_text_left_of(lines[j], row_boundary)
        if not text:
            continue
        if not (w2_is_letterish(text) or W2_CITY_STATE_ZIP_RE.match(text)):
            continue
        if not name_row_used:
            first, last, suffix = w2_split_name_row(lines[j], last_x0, suff_x0, row_boundary)
            name_row_used = True
        else:
            plain_block.append(text)

    if not name_row_used:
        print(f"  page {page_num} ({column_label}): 'e' name caption found but box e/f's own column is "
              f"empty all the way down -- name/address left blank")
        return w2_empty_name_address_fields()

    fields = w2_split_address_block(plain_block, 0, page_num, column_label)
    fields.update({"First Name": first, "Last Name": last, "Suffix": suffix})
    return fields


def w2_extract_employee(lines, plain_lines, page_num, column_label):
    ssn = w2_find_ssn(plain_lines, page_num, column_label)
    fields = w2_find_name_address(lines, page_num, column_label)
    return {"Page": page_num, "SSN": ssn, **fields}


def w2_is_letterish(line):
    letters = sum(c.isalpha() for c in line)
    digits = sum(c.isdigit() for c in line)
    return letters >= 3 and letters > digits


def w2_find_ssn_by_shape(plain_lines, page_num, column_label):
    for line in plain_lines:
        m = W2_SSN_VALUE_RE.search(line)
        if m:
            return w2_normalize_ssn(m)
        m = W2_SSN_VALUE_RE.search(line.translate(W2_OCR_DIGIT_FIX))
        if m:
            print(f"  page {page_num} ({column_label}): [shape-fallback] SSN recovered via OCR digit-lookalike correction")
            return w2_normalize_ssn(m)
    print(f"  page {page_num} ({column_label}): [shape-fallback] no SSN-shaped value found in this cell")
    return ""


def w2_find_name_address_by_shape(plain_lines, page_num, column_label):
    csz_indices = [i for i, l in enumerate(plain_lines) if W2_CITY_STATE_ZIP_RE.match(l)]
    if not csz_indices:
        print(f"  page {page_num} ({column_label}): [shape-fallback] no city/state/zip-shaped line found")
        return w2_empty_name_address_fields()

    blocks = []
    for csz_idx in csz_indices:
        block = [csz_idx]
        j = csz_idx - 1
        while j >= 0 and len(block) < 4:
            line = plain_lines[j].strip()
            if not line or W2_CITY_STATE_ZIP_RE.match(line) or not w2_is_letterish(line):
                break
            block.insert(0, j)
            j -= 1
        if len(block) >= 2:
            blocks.append(block)

    if not blocks:
        print(f"  page {page_num} ({column_label}): [shape-fallback] no preceding name/street line(s)")
        return w2_empty_name_address_fields()

    ein_idx = next((i for i, l in enumerate(plain_lines) if W2_EIN_VALUE_RE.search(l)), None)

    if ein_idx is None:
        if len(blocks) > 1:
            print(f"  page {page_num} ({column_label}): [shape-fallback] no EIN found to anchor on, using the "
                  f"last block as the employee's; spot-check this record")
            chosen = blocks[-1]
        else:
            print(f"  page {page_num} ({column_label}): [shape-fallback] only one block found and no EIN to "
                  f"confirm employer vs employee -- skipping")
            return w2_empty_name_address_fields()
    else:
        employee_like = [b for b in blocks if b[0] - ein_idx > W2_EMPLOYER_BLOCK_MAX_GAP_LINES]
        if not employee_like:
            print(f"  page {page_num} ({column_label}): [shape-fallback] every block sits within "
                  f"{W2_EMPLOYER_BLOCK_MAX_GAP_LINES} lines of the EIN -- that's the employer's, not the employee's")
            return w2_empty_name_address_fields()
        if len(employee_like) > 1:
            print(f"  page {page_num} ({column_label}): [shape-fallback] {len(employee_like)} candidate employee "
                  f"blocks found -- using the last one; spot-check this record")
        chosen = employee_like[-1]

    name = plain_lines[chosen[0]].strip()
    csz_line = plain_lines[chosen[-1]].strip()
    m = W2_CITY_STATE_ZIP_RE.match(csz_line)
    first, last, suffix = w2_split_name_by_whitespace(name)
    return {
        "First Name": first, "Last Name": last, "Suffix": suffix,
        "Street Address": " ".join(plain_lines[i].strip() for i in chosen[1:-1]),
        "City": m.group("city").rstrip(","), "State": m.group("state"), "Zip Code": m.group("zip"),
    }


def w2_extract_employee_by_shape(plain_lines, page_num, column_label):
    ssn = w2_find_ssn_by_shape(plain_lines, page_num, column_label)
    fields = w2_find_name_address_by_shape(plain_lines, page_num, column_label)
    return {"Page": page_num, "SSN": ssn, **fields}


def w2_caption_line_groups(words, caption_re, y_tol=3):
    groups = {}
    for w in words:
        x0, y0, _, _, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        groups.setdefault(key, []).append((x0, text))
    result = []
    for key, items in sorted(groups.items()):
        joined = normalize_text(" ".join(t for _, t in sorted(items, key=lambda t: t[0])))
        count = len(caption_re.findall(joined))
        if count:
            result.append((key, count))
    return result


def w2_build_grid_cells(page, words, groups, split_fraction, page_num):
    row_ys = [y for y, _ in groups]
    row_counts = [c for _, c in groups]
    num_rows = len(row_ys)

    if num_rows <= 1:
        row_bounds = [(float("-inf"), float("inf"))]
    else:
        page_height = page.rect.height
        edges = [page_height * i / num_rows for i in range(1, num_rows)]
        edges = [float("-inf")] + edges + [float("inf")]
        row_bounds = [(edges[i], edges[i + 1]) for i in range(num_rows)]

    if num_rows == 1:
        row_label = lambda idx: ""
    elif num_rows == 2:
        row_label = lambda idx: ("Top", "Bottom")[idx]
    else:
        row_label = lambda idx: f"Row{idx + 1}"

    split_x = page.rect.width * split_fraction
    cells = []
    for idx, (y_lo, y_hi) in enumerate(row_bounds):
        row_words = words if num_rows <= 1 else [w for w in words if y_lo <= (w[1] + w[3]) / 2 < y_hi]
        count = row_counts[idx] if idx < len(row_counts) else 1
        prefix = row_label(idx)

        if count >= 2:
            if count > 2:
                print(f"  page {page_num}: row {idx + 1} has {count} SSN captions (expected 1 or 2) -- only "
                      f"the first 2 columns will be split out")
            cells.append((prefix + "Left", [w for w in row_words if (w[0] + w[2]) / 2 < split_x]))
            cells.append((prefix + "Right", [w for w in row_words if (w[0] + w[2]) / 2 >= split_x]))
        else:
            cells.append((prefix if prefix else "Single", row_words))

    return cells


W2_WAGE_BOX_DEFS = [
    ("Box 1 Wages Tips Other Comp", ["wages, tips, other comp", "wages, tips, other", "1 wages"]),
    ("Box 2 Federal Tax Withheld", ["federal income tax withheld", "fed income tax", "2 federal"]),
    ("Box 3 SS Wages", ["social security wages", "3 social security wages", "ss wages"]),
    ("Box 4 SS Tax Withheld", ["social security tax withheld", "4 social security tax", "ss tax withheld"]),
    ("Box 5 Medicare Wages", ["medicare wages and tips", "5 medicare wages", "medicare wages"]),
    ("Box 6 Medicare Tax Withheld", ["medicare tax withheld", "6 medicare tax", "medicare tax"]),
]
W2_WAGE_COLUMNS = [key for key, _ in W2_WAGE_BOX_DEFS]
W2_AMOUNT_RE = re.compile(r"\$?\s*(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)")


def w2_parse_amount(text: str) -> str:
    m = W2_AMOUNT_RE.search(text)
    return m.group(0).replace(" ", "") if m else ""


W2_STRICT_AMOUNT_RE = re.compile(r"^\$?(?:\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d{3,})$")
W2_BOX_NUMBER_RE = re.compile(r"^(?:\d{1,2}[a-d]?|[a-f])$")


def w2_extract_wages_from_lines(lines) -> dict:
    """lines: list of lines, each a list of (x0, x1, text) tuples.

    Boxes sit side by side on one text row (caption row, then a value row
    below), so each label's value is searched on its own line to the right of
    the label, then on the next two lines, but only within the label's own
    column -- from the box number before the label up to the next box number
    on the same row. Only money-shaped tokens count, so a neighbouring box
    number such as "2" is never read as an amount."""
    results = {key: "" for key, _ in W2_WAGE_BOX_DEFS}
    for box_key, label_list in W2_WAGE_BOX_DEFS:
        for li, ln in enumerate(lines):
            words = [t for _, _, t in ln]
            lowers = [w.lower() for w in words]
            text = " ".join(lowers)
            label = next((lbl for lbl in label_list if lbl in text), None)
            if label is None:
                continue
            pos = text.find(label)
            starts, off = [], 0
            for w in lowers:
                starts.append(off)
                off += len(w) + 1
            wi = max(k for k in range(len(starts)) if starts[k] <= pos)
            end_char = pos + len(label)
            end_wi = next((k for k in range(len(starts)) if starts[k] >= end_char), len(words))
            x_start = ln[wi][0]
            if wi > 0 and W2_BOX_NUMBER_RE.match(words[wi - 1]):
                x_start = ln[wi - 1][0]
            x_end = float("inf")
            for k in range(end_wi, len(words) - 1):
                if W2_BOX_NUMBER_RE.match(words[k]) and words[k + 1][:1].isalpha():
                    x_end = ln[k][0]
                    break
            candidates = [ln[k] for k in range(end_wi, len(words)) if ln[k][0] < x_end]
            for below in lines[li + 1: li + 3]:
                candidates += [w for w in below if x_start - 6 <= w[0] < x_end]
            amount = next((w2_parse_amount(t) for _, _, t in candidates if W2_STRICT_AMOUNT_RE.match(t)), "")
            if amount:
                results[box_key] = amount
                break
    return results


def w2_extract_wages_for_cell(column_words) -> dict:
    return w2_extract_wages_from_lines(group_words_into_lines(column_words))


def w2_process_page(page, page_num, split_fraction, want_wages):
    text = page.get_text()
    if len(text.strip()) < W2_MIN_TEXT_CHARS_PER_PAGE:
        print(f"  page {page_num}: only {len(text.strip())} chars of text -- scanned/image-only page, skipping")
        return []

    norm_text = normalize_text(text)
    caption_count = len(W2_SSN_CAPTION_RE.findall(norm_text))
    shape_fallback = False
    marker_re = W2_SSN_CAPTION_RE
    marker_count = caption_count

    if caption_count == 0:
        value_count = len(W2_SSN_VALUE_RE.findall(norm_text))
        if value_count == 0:
            print(f"  page {page_num}: no 'Employee's SSN' caption and no SSN-shaped value found -- skipping")
            return []
        print(f"  page {page_num}: no SSN caption in the OCR text -- falling back to locating employees by "
              f"SSN value shape/position")
        shape_fallback = True
        marker_re = W2_SSN_VALUE_RE
        marker_count = value_count

    words = page.get_text("words")
    records = []

    if marker_count == 1:
        cells = [("Single", words)]
    else:
        groups = w2_caption_line_groups(words, marker_re)
        if not groups:
            print(f"  page {page_num}: {marker_count} SSN {'values' if shape_fallback else 'captions'} found "
                  f"but none could be grouped into row bands -- treating the whole page as one employee")
            cells = [("Single", words)]
        else:
            cells = w2_build_grid_cells(page, words, groups, split_fraction, page_num)

    for column_label, column_words in cells:
        lines = group_words_into_lines(column_words)
        plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]

        if shape_fallback:
            has_ssn_here = any(W2_SSN_VALUE_RE.search(l) or W2_SSN_VALUE_RE.search(l.translate(W2_OCR_DIGIT_FIX))
                                for l in plain_lines)
            if not has_ssn_here:
                print(f"  page {page_num} ({column_label}): [shape-fallback] no SSN-shaped value in this cell")
                continue
            rec = w2_extract_employee_by_shape(plain_lines, page_num, column_label)
        else:
            if not any(W2_SSN_CAPTION_RE.search(l) for l in plain_lines):
                print(f"  page {page_num} ({column_label}): no SSN caption on this cell")
                continue
            rec = w2_extract_employee(lines, plain_lines, page_num, column_label)

        if want_wages:
            rec.update(w2_extract_wages_for_cell(column_words))
        records.append(rec)

    return records


def w2_process_pdf(path: Path, split_fraction, want_wages):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        records.extend(w2_process_page(page, i, split_fraction, want_wages))
    doc.close()
    return records


def w2_dedupe_records(records):
    seen = set()
    deduped = []
    for rec in records:
        ssn = rec.get("SSN", "")
        name_addr = (rec.get("First Name", ""), rec.get("Last Name", ""), rec.get("Suffix", ""),
                     rec.get("Street Address", ""), rec.get("City", ""), rec.get("State", ""),
                     rec.get("Zip Code", ""))
        if ssn:
            key = ("ssn", ssn)
        elif any(name_addr):
            key = ("name_addr",) + name_addr
        else:
            deduped.append(rec)
            continue
        if key in seen:
            continue
        seen.add(key)
        deduped.append(rec)
    return deduped


# ===========================================================================
# 1095c — Employee identity/address fields from IRS Form 1095-C PDFs
# Ported from extract_1095c.py (pymupdf/fitz, Format A / Format B detection).
# ===========================================================================

C1095_MIN_TEXT_CHARS_PER_PAGE = 20

C1095_OUTPUT_COLUMNS = ["Page", "Format", "First Name", "Middle Name", "Last Name",
                         "Street Address", "City", "State", "Zip Code", "Country", "SSN"]

C1095_CITY_STATE_ZIP_RE = re.compile(
    r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b\s*(?P<country>.*)$"
)
C1095_SSN_VALUE_RE = re.compile(r"[\dXx*]{3}[-\s]?[\dXx*]{2}[-\s]?[\dXx*]{4}")
C1095_FORMAT_A_NAME_LABEL_RE = re.compile(r"EMPLOYEE.s name,\s*address,\s*ZIP/?postal code\s*&?\s*country", re.IGNORECASE)
C1095_FORMAT_A_SSN_LABEL_RE = re.compile(r"EMPLOYEE.s social security number\s*\(SSN\)", re.IGNORECASE)
C1095_FORMAT_A_NAME_BLOCK_END_RE = re.compile(r"Do not attach to your tax return", re.IGNORECASE)
C1095_FORMAT_A_SSN_BLOCK_END_RE = re.compile(
    r"Department of the Treasury|Covered Individuals|APPLICABLE LARGE EMPLOYER", re.IGNORECASE
)
C1095_FORMAT_B_FIELDS = [
    ("employee_name", re.compile(r"^\s*1\s+Name of employee", re.IGNORECASE)),
    ("ssn", re.compile(r"^\s*2\s+Social security number\s*\(SSN\)", re.IGNORECASE)),
    ("street", re.compile(r"^\s*3\s+Street address(?:\s*\(including apartment no\.?\))?", re.IGNORECASE)),
    ("city", re.compile(r"^\s*4\s+City or town", re.IGNORECASE)),
    ("state", re.compile(r"^\s*5\s+State or province", re.IGNORECASE)),
    ("zip", re.compile(r"^\s*6\s+Country and ZIP(?:\s+or foreign postal code)?", re.IGNORECASE)),
    ("_end", re.compile(r"^\s*7\s+Name of employer", re.IGNORECASE)),
]
C1095_ZIP_CAPTION_LEFTOVER_RE = re.compile(r"^(?:\b(?:or|foreign|postal|code)\b\s*)+", re.IGNORECASE)
C1095_DIAGNOSTIC_TERMS = ["1095-C", "employee", "employer", "social security", "coverage"]


def c1095_split_employee_name(full_name):
    full_name = re.sub(r"\s+", " ", (full_name or "").strip())
    if not full_name:
        return "", "", "", False
    parts = full_name.split(" ")
    if len(parts) == 2:
        return parts[0], "", parts[1], False
    if len(parts) == 1:
        return parts[0], "", "", False
    return parts[0], " ".join(parts[1:-1]), parts[-1], len(parts) != 3


def c1095_find_label_span(plain_lines, label_re, start=0, max_window=3):
    for i in range(start, len(plain_lines)):
        for span in range(1, max_window + 1):
            if i + span > len(plain_lines):
                break
            joined = " ".join(plain_lines[i:i + span])
            m = label_re.match(joined)
            if m:
                return i, i + span - 1, joined[m.end():].strip()
    return None, None, None


def c1095_value_after_span(plain_lines, span_end_idx, trailing_text, boundary_start_idx):
    between = [l.strip() for l in plain_lines[span_end_idx + 1:boundary_start_idx] if l.strip()]
    combined = ([trailing_text] if trailing_text else []) + between
    return " ".join(combined)


def c1095_log_no_match_diagnostics(plain_lines, page_num):
    joined = " ".join(plain_lines).lower()
    hits = {term: joined.count(term.lower()) for term in C1095_DIAGNOSTIC_TERMS}
    print(f"  page {page_num}: neither Format A nor Format B employee block found -- skipping")
    print(f"    diagnostic (counts only, no PHI): {len(plain_lines)} line(s) extracted; "
          + ", ".join(f"'{term}': {count}" for term, count in hits.items()))


def c1095_extract_format_a(plain_lines, page_num):
    name_start, name_end, name_trailing = c1095_find_label_span(plain_lines, C1095_FORMAT_A_NAME_LABEL_RE)
    if name_start is None:
        return None

    end_start, _, _ = c1095_find_label_span(plain_lines, C1095_FORMAT_A_NAME_BLOCK_END_RE, name_end + 1)
    end_idx = end_start if end_start is not None else min(name_end + 6, len(plain_lines))

    if C1095_FORMAT_A_SSN_LABEL_RE.search(name_trailing or ""):
        name_trailing = ""  # the SSN caption sits on the same row as the name caption
    block = [l.strip() for l in [name_trailing] + plain_lines[name_end + 1:end_idx] if l.strip()]
    if not block:
        print(f"  page {page_num} (Format A): name/address block empty after label -- skipping")
        return None

    full_name = block[0]
    inline_ssn = ""
    m_inline = C1095_SSN_VALUE_RE.search(full_name)
    if m_inline:  # "Jane Q Doe 000-12-3456" -- name and SSN share the value row
        inline_ssn = m_inline.group()
        full_name = (full_name[:m_inline.start()] + full_name[m_inline.end():]).strip()
    street = block[1] if len(block) > 1 else ""
    city = state = zip_code = country = ""
    if len(block) > 2:
        m = C1095_CITY_STATE_ZIP_RE.match(block[2])
        if m:
            city, state, zip_code, country = m.group("city").rstrip(","), m.group("state"), m.group("zip"), m.group("country").strip()
        else:
            city = block[2]
            print(f"  page {page_num} (Format A): could not split city/state/zip line -- placing raw text in City")

    ssn = ""
    ssn_start, ssn_end, ssn_trailing = c1095_find_label_span(plain_lines, C1095_FORMAT_A_SSN_LABEL_RE, end_idx)
    if inline_ssn:
        ssn = inline_ssn
    elif ssn_start is not None:
        ssn_boundary_start, _, _ = c1095_find_label_span(plain_lines, C1095_FORMAT_A_SSN_BLOCK_END_RE, ssn_end + 1)
        if ssn_boundary_start is None:
            ssn_boundary_start = min(ssn_end + 4, len(plain_lines))
        ssn_text = c1095_value_after_span(plain_lines, ssn_end, ssn_trailing, ssn_boundary_start)
        ssn_match = C1095_SSN_VALUE_RE.search(ssn_text)
        ssn = ssn_match.group() if ssn_match else ""
    else:
        print(f"  page {page_num} (Format A): SSN label not found")

    return {
        "Page": page_num, "Format": "A", "full_name": full_name,
        "Street Address": street, "City": city, "State": state,
        "Zip Code": zip_code, "Country": country, "SSN": ssn,
    }


def c1095_extract_format_b(plain_lines, page_num):
    spans = {}
    search_from = 0
    for key, label_re in C1095_FORMAT_B_FIELDS:
        start, end, trailing = c1095_find_label_span(plain_lines, label_re, search_from)
        if start is None:
            if key == "employee_name":
                return None
            continue
        spans[key] = (start, end, trailing)
        search_from = end + 1

    if "employee_name" not in spans:
        return None

    ordered_keys = [k for k, _ in C1095_FORMAT_B_FIELDS]
    values = {}
    for i, key in enumerate(ordered_keys):
        if key == "_end" or key not in spans:
            continue
        _, span_end, trailing = spans[key]
        next_start = None
        for later_key in ordered_keys[i + 1:]:
            if later_key in spans:
                next_start = spans[later_key][0]
                break
        if next_start is None:
            next_start = min(span_end + 3, len(plain_lines))
        values[key] = c1095_value_after_span(plain_lines, span_end, trailing, next_start)

    missing = [k for k in ("employee_name", "ssn", "street", "city", "state", "zip") if not values.get(k)]
    if missing:
        print(f"  page {page_num} (Format B): fields not matched: {', '.join(missing)}")

    zip_text = C1095_ZIP_CAPTION_LEFTOVER_RE.sub("", values.get("zip", "")).strip()
    zip_match = re.search(r"\d{5}(?:-\d{4})?", zip_text)
    zip_code = zip_match.group() if zip_match else ""
    country = zip_text[:zip_match.start()].strip() if zip_match else zip_text.strip()
    ssn_match = C1095_SSN_VALUE_RE.search(values.get("ssn", ""))

    return {
        "Page": page_num, "Format": "B", "full_name": values.get("employee_name", ""),
        "Street Address": values.get("street", ""), "City": values.get("city", ""),
        "State": values.get("state", ""), "Zip Code": zip_code, "Country": country,
        "SSN": ssn_match.group() if ssn_match else "",
    }


def c1095_dump_debug_lines(plain_lines, pdf_stem, page_num, output_dir):
    debug_path = output_dir / f"{pdf_stem}_debug.txt"
    with open(debug_path, "a", encoding="utf-8") as f:
        f.write(f"\n===== page {page_num} =====\n")
        for i, line in enumerate(plain_lines):
            f.write(f"[{i}] {line}\n")
    return debug_path


def c1095_process_page(page, page_num, debug=False, pdf_stem=None, output_dir=None):
    text_len = len(page.get_text().strip())
    if text_len < C1095_MIN_TEXT_CHARS_PER_PAGE:
        print(f"  page {page_num}: only {text_len} chars of text -- scanned/image-only page, skipping")
        return None

    words = page.get_text("words")
    lines = group_words_into_lines(words)
    plain_lines = [" ".join(t for _, _, t in ln) for ln in lines]

    if debug:
        c1095_dump_debug_lines(plain_lines, pdf_stem, page_num, output_dir)

    record = c1095_extract_format_a(plain_lines, page_num) or c1095_extract_format_b(plain_lines, page_num)
    if record is None:
        c1095_log_no_match_diagnostics(plain_lines, page_num)
        return None

    first, middle, last, nonstandard = c1095_split_employee_name(record.pop("full_name"))
    record["First Name"], record["Middle Name"], record["Last Name"] = first, middle, last
    if nonstandard:
        print(f"  page {page_num}: employee name has more than 3 words -- review this row")
    return record


def c1095_process_pdf(path: Path, debug=False, output_dir=None):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        record = c1095_process_page(page, i, debug=debug, pdf_stem=path.stem, output_dir=output_dir)
        if record:
            records.append(record)
    doc.close()
    return records


def c1095_dedupe_complete_rows(records):
    compare_cols = [c for c in C1095_OUTPUT_COLUMNS if c != "Page"]
    seen = set()
    kept = []
    for rec in records:
        key = tuple((str(rec.get(c, "") or "")).strip().lower() for c in compare_cols)
        if key in seen:
            continue
        seen.add(key)
        kept.append(rec)
    return kept


# ===========================================================================
# 401k — employee 401k data into the standard 38-column PII template
# Ported from extract_401k.py (pdfplumber + pypdf AcroForm + regex fallback).
# ===========================================================================

K401_COL_DOCID = 1
K401_COL_LAST_NAME = 2
K401_COL_FIRST_NAME = 3
K401_COL_MIDDLE_NAME = 4
K401_COL_SUFFIX = 5
K401_COL_SUBJECT_TYPE = 6
K401_COL_GOV_ID_TAG = 18
K401_COL_SSN = 19
K401_COL_WORK_INFO = 37
K401_COL_EE_ID = 38
K401_TEMPLATE_COL_COUNT = 38

K401_SSN_RE = re.compile(r'\b(\d{3})-(\d{2})-(\d{4})\b')
K401_SSN_RE_XPAD = re.compile(r'(?<!\d)(\d{3})-(\d{2})-([\dxX*]{1,4})(?![\dxX*])', re.I)
K401_AMOUNT_SIGNED_RE = re.compile(r'\(?-?\$?\s*(?:\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?)\)?')
K401_AMOUNT_RE = re.compile(r'\$?\s*([\d]{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)')
K401_NAME_RE = re.compile(r'\b([A-Z][a-z]+(?:[ \t]+[A-Z][a-z]*\.?){1,4})\b')
K401_NAME_RE_UPPER = re.compile(r'\b([A-Z]{2,}(?:[ \t]+[A-Z]{2,}){1,4})\b')
K401_NAME_RE_LASTFIRST = re.compile(r'\b([A-Za-z]{2,}),[ \t]+([A-Za-z]{2,}(?:[ \t]+[A-Za-z]{1,}\.?)?)\b')

K401_SSN_PARTIAL_PATTERNS = [
    (re.compile(r'\b(\d{3})-(\d{2})-(\d{1,3})\b(?!\d)'), "short_last"),
    (re.compile(r'\b(\d{3})-(\d{2})\b(?!\s*-?\s*\d)'), "front5"),
    (re.compile(r'[xX*]{3}[- ][xX*]{2}[- ](\d{4})\b'), "last4"),
    (re.compile(r'\b(\d{3})[- ][xX*]{2}[- ](\d{4})\b'), "mid_masked"),
]

K401_SUFFIXES = {
    "jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v",
    "esq", "esq.", "phd", "md", "dds", "cpa", "rn",
}

K401_COL_MAP: dict = {}


def _k401_add(canonical, *aliases):
    for a in aliases:
        K401_COL_MAP[a.lower()] = canonical


_k401_add("first_name", "first name", "first", "fname", "given name", "employee first name", "participant first name")
_k401_add("last_name", "last name", "last", "lname", "surname", "family name", "employee last name", "participant last name")
_k401_add("middle_name", "middle name", "middle", "middle initial", "mi", "mname", "employee middle name")
_k401_add("suffix", "suffix", "name suffix", "generational suffix", "jr", "sr")
_k401_add("name", "name", "employee name", "emp name", "employee", "participant name", "participant",
          "last, first", "last/first", "full name")
_k401_add("ssn", "ssn", "social security", "social security number", "soc sec", "ss number", "ss#", "ssn#",
          "tax id", "tin", "employee ssn", "participant ssn", "employee id / ssn")
_k401_add("gross_salary", "gross salary", "gross pay", "gross wages", "gross compensation", "annual salary",
          "annual compensation", "compensation", "wages", "salary", "ytd gross", "ytd wages", "total compensation")
_k401_add("ee_contribution", "ee contribution", "employee contribution", "ee deferral", "employee deferral",
          "employee 401k", "elective deferral", "pre-tax contribution", "pre tax contribution",
          "roth contribution", "employee amount", "ee amount", "employee contrib")
_k401_add("er_match", "er match", "employer match", "employer contribution", "company match",
          "company contribution", "er contribution", "matching contribution", "employer match amount")
_k401_add("ytd_contribution", "ytd contribution", "ytd total", "year to date", "ytd", "ytd deferral", "total ytd")
_k401_add("ee_id", "employee id", "ee id", "emp id", "employee number", "emp number",
          "employee identification", "employee identification number", "ee#")


def k401_map_column(raw_header: str):
    normalised = re.sub(r'\s+', ' ', (raw_header or "").strip().lower().rstrip(":"))
    if normalised in K401_COL_MAP:
        return K401_COL_MAP[normalised]
    best, best_len = None, 0
    for alias, canonical in K401_COL_MAP.items():
        if alias in normalised and len(alias) > best_len:
            best, best_len = canonical, len(alias)
    return best


def k401_fmt_ssn(raw: str) -> str:
    raw = _clean(raw)
    if not raw:
        return raw
    if re.match(r'^[\dX]{3}-[\dX]{2}-[\dX]{4}$', raw, re.I):
        return raw.upper()
    m = re.match(r'^(\d{3})-(\d{2})-(\d{4})$', raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.match(r'^(\d{3})-(\d{2})-(\d{1,3})$', raw)
    if m:
        serial = m.group(3).ljust(4, 'X')
        return f"{m.group(1)}-{m.group(2)}-{serial}"
    m = re.match(r'^(\d{3})-(\d{2})$', raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}-XXXX"
    m = re.match(r'^(\d{3})-[xX*]+-(\d{4})$', raw)
    if m:
        return f"{m.group(1)}-XX-{m.group(2)}"
    m = re.match(r'^[xX*]+-(\d{2})-(\d{4})$', raw)
    if m:
        return f"XXX-{m.group(1)}-{m.group(2)}"
    m = re.match(r'^[xX*]+-[xX*]+-(\d{4})$', raw)
    if m:
        return f"XXX-XX-{m.group(1)}"
    d = re.sub(r'\D', '', raw)
    n = len(d)
    if n >= 9:
        return f"{d[:3]}-{d[3:5]}-{d[5:9]}"
    if n == 8:
        return f"{d[0:3]}-{d[3:5]}-{d[5:8]}X"
    if n == 7:
        return f"{d[0:3]}-{d[3:5]}-{d[5:7]}XX"
    if n == 6:
        return f"{d[0:3]}-{d[3:5]}-{d[5]}XXX"
    if n == 5:
        return f"{d[0:3]}-{d[3:5]}-XXXX"
    if n == 4:
        return f"{d[0:3]}-{d[3]}X-XXXX"
    if n == 3:
        return f"{d}-XX-XXXX"
    return raw


def k401_parse_amount(text: str) -> str:
    m = K401_AMOUNT_SIGNED_RE.search(_clean(text))
    return m.group(0).replace(" ", "") if m else ""


K401_LABEL_STRIP_RE = re.compile(
    r'\b(?:EMPLOYEE|EMPLOYER|PARTICIPANT|MEMBER|CLAIMANT|'
    r'SALARY|COMPENSATION|CONTRIBUTION|DEFERRAL|MATCH|'
    r'GROSS|NET|TOTAL|AMOUNT|NUMBER|ID|SSN|TIN|TAX|'
    r'PLAN|COMPANY|PAGE|DATE|ADDRESS|STATE|CITY|ZIP|'
    r'PHONE|EMAIL|FAX|FOR|AND|THE|INC|LLC|CORP|CO|LTD)\b',
    re.I,
)


def k401_find_name_in_text(text: str) -> str:
    text = _clean(text)
    if not text or K401_AMOUNT_RE.fullmatch(text) or K401_SSN_RE.search(text):
        return ''

    def _ok(s):
        words = s.split()
        return len(words) >= 2 and not any(c.isdigit() for c in s)

    m = K401_NAME_RE_LASTFIRST.search(text)
    if m:
        candidate = f"{m.group(1)}, {m.group(2)}"
        if _ok(candidate.replace(",", "")):
            return candidate

    stripped = K401_LABEL_STRIP_RE.sub('', text)
    stripped = re.sub(r'\b\d[\d,.$%-]*\b', '', stripped)
    stripped = re.sub(r'\s+', ' ', stripped).strip()
    m = K401_NAME_RE_UPPER.search(stripped)
    if m:
        candidate = m.group(1)
        if _ok(candidate):
            return candidate

    m = K401_NAME_RE.search(text)
    if m:
        candidate = m.group(1)
        if _ok(candidate):
            return candidate

    return ''


def k401_ssn_from_cell(cell_text: str) -> str:
    cell_text = _clean(cell_text)
    if not cell_text:
        return ''
    m = K401_SSN_RE.search(cell_text)
    if m:
        return k401_fmt_ssn(m.group(0))
    m = K401_SSN_RE_XPAD.search(cell_text)
    if m:
        return k401_fmt_ssn(m.group(0))
    for pat, kind in K401_SSN_PARTIAL_PATTERNS:
        pm = pat.search(cell_text)
        if pm:
            if kind == "short_last":
                partial = f"{pm.group(1)}-{pm.group(2)}-{pm.group(3)}"
            elif kind == "front5":
                partial = f"{pm.group(1)}-{pm.group(2)}"
            elif kind == "last4":
                partial = f"XXX-XX-{pm.group(1)}"
            elif kind == "mid_masked":
                partial = f"{pm.group(1)}-XX-{pm.group(2)}"
            else:
                partial = pm.group(0)
            return k401_fmt_ssn(partial)
    return ''


def k401_ssn_real_digits(ssn: str) -> int:
    return sum(1 for c in ssn if c.isdigit())


def k401_drop_redacted_duplicates(records: list) -> list:
    if len(records) <= 1:
        return records
    prefix_groups = defaultdict(list)
    for i, rec in enumerate(records):
        if not rec.ssn:
            continue
        m = re.match(r'^(\d{3}-\d{2})-', rec.ssn)
        if m:
            prefix_groups[m.group(1)].append(i)

    drop = set()
    for prefix, idxs in prefix_groups.items():
        if len(idxs) < 2:
            continue
        full_idxs = [i for i in idxs if k401_ssn_real_digits(records[i].ssn) == 9]
        if not full_idxs:
            continue
        for i in idxs:
            if i not in full_idxs:
                drop.add(i)
    return [r for i, r in enumerate(records) if i not in drop]


def k401_drop_exact_duplicates(records: list) -> list:
    seen = set()
    out = []
    for rec in records:
        fn = rec.first_name.strip().lower()
        ln = rec.last_name.strip().lower()
        ssn = rec.ssn.strip()
        key = (fn, ln, ssn)
        if ssn or (fn and ln):
            if key in seen:
                continue
            seen.add(key)
        out.append(rec)
    return out


def k401_first_name_prefix_match(a: str, b: str) -> bool:
    a = a.strip().lower()
    b = b.strip().lower()
    if not a or not b:
        return False
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    return longer.startswith(shorter)


def k401_merge_exact_duplicates(records: list) -> list:
    if len(records) <= 1:
        return records
    out = []
    for rec in records:
        fn = rec.first_name.strip()
        ln = rec.last_name.strip().lower()
        ssn = rec.ssn.strip()
        if not (fn and ln and ssn):
            out.append(rec)
            continue
        merged = False
        for primary in out:
            if primary.last_name.strip().lower() != ln:
                continue
            if primary.ssn.strip() != ssn:
                continue
            if not k401_first_name_prefix_match(fn, primary.first_name):
                continue
            if len(fn.strip()) > len(primary.first_name.strip()):
                primary.first_name = fn.strip()
                if rec.middle_name:
                    primary.middle_name = rec.middle_name
                if rec.suffix:
                    primary.suffix = rec.suffix
            else:
                if not primary.middle_name and rec.middle_name:
                    primary.middle_name = rec.middle_name
                if not primary.suffix and rec.suffix:
                    primary.suffix = rec.suffix
            merged = True
            break
        if not merged:
            out.append(rec)
    return out


def k401_parse_name_parts(full_name: str):
    name = _clean(full_name)
    if not name:
        return ("", "", "", "")
    if "," not in name:
        return ("", "", name, "")
    last_part, _, rest_part = name.partition(",")
    last = last_part.strip()
    rest = rest_part.strip().split()
    suffix = ""
    if rest and rest[-1].lower().rstrip(".") in K401_SUFFIXES:
        suffix = rest.pop()
    first = rest[0] if rest else ""
    middle = " ".join(w for w in rest[1:]) if len(rest) > 1 else ""
    return (first, middle, last, suffix)


class K401EmployeeRecord:
    __slots__ = (
        "source", "page", "doc_id", "first_name", "middle_name", "last_name", "suffix",
        "_full_name", "ssn", "gross_salary", "ee_contribution", "er_match", "ytd_contribution", "ee_id",
    )

    def __init__(self, source="", page="", doc_id=""):
        self.source = source
        self.page = page
        self.doc_id = doc_id
        self.first_name = ""
        self.middle_name = ""
        self.last_name = ""
        self.suffix = ""
        self._full_name = ""
        self.ssn = ""
        self.gross_salary = ""
        self.ee_contribution = ""
        self.er_match = ""
        self.ytd_contribution = ""
        self.ee_id = ""

    def set_field(self, canonical: str, value: str) -> None:
        value = _clean(value)
        if not value:
            return
        if canonical == "name" and not self._full_name:
            self._full_name = value
            first, middle, last, suffix = k401_parse_name_parts(value)
            if not self.first_name: self.first_name = first
            if not self.middle_name: self.middle_name = middle
            if not self.last_name: self.last_name = last
            if not self.suffix: self.suffix = suffix
        elif canonical == "first_name" and not self.first_name:
            self.first_name = value
        elif canonical == "middle_name" and not self.middle_name:
            self.middle_name = value
        elif canonical == "last_name" and not self.last_name:
            self.last_name = value
        elif canonical == "suffix" and not self.suffix:
            self.suffix = value
        elif canonical == "ssn" and not self.ssn:
            self.ssn = k401_fmt_ssn(value)
        elif canonical == "gross_salary" and not self.gross_salary:
            self.gross_salary = k401_parse_amount(value) or value
        elif canonical == "ee_contribution" and not self.ee_contribution:
            self.ee_contribution = k401_parse_amount(value) or value
        elif canonical == "er_match" and not self.er_match:
            self.er_match = k401_parse_amount(value) or value
        elif canonical == "ytd_contribution" and not self.ytd_contribution:
            self.ytd_contribution = k401_parse_amount(value) or value
        elif canonical == "ee_id" and not self.ee_id:
            self.ee_id = value

    def _work_info(self) -> str:
        parts = []
        if self.gross_salary:
            parts.append(f"Salary or Compensation Information: {self.gross_salary}")
        if self.ee_contribution:
            parts.append(f"Employee Contribution: {self.ee_contribution}")
        if self.er_match:
            parts.append(f"Employer Match: {self.er_match}")
        if self.ytd_contribution:
            parts.append(f"YTD Contribution: {self.ytd_contribution}")
        return " | ".join(parts)

    def is_meaningful(self) -> bool:
        return bool(self.ssn or self.first_name or self.last_name or self._full_name)

    def to_template_row(self) -> list:
        row = [""] * K401_TEMPLATE_COL_COUNT
        row[K401_COL_DOCID - 1] = self.doc_id
        row[K401_COL_LAST_NAME - 1] = self.last_name
        row[K401_COL_FIRST_NAME - 1] = self.first_name
        row[K401_COL_MIDDLE_NAME - 1] = self.middle_name
        row[K401_COL_SUFFIX - 1] = self.suffix
        row[K401_COL_SUBJECT_TYPE - 1] = "Employee"
        if self.ssn:
            row[K401_COL_GOV_ID_TAG - 1] = "Social Security Number (SSN)"
            row[K401_COL_SSN - 1] = self.ssn
        work = self._work_info()
        if work:
            row[K401_COL_WORK_INFO - 1] = work
        if self.ee_id:
            row[K401_COL_EE_ID - 1] = self.ee_id
        return row


def k401_resolve(obj):
    return obj.get_object() if hasattr(obj, "get_object") else obj


def k401_extract_acroform(pdf_path: Path, doc_id: str):
    if not HAS_PYPDF:
        return []
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return []
    try:
        root = k401_resolve(reader.trailer["/Root"])
        acroform = k401_resolve(root["/AcroForm"])
        fields = acroform.get("/Fields") or []
    except Exception:
        return []

    raw_fields = []

    def _traverse(ref):
        try:
            f = k401_resolve(ref)
        except Exception:
            return
        t = _clean(str(f.get("/T") or ""))
        v = _clean(str(f.get("/V") or ""))
        if v and v not in ("/Off", "Off"):
            raw_fields.append((t, v))
        for kid in (f.get("/Kids") or []):
            _traverse(kid)

    for ref in fields:
        _traverse(ref)

    if not raw_fields:
        return []

    rec = K401EmployeeRecord(source=pdf_path.name, page="1", doc_id=doc_id)
    for fname, fval in raw_fields:
        canonical = k401_map_column(fname)
        if canonical:
            rec.set_field(canonical, fval)
        else:
            m = K401_SSN_RE.search(fval) or K401_SSN_RE_XPAD.search(fval)
            if m:
                rec.set_field("ssn", m.group(0))

    return [rec] if rec.is_meaningful() else []


def k401_extract_tables(pdf_path: Path, doc_id: str):
    if not HAS_PDFPLUMBER:
        return []
    records = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                tables = page.extract_tables()
            except Exception:
                continue

            for tbl in (tables or []):
                if not tbl:
                    continue

                col_roles = {}
                for row_i, row in enumerate(tbl):
                    tentative = {}
                    for col_i, cell in enumerate(row):
                        canonical = k401_map_column(str(cell or ""))
                        if canonical:
                            tentative[col_i] = canonical
                    if len(tentative) >= 2:
                        col_roles = tentative
                        break

                for row_idx, row in enumerate(tbl):
                    if not row or all(not _clean(str(c)) for c in row):
                        continue

                    ssn_val = ''
                    ssn_col = -1
                    for col_i, cell in enumerate(row):
                        found = k401_ssn_from_cell(str(cell or ""))
                        if found:
                            ssn_val = found
                            ssn_col = col_i
                            break

                    if not ssn_val:
                        continue

                    rec = K401EmployeeRecord(source=pdf_path.name, page=str(page_num), doc_id=doc_id)
                    rec.ssn = ssn_val

                    for col_i, canonical in col_roles.items():
                        if col_i == ssn_col or col_i >= len(row):
                            continue
                        rec.set_field(canonical, str(row[col_i] or ""))

                    if not rec._full_name and not rec.first_name:
                        ssn_cell_text = _clean(str(row[ssn_col] or ""))
                        name_part = re.sub(r'\d{3}-\d{2}-[\dxX*]{1,4}', '', ssn_cell_text, flags=re.I).strip()
                        name = k401_find_name_in_text(name_part)
                        if name:
                            rec.set_field("name", name)

                    if not rec._full_name and not rec.first_name:
                        for col_i, cell in enumerate(row):
                            if col_i == ssn_col:
                                continue
                            cell_text = _clean(str(cell or ""))
                            if not cell_text:
                                continue
                            name = k401_find_name_in_text(cell_text)
                            if name:
                                rec.set_field("name", name)
                                break

                    if not rec._full_name and not rec.first_name:
                        search_rows = []
                        for delta in (-1, -2, -3, 1):
                            ri = row_idx + delta
                            if 0 <= ri < len(tbl):
                                search_rows.append(tbl[ri])
                        for nearby_row in search_rows:
                            if rec._full_name or rec.first_name:
                                break
                            for col_i, cell in enumerate(nearby_row):
                                cell_text = _clean(str(cell or ""))
                                name = k401_find_name_in_text(cell_text)
                                if name:
                                    rec.set_field("name", name)
                                    break

                    if rec.is_meaningful():
                        records.append(rec)

    return records


K401_LABEL_PATS = {
    "name": re.compile(r'(?:employee|participant|emp(?:loyee)?)\s+name\s*[:\-]?\s*([A-Za-z][A-Za-z ,\.\'\-]+)', re.I),
    "first_name": re.compile(r'first\s+name\s*[:\-]?\s*(.+)', re.I),
    "last_name": re.compile(r'last\s+name\s*[:\-]?\s*(.+)', re.I),
    "middle_name": re.compile(r'middle\s+(?:name|initial)\s*[:\-]?\s*(.+)', re.I),
    "suffix": re.compile(r'(?:name\s+)?suffix\s*[:\-]?\s*(.+)', re.I),
    "ssn": re.compile(r'(?:ssn|ss#|social\s+security(?:\s+number)?|tax\s+id)\s*[:\-#]?\s*([\d\- ]+)', re.I),
    "gross_salary": re.compile(r'(?:gross\s+(?:salary|pay|wages?|comp(?:ensation)?)|annual\s+salary|salary)\s*[:\-]?\s*((?=[\$\(\-]*\d)[\$\d,\.\(\)\-]+)', re.I),
    "ee_contribution": re.compile(r'(?:ee\s+(?:contribution|deferral)|employee\s+(?:contribution|deferral|401k)|elective\s+deferral)\s*[:\-]?\s*((?=[\$\(\-]*\d)[\$\d,\.\(\)\-]+)', re.I),
    "er_match": re.compile(r'(?:er\s+match|employer\s+(?:match|contribution)|company\s+match)\s*[:\-]?\s*((?=[\$\(\-]*\d)[\$\d,\.\(\)\-]+)', re.I),
    "ytd_contribution": re.compile(r'(?:ytd\s+(?:contribution|total|deferral)|year\s+to\s+date)\s*[:\-]?\s*((?=[\$\(\-]*\d)[\$\d,\.\(\)\-]+)', re.I),
    "ee_id": re.compile(r'(?:employee\s+(?:id|number|identification)|ee\s+id|emp\s+(?:id|number))\s*[:\-#]?\s*([A-Z0-9\-]+)', re.I),
}


def k401_collect_ssn_positions(lines):
    positions = []
    for line_i, line in enumerate(lines):
        found_on_line = False
        for m in K401_SSN_RE.finditer(line):
            positions.append((line_i, k401_fmt_ssn(m.group(0))))
            found_on_line = True
        if not found_on_line:
            for m in K401_SSN_RE_XPAD.finditer(line):
                positions.append((line_i, k401_fmt_ssn(m.group(0))))
                found_on_line = True
        if not found_on_line:
            for pat, kind in K401_SSN_PARTIAL_PATTERNS:
                pm = pat.search(line)
                if pm:
                    if kind == "short_last":
                        partial = f"{pm.group(1)}-{pm.group(2)}-{pm.group(3)}"
                    elif kind == "front5":
                        partial = f"{pm.group(1)}-{pm.group(2)}"
                    elif kind == "last4":
                        partial = f"XXX-XX-{pm.group(1)}"
                    elif kind == "mid_masked":
                        partial = f"{pm.group(1)}-XX-{pm.group(2)}"
                    else:
                        partial = pm.group(0)
                    positions.append((line_i, k401_fmt_ssn(partial)))
                    break
    return positions


def k401_extract_regex(pdf_path: Path, doc_id: str):
    if not HAS_PDFPLUMBER:
        return []
    records = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            lines = text.splitlines()
            ssn_positions = k401_collect_ssn_positions(lines)

            if not ssn_positions:
                rec = K401EmployeeRecord(source=pdf_path.name, page=str(page_num), doc_id=doc_id)
                for canonical, pat in K401_LABEL_PATS.items():
                    for line in lines:
                        m = pat.search(line)
                        if m:
                            rec.set_field(canonical, m.group(1))
                            break
                if rec.is_meaningful():
                    records.append(rec)
                continue

            for ssn_line_i, ssn_val in ssn_positions:
                rec = K401EmployeeRecord(source=pdf_path.name, page=str(page_num), doc_id=doc_id)
                rec.ssn = ssn_val

                raw_line = lines[ssn_line_i]
                cleaned_line = re.sub(r'\d{3}-\d{2}-[\dxX*]{1,4}', '', raw_line, flags=re.I).strip()
                name = k401_find_name_in_text(cleaned_line)
                if name:
                    rec.set_field("name", name)

                # Only lines closest to THIS employee's SSN belong to this record (labels after an SSN lean
                # toward the preceding SSN), so a second employee never inherits the first one's name/ID/amounts.
                ssn_idx = [i for i, _ in ssn_positions]

                def owner(j, ssn_idx=ssn_idx):
                    return min(ssn_idx, key=lambda s: (j - s) * 0.5 if j >= s else (s - j))

                w_start = max(0, ssn_line_i - 20)
                w_end = min(len(lines), ssn_line_i + 20)
                window = "\n".join(lines[j] for j in range(w_start, w_end) if owner(j) == ssn_line_i)

                _numeric_canonicals = {"gross_salary", "ee_contribution", "er_match", "ytd_contribution"}
                for canonical, pat in K401_LABEL_PATS.items():
                    if canonical == "ssn":
                        continue
                    m2 = pat.search(window)
                    if m2:
                        candidate = _clean(m2.group(1))
                        if K401_SSN_RE.search(candidate):
                            continue
                        # Numeric fields are SUPPOSED to look like an amount; only the
                        # name-like fields need the "doesn't look like an amount" guard
                        # (keeps a stray dollar figure from being mistaken for a name).
                        if canonical not in _numeric_canonicals and K401_AMOUNT_RE.fullmatch(candidate):
                            continue
                        rec.set_field(canonical, candidate[:80])

                if not rec._full_name and not rec.first_name:
                    for offset in (-1, 1, -2, 2, -3, 3, -4, 4, -5, 5, -6, 6, -7, 7, -8, 8, -9, 9, -10, 10):
                        check_i = ssn_line_i + offset
                        if not (0 <= check_i < len(lines)) or owner(check_i) != ssn_line_i:
                            continue
                        check_line = re.sub(r'\d{3}-\d{2}-[\dxX*]{1,4}', '', lines[check_i], flags=re.I).strip()
                        name = k401_find_name_in_text(check_line)
                        if name:
                            rec.set_field("name", name)
                            break

                if rec.is_meaningful():
                    records.append(rec)

    return records


def k401_extract(pdf_path: Path, doc_id: str):
    records = k401_extract_acroform(pdf_path, doc_id)
    if records:
        records = k401_drop_redacted_duplicates(records)
        records = k401_drop_exact_duplicates(records)
        records = k401_merge_exact_duplicates(records)
        print(f"    [AcroForm] {len(records)} employee record(s) found.")
        return records

    tbl_records = k401_extract_tables(pdf_path, doc_id)
    text_records = k401_extract_regex(pdf_path, doc_id)
    records = k401_merge_exact_duplicates(k401_drop_exact_duplicates(k401_drop_redacted_duplicates(tbl_records + text_records)))

    method_parts = []
    if tbl_records:
        method_parts.append(f"Table:{len(tbl_records)}")
    if text_records:
        method_parts.append(f"Text:{len(text_records)}")

    if records:
        method_str = " + ".join(method_parts) if method_parts else "no-header scan"
        print(f"    [{method_str}] {len(records)} employee record(s) found.")
    else:
        print("    [!] No employee records found -- check PDF layout.")

    return records


def sniff_pdf_text(pdf_path: Path, max_pages: int = 3):
    """Return (lowercased sampled text, has_text_layer) for format detection."""
    if not HAS_FITZ:
        return "", False
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return "", False
    parts = []
    total_chars = 0
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        t = page.get_text()
        parts.append(t)
        total_chars += len(t.strip())
    doc.close()
    return "\n".join(parts).replace("\u00a0", " ").replace("\u00ad", "-").lower(), total_chars >= 40


def _sniff_w2(text: str, has_text: bool) -> bool:
    return bool(
        "wage and tax statement" in text
        or "wages, tips, other comp" in text
        or re.search(r"employee.?s social security number", text)
    )


def _sniff_1095c(text: str, has_text: bool) -> bool:
    return "1095-c" in text or "1095c" in text or "offer and coverage" in text


def _sniff_401k(text: str, has_text: bool) -> bool:
    if "gross to net" in text or "earnings statement" in text:
        return False
    return bool(
        re.search(r"401\s*\(?k\)?", text)
        and ("contribution report" in text or "employer match" in text or "elective deferral" in text)
    )


# FORMAT_SNIFFERS is defined after the 1099 sniffer further down.


def detect_format_for_file(pdf_path: Path, selected_formats: list):
    """Return the first supported form (in FORMAT_PRIORITY order) whose identifiers appear on the first
    pages of this PDF, or None."""
    text, has_text = sniff_pdf_text(pdf_path)
    for fmt in FORMAT_PRIORITY:
        if fmt in selected_formats and FORMAT_SNIFFERS[fmt](text, has_text):
            return fmt
    return None


# ===========================================================================
# Form 1099 and payroll statements (label-then-value scanning on page text)
# Ported from extract_1099.py / extract_adp.py / extract_w2.py.
# These label spellings are first-pass placeholders -- confirm against a real
# sample of each form before trusting the output.
# ===========================================================================

LBL_SSN_RE = re.compile(r"(?:\d{3}-\d{2}-\d{4}|[Xx\*]{3}-[Xx\*]{2}-\d{4}|\d{9})")
LBL_EIN_RE = re.compile(r"\d{2}-\d{7}")
LBL_TIN_RE = re.compile(r"(?:\d{2}-\d{7}|" + LBL_SSN_RE.pattern + r")")
LBL_FREE_TEXT_RE = re.compile(r"[A-Za-z0-9][A-Za-z0-9,.'&\-]*(?: [A-Za-z0-9,.'&\-]+)*")
LBL_ID_RE = re.compile(r"(?!\d{1,2}(?:\s|$))(?=[A-Za-z0-9\-]*\d)[A-Za-z0-9\-]+")
LBL_MONEY_RE = re.compile(r"\(?-?\$?\s?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d+\)?|\(?-?\$?\s?\d{1,3}(?:,\d{3})+\)?")
LBL_DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}|\d{4}-\d{2}-\d{2}|[A-Za-z]{3,9}\.? \d{1,2}, \d{4}")
LBL_HOURS_RE = re.compile(r"\d[\d,]*\.\d+")
LBL_YEAR_RE = re.compile(r"20\d{2}")

# unified column name -> (label patterns, value shape)
F1099_FIELDS = {
    "Tax Year":                  ([r"Form\s+1099-?(?:NEC|MISC)"], LBL_YEAR_RE),
    "Employer TIN":              ([r"PAYER'?S TIN"], LBL_TIN_RE),
    "SSN":                       ([r"RECIPIENT'?S TIN"], LBL_TIN_RE),
    "Account Number":            ([r"[Aa]ccount number(?:\s*\(see instructions\))?"], LBL_ID_RE),
    "1099 Box 1 Nonemployee Comp": ([r"\b1\b\s*Nonemployee compensation"], LBL_MONEY_RE),
    "Federal Tax Withheld":      ([r"\b4\b\s*Federal income tax withheld"], LBL_MONEY_RE),
    "State Tax Withheld":        ([r"\b5\b\s*State tax withheld"], LBL_MONEY_RE),
    "1099 Box 6 State/Payer State No": ([r"\b6\b\s*State/Payer'?s state no"], LBL_ID_RE),
    "1099 Box 7 State Income":   ([r"\b7\b\s*State income"], LBL_MONEY_RE),
}


W2_EMPLOYER_FIELDS = {
    "Tax Year":      ([r"Form\s+W-?2", r"Wage and Tax Statement"], LBL_YEAR_RE),
    "Employer TIN":  ([r"[Ee]mployer identification number", r"\bb\b\s*Employer"], LBL_EIN_RE),
}

F1099_PAYER_BLOCK = [r"PAYER'?S name, street address[^\n]*"]
F1099_RECIPIENT_BLOCK = [r"RECIPIENT'?S name[^\n]*"]
W2_EMPLOYER_BLOCK = [r"[Ee]mployer'?s name, address[^\n]*"]


def label_extract_field(text: str, label_patterns, value_re, max_gap: int = 120) -> str:
    """Find the first label spelling in `text`; return the value that follows it on the same line, or --
    when the label ends its line -- a value that starts the next non-empty line. Never reaches further,
    so an empty field cannot pick up a neighbouring label's value."""
    for pattern in label_patterns:
        label_match = re.search(pattern, text)
        if not label_match:
            continue
        rest = text[label_match.end():]
        same_line = rest.split("\n", 1)[0].lstrip(" :\t")
        value_match = value_re.search(same_line[:max_gap])
        if value_match:
            return value_match.group(0).strip()
        for nxt in rest.split("\n")[1:4]:
            nxt = nxt.strip()
            if not nxt:
                continue
            value_match = value_re.match(nxt)
            return value_match.group(0).strip() if value_match else ""
    return ""


LBL_SKIP_CAPTION_RE = re.compile(
    r"(?i)\b(?:street address|city or town|state or province|foreign postal|telephone no|apt\. no|"
    r"including (?:apt|room)|see instructions|and zip|or foreign|postal code)\b")
LBL_CAPTION_LINE_RE = re.compile(
    r"(?i)\b(?:tin|number|ein|copy|form|compensation|wages|tips|withheld|income|account|payer|recipient|"
    r"employer|employee|federal|state|zip|control|social|medicare|nonemployee|box)\b")


def label_extract_block(text: str, label_patterns, max_lines: int = 5) -> list:
    """Lines of the name/address block that follows a caption: skips wrapped caption lines
    ("Street address (including apt. no.)", "City or town, state or province ..."), stops at the next
    field caption, and returns at most max_lines value lines."""
    for pattern in label_patterns:
        m = re.search(pattern, text)
        if not m:
            continue
        block = []
        for line in text[m.end():].split("\n"):
            line = _clean(line)
            if not line:
                continue
            if LBL_SKIP_CAPTION_RE.search(line):
                continue
            if LBL_CAPTION_LINE_RE.search(line):
                break
            block.append(line)
            if len(block) == max_lines:
                break
        if block:
            return block
    return []


def parse_name_address_block(block: list) -> dict:
    """[name, street, 'City, ST 12345'] -> Full Name / Street Address / City / State / Zip Code."""
    out = {}
    if not block:
        return out
    out["Full Name"] = block[0]
    rest = block[1:]
    for i, line in enumerate(rest):  # the first "City, ST 12345" line ends the address (a phone or id may follow)
        m = W2_CITY_STATE_ZIP_RE.match(line)
        if m:
            out.update({"City": m.group("city").rstrip(","), "State": m.group("state"), "Zip Code": m.group("zip")})
            rest = rest[:i]
            break
    if rest:
        out["Street Address"] = ", ".join(rest)
    return out


def label_extract_record(text: str, fields: dict) -> dict:
    return {col: label_extract_field(text, labels, vre) for col, (labels, vre) in fields.items()}


def page_texts(pdf_path: Path):
    """Yield (page_number, text) for each page's text layer (no OCR)."""
    doc = fitz.open(pdf_path)
    try:
        for i, page in enumerate(doc, start=1):
            yield i, page.get_text()
    finally:
        doc.close()


# ===========================================================================
# Unified output -- one CSV per PDF, one schema for every document type.
# Columns: File Name, Page Number, the PII/PHI identifiers, then the
# employer, pay-period, wage/tax/hour and payroll-change values.
# ===========================================================================

PII_COLUMNS = [
    "File Name", "Page Number",
    "First Name", "Middle Name", "Last Name", "Suffix", "Full Name",
    "Street Address", "City", "State", "Zip Code", "Country",
    "SSN", "Employee ID", "Account Number",
    "Health Plan ID", "Member #", "Claim #", "Patient Acct #",
    "Service Date", "CPT Code", "Date of Birth",
    # employer details
    "Tax Year", "Employer Name", "Employer Address", "Employer TIN",
    "Position ID", "Home Department", "Home Cost Number",
    # pay period
    "Pay Date", "Period Begin", "Period End",
    # wages / pay amounts
    "Gross Pay", "Net Pay", "Total Earnings", "Total Deductions", "Total Taxes",
    "Gross Salary", "Employee Contribution", "Employer Match", "YTD Contribution",
    "1099 Box 1 Nonemployee Comp", "1099 Box 6 State/Payer State No", "1099 Box 7 State Income",
    # taxes withheld
    "Federal Tax Withheld", "State Tax Withheld", "Social Security Withheld", "Medicare Withheld",
    "Box 1 Wages Tips Other Comp", "Box 2 Federal Tax Withheld", "Box 3 SS Wages",
    "Box 4 SS Tax Withheld", "Box 5 Medicare Wages", "Box 6 Medicare Tax Withheld",
    # hours
    "Total Hrs/Units", "Regular Hours", "Overtime Hours",
    # payroll changes
    "Changed Field", "Changed From", "Changed To",
]


NAME_SUFFIXES = {"jr": "Jr", "sr": "Sr", "ii": "II", "iii": "III", "iv": "IV", "v": "V"}
NAME_FILLER_WORDS = {"medi-cal", "medicare", "commercial", "medi-medi", "payee", "payee:"}
DATE_OF_BIRTH_COLUMNS = {"Date of Birth"}


def normalize_ssn(value) -> str:
    """9 digits -> ###-##-####; masked values -> XXX-XX-####. An EIN-shaped value (##-#######) is left as is."""
    text = _clean(value)
    if re.fullmatch(r"\d{2}-\d{7}", text):
        return text
    chars = re.sub(r"[^0-9Xx*]", "", text)
    if len(chars) == 9:
        chars = chars.replace("*", "X").replace("x", "X")
        return f"{chars[:3]}-{chars[3:5]}-{chars[5:]}"
    return text


def normalize_date(value) -> str:
    """Date of Birth -> MM/DD/YYYY; unparseable text is left unchanged."""
    text = _clean(value)
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%b %d, %Y", "%B %d, %Y", "%b. %d, %Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    return text


def split_full_name(name: str):
    """'Doe, Jane Q', 'Doe Jr, John', 'Van Der Berg, Hans III' -> name parts as printed, or None when there is no comma."""
    name = _clean(name)
    if "," not in name:
        return None
    last_part, _, rest = name.partition(",")
    last_tokens, rest_tokens = last_part.split(), rest.split()
    for i, tok in enumerate(rest_tokens):  # drop trailing plan/provider words ("Medicare", "Payee: ACME")
        if tok.lower() in NAME_FILLER_WORDS:
            rest_tokens = rest_tokens[:i]
            break
    suffix = ""
    if rest_tokens and rest_tokens[-1].strip(".").lower() in NAME_SUFFIXES:
        suffix = rest_tokens.pop()
    elif len(last_tokens) > 1 and last_tokens[-1].strip(".").lower() in NAME_SUFFIXES:
        suffix = last_tokens.pop()
    if not rest_tokens or not last_tokens:
        return None
    return {"Last Name": " ".join(last_tokens), "First Name": rest_tokens[0],
            "Middle Name": " ".join(rest_tokens[1:]), "Suffix": suffix}


def finalize_row(row: dict) -> dict:
    """Only SSN and Date of Birth are reformatted. Everything else stays exactly as the document prints it;
    'Last, First M' is only split across columns (the text itself is unchanged)."""
    if row.get("Full Name") and not row.get("Last Name"):
        parts = split_full_name(row["Full Name"])
        if parts:
            row.pop("Full Name")
            row.update({k: v for k, v in parts.items() if v})
    # "First Middle Last Suffix" read as First/Middle/Last (1095-C): move a trailing suffix word out of Last
    last = row.get("Last Name", "")
    if last.strip(".").lower() in NAME_SUFFIXES and not row.get("Suffix"):
        row["Suffix"] = last
        tokens = row.get("Middle Name", "").split()
        if tokens:
            row["Last Name"], row["Middle Name"] = tokens[-1], " ".join(tokens[:-1])
        else:
            row.pop("Last Name")
    if row.get("SSN"):
        row["SSN"] = normalize_ssn(row["SSN"])
    for col in DATE_OF_BIRTH_COLUMNS:
        if row.get(col):
            row[col] = normalize_date(row[col])
    for col in ("Patient Acct #", "Claim #", "Member #", "CPT Code"):
        if row.get(col):
            row[col] = row[col].replace(":", ";")  # merged-value delimiter written by the claims engines
    return {k: v for k, v in row.items() if v}


def format_address(street, city, state, zip_code) -> str:
    """'street, City, ST 12345' from whichever parts are present."""
    locality = " ".join(x for x in (f"{city}," if city and (state or zip_code) else city, state, zip_code) if x)
    return ", ".join(x for x in (street, locality) if x)


def _row(**fields) -> dict:
    out = {}
    for k, v in fields.items():
        if k in PII_COLUMNS and (text := _clean(v)):
            out[k] = text
    return finalize_row(out)


def w2_column_of_label(lines, label_re):
    """Find the caption row for a box and return (line index, x where the box column starts, x where the
    next box on that row starts, index of the first word after the caption) -- or None."""
    for li, ln in enumerate(lines):
        words = [t for _, _, t in ln]
        text = " ".join(words)
        m = label_re.search(text)
        if not m:
            continue
        starts, off = [], 0
        for w in words:
            starts.append(off)
            off += len(w) + 1
        wi = max(k for k in range(len(starts)) if starts[k] <= m.start())
        end_wi = next((k for k in range(len(starts)) if starts[k] >= m.end()), len(words))
        x_start = ln[wi][0]
        if wi > 0 and re.fullmatch(r"[a-f]", words[wi - 1]):
            x_start = ln[wi - 1][0]
        x_end = float("inf")
        for k in range(end_wi, len(words) - 1):
            if W2_BOX_NUMBER_RE.match(words[k]) and words[k + 1][:1].isalpha():
                x_end = ln[k][0]
                break
        return li, x_start, x_end, end_wi
    return None


def w2_employer_from_lines(lines) -> dict:
    """Employer EIN, name and address from the 'b' and 'c' boxes, reading only the words inside each box's
    own column (boxes sit side by side, so whole text rows mix in the neighbouring boxes)."""
    out = {}
    pos = w2_column_of_label(lines, re.compile(r"employer identification number", re.IGNORECASE))
    if pos:
        li, xs, xe, ewi = pos
        candidates = [w for w in lines[li][ewi:] if w[0] < xe]
        for below in lines[li + 1: li + 3]:
            candidates += [w for w in below if xs - 6 <= w[0] < xe]
        ein = next((t for _, _, t in candidates if re.fullmatch(r"\d{2}-\d{7}", t)), "")
        if ein:
            out["Employer TIN"] = ein
    pos = w2_column_of_label(lines, re.compile(r"employer.s name, address", re.IGNORECASE))
    if pos:
        li, xs, xe, _ = pos
        block = []
        for below in lines[li + 1: li + 10]:
            col = [w for w in below if xs - 6 <= w[0] < xe]
            if not col:
                continue
            text = normalize_text(" ".join(t for _, _, t in col))
            if re.match(r"^[a-f]\s+[A-Z]", text) or w2_find_right_column_boundary_strict(col) == col[0][0]:
                break  # the next box's caption ("d Control number", "e Employee's name ...")
            if LBL_SKIP_CAPTION_RE.search(text):
                continue
            block.append(text)
            if len(block) == 5:
                break
        parsed = parse_name_address_block(block)
        if parsed.get("Full Name"):
            out["Employer Name"] = parsed["Full Name"]
            addr = format_address(parsed.get("Street Address"), parsed.get("City"), parsed.get("State"),
                                  parsed.get("Zip Code"))
            if addr:
                out["Employer Address"] = addr
    return out


def extract_w2_rows(pdf: Path) -> list:
    employer_by_page = {}
    doc = fitz.open(pdf)
    try:
        for page_num, page in enumerate(doc, start=1):
            text = page.get_text()
            emp = label_extract_record(text, W2_EMPLOYER_FIELDS)
            blk = parse_name_address_block(label_extract_block(text, W2_EMPLOYER_BLOCK))
            if blk.get("Full Name"):
                emp["Employer Name"] = blk["Full Name"]
                addr = format_address(blk.get("Street Address"), blk.get("City"), blk.get("State"), blk.get("Zip Code"))
                if addr:
                    emp["Employer Address"] = addr
            if not (emp.get("Employer Name") and emp.get("Employer TIN")):
                by_position = w2_employer_from_lines(group_words_into_lines(page.get_text("words")))
                for key, value in by_position.items():
                    if value and not emp.get(key):
                        emp[key] = value
            employer_by_page[page_num] = emp
    finally:
        doc.close()
    recs = w2_dedupe_records(w2_process_pdf(pdf, 0.5, True))
    rows = []
    for r in recs:
        emp = employer_by_page.get(r.get("Page"), {})
        rows.append(_row(**{**emp, "Page Number": r.get("Page"), "First Name": r.get("First Name"),
                            "Last Name": r.get("Last Name"), "Suffix": r.get("Suffix"),
                            "Street Address": r.get("Street Address"), "City": r.get("City"),
                            "State": r.get("State"), "Zip Code": r.get("Zip Code"), "SSN": r.get("SSN"),
                            **{c: r.get(c) for c in W2_WAGE_COLUMNS}}))
    return rows


def extract_1095c_rows(pdf: Path) -> list:
    recs = c1095_dedupe_complete_rows(c1095_process_pdf(pdf, debug=False, output_dir=pdf.parent))
    return [_row(**{"Page Number": r.get("Page"), "First Name": r.get("First Name"),
                    "Middle Name": r.get("Middle Name"), "Last Name": r.get("Last Name"),
                    "Street Address": r.get("Street Address"), "City": r.get("City"),
                    "State": r.get("State"), "Zip Code": r.get("Zip Code"),
                    "Country": r.get("Country"), "SSN": r.get("SSN")})
            for r in recs]


def extract_1099_rows(pdf: Path) -> list:
    rows = []
    for page_num, text in page_texts(pdf):
        vals = label_extract_record(text, F1099_FIELDS)
        recipient = parse_name_address_block(label_extract_block(text, F1099_RECIPIENT_BLOCK))
        payer = parse_name_address_block(label_extract_block(text, F1099_PAYER_BLOCK))
        if payer.get("Full Name"):
            vals["Employer Name"] = payer["Full Name"]
            vals["Employer Address"] = format_address(
                payer.get("Street Address"), payer.get("City"), payer.get("State"), payer.get("Zip Code"))
        vals.update(recipient)
        if not any(v for k, v in vals.items() if k != "Tax Year"):
            continue
        rows.append(_row(**{"Page Number": page_num, **vals}))
    return rows


def extract_401k_rows(pdf: Path) -> list:
    def _name(r):
        if not r.first_name and " " in r.last_name:  # "First Last" without a comma: don't guess the split
            return {"Full Name": r.last_name}
        return {"First Name": r.first_name, "Middle Name": r.middle_name, "Last Name": r.last_name}
    return [_row(**{"Page Number": r.page, **_name(r), "Suffix": r.suffix, "SSN": r.ssn, "Employee ID": r.ee_id,
                    "Gross Salary": r.gross_salary, "Employee Contribution": r.ee_contribution,
                    "Employer Match": r.er_match, "YTD Contribution": r.ytd_contribution})
            for r in k401_extract(pdf, pdf.stem)]


EXTRACTORS = {
    "1095c": extract_1095c_rows,
    "w2": extract_w2_rows,
    "1099": extract_1099_rows,
    "401k": extract_401k_rows,
}


def _sniff_1099(text: str, has_text: bool) -> bool:
    return bool(re.search(r"1099-?(?:nec|misc)", text) or "nonemployee compensation" in text)


FORMAT_CHOICES = [
    ("1095c", "Form 1095-C"),
    ("w2", "W-2"),
    ("1099", "Form 1099-NEC / 1099-MISC"),
    ("401k", "401k contribution report"),
]
FORMAT_LABELS = dict(FORMAT_CHOICES)
FORMAT_PRIORITY = [key for key, _ in FORMAT_CHOICES]
FORMAT_SNIFFERS = {"1095c": _sniff_1095c, "w2": _sniff_w2, "1099": _sniff_1099, "401k": _sniff_401k}


# ===========================================================================
# CLI
# ===========================================================================

def dedupe_rows(rows: list) -> list:
    """Drop rows identical in every column except Page Number (Copy B/C duplicates, repeated pages)."""
    seen, kept = set(), []
    for row in rows:
        key = tuple(sorted((k, v) for k, v in row.items() if k != "Page Number"))
        if key in seen:
            continue
        seen.add(key)
        kept.append(row)
    return kept


def _write_text_xlsx(headers: list, matrix: list, out_path: Path) -> None:
    """Write a header row plus rows of strings; every cell is stored as Excel Text ("@")."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted"
    bold = Font(bold=True)
    widths = [len(h) for h in headers]
    for ci, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=ci, value=header)
        cell.font = bold
        cell.number_format = "@"
    for ri, values in enumerate(matrix, start=2):
        for ci, text in enumerate(values, start=1):
            cell = ws.cell(row=ri, column=ci, value=text)
            cell.data_type = "s"  # never a formula, even when the text starts with "=" or "-"
            cell.number_format = "@"
            if len(text) > widths[ci - 1]:
                widths[ci - 1] = len(text)
    for ci, width in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = min(width + 2, 60)
    ws.freeze_panes = "A2"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return headers


def write_pii_xlsx(rows: list, out_path: Path) -> None:
    """Write rows with only the columns that have a value in at least one row (File Name always first),
    kept in PII_COLUMNS order. Every cell is stored as Text, so Excel never strips leading zeros or
    reformats numbers, dates or SSNs."""
    columns = [c for c in PII_COLUMNS if c == "File Name" or any(row.get(c) for row in rows)]
    return _write_text_xlsx(columns, [[str(row.get(c, "")) for c in columns] for row in rows], out_path)


# Used only as the last resort, in place of the plain rows/tables copy, and only when --ai is given.
# Each page's text layer plus a picture of the page is sent to YOUR Azure OpenAI deployment (Entra ID
# sign-in, no API key). Claims / patient-info documents are never sent (AI_PHI_FORMATS).
# Adapted from PDFWithTableConvertToExcelTool/doc_reader_v2.py.


SUMMARY_HEADERS = ["File Name", "Document Type", "Pages", "Time Taken (s)", "Records", "Extracted Columns",
                   "Status", "Output File", "Reason Not Extracted"]


def pdf_page_count(pdf: Path) -> str:
    try:
        with fitz.open(pdf) as doc:
            return str(doc.page_count)
    except Exception:
        return ""  # unreadable or password protected


def write_summary_xlsx(entries: list, out_path: Path) -> None:
    """Run summary (column names and counts only -- no extracted values): sheet 'Summary' lists every
    document; sheet 'Not Extracted' lists the ones that produced no output and why."""
    wb = openpyxl.Workbook()
    bold = Font(bold=True)

    def fill(ws, headers, matrix):
        widths = [len(h) for h in headers]
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=1, column=ci, value=h)
            c.font, c.number_format = bold, "@"
        for ri, values in enumerate(matrix, start=2):
            for ci, text in enumerate(values, start=1):
                c = ws.cell(row=ri, column=ci, value=text)
                c.data_type, c.number_format = "s", "@"
                widths[ci - 1] = max(widths[ci - 1], len(text))
        for ci, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = min(w + 2, 70)
        ws.freeze_panes = "A2"

    summary = wb.active
    summary.title = "Summary"
    fill(summary, SUMMARY_HEADERS, [[e.get(h, "") for h in SUMMARY_HEADERS] for e in entries])
    missed = [e for e in entries if e["Status"] != "Extracted"]
    not_extracted = wb.create_sheet("Not Extracted")
    fill(not_extracted, ["File Name", "Document Type", "Pages", "Reason Not Extracted"],
         [[e["File Name"], e.get("Document Type", ""), e.get("Pages", ""), e.get("Reason Not Extracted", "")]
          for e in missed] or [["(none)", "", "", ""]])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract W-2, 401k, 1095-C and 1099 PDFs into one <filename>_extracted.xlsx per PDF, "
                    "plus a run summary workbook.")
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("-o", "--output-dir", default=None,
                        help="Folder for the <filename>_extracted.xlsx files and the summary (default: next to the input)")
    parser.add_argument("--recursive", action="store_true", help="Recurse into subfolders")
    parser.add_argument("--format", default="auto", choices=["auto"] + FORMAT_PRIORITY,
                        help="Force one form type for every file instead of auto-detecting")
    args = parser.parse_args()

    if not HAS_FITZ:
        sys.exit("ERROR: pymupdf is required. Run: pip install pymupdf")
    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required. Run: pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")
    if not HAS_PYPDF:
        print("WARNING: pypdf not installed -- 401k AcroForm extraction skipped. Run: pip install pypdf")

    input_path = Path(args.input)
    if not input_path.exists():
        sys.exit(f"ERROR: path not found -- {input_path}")
    pdf_files = list(_iter_pdfs(input_path, args.recursive))
    if not pdf_files:
        sys.exit(f"ERROR: no .pdf files found at {input_path}")

    failed, written = [], 0
    summary = []
    for pdf in tqdm(pdf_files, desc="Extracting", unit="file"):
        t0 = time.perf_counter()
        entry = {"File Name": pdf.name, "Pages": pdf_page_count(pdf), "Records": "0", "Extracted Columns": "",
                 "Status": "Not extracted", "Output File": "", "Reason Not Extracted": "", "Document Type": ""}

        def finish(**fields):
            entry.update(fields)
            entry["Time Taken (s)"] = f"{time.perf_counter() - t0:.2f}"
            summary.append(entry)

        if not entry["Pages"]:
            failed.append(f"{pdf.name}: file could not be opened (corrupt or password protected)")
            finish(**{"Document Type": "Unreadable PDF",
                      "Reason Not Extracted": "File could not be opened (corrupt or password protected)"})
            continue

        fmt = args.format if args.format != "auto" else detect_format_for_file(pdf, FORMAT_PRIORITY)
        if fmt is None:
            finish(**{"Document Type": "Not a supported form",
                      "Reason Not Extracted": "Not a W-2, 401k, 1095-C or 1099 (or no usable text layer)"})
            continue

        try:
            rows = dedupe_rows([{"File Name": pdf.name, **r} for r in EXTRACTORS[fmt](pdf)])
        except Exception as exc:
            failed.append(f"{pdf.name}: {exc}")
            finish(**{"Document Type": FORMAT_LABELS[fmt], "Reason Not Extracted": f"Error: {exc}"})
            continue
        if not rows:
            print(f"  {pdf.name} -> [{FORMAT_LABELS[fmt]}] no records -- no file written")
            finish(**{"Document Type": FORMAT_LABELS[fmt],
                      "Reason Not Extracted": "Form recognised but no records could be read from it"})
            continue

        out_dir = Path(args.output_dir) if args.output_dir else pdf.parent
        out_path = out_dir / f"{pdf.stem}_extracted.xlsx"
        columns = write_pii_xlsx(rows, out_path)
        written += 1
        print(f"  {pdf.name} -> [{FORMAT_LABELS[fmt]}] {len(rows)} record(s) -> {out_path.name}")
        finish(**{"Document Type": FORMAT_LABELS[fmt], "Records": str(len(rows)),
                  "Extracted Columns": ", ".join(columns), "Status": "Extracted", "Output File": out_path.name})

    if written:
        print(f"\nDone. {written} Excel file(s) written.")
        print("Reminder: these files contain PII. Save them only in the appropriate Global Insider "
              "folder; Claude conversations are not a secure record.")
    else:
        print("\nNo records extracted. No file written.")

    # the run summary is always written
    folder = Path(args.output_dir) if args.output_dir else (input_path if input_path.is_dir() else input_path.parent)
    summary_path = folder / f"{datetime.now():%y%m%d} AS govt forms extraction summary.xlsx"
    write_summary_xlsx(summary, summary_path)
    missed = sum(1 for e in summary if e["Status"] != "Extracted")
    print(f"Summary: {len(summary)} document(s), {len(summary) - missed} extracted, {missed} not extracted "
          f"-> {summary_path.name}")

    if failed:
        print("\nErrors:")
        for e in failed:
            print(f"  [!] {e}")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
