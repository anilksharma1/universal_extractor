r"""
universal_pdf_extractor.py — single entry point for every document-specific
PDF extractor in this project, dispatched by subcommand:

    python universal_pdf_extractor.py w2            <input> [options]
    python universal_pdf_extractor.py 1095c          <input> <bde_template.xlsx> [options]
    python universal_pdf_extractor.py claims         <input> [options]
    python universal_pdf_extractor.py patient-info   <input> [options]
    python universal_pdf_extractor.py gross-pay      <input> [options]
    python universal_pdf_extractor.py 401k           <input> [options]
    python universal_pdf_extractor.py creditor-list  <input> [options]
    python universal_pdf_extractor.py generic        <input> [options]
    python universal_pdf_extractor.py bucket11       <input> [options]

Every subcommand writes its primary output next to the source PDF (or into
-o/--output-dir) as "<stem>_extracted.xlsx"; secondary outputs (template
mappings, processing summaries, combined workbooks) use that same stem with
an additional suffix, e.g. "<stem>_extracted_template.xlsx".

Run `python universal_pdf_extractor.py <subcommand> --help` for that
subcommand's own options.

DEPENDENCIES (install what each subcommand needs)
    pip install pymupdf pdfplumber openpyxl pypdf pandas tqdm pytesseract pillow

This project processes PHI/PII (SSNs, patient/employee names, addresses,
wages, claims). Console/debug output across every subcommand is written to
avoid printing actual field values where the original scripts redacted them
(claims, 1095c) — see each subcommand's own --debug behavior.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
import unicodedata
from pathlib import Path
from collections import defaultdict

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
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

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

try:
    import pytesseract
    from PIL import Image
    HAS_OCR = True
except ImportError:
    HAS_OCR = False


# ===========================================================================
# Shared utilities — used across every subcommand
# ===========================================================================

def _clean(text) -> str:
    return " ".join(str(text or "").split()).strip()


def _csv_safe(value) -> str:
    """Prefix a leading =+-@ with an apostrophe so Excel/CSV never treats an
    extracted field as a formula."""
    s = str(value or "")
    if s[:1] in ("=", "+", "-", "@"):
        return "'" + s
    return s


def _atomic_replace(tmp_path: Path, out_path: Path) -> bool:
    """Write-then-rename with retry so a file open in Excel doesn't abort a
    whole batch run — used by every subcommand's workbook writer."""
    for attempt in range(5):
        try:
            os.replace(tmp_path, out_path)
            return True
        except PermissionError:
            if attempt < 4:
                time.sleep(1)
            else:
                print(f"  WARNING: could not write {out_path.name} after 5 attempts "
                      "(file locked) -- skipping")
                try:
                    tmp_path.unlink(missing_ok=True)
                except OSError:
                    pass
                return False
    return False


def _iter_pdfs(target: Path, recursive: bool = False):
    if target.is_file():
        yield target
        return
    pattern = "**/*.pdf" if recursive else "*.pdf"
    yield from sorted(target.glob(pattern))


def _extracted_path(pdf_path: Path, output_dir, suffix: str = "_extracted.xlsx") -> Path:
    dest_dir = Path(output_dir) if output_dir else pdf_path.parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    return dest_dir / f"{pdf_path.stem}{suffix}"


def mask_shape(s: str) -> str:
    """Digit/letter shape mask for safe-to-share debug output (PHI scripts)."""
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def normalize_text(s: str) -> str:
    """Fold OCR's curly quotes/dashes and odd whitespace to plain ASCII."""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate({0x2018: "'", 0x2019: "'", 0x00B4: "'", 0x0060: "'",
                      0x201C: '"', 0x201D: '"', 0x00A0: " ",
                      0x2013: "-", 0x2014: "-"})
    return re.sub(r"[ \t]+", " ", s).strip()


def _style_header_row(ws, row_idx: int, col_count: int) -> None:
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _autofit(ws) -> None:
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 80)


def _autofit_widths(ws, widths) -> None:
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


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

W2_CSV_COLUMNS = ["Page", "First Name", "Last Name", "Suffix",
                  "Street Address", "City", "State", "Zip Code", "SSN"]

W2_NAME_SUFFIXES = {"JR", "SR", "II", "III", "IV", "V"}

W2_SSN_CAPTION_RE = re.compile(r"Employee.?s\s+(?:social security number|SSN)\b", re.IGNORECASE)
W2_NAME_ONLY_CAPTION_RE = re.compile(r"Employee.?s\s+first name and initial\b", re.IGNORECASE)
W2_COMBINED_NAME_ADDR_CAPTION_RE = re.compile(r"Employee.?s\s+name,\s*address", re.IGNORECASE)
W2_ADDRESS_CAPTION_RE = re.compile(r"\bf\s+Employee.?s address\b", re.IGNORECASE)
W2_NAME_CAPTION_RE = re.compile(
    r"Employee.?s\s+(?:first name and initial|name,\s*address)|\bf\s+Employee.?s address\b",
    re.IGNORECASE,
)
W2_STOP_LABEL_RE = re.compile(
    r"^(?:\d{1,2}\s+State\b|Employer.s state ID|Local income tax|Locality name|"
    r"Form\s*W-?2\b|Wage\s*(?:&|and)\s*Tax Statement|Copy\s+[A-Z0-9]|"
    r"For (?:Official|Privacy)|Department of the Treasury|20\d{2}$)",
    re.IGNORECASE,
)
W2_BOX_NUMERIC_LABEL_RE = re.compile(r"^1[1-4][a-d]?$")
W2_SSN_VALUE_RE = re.compile(r"\b(\d{3})[-\s]?(\d{2})[-\s]?(\d{4})\b")
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


def w2_group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
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
    for x0, x1, text in word_line:
        if W2_BOX_NUMERIC_LABEL_RE.match(text.strip(".:,")):
            return x0
    return None


def w2_row_text_left_of(word_line, max_x):
    words = word_line if max_x is None else [w for w in word_line if w[0] < max_x]
    return normalize_text(" ".join(t for _, _, t in words))


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

    for i, line in enumerate(plain_lines):
        if W2_COMBINED_NAME_ADDR_CAPTION_RE.search(line):
            return _w2_find_name_address_combined_box(plain_lines, i, page_num, column_label)

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
    if m:
        raw = m.group(1).replace(",", "")
        try:
            return f"{float(raw):.2f}"
        except ValueError:
            return ""
    return ""


def w2_extract_wages_from_text(cell_text: str) -> dict:
    lines = cell_text.splitlines()
    results = {key: "" for key, _ in W2_WAGE_BOX_DEFS}
    for box_key, label_list in W2_WAGE_BOX_DEFS:
        for i, line in enumerate(lines):
            line_lower = line.lower()
            match_pos = next((line_lower.find(lbl) for lbl in label_list if lbl in line_lower), -1)
            if match_pos == -1:
                continue
            label_len = next(len(lbl) for lbl in label_list if lbl in line_lower)
            # Search only after the label text itself, so a leading box number
            # ("1 Wages, tips...") is never mistaken for the dollar amount.
            amount = w2_parse_amount(line[match_pos + label_len:])
            if not amount and i + 1 < len(lines):
                amount = w2_parse_amount(lines[i + 1])
            if amount:
                results[box_key] = amount
                break
    return results


def w2_extract_wages_from_words(words) -> dict:
    """words: fitz (x0, y0, x1, y1, text, ...) tuples for one employee's cell."""
    results = {key: "" for key, _ in W2_WAGE_BOX_DEFS}
    LINE_TOL = 6

    def words_near(target_top, x_min=0, x_max=9999, y_slack=LINE_TOL):
        return [w for w in words if abs(w[1] - target_top) <= y_slack and x_min <= w[0] <= x_max]

    for box_key, label_list in W2_WAGE_BOX_DEFS:
        if results[box_key]:
            continue
        for w in words:
            text_lower = w[4].lower()
            if any(lbl in text_lower for lbl in label_list):
                nearby = sorted(words_near(w[1], x_min=w[2]), key=lambda ww: ww[0])
                for cand in nearby:
                    amt = w2_parse_amount(cand[4])
                    if amt:
                        results[box_key] = amt
                        break
                break
    return results


def w2_extract_wages_for_cell(column_words) -> dict:
    cell_text = "\n".join(" ".join(t for _, _, t in ln) for ln in w2_group_words_into_lines(column_words))
    wages_words = w2_extract_wages_from_words(column_words)
    wages_text = w2_extract_wages_from_text(cell_text)
    return {key: wages_words.get(key) or wages_text.get(key) or "" for key, _ in W2_WAGE_BOX_DEFS}


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
        lines = w2_group_words_into_lines(column_words)
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


def w2_debug_page(path: Path, page_num: int):
    doc = fitz.open(path)
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return
    page = doc[page_num - 1]

    text = normalize_text(page.get_text())
    print(f"--- {path.name} page {page_num} of {len(doc)} ---")
    print(f"text layer: {len(text.strip())} chars")
    ssn_caption_count = len(W2_SSN_CAPTION_RE.findall(text))
    name_caption_count = len(W2_NAME_CAPTION_RE.findall(text))
    ssn_value_count = len(W2_SSN_VALUE_RE.findall(text))
    print(f"SSN captions found on page: {ssn_caption_count}")
    print(f"Name/address captions found on page: {name_caption_count}")
    print(f"SSN-shaped values found on page: {ssn_value_count}")

    words = page.get_text("words")
    lines = w2_group_words_into_lines(words)
    plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]
    print(f"lines detected: {len(plain_lines)}")
    for i, line in enumerate(plain_lines):
        if W2_SSN_CAPTION_RE.search(line):
            tag, shown = "  <-- SSN caption", line
        elif W2_NAME_ONLY_CAPTION_RE.search(line):
            tag, shown = "  <-- name caption (box e)", line
        elif W2_ADDRESS_CAPTION_RE.search(line):
            tag, shown = "  <-- address caption (box f)", line
        elif W2_COMBINED_NAME_ADDR_CAPTION_RE.search(line):
            tag, shown = "  <-- combined name/address caption", line
        elif W2_STOP_LABEL_RE.search(line):
            tag, shown = "  <-- stop label", line
        else:
            tag, shown = "", mask_shape(line)
        print(f"  [{i:>3}] {shown}{tag}")
    doc.close()


def w2_write_workbook(records, out_path: Path, columns):
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted"
    ws.append(columns)
    _style_header_row(ws, 1, len(columns))
    for rec in records:
        ws.append([_csv_safe(rec.get(col, "")) for col in columns])
    _autofit(ws)
    wb.save(tmp_path)
    wb.close()
    _atomic_replace(tmp_path, out_path)


def cmd_w2(args) -> int:
    if not HAS_FITZ:
        sys.exit("ERROR: pymupdf is required. Run: pip install pymupdf")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")

    input_path = Path(args.input)
    pdf_files = list(_iter_pdfs(input_path, getattr(args, "recursive", False))) if input_path.is_dir() else [input_path]
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return 0

    if args.debug:
        for pdf in pdf_files:
            w2_debug_page(pdf, args.debug_page)
        return 0

    columns = W2_CSV_COLUMNS + (W2_WAGE_COLUMNS if args.wages else [])
    combined_records = []
    total_written = 0

    for pdf in tqdm(pdf_files, desc="Extracting", unit="file"):
        records = w2_process_pdf(pdf, args.split_fraction, args.wages)
        if not records:
            print(f"  no records extracted from {pdf.name} -- no output written for this file")
            continue

        deduped = w2_dedupe_records(records)
        removed = len(records) - len(deduped)
        if removed:
            print(f"  removed {removed} duplicate record(s) (same employee found on multiple copies/pages)")

        out_path = _extracted_path(pdf, args.output_dir)
        w2_write_workbook(deduped, out_path, columns)
        total_written += len(deduped)
        print(f"  -> wrote {len(deduped)} record(s) to {out_path}")

        for rec in deduped:
            combined_records.append({"Document ID": pdf.stem, **rec})

    if args.combined and combined_records:
        combined_dir = args.output_dir if args.output_dir else (input_path if input_path.is_dir() else input_path.parent)
        combined_path = Path(combined_dir) / args.combined_name
        w2_write_workbook(combined_records, combined_path, ["Document ID"] + columns)
        print(f"\nWrote combined workbook: {len(combined_records)} record(s) -> {combined_path}")

    print(f"\nDone. {total_written} record(s) written across {len(pdf_files)} file(s).")
    return 0


# ===========================================================================
# 1095c — Employee identity/address fields from IRS Form 1095-C PDFs
# Ported from extract_1095c.py (pymupdf/fitz, Format A / Format B detection).
# ===========================================================================

C1095_MIN_TEXT_CHARS_PER_PAGE = 20

C1095_TEMPLATE_SHEET_NAME = "Import Template"
C1095_TEMPLATE_COL = {
    "Doc ID": 1, "First Name": 2, "Middle Name": 3, "Last Name": 4,
    "Entity Type_Employee": 7, "Street Address": 11, "City": 12,
    "State": 13, "Zip Code": 14, "Country": 19,
    "Social Security Number (SSN)": 23,
}
C1095_TEMPLATE_PAGE_COL = 66

C1095_OUTPUT_COLUMNS = ["Page", "Format", "First Name", "Middle Name", "Last Name",
                         "Street Address", "City", "State", "Zip Code", "Country", "SSN"]

C1095_CITY_STATE_ZIP_RE = re.compile(
    r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b\s*(?P<country>.*)$"
)
C1095_SSN_VALUE_RE = re.compile(r"[\dXx\*\-]{9,}")
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


def c1095_group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


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

    block = [l.strip() for l in [name_trailing] + plain_lines[name_end + 1:end_idx] if l.strip()]
    if not block:
        print(f"  page {page_num} (Format A): name/address block empty after label -- skipping")
        return None

    full_name = block[0]
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
    if ssn_start is not None:
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
    lines = c1095_group_words_into_lines(words)
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


def c1095_build_workbook(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "extracted"
    ws.append(C1095_OUTPUT_COLUMNS)
    _style_header_row(ws, 1, len(C1095_OUTPUT_COLUMNS))
    for rec in records:
        ws.append([rec.get(c, "") for c in C1095_OUTPUT_COLUMNS])
    _autofit(ws)
    return wb


def c1095_build_template_workbook(template_path, records, doc_id):
    records = c1095_dedupe_complete_rows(records)
    wb = openpyxl.load_workbook(template_path)
    ws = wb[C1095_TEMPLATE_SHEET_NAME]
    ws.cell(row=1, column=C1095_TEMPLATE_PAGE_COL, value="Page #")
    ws.cell(row=2, column=C1095_TEMPLATE_PAGE_COL, value="Page #")

    row_ptr = 3
    for rec in records:
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Doc ID"], value=doc_id)
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["First Name"], value=rec.get("First Name", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Middle Name"], value=rec.get("Middle Name", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Last Name"], value=rec.get("Last Name", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Entity Type_Employee"], value=True)
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Street Address"], value=rec.get("Street Address", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["City"], value=rec.get("City", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["State"], value=rec.get("State", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Zip Code"], value=rec.get("Zip Code", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Country"], value=rec.get("Country", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_COL["Social Security Number (SSN)"], value=rec.get("SSN", ""))
        ws.cell(row=row_ptr, column=C1095_TEMPLATE_PAGE_COL, value=rec.get("Page", ""))
        row_ptr += 1
    return wb


def cmd_1095c(args) -> int:
    if not HAS_FITZ:
        sys.exit("ERROR: pymupdf is required. Run: pip install pymupdf")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return 0

    output_dir = Path(args.output_dir) if args.output_dir else input_path.parent if input_path.is_file() else input_path
    output_dir.mkdir(parents=True, exist_ok=True)

    tmpl_wb = openpyxl.load_workbook(args.template, read_only=True)
    if C1095_TEMPLATE_SHEET_NAME not in tmpl_wb.sheetnames:
        print(f"Sheet '{C1095_TEMPLATE_SHEET_NAME}' not found in template: {args.template}")
        return 1
    tmpl_wb.close()

    for pdf in pdf_files:
        print(f"Processing {pdf.name} ...")
        records = c1095_process_pdf(pdf, debug=args.debug, output_dir=output_dir)
        if not records:
            print(f"  no records extracted from {pdf.name}")
            continue

        extracted_path = _extracted_path(pdf, output_dir)
        template_path = _extracted_path(pdf, output_dir, "_extracted_template.xlsx")

        c1095_build_workbook(records).save(extracted_path)
        c1095_build_template_workbook(args.template, records, pdf.name).save(template_path)

        print(f"  -> {extracted_path.name} ({len(records)} record(s))")
        print(f"  -> {template_path.name}")

    return 0


# ===========================================================================
# claims — Patient/claim identity fields from ABBYY-searchable remittance PDFs
# Ported from extract_claims_info.py (pymupdf/fitz, Format A/B auto-detect).
# ===========================================================================

from datetime import datetime as _datetime

CLAIMS_NAME_RE = re.compile(r"([A-Z][A-Za-z'\-]+(?:\s+[A-Z][A-Za-z'\-]+)*),[^A-Za-z]*([A-Za-z][A-Za-z'\-\. ]*)")
CLAIMS_NAME_RE_LOOSE_SEP = re.compile(r"([A-Z][A-Za-z'\-]+(?:\s+[A-Z][A-Za-z'\-]+)*)[,.][^A-Za-z]*([A-Za-z][A-Za-z'\-\. ]*)")
CLAIMS_DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")
CLAIMS_NAME_PRECEDING_FILLER_WORDS = {"medi-cal", "medicare", "commercial", "medi-medi", "pat"}

CLAIMS_DEBUG_SAFE_WORDS = {
    "health", "plan", "id", "claim", "status", "paid", "pat", "acct", "payee",
    "service", "date", "revenue", "code", "cpt", "billed", "amount", "not",
    "allowed", "contract", "adjust", "interest", "penalty", "deduct", "coin",
    "copay", "net", "amt", "user", "comments", "member", "totals", "line",
    "ver", "received", "from", "to", "proc", "mod", "qty", "patient", "name",
    "provider", "business", "medi", "cal", "no", "pcp", "network", "reason",
    "of", "by", "program", "all", "programs", "combined", "check", "vendor",
    "group", "community", "clinics", "connect", "e", "p", "s", "t", "ra",
}

CLAIMS_HEALTH_PLAN_RE = re.compile(r"Health\s*Plan\s*ID\s*:\s*(\S+)")
CLAIMS_HEALTH_PLAN_LABEL_RE = re.compile(r"Health\s*Plan\s*ID\s*:")
CLAIMS_PAT_ACCT_RE = re.compile(r"Pat\.?\s*Acct\s*#\s*(\S+)")
CLAIMS_CLAIM_LABEL_RE = re.compile(r"Claim\s*:")
CLAIMS_CLAIM_CONTINUATION_RE = re.compile(r"^[A-Za-z0-9-]{1,3}\d+$")
CLAIMS_PATIENT_ACCT_RE = re.compile(r"Patient\s+Acct\.?\s*#\s*(\S+)")
CLAIMS_MEMBER_TOTALS_RE = re.compile(r"Member Totals\s*[:;]")
CLAIMS_PROC_TOKEN_RE = re.compile(r"[A-Za-z]?\d{3,5}")
CLAIMS_DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y")
CLAIMS_PROVIDER_NAME_HEADER_RE = re.compile(r"Provider\s*Name")

CLAIMS_Y_TOL = 5
CLAIMS_COL_X_TOL = 2
CLAIMS_LOOKBACK_NAME = 2
CLAIMS_LOOKAHEAD_ACCT = 4
CLAIMS_LOOKAHEAD_HEADER = 8
CLAIMS_LOOKAHEAD_DATA_ROWS = 25

CLAIMS_COLUMNS = ["Patient Name", "Health Plan ID", "Member #", "Claim #", "Pat. Acct #", "Service Date", "CPT Code", "Page"]
CLAIMS_COMBINED_COLUMNS = ["Source File"] + CLAIMS_COLUMNS


def claims_find_name_match(text, pos=0, pattern=CLAIMS_NAME_RE):
    while True:
        m = pattern.search(text, pos)
        if not m:
            return None
        first_word = m.group(1).split()[0]
        if first_word.lower() not in CLAIMS_NAME_PRECEDING_FILLER_WORDS:
            return m
        pos = m.start(1) + len(first_word)


def claims_redact_word(word):
    core = word.strip(":#.,()/-'’")
    if core.lower() in CLAIMS_DEBUG_SAFE_WORDS:
        return word
    return "".join("#" if ch.isdigit() else ("X" if ch.isalpha() else ch) for ch in word)


def claims_redact_line_for_debug(line):
    return " ".join(claims_redact_word(t) for _, _, t in line)


def claims_redact_value(value):
    return "".join("#" if ch.isdigit() else ("X" if ch.isalpha() else ch) for ch in str(value))


def claims_find_provider_name_x0(lines):
    for line in lines:
        if CLAIMS_PROVIDER_NAME_HEADER_RE.search(claims_line_text(line)):
            x0 = claims_word_x0(line, re.compile(r"^Provider$"))
            if x0 is not None:
                return x0
    return None


def claims_parse_date(date_str):
    for fmt in CLAIMS_DATE_FORMATS:
        try:
            return _datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def claims_group_words_into_lines(words):
    entries = sorted(
        ((x0, x1, text, (y0 + y1) / 2) for x0, y0, x1, y1, text, *_ in words),
        key=lambda e: (e[3], e[0]),
    )
    lines = []
    current = []
    last_y = None
    for x0, x1, text, ymid in entries:
        if last_y is not None and abs(ymid - last_y) > CLAIMS_Y_TOL:
            lines.append(sorted(current, key=lambda t: t[0]))
            current = []
        current.append((x0, x1, text))
        last_y = ymid
    if current:
        lines.append(sorted(current, key=lambda t: t[0]))
    return lines


def claims_line_text(line):
    return " ".join(t for _, _, t in line)


def claims_build_lines_for_page(page, mirror):
    words = page.get_text("words")
    if mirror:
        page_width = page.rect.width
        words = [(page_width - x1, y0, page_width - x0, y1, text, *rest)
                  for x0, y0, x1, y1, text, *rest in words]
    return claims_group_words_into_lines(words)


def claims_page_looks_readable(lines):
    text = "\n".join(claims_line_text(l) for l in lines)
    return bool(
        CLAIMS_HEALTH_PLAN_RE.search(text) or CLAIMS_MEMBER_TOTALS_RE.search(text)
        or CLAIMS_PATIENT_ACCT_RE.search(text) or CLAIMS_PAT_ACCT_RE.search(text)
    )


def claims_score_complete_records(lines, line_pages):
    b_rows = claims_extract_records(lines, line_pages, quiet=True)
    a_rows = claims_extract_tabular_records(lines, line_pages, quiet=True)
    complete_b = sum(
        1 for r in b_rows
        if all(r.get(c) for c in ("Patient Name", "Claim #", "Pat. Acct #", "Service Date", "CPT Code"))
    )
    complete_a = sum(
        1 for r in a_rows
        if all(r.get(c) for c in ("Patient Name", "Pat. Acct #", "Service Date"))
    )
    return complete_b + complete_a


def claims_get_page_lines(page, debug=False, page_num=None):
    normal = claims_build_lines_for_page(page, mirror=False)
    oriented = normal
    if not claims_page_looks_readable(normal):
        mirrored = claims_build_lines_for_page(page, mirror=True)
        if claims_page_looks_readable(mirrored):
            oriented = mirrored
            if debug:
                print(f"  [debug] page {page_num}: text layer is mirrored left-right -- corrected automatically")

    page_nums = [page_num] * len(oriented)
    reversed_lines = list(reversed(oriented))

    forward_strict = claims_count_backward_names(oriented)
    reversed_strict = claims_count_backward_names(reversed_lines)
    if forward_strict == 0 and reversed_strict == 0:
        forward_score = claims_score_complete_records(oriented, page_nums)
        if forward_score > 0:
            return oriented
        reversed_score = claims_score_complete_records(reversed_lines, page_nums)
        if reversed_score > 0:
            if debug:
                print(f"  [debug] page {page_num}: top-to-bottom line order was reversed -- corrected automatically")
            return reversed_lines
        return oriented

    if reversed_strict > forward_strict:
        if debug:
            print(f"  [debug] page {page_num}: top-to-bottom line order was reversed -- corrected automatically")
        return reversed_lines
    return oriented


def claims_word_x0(line, token_re):
    for x0, x1, text in line:
        if token_re.match(text):
            return x0
    return None


def claims_bucket_row(line, col_defs):
    buckets = {label: [] for label, _ in col_defs}
    for x0, x1, text in line:
        label = col_defs[0][0]
        for cand_label, start in col_defs:
            if x0 + CLAIMS_COL_X_TOL >= start:
                label = cand_label
        buckets[label].append(text)
    return {label: " ".join(words).strip() for label, words in buckets.items()}


def claims_extract_name(lines, hp_idx, allow_forward_fallback=True):
    same_line = claims_line_text(lines[hp_idx]).split("Health Plan ID:")[0]
    m = claims_find_name_match(same_line.split("Payee:")[0], pattern=CLAIMS_NAME_RE_LOOSE_SEP)
    if m:
        return f"{m.group(1).strip()}, {m.group(2).strip()}"

    lookback_stop = max(hp_idx - 1 - CLAIMS_LOOKBACK_NAME, -1)
    for i in range(hp_idx - 1, lookback_stop, -1):
        text = claims_line_text(lines[i])
        if CLAIMS_HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        above = text.split("Payee:")[0]
        m = claims_find_name_match(above, pattern=CLAIMS_NAME_RE_LOOSE_SEP)
        if m:
            return f"{m.group(1).strip()}, {m.group(2).strip()}"

    if not allow_forward_fallback:
        return ""

    lookahead_stop = min(hp_idx + 1 + CLAIMS_LOOKBACK_NAME, len(lines))
    for i in range(hp_idx + 1, lookahead_stop):
        text = claims_line_text(lines[i])
        if CLAIMS_HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        above = text.split("Payee:")[0]
        m = claims_find_name_match(above, pattern=CLAIMS_NAME_RE_LOOSE_SEP)
        if m:
            return f"{m.group(1).strip()}, {m.group(2).strip()}"
    return ""


def claims_count_backward_names(lines):
    return sum(
        1 for i, line in enumerate(lines)
        if CLAIMS_HEALTH_PLAN_RE.search(claims_line_text(line)) and claims_extract_name(lines, i, allow_forward_fallback=False)
    )


def claims_extract_health_plan_id(text):
    label = CLAIMS_HEALTH_PLAN_LABEL_RE.search(text)
    if not label:
        return ""
    for tok in text[label.end():].split():
        if any(ch.isdigit() for ch in tok):
            return tok
    return ""


def claims_extract_claim(lines, hp_idx):
    text = claims_line_text(lines[hp_idx])
    label = CLAIMS_CLAIM_LABEL_RE.search(text)
    if not label:
        return ""
    tokens = text[label.end():].split()
    if not tokens:
        return ""
    claim = tokens[0]
    if len(tokens) > 1 and CLAIMS_CLAIM_CONTINUATION_RE.match(tokens[1]):
        claim += tokens[1]
    return claim


def claims_extract_acct(lines, hp_idx):
    for i in range(hp_idx, min(hp_idx + CLAIMS_LOOKAHEAD_ACCT, len(lines))):
        m = CLAIMS_PAT_ACCT_RE.search(claims_line_text(lines[i]))
        if m:
            return m.group(1)
    lookback_stop = max(hp_idx - 1 - CLAIMS_LOOKAHEAD_ACCT, -1)
    for i in range(hp_idx - 1, lookback_stop, -1):
        text = claims_line_text(lines[i])
        if CLAIMS_HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        m = CLAIMS_PAT_ACCT_RE.search(text)
        if m:
            return m.group(1)
    return ""


def claims_find_service_header(lines, hp_idx):
    def header_col_defs(line):
        cpt_idx = next((idx for idx, (_, _, text) in enumerate(line) if re.match(r"^CPT-?$", text)), None)
        if cpt_idx is None or cpt_idx + 1 >= len(line):
            return None
        svc_x0 = line[0][0]
        cpt_x0 = line[cpt_idx][0]
        billed_x0 = line[cpt_idx + 1][0]
        rev_x0 = claims_word_x0(line, re.compile(r"^Revenue$"))
        return sorted(
            {"service_date": svc_x0, "revenue_code": rev_x0 if rev_x0 is not None else cpt_x0,
             "cpt": cpt_x0, "amounts": billed_x0}.items(),
            key=lambda kv: kv[1],
        )

    lookback_stop = max(hp_idx - 1 - CLAIMS_LOOKAHEAD_HEADER, -1)
    for i in range(hp_idx - 1, lookback_stop, -1):
        text = claims_line_text(lines[i])
        if CLAIMS_HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        col_defs = header_col_defs(lines[i])
        if col_defs:
            return i, col_defs

    for i in range(hp_idx, min(hp_idx + CLAIMS_LOOKAHEAD_HEADER, len(lines))):
        col_defs = header_col_defs(lines[i])
        if col_defs:
            return i, col_defs
    return None, None


def claims_extract_service_date_and_cpt(lines, header_idx, col_defs):
    latest_date, cpt = None, ""
    for i in range(header_idx + 1, min(header_idx + 1 + CLAIMS_LOOKAHEAD_DATA_ROWS, len(lines))):
        if CLAIMS_HEALTH_PLAN_RE.search(claims_line_text(lines[i])):
            break
        bucket = claims_bucket_row(lines[i], col_defs)
        date_match = CLAIMS_DATE_RE.search(bucket.get("service_date", ""))
        if date_match:
            parsed = claims_parse_date(date_match.group())
            if parsed and (latest_date is None or parsed > latest_date):
                latest_date = parsed
        cpt_candidate = bucket.get("cpt", "")
        if not cpt and cpt_candidate and any(ch.isdigit() for ch in cpt_candidate):
            cpt = cpt_candidate
    service_date = latest_date.strftime("%m/%d/%Y") if latest_date else ""
    return service_date, cpt


def claims_extract_records(lines, line_pages, quiet=False):
    rows = []
    for i, line in enumerate(lines):
        hp_match = CLAIMS_HEALTH_PLAN_RE.search(claims_line_text(line))
        if not hp_match:
            continue

        health_plan_id = claims_extract_health_plan_id(claims_line_text(line))
        name = claims_extract_name(lines, i)
        claim = claims_extract_claim(lines, i)
        acct = claims_extract_acct(lines, i)
        header_idx, col_defs = claims_find_service_header(lines, i)
        service_date, cpt = ("", "")
        if header_idx is not None:
            service_date, cpt = claims_extract_service_date_and_cpt(lines, max(header_idx, i), col_defs)

        missing = [
            label for label, val in [
                ("Patient Name", name), ("Health Plan ID", health_plan_id),
                ("Claim #", claim), ("Pat. Acct #", acct),
                ("Service Date", service_date), ("CPT Code", cpt),
            ] if not val
        ]
        if missing and not quiet:
            print(f"    record at Health Plan ID {claims_redact_value(health_plan_id or hp_match.group(1))} "
                  f"(page {line_pages[i]}): missing {', '.join(missing)}")

        rows.append({
            "Patient Name": name, "Health Plan ID": health_plan_id, "Member #": "",
            "Claim #": claim, "Pat. Acct #": acct, "Service Date": service_date,
            "CPT Code": cpt, "Page": line_pages[i],
        })
    return rows


def claims_last_date_and_proc(text):
    dates = list(CLAIMS_DATE_RE.finditer(text))
    if not dates:
        return None, ""
    proc_match = CLAIMS_PROC_TOKEN_RE.search(text, dates[-1].end())
    return dates[-1].group(), (proc_match.group() if proc_match else "")


def claims_extract_tabular_records(lines, line_pages, quiet=False):
    rows = []
    skipped = []
    current = None
    provider_name_x0 = claims_find_provider_name_x0(lines)

    for i, line in enumerate(lines):
        text = claims_line_text(line)
        if provider_name_x0 is not None:
            name_search_text = claims_line_text([w for w in line if w[0] < provider_name_x0 - CLAIMS_COL_X_TOL])
        else:
            name_search_text = text

        name_match = claims_find_name_match(name_search_text)
        prefix = name_search_text[:name_match.start(1)].strip() if name_match else ""
        if name_match and not any(ch.isdigit() for ch in prefix):
            name_match = None

        if name_match:
            name = f"{name_match.group(1).strip()}, {name_match.group(2).strip()}"
            prefix_tokens = prefix.split()
            claim = prefix_tokens[0] if prefix_tokens else ""
            member_id = ""
            if i > 0:
                prev_text = claims_line_text(lines[i - 1]).strip()
                prev_tokens = prev_text.split()
                if (prev_tokens and "Totals" not in prev_text and "Acct" not in prev_text
                        and any(ch.isdigit() for ch in prev_tokens[0])):
                    member_id = prev_tokens[0]
            current = {
                "name": name, "member_id": member_id, "claim": claim,
                "acct": "", "latest_date": None, "page": line_pages[i],
            }

        if current is None:
            continue

        acct_match = CLAIMS_PATIENT_ACCT_RE.search(text)
        if acct_match:
            current["acct"] = acct_match.group(1)
            continue

        if CLAIMS_MEMBER_TOTALS_RE.search(text):
            required_missing = [
                label for label, val in [("Pat. Acct #", current["acct"]), ("Service Date", current["latest_date"])]
                if not val
            ]
            optional_missing = [
                label for label, val in [("Member #", current["member_id"]), ("Claim #", current["claim"])]
                if not val
            ]
            if required_missing:
                skipped.append((current["name"], current["page"], required_missing + optional_missing))
            else:
                if optional_missing and not quiet:
                    print(f"    tabular record for '{claims_redact_value(current['name'])}' (page {current['page']}): "
                          f"missing {', '.join(optional_missing)}")
                rows.append({
                    "Patient Name": current["name"], "Health Plan ID": "",
                    "Member #": current["member_id"], "Claim #": current["claim"],
                    "Pat. Acct #": current["acct"], "Service Date": current["latest_date"].strftime("%m/%d/%Y"),
                    "CPT Code": "", "Page": current["page"],
                })
            current = None
            continue

        date_str, proc = claims_last_date_and_proc(text)
        if date_str and proc:
            parsed = claims_parse_date(date_str)
            if parsed and (current["latest_date"] is None or parsed > current["latest_date"]):
                current["latest_date"] = parsed

    if not quiet:
        for name, page, missing in skipped:
            print(f"    tabular record for '{claims_redact_value(name)}' (page {page}): missing {', '.join(missing)} -- skipped")
    return rows


def claims_merge_duplicate_rows(rows):
    groups = {}
    order = []
    for idx, row in enumerate(rows):
        hp, name = row.get("Health Plan ID", ""), row.get("Patient Name", "")
        key = (hp, name) if hp and name else ("_unique", idx)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    merged = []
    for key in order:
        group = groups[key]
        if len(group) == 1:
            merged.append(group[0])
            continue
        combined = {}
        for col in CLAIMS_COLUMNS:
            if col == "Service Date":
                dates = [d for d in (claims_parse_date(str(row.get(col, ""))) for row in group) if d]
                combined[col] = max(dates).strftime("%m/%d/%Y") if dates else ""
                continue
            seen = []
            for row in group:
                val = str(row.get(col, ""))
                if val and val not in seen:
                    seen.append(val)
            combined[col] = ";".join(seen)
        merged.append(combined)
    return merged


def claims_page_sort_key(row):
    try:
        return int(str(row.get("Page", "")).split(";")[0])
    except ValueError:
        return 0


def claims_extract_pdf(path: Path, debug=False, only_page=None):
    doc = fitz.open(path)
    all_lines = []
    line_pages = []
    for page_num, page in enumerate(doc, start=1):
        if only_page is not None and page_num != only_page:
            continue
        text_len = len(page.get_text().strip())
        if text_len == 0:
            print(f"  page {page_num}: no text layer found -- this page needs OCR before it can be extracted")
            continue
        page_lines = claims_get_page_lines(page, debug=debug, page_num=page_num)
        all_lines.extend(page_lines)
        line_pages.extend([page_num] * len(page_lines))
    doc.close()

    if debug:
        print(f"  [debug] {len(all_lines)} line(s) read (names/IDs/dates masked with X/# below):")
        for i, line in enumerate(all_lines):
            print(f"  [debug] line {i} (page {line_pages[i]}): {claims_redact_line_for_debug(line)!r}")

    rows = (claims_extract_records(all_lines, line_pages, quiet=not debug)
            + claims_extract_tabular_records(all_lines, line_pages, quiet=not debug))
    rows = claims_merge_duplicate_rows(rows)
    rows.sort(key=claims_page_sort_key)
    return rows


def claims_build_workbook(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "extracted"
    ws.append(CLAIMS_COLUMNS)
    _style_header_row(ws, 1, len(CLAIMS_COLUMNS))
    for row in rows:
        ws.append([row.get(c, "") for c in CLAIMS_COLUMNS])
    _autofit(ws)
    return wb


def claims_build_combined_workbook(source_and_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "combined"
    ws.append(CLAIMS_COMBINED_COLUMNS)
    _style_header_row(ws, 1, len(CLAIMS_COMBINED_COLUMNS))
    for source, row in source_and_rows:
        ws.append([source] + [row.get(c, "") for c in CLAIMS_COLUMNS])
    _autofit(ws)
    return wb


def cmd_claims(args) -> int:
    if not HAS_FITZ:
        sys.exit("ERROR: pymupdf is required. Run: pip install pymupdf")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return 0

    combined_rows = []
    for pdf in pdf_files:
        print(f"Processing {pdf.name} ...")
        rows = claims_extract_pdf(pdf, debug=args.debug, only_page=args.page)
        if not rows:
            print(f"  no 'Health Plan ID:' records found")
            continue

        output_xlsx = _extracted_path(pdf, args.output_dir)
        try:
            claims_build_workbook(rows).save(output_xlsx)
        except PermissionError:
            print(f"  could not write {output_xlsx.name} -- it's likely open in Excel; close it and re-run")
            continue
        print(f"  -> {output_xlsx.name} ({len(rows)} record(s))")
        combined_rows.extend((pdf.name, row) for row in rows)

    if combined_rows:
        combined_output_dir = Path(args.output_dir) if args.output_dir else pdf_files[0].parent
        combined_output_dir.mkdir(parents=True, exist_ok=True)
        combined_xlsx = combined_output_dir / "combined_extracted.xlsx"
        try:
            claims_build_combined_workbook(combined_rows).save(combined_xlsx)
        except PermissionError:
            print(f"could not write {combined_xlsx.name} -- it's likely open in Excel; close it and re-run")
        else:
            print(f"-> {combined_xlsx.name} ({len(combined_rows)} record(s))")

    return 0


# ===========================================================================
# patient-info — patient info from scanned/digital remittance PDFs, with OCR
# Ported from extract_patient_info.py (pymupdf/fitz + pytesseract OCR pass).
# ===========================================================================

import bisect as _bisect

PATIENT_DPI = 300
PATIENT_MIN_TEXT_CHARS_PER_PAGE = 20
PATIENT_NAME_RE = r"[A-Z][A-Za-z'\-]+,\s*[A-Z][A-Za-z'\-\.]*(?: [A-Za-z][A-Za-z'\-\.]*){0,4}"
PATIENT_DATE_RE = r"\d{1,2}/\d{1,2}/\d{2,4}"
PATIENT_NAME_SEARCH_RE = re.compile(PATIENT_NAME_RE)
PATIENT_DATE_SEARCH_RE = re.compile(PATIENT_DATE_RE)
PATIENT_WINDOW_BEFORE = 200
PATIENT_WINDOW_AFTER = 500

PATIENT_HEALTH_PLAN_RE = re.compile(r"Health Plan ID:\s*(?P<planid>[\w\d]+)")
PATIENT_CLAIM_LABEL_RE = re.compile(r"Claim:\s*(?P<claim>[\w\d]+)")
PATIENT_ACCT_LABEL_RE = re.compile(r"Pat\.?\s*Acct\s*#\s*(?P<acct>[\w\d]+)")
PATIENT_SVC_HEADER_RE = re.compile(r"Service Date.{0,80}?CPT-?", re.DOTALL)

PATIENT_NAME_LINE_RE = PATIENT_NAME_SEARCH_RE
PATIENT_MEMBER_TOTALS_RE = re.compile(r"Member Totals\s*:")
PATIENT_PATIENT_ACCT_RE = re.compile(r"Patient Acct\.?\s*#\s*(?P<acct>[\w\d]+)")
PATIENT_MEMBER_RE = re.compile(r"(?P<member>\d{9,})")
PATIENT_CLAIM_LINE_RE = re.compile(r"(?P<claim>\d{6,})")
PATIENT_PROC_TOKEN_RE = re.compile(r"[A-Za-z]?\d{3,5}")
PATIENT_DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y")


def patient_deskew_page_image(pil_img):
    try:
        osd = pytesseract.image_to_osd(pil_img)
        angle = int(re.search(r"Rotate: (\d+)", osd).group(1))
    except pytesseract.TesseractError:
        angle = 0
    if angle:
        pil_img = pil_img.rotate(-angle, expand=True)
    return pil_img


def patient_make_searchable_pdf(input_path: Path, output_path: Path):
    doc = fitz.open(input_path)
    out_pdf = fitz.open()

    pages = tqdm(list(enumerate(doc)), desc=f"  {input_path.name}", unit="page", leave=False)
    for page_index, page in pages:
        text = page.get_text().strip()

        if len(text) >= PATIENT_MIN_TEXT_CHARS_PER_PAGE:
            out_pdf.insert_pdf(doc, from_page=page_index, to_page=page_index)
            continue

        try:
            pix = page.get_pixmap(dpi=PATIENT_DPI)
            pil_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            pil_img = patient_deskew_page_image(pil_img)
            ocr_pdf_bytes = pytesseract.image_to_pdf_or_hocr(pil_img, extension="pdf")
            ocr_page = fitz.open("pdf", ocr_pdf_bytes)
            out_pdf.insert_pdf(ocr_page)
            ocr_page.close()
        except Exception as exc:
            print(f"    {input_path.name} page {page_index + 1}: OCR failed ({exc}), keeping original page as-is")
            out_pdf.insert_pdf(doc, from_page=page_index, to_page=page_index)

    if out_pdf.page_count != doc.page_count:
        print(f"    WARNING: {input_path.name}: searchable PDF has {out_pdf.page_count} pages, "
              f"expected {doc.page_count} -- some content may be missing")

    out_pdf.save(output_path)
    out_pdf.close()
    doc.close()


def patient_extract_pdf_text(path: Path):
    doc = fitz.open(path)
    parts = [page.get_text(sort=True) for page in doc]
    doc.close()

    page_offsets = []
    offset = 0
    for part in parts:
        page_offsets.append(offset)
        offset += len(part) + 1

    return "\n".join(parts), page_offsets


def patient_page_number_at(char_offset, page_offsets):
    return _bisect.bisect_right(page_offsets, char_offset)


def patient_extract_format_b(text, page_offsets):
    rows = []
    skipped = []
    for i, hp_match in enumerate(PATIENT_HEALTH_PLAN_RE.finditer(text), start=1):
        before = text[max(0, hp_match.start() - PATIENT_WINDOW_BEFORE):hp_match.start()]
        after = text[hp_match.end():hp_match.end() + PATIENT_WINDOW_AFTER]

        name_line = before.rstrip("\n").rsplit("\n", 1)[-1]
        name_line = name_line.split("Payee:", 1)[0]
        name_match = PATIENT_NAME_SEARCH_RE.search(name_line)

        claim_match = PATIENT_CLAIM_LABEL_RE.search(after)
        acct_match = PATIENT_ACCT_LABEL_RE.search(after)
        svc_header_match = PATIENT_SVC_HEADER_RE.search(after)
        missing = [
            field for field, m in [
                ("Patient Name", name_match), ("Claim", claim_match),
                ("Pat. Acct #", acct_match), ("Service Date/CPT header", svc_header_match),
            ] if not m
        ]
        if missing:
            skipped.append((i, missing))
            continue

        date_match = PATIENT_DATE_SEARCH_RE.search(after, svc_header_match.end())
        if not date_match:
            skipped.append((i, ["Service Date value"]))
            continue

        dollar_match = re.search(r"\$", after[date_match.end():])
        cpt_end = date_match.end() + dollar_match.start() if dollar_match else date_match.end() + 40
        cpt_value = re.sub(r"\s+", " ", after[date_match.end():cpt_end]).strip()

        rows.append({
            "Format": "B", "Patient Name": name_match.group().strip(),
            "Health Plan ID": hp_match.group("planid"), "Claim #": claim_match.group("claim"),
            "Member #": "", "Pat. Acct #": acct_match.group("acct"),
            "Service Date": date_match.group(), "CPT": cpt_value,
            "Page": patient_page_number_at(hp_match.start(), page_offsets),
        })

    for i, missing in skipped:
        print(f"    Format B record #{i}: skipped -- missing/unmatched: {', '.join(missing)}")
    return rows


def patient_line_latest_date_and_proc(line):
    dates = list(PATIENT_DATE_SEARCH_RE.finditer(line))
    if not dates:
        return "", ""
    proc_match = PATIENT_PROC_TOKEN_RE.search(line, dates[-1].end())
    return dates[-1].group(), (proc_match.group() if proc_match else "")


def patient_parse_date(date_str):
    for fmt in PATIENT_DATE_FORMATS:
        try:
            return _datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def patient_extract_format_a(text, page_offsets):
    rows = []
    skipped = []
    current = None
    record_count = 0

    offset = 0
    for raw_line in text.split("\n"):
        line_start = offset
        offset += len(raw_line) + 1
        line = raw_line

        name_match = PATIENT_NAME_LINE_RE.search(line)
        if name_match:
            record_count += 1
            member_match = PATIENT_MEMBER_RE.search(line)
            claim_match = PATIENT_CLAIM_LINE_RE.search(line)
            svc_date, proc_value = patient_line_latest_date_and_proc(line)
            current = {
                "name": name_match.group().strip(),
                "member": member_match.group("member") if member_match else "",
                "claim": claim_match.group("claim") if claim_match else "",
                "accts": [], "svc_date": svc_date, "proc": proc_value, "line_start": line_start,
            }
            continue

        if current is None:
            continue

        acct_match = PATIENT_PATIENT_ACCT_RE.search(line)
        if acct_match:
            acct = acct_match.group("acct")
            if acct not in current["accts"]:
                current["accts"].append(acct)
            continue

        if PATIENT_MEMBER_TOTALS_RE.search(line):
            missing = [
                field for field, v in [
                    ("Member #/Claim #", current["member"] or current["claim"]),
                    ("Patient Acct #", current["accts"]), ("Service Date", current["svc_date"]),
                ] if not v
            ]
            if missing:
                skipped.append((record_count, missing))
            else:
                rows.append({
                    "Format": "A", "Patient Name": current["name"], "Health Plan ID": "",
                    "Claim #": current["claim"], "Member #": current["member"],
                    "Pat. Acct #": ":".join(current["accts"]), "Service Date": current["svc_date"],
                    "Proc": current["proc"], "Page": patient_page_number_at(current["line_start"], page_offsets),
                })
            current = None
            continue

        svc_date, _proc = patient_line_latest_date_and_proc(line)
        if svc_date and (patient_parse_date(svc_date) or _datetime.min) > (patient_parse_date(current["svc_date"]) or _datetime.min):
            current["svc_date"] = svc_date

    for i, missing in skipped:
        print(f"    Format A record #{i}: skipped -- missing/unmatched: {', '.join(missing)}")
    return rows


def patient_merge_same_patient(rows):
    groups = {}
    order = []
    passthrough = []

    for row in rows:
        plan_id = row.get("Health Plan ID", "")
        if not plan_id:
            passthrough.append(row)
            continue
        key = (row.get("Patient Name", ""), plan_id)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    merged = []
    for key in order:
        group = groups[key]
        if len(group) == 1:
            merged.append(group[0])
            continue

        accts = []
        for r in group:
            acct = r.get("Pat. Acct #", "")
            if acct and acct not in accts:
                accts.append(acct)

        latest_row = max(group, key=lambda r: patient_parse_date(r.get("Service Date", "")) or _datetime.min)

        combined = dict(group[0])
        combined["Pat. Acct #"] = ":".join(accts)
        combined["Service Date"] = latest_row.get("Service Date", "")
        merged.append(combined)

    return passthrough + merged


PATIENT_COLUMNS = ["Format", "Patient Name", "Health Plan ID", "Claim #", "Member #", "Pat. Acct #", "Service Date", "Proc", "CPT", "Page"]


def patient_build_workbook(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "extracted"
    ws.append(PATIENT_COLUMNS)
    _style_header_row(ws, 1, len(PATIENT_COLUMNS))
    for row in rows:
        ws.append([row.get(c, "") for c in PATIENT_COLUMNS])
    _autofit(ws)
    return wb


def cmd_patient_info(args) -> int:
    if not HAS_FITZ:
        sys.exit("ERROR: pymupdf is required. Run: pip install pymupdf")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")
    if not HAS_OCR:
        sys.exit("ERROR: pytesseract and pillow are required. Run: pip install pytesseract pillow")

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return 0

    output_dir = Path(args.output_dir) if args.output_dir else Path(".")
    output_dir.mkdir(parents=True, exist_ok=True)

    searchable_paths = {}
    for pdf in tqdm(pdf_files, desc="Pass 1/2: making files searchable", unit="file"):
        searchable_path = output_dir / f"{pdf.stem}_searchable.pdf"
        if searchable_path.exists():
            expected_pages = fitz.open(pdf).page_count
            actual_pages = fitz.open(searchable_path).page_count
            if actual_pages != expected_pages:
                print(f"  {pdf.name}: existing searchable PDF has {actual_pages} pages, expected {expected_pages} -- redoing OCR")
                patient_make_searchable_pdf(pdf, searchable_path)
            else:
                print(f"  {pdf.name}: searchable PDF already exists ({actual_pages} pages), skipping OCR")
        else:
            patient_make_searchable_pdf(pdf, searchable_path)
        searchable_paths[pdf] = searchable_path

    for pdf in tqdm(pdf_files, desc="Pass 2/2: extracting patient info", unit="file"):
        text, page_offsets = patient_extract_pdf_text(searchable_paths[pdf])
        rows = patient_extract_format_b(text, page_offsets) + patient_extract_format_a(text, page_offsets)
        rows = patient_merge_same_patient(rows)
        if not rows:
            print(f"  {pdf.name}: no records matched either format")

        output_xlsx = _extracted_path(pdf, output_dir)
        patient_build_workbook(rows).save(output_xlsx)
        print(f"  {pdf.name} -> {output_xlsx.name} ({len(rows)} record(s))")

    return 0


# ===========================================================================
# gross-pay — "Employee Gross To Net" report data
# Merged from emp_gross_pay.py (SSN formatting, tqdm progress, template
# auto-detect) and employee_gross_to_pay.py (fuller column set: Payment Date,
# Total Hrs/Units, Total Earnings, Total Deductions, Total Taxes, Net Pay) --
# the two were forks of the same rotation-aware pymupdf header/row-detection
# core, reconciled here into one extractor with the fuller column set plus
# the narrower fork's SSN formatting and template auto-detection.
# ===========================================================================

GP_SSN_PATTERN = re.compile(r"\d{3}-\d{2}-\d{4}|[*Xx]{3}-[*Xx]{2}-\d{4}")
GP_EEID_PATTERN = re.compile(r"^[A-Za-z0-9-]{2,15}$")

GP_HEADER_DEFS = [
    ("EE ID", ["EE", "ID"]),
    ("Employee Name", ["Employee", "Name"]),
    ("SSN", ["SSN"]),
    ("Payment Date", ["Payment", "Date"]),
    ("Total Hrs/Units", ["Total", "Hrs/Units"]),
    ("Total Earnings", ["Total", "Earnings"]),
    ("Total Deductions", ["Total", "Deductions"]),
    ("Total Taxes", ["Total", "Taxes"]),
    ("Net Pay", ["Net", "Pay"]),
]
GP_REPORT_COLUMNS = [label for label, _ in GP_HEADER_DEFS]

GP_TEMPLATE_SHEET_NAME = "Import Template"
GP_TEMPLATE_COL = {
    "Doc ID": 1, "First Name": 2, "Middle Name": 3, "Last Name": 4,
    "Entity Type_Employee": 7, "Social Security Number (SSN)": 23,
}


def gp_group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def gp_find_token_sequence_start(line, tokens):
    words = [t for _, _, t in line]
    n = len(tokens)
    for i in range(len(words) - n + 1):
        if words[i:i + n] == tokens:
            return line[i][0]
    for i, w in enumerate(words):
        if w == tokens[0]:
            return line[i][0]
    return None


def gp_find_header_bounds(lines):
    for line in lines:
        texts = [t for _, _, t in line]
        if "EE" in texts and "SSN" in texts:
            positions = {}
            for label, tokens in GP_HEADER_DEFS:
                pos = gp_find_token_sequence_start(line, tokens)
                if pos is not None:
                    positions[label] = pos
            if "EE ID" in positions and "SSN" in positions:
                return positions
    return None


def gp_extract_rows(lines, bounds, page_num):
    ordered = sorted(bounds.items(), key=lambda kv: kv[1])
    col_names = [c for c, _ in ordered]
    col_starts = [x for _, x in ordered]

    records = []
    header_seen = False
    for line in lines:
        texts = [t for _, _, t in line]
        if "EE" in texts and "SSN" in texts:
            header_seen = True
            continue
        if not header_seen:
            continue

        buckets = {name: [] for name in col_names}
        for x0, x1, text in line:
            col = col_names[0]
            for name, start in zip(col_names, col_starts):
                if x0 + 1 >= start:
                    col = name
            buckets[col].append(text)

        row = {name: " ".join(buckets[name]).strip() for name in col_names}
        if GP_EEID_PATTERN.match(row.get("EE ID", "")) or GP_SSN_PATTERN.search(row.get("SSN", "")):
            row["Page"] = page_num
            records.append(row)
    return records


def gp_process_pdf(path: Path):
    doc = fitz.open(path)
    all_records = []
    for i, page in enumerate(doc, start=1):
        rotation = page.rotation
        if rotation != 0:
            print(f"[{path.name}] page {i}: detected rotation {rotation} deg, normalizing for extraction")
        words = page.get_text("words")
        lines = gp_group_words_into_lines(words)
        bounds = gp_find_header_bounds(lines)
        if not bounds:
            print(f"[{path.name}] page {i}: header row not found, skipping")
            continue
        all_records.extend(gp_extract_rows(lines, bounds, i))
    doc.close()
    return all_records


def gp_format_ssn(raw: str) -> str:
    raw = (raw or "").strip()
    match = GP_SSN_PATTERN.search(raw)
    if match:
        return match.group(0)
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 9:
        return f"{digits[0:3]}-{digits[3:5]}-{digits[5:9]}"
    return digits


def gp_split_last_first_middle(name: str):
    name = (name or "").strip()
    if not name:
        return "", "", ""
    if "," in name:
        last, rest = name.split(",", 1)
        parts = rest.strip().split()
        first = parts[0] if parts else ""
        middle = " ".join(parts[1:]) if len(parts) > 1 else ""
        return last.strip(), first, middle
    print(f"  [warn] name '{name}' has no comma ('Last, First' expected) -- placed entirely in Last Name")
    return name, "", ""


def gp_find_template(folder: Path) -> Path:
    candidates = [
        p for p in folder.glob("*.xlsx")
        if not p.name.startswith("~$")
        and "template" in p.name.lower()
        and "_extracted" not in p.name.lower()
    ]
    if not candidates:
        raise FileNotFoundError(
            f"No BDE template file found in {folder} "
            f"(expected an .xlsx with 'template' in its name, e.g. 'BDE Template_Community Health.xlsx')"
        )
    if len(candidates) > 1:
        raise FileNotFoundError(
            f"Multiple possible template files found in {folder}: {[c.name for c in candidates]}. Keep only one."
        )
    return candidates[0]


def gp_build_extracted_workbook(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted"
    headers = ["Page"] + GP_REPORT_COLUMNS
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))
    for rec in records:
        row = [rec.get("Page", "")] + [rec.get(c, "") for c in GP_REPORT_COLUMNS]
        ssn_idx = 1 + GP_REPORT_COLUMNS.index("SSN")
        row[ssn_idx] = gp_format_ssn(row[ssn_idx])
        ws.append([_csv_safe(v) for v in row])
    _autofit(ws)
    return wb


def gp_build_template_workbook(template_path, records):
    wb = openpyxl.load_workbook(template_path)
    ws = wb[GP_TEMPLATE_SHEET_NAME]

    row_ptr = 3
    for rec in records:
        last, first, middle = gp_split_last_first_middle(rec.get("Employee Name", ""))
        ws.cell(row=row_ptr, column=GP_TEMPLATE_COL["Doc ID"], value=rec.get("Doc ID", ""))
        ws.cell(row=row_ptr, column=GP_TEMPLATE_COL["First Name"], value=first)
        ws.cell(row=row_ptr, column=GP_TEMPLATE_COL["Middle Name"], value=middle)
        ws.cell(row=row_ptr, column=GP_TEMPLATE_COL["Last Name"], value=last)
        ws.cell(row=row_ptr, column=GP_TEMPLATE_COL["Entity Type_Employee"], value=True)
        ssn_cell = ws.cell(row=row_ptr, column=GP_TEMPLATE_COL["Social Security Number (SSN)"],
                            value=gp_format_ssn(rec.get("SSN", "")))
        ssn_cell.number_format = "@"
        row_ptr += 1
    return wb


def cmd_gross_pay(args) -> int:
    if not HAS_FITZ:
        sys.exit("ERROR: pymupdf is required. Run: pip install pymupdf")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return 0

    search_dir = input_path.parent if input_path.is_file() else input_path
    if args.template:
        template_path = Path(args.template)
        if not template_path.exists():
            print(f"Template not found: {template_path}")
            return 1
    else:
        try:
            template_path = gp_find_template(search_dir)
        except FileNotFoundError as e:
            print(e)
            return 1
    print(f"Using template: {template_path.name}")

    check_wb = openpyxl.load_workbook(template_path, read_only=True)
    if GP_TEMPLATE_SHEET_NAME not in check_wb.sheetnames:
        print(f"Sheet '{GP_TEMPLATE_SHEET_NAME}' not found in template: {template_path}")
        return 1
    check_wb.close()

    for pdf in tqdm(pdf_files, desc="Extracting", unit="file"):
        records = gp_process_pdf(pdf)
        if not records:
            print(f"  no matching rows extracted from {pdf.name}")
            continue
        for r in records:
            r["Doc ID"] = pdf.name

        extracted_path = _extracted_path(pdf, args.output_dir)
        gp_build_extracted_workbook(records).save(extracted_path)

        template_out_path = _extracted_path(pdf, args.output_dir, "_extracted_template.xlsx")
        gp_build_template_workbook(template_path, records).save(template_out_path)

        print(f"  {pdf.name} -> {extracted_path.name}, {template_out_path.name}")

    return 0


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
    m = K401_AMOUNT_RE.search(_clean(text))
    return m.group(1).replace(',', '') if m else ""


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
            return candidate.title()

    stripped = K401_LABEL_STRIP_RE.sub('', text)
    stripped = re.sub(r'\b\d[\d,.$%-]*\b', '', stripped)
    stripped = re.sub(r'\s+', ' ', stripped).strip()
    m = K401_NAME_RE_UPPER.search(stripped)
    if m:
        candidate = m.group(1)
        if _ok(candidate):
            return candidate.title()

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


def k401_deduplicate_extracted(records: list) -> list:
    if len(records) <= 1:
        return list(records)
    out = []
    for rec in records:
        fn = rec.first_name.strip()
        ln = rec.last_name.strip().lower()
        ssn = rec.ssn.strip()
        matched = False
        for primary in out:
            if primary.ssn.strip() != ssn:
                continue
            if primary.last_name.strip().lower() != ln:
                continue
            if not k401_first_name_prefix_match(fn, primary.first_name.strip()):
                continue
            if len(fn.strip()) > len(primary.first_name.strip()):
                primary.first_name = fn.strip()
                primary.middle_name = rec.middle_name or primary.middle_name
                primary.suffix = rec.suffix or primary.suffix
            else:
                if not primary.middle_name and rec.middle_name:
                    primary.middle_name = rec.middle_name
                if not primary.suffix and rec.suffix:
                    primary.suffix = rec.suffix
            matched = True
            break
        if not matched:
            out.append(rec)
    return out


def k401_parse_name_parts(full_name: str):
    name = _clean(full_name)
    if not name:
        return ("", "", "", "")
    if "," not in name:
        return ("", "", name.title(), "")
    last_part, _, rest_part = name.partition(",")
    last = last_part.strip().title()
    rest = rest_part.strip().split()
    suffix = ""
    if rest and rest[-1].lower().rstrip(".") in K401_SUFFIXES:
        suffix = rest.pop().title()
    first = rest[0].title() if rest else ""
    middle = " ".join(w.title() for w in rest[1:]) if len(rest) > 1 else ""
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
            self.first_name = value.title()
        elif canonical == "middle_name" and not self.middle_name:
            self.middle_name = value.title()
        elif canonical == "last_name" and not self.last_name:
            self.last_name = value.title()
        elif canonical == "suffix" and not self.suffix:
            self.suffix = value.title()
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
    "gross_salary": re.compile(r'(?:gross\s+(?:salary|pay|wages?|comp(?:ensation)?)|annual\s+salary|salary)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "ee_contribution": re.compile(r'(?:ee\s+(?:contribution|deferral)|employee\s+(?:contribution|deferral|401k)|elective\s+deferral)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "er_match": re.compile(r'(?:er\s+match|employer\s+(?:match|contribution)|company\s+match)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "ytd_contribution": re.compile(r'(?:ytd\s+(?:contribution|total|deferral)|year\s+to\s+date)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
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

                w_start = max(0, ssn_line_i - 20)
                w_end = min(len(lines), ssn_line_i + 20)
                window = "\n".join(lines[w_start:w_end])

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
                        if not (0 <= check_i < len(lines)):
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
        print(f"    [!] No employee records found -- check PDF layout.")

    return records


K401_FALLBACK_HEADER = [
    "DOCID", "Last Name", "First Name", "Middle Name", "Suffix",
    "Data Subject Type", "Residential Address", "State of Residence (if US)",
    "Country of Residence", "City", "Province of Residence (if Canada)",
    "Zip Code", "Address Comments", "Phone Number", "Email Address - Personal",
    "PI Notes", "Contact Information", "Government- Issued Identification",
    "Social Security Number", "Passport Number", "Passport Country",
    "Driver's License Number", "DL Issuing Country",
    "DL Issuing Province (if Canada)", "DL Issuing State (if US)",
    "Government-Issued ID Number", "Government ID Issuing Country",
    "Health Related Information", "Birth Information",
    "Full Date of Birth (MM/DD/YYYY)", "Financial Account Information",
    "Access Credentials (Non-Financial Account)", "Biometric Data",
    "Demographic Information", "Family Information",
    "Student-Related Information", "Work-Related Information",
    "Employee Identification Number",
]


def k401_load_template_rows(template_path: Path):
    try:
        wb = openpyxl.load_workbook(str(template_path))
        ws = wb.active
        header_row = [_clean(str(ws.cell(row=1, column=c).value or "")) for c in range(1, K401_TEMPLATE_COL_COUNT + 1)]
        wb.close()
        return header_row
    except Exception:
        return []


def k401_write_sheet(ws, header, records):
    ws.append(header)
    for rec in records:
        ws.append([_csv_safe(v) for v in rec.to_template_row()])
    if not records:
        ws.append(["No employee records extracted."])


def k401_write_workbook(out_path: Path, records, template_path=None):
    tmp = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()

    header_row = k401_load_template_rows(template_path) if (template_path and template_path.exists()) else []
    hdr = header_row if header_row else K401_FALLBACK_HEADER

    ws = wb.active
    ws.title = "Employees"
    k401_write_sheet(ws, hdr, records)

    ws_ext = wb.create_sheet(title="Extracted")
    k401_write_sheet(ws_ext, hdr, k401_deduplicate_extracted(records))

    wb.save(tmp)
    wb.close()
    _atomic_replace(tmp, out_path)


def cmd_401k(args) -> int:
    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required. Run: pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")
    if not HAS_PYPDF:
        print("WARNING: pypdf not installed -- AcroForm extraction skipped. Run: pip install pypdf")

    target_str = args.target
    if not target_str:
        target_str = input("Enter path to PDF file or folder: ").strip().strip('"').strip("'")
    if not target_str:
        sys.exit("ERROR: no path provided.")

    target = Path(target_str)
    if not target.exists():
        sys.exit(f"ERROR: path not found -- {target}")

    if args.template:
        template_path = Path(args.template)
    else:
        script_dir = Path(__file__).parent
        target_dir = target if target.is_dir() else target.parent
        candidates = [script_dir / "Latest Template.xlsx", target_dir / "Latest Template.xlsx"]
        template_path = next((p for p in candidates if p.exists()), None)

    if template_path and template_path.exists():
        print(f"Template: {template_path.name}")
    else:
        print("WARNING: 'Latest Template.xlsx' not found -- using built-in headers. Use --template to specify its path.")
        template_path = None

    files = list(_iter_pdfs(target, args.recursive))
    if not files:
        sys.exit(f"ERROR: no .pdf files found at {target}")

    output_dir = args.output_dir if args.output_dir else (target if target.is_dir() else target.parent)
    print(f"\nFound {len(files)} PDF file(s). Starting extraction...\n")

    total_records = 0
    errors = []
    for pdf_path in tqdm(files, desc="Processing", unit="doc"):
        doc_id = pdf_path.stem
        try:
            records = k401_extract(pdf_path, doc_id)
            out_path = _extracted_path(pdf_path, output_dir)
            k401_write_workbook(out_path, records, template_path)
            total_records += len(records)
        except Exception as exc:
            errors.append(f"{pdf_path.name}: {exc}")

    if errors:
        print("\nErrors:")
        for e in errors:
            print(f"  [!] {e}")

    print(f"\nDone. Total docs: {len(files)} | Processed: {len(files) - len(errors)} | "
          f"Failed: {len(errors)} | Employee records: {total_records}")
    return 0


# ===========================================================================
# creditor-list — creditor names/addresses from bankruptcy mailing-matrix PDFs
# Ported from extract_creditor_list.py (pdfplumber, 3-column grid + gap detection).
# ===========================================================================

CREDITOR_LINE_TOL = 3.0
CREDITOR_GAP_FACTOR = 0.6
CREDITOR_COMPANY_FIRST_NAMES = {"inc", "inc.", "llc", "llc."}
CREDITOR_NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v", "esq", "md", "phd"}
CREDITOR_ENTITY_TOKENS = {"inc", "llc", "corp", "ltd", "co", "ag", "lp", "llp", "plc", "pc", "pa", "na", "nv", "sa", "gmbh"}
CREDITOR_COUNTRY_NAMES = {
    "canada", "mexico", "australia", "united kingdom", "uk", "england",
    "germany", "france", "italy", "spain", "japan", "china", "india",
}
CREDITOR_RE_US_ZIP = re.compile(r'\b([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\s*$')
CREDITOR_RE_CA_POST = re.compile(r'\b([A-Z]{2})\s+([A-Z]\d[A-Z]\s*\d[A-Z]\d)\s*$')


def creditor_is_person(first_name: str) -> bool:
    fn = first_name.strip().lower()
    return bool(fn) and fn not in CREDITOR_COMPANY_FIRST_NAMES


def creditor_parse_name(cell: str) -> dict:
    blank = {"Last Name": "", "First Name": "", "Middle Name": "", "Suffix": ""}
    s = _clean(cell)
    if not s:
        return blank

    tokens = s.split()
    if "," not in s:
        return {**blank, "Last Name": s}

    last_part, _, rest = s.partition(",")
    last_toks = last_part.split()
    rest_toks = rest.split()

    if last_toks and last_toks[-1].strip(".,").lower() in CREDITOR_ENTITY_TOKENS:
        return {**blank, "Last Name": s.rstrip(",")}

    suffix = ""
    if rest_toks and rest_toks[-1].strip(".").lower() in CREDITOR_NAME_SUFFIXES:
        suffix = rest_toks[-1]
        rest_toks = rest_toks[:-1]
    elif last_toks and last_toks[-1].strip(".").lower() in CREDITOR_NAME_SUFFIXES:
        suffix = last_toks[-1]
        last_toks = last_toks[:-1]

    return {
        "Last Name": " ".join(last_toks),
        "First Name": rest_toks[0] if rest_toks else "",
        "Middle Name": " ".join(rest_toks[1:]) if len(rest_toks) > 1 else "",
        "Suffix": suffix,
    }


def creditor_parse_address(lines) -> dict:
    result = {"street": "", "city": "", "state": "", "zipcode": ""}
    clean_lines = [_clean(ln) for ln in lines if _clean(ln)]
    if not clean_lines:
        return result

    if clean_lines[-1].lower() in CREDITOR_COUNTRY_NAMES:
        clean_lines = clean_lines[:-1]
    if not clean_lines:
        return result

    last = clean_lines[-1]
    m = CREDITOR_RE_US_ZIP.search(last) or CREDITOR_RE_CA_POST.search(last)
    if m:
        result["state"] = m.group(1)
        result["zipcode"] = m.group(2).replace(" ", "")
        result["city"] = last[:m.start()].rstrip(", ").strip()
        result["street"] = ", ".join(clean_lines[:-1])
    else:
        result["street"] = ", ".join(clean_lines)
    return result


def creditor_group_lines(words):
    if not words:
        return []
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines = [[ordered[0]]]
    cur_top = ordered[0]["top"]
    for w in ordered[1:]:
        if abs(w["top"] - cur_top) <= CREDITOR_LINE_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])
            cur_top = w["top"]
    for ln in lines:
        ln.sort(key=lambda w: w["x0"])
    return lines


def creditor_remove_page_headers(lines, page_width: float):
    threshold = page_width * 0.55
    return [
        ln for ln in lines
        if ln and (max(w["x1"] for w in ln) - min(w["x0"] for w in ln)) <= threshold
    ]


def creditor_assign_columns(lines, n_cols=3):
    all_x0 = [w["x0"] for ln in lines for w in ln]
    all_x1 = [w["x1"] for ln in lines for w in ln]
    if not all_x0:
        return [[] for _ in range(n_cols)]

    page_left = min(all_x0)
    page_right = max(all_x1)
    col_width = (page_right - page_left) / n_cols

    buckets = [defaultdict(list) for _ in range(n_cols)]
    for ln in lines:
        for w in ln:
            cx = (w["x0"] + w["x1"]) / 2
            col_idx = min(int((cx - page_left) / col_width), n_cols - 1)
            buckets[col_idx][w["top"]].append(w)

    result = []
    for ci in range(n_cols):
        col_lines = []
        for top_y in sorted(buckets[ci]):
            col_lines.append(sorted(buckets[ci][top_y], key=lambda w: w["x0"]))
        result.append(col_lines)
    return result


def creditor_parse_entries_by_gap(col_lines, page_num):
    if not col_lines:
        return []

    tops = [min(w["top"] for w in ln) for ln in col_lines]
    bottoms = [max(w["bottom"] for w in ln) for ln in col_lines]
    heights = [b - t for t, b in zip(tops, bottoms)]
    avg_height = sum(heights) / len(heights) if heights else 12.0
    threshold = avg_height * CREDITOR_GAP_FACTOR

    groups = [[col_lines[0]]]
    for i in range(len(col_lines) - 1):
        gap = tops[i + 1] - bottoms[i]
        if gap > threshold:
            groups.append([col_lines[i + 1]])
        else:
            groups[-1].append(col_lines[i + 1])

    entries = []
    for group in groups:
        text_lines = [_clean(" ".join(w["text"] for w in ln)) for ln in group]
        text_lines = [l for l in text_lines if l]
        if not text_lines:
            continue
        name = text_lines[0]
        addr_lines = text_lines[1:]
        addr = creditor_parse_address(addr_lines)
        full_addr = ", ".join(l for l in addr_lines if l)
        entries.append({"full_name": name, "address": full_addr, "page_num": page_num, **addr})

    return entries


def creditor_extract_pdf(pdf_path: Path):
    all_entries = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if not words:
                continue
            lines = creditor_group_lines(words)
            lines = creditor_remove_page_headers(lines, float(page.width))
            columns = creditor_assign_columns(lines, n_cols=3)
            for col_lines in columns:
                all_entries.extend(creditor_parse_entries_by_gap(col_lines, page_num))
    return all_entries


CREDITOR_PARSED_HEADERS = [
    "First Name", "Last Name", "Middle Name", "Suffix",
    "Street Address", "City", "State", "ZIP Code", "Page Number",
]


def creditor_parsed_row(e: dict) -> list:
    nm = creditor_parse_name(e["full_name"])
    return [
        _csv_safe(nm["First Name"]), _csv_safe(nm["Last Name"]), _csv_safe(nm["Middle Name"]), _csv_safe(nm["Suffix"]),
        _csv_safe(e["street"]), _csv_safe(e["city"]), _csv_safe(e["state"]), _csv_safe(e["zipcode"]), e["page_num"],
    ]


def creditor_write_sheet1(wb, entries):
    ws = wb.active
    ws.title = "extracted"
    ws.row_dimensions[1].height = 28
    for col, h in enumerate(["Full Name", "Address Information", "Page Number"], 1):
        ws.cell(row=1, column=col, value=h)
    _style_header_row(ws, 1, 3)
    for r, e in enumerate(entries, 2):
        ws.cell(row=r, column=1, value=_csv_safe(e["full_name"]))
        ws.cell(row=r, column=2, value=_csv_safe(e["address"]))
        ws.cell(row=r, column=3, value=e["page_num"])
    _autofit_widths(ws, [42, 62, 14])


def creditor_write_sheet2(wb, entries):
    ws = wb.create_sheet("parsed names addresses")
    ws.row_dimensions[1].height = 28
    for col, h in enumerate(CREDITOR_PARSED_HEADERS, 1):
        ws.cell(row=1, column=col, value=h)
    _style_header_row(ws, 1, len(CREDITOR_PARSED_HEADERS))
    for r, e in enumerate(entries, 2):
        for col, v in enumerate(creditor_parsed_row(e), 1):
            ws.cell(row=r, column=col, value=v)
    _autofit_widths(ws, [22, 35, 16, 10, 45, 22, 10, 14, 14])


def creditor_write_sheet3(wb, entries):
    ws = wb.create_sheet("standard data")
    ws.row_dimensions[1].height = 28
    for col, h in enumerate(CREDITOR_PARSED_HEADERS, 1):
        ws.cell(row=1, column=col, value=h)
    _style_header_row(ws, 1, len(CREDITOR_PARSED_HEADERS))
    r = 2
    for e in entries:
        row = creditor_parsed_row(e)
        first_name = row[0]
        if not creditor_is_person(first_name):
            continue
        for col, v in enumerate(row, 1):
            ws.cell(row=r, column=col, value=v)
        r += 1
    _autofit_widths(ws, [22, 35, 16, 10, 45, 22, 10, 14, 14])


def creditor_write_workbook(out_path: Path, entries):
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    creditor_write_sheet1(wb, entries)
    creditor_write_sheet2(wb, entries)
    creditor_write_sheet3(wb, entries)
    wb.save(tmp_path)
    wb.close()
    _atomic_replace(tmp_path, out_path)


def cmd_creditor_list(args) -> int:
    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required. Run: pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")

    target_str = args.target
    if not target_str:
        target_str = input("Enter path to PDF file or folder: ").strip().strip("\"'")
    if not target_str:
        sys.exit("ERROR: no path provided.")

    target = Path(target_str)
    if not target.exists():
        sys.exit(f"ERROR: path not found -- {target}")

    pdfs = list(_iter_pdfs(target, args.recursive))
    if not pdfs:
        sys.exit(f"ERROR: no PDF files found at {target}")

    print(f"Found {len(pdfs)} PDF file(s). Extracting...")

    for pdf_path in pdfs:
        print(f"  Processing: {pdf_path.name}")
        try:
            entries = creditor_extract_pdf(pdf_path)
            out_path = _extracted_path(pdf_path, args.output_dir)
            if entries:
                person_count = sum(1 for e in entries if creditor_is_person(creditor_parse_name(e["full_name"])["First Name"]))
                creditor_write_workbook(out_path, entries)
                print(f"    {len(entries)} entr(y/ies) extracted -> {out_path.name} "
                      f"({person_count} individual(s), companies filtered out of standard data)")
            else:
                print(f"    no entries extracted -- no output written")
        except Exception as exc:
            print(f"    ERROR: {exc}")

    return 0


# ===========================================================================
# generic — dynamically extract headers/tables/form-fields from any PDF
# Ported from extract_pdf.py (pdfplumber + pypdf AcroForm).
# ===========================================================================

GENERIC_MEDIAN_SIZE_FALLBACK = 10.0
GENERIC_OUTPUT_COLS = ["Source File", "Page", "Category", "Field Name", "Value"]
GENERIC_FT_LABELS = {"/Tx": "Text", "/Ch": "Dropdown / List", "/Btn": "Checkbox / Radio", "/Sig": "Signature"}


def generic_is_mostly_upper(text: str) -> bool:
    letters = [c for c in text if c.isalpha()]
    if len(letters) < 3:
        return False
    return sum(1 for c in letters if c.isupper()) / len(letters) >= 0.75


def generic_looks_like_label(text: str) -> bool:
    stripped = text.strip().rstrip(":")
    if len(stripped.split()) > 8:
        return False
    if text.strip().endswith(":"):
        return True
    if generic_is_mostly_upper(stripped) and len(stripped.split()) <= 5:
        return True
    return False


def generic_median(values):
    if not values:
        return GENERIC_MEDIAN_SIZE_FALLBACK
    sorted_v = sorted(values)
    mid = len(sorted_v) // 2
    return sorted_v[mid]


def generic_word_font_sizes(words):
    sizes = []
    for w in words:
        try:
            s = float(w.get("size") or 0)
            if s > 0:
                sizes.append(s)
        except (TypeError, ValueError):
            pass
    return sizes


def generic_is_bold(word: dict) -> bool:
    font_name = str(word.get("fontname") or "").lower()
    return "bold" in font_name or "bd" in font_name


def generic_word_is_header(word: dict, median_size: float) -> bool:
    try:
        size = float(word.get("size") or 0)
    except (TypeError, ValueError):
        size = 0.0
    if size > median_size * 1.15:
        return True
    return generic_is_bold(word)


def generic_word_in_bbox(word: dict, bboxes) -> bool:
    try:
        wx = (float(word["x0"]) + float(word["x1"])) / 2
        wy = (float(word["top"]) + float(word["bottom"])) / 2
    except (KeyError, TypeError, ValueError):
        return False
    for (x0, top, x1, bottom) in bboxes:
        if x0 <= wx <= x1 and top <= wy <= bottom:
            return True
    return False


def generic_extract_page_sections(page, table_bboxes):
    all_words = page.extract_words(extra_attrs=["fontname", "size"])
    words = [w for w in all_words if not generic_word_in_bbox(w, table_bboxes)]

    sizes = generic_word_font_sizes(words)
    median_size = generic_median(sizes) if sizes else GENERIC_MEDIAN_SIZE_FALLBACK

    lines_by_y = defaultdict(list)
    for w in words:
        y_bucket = round(float(w.get("top") or 0) / 2) * 2
        lines_by_y[y_bucket].append(w)

    ordered_lines = [lines_by_y[y] for y in sorted(lines_by_y)]

    sections = []
    current_header = ""
    current_value_parts = []

    def flush():
        nonlocal current_header, current_value_parts
        if current_header:
            value = _clean(" ".join(current_value_parts))
            sections.append({"header": current_header, "value": value})
        current_header = ""
        current_value_parts = []

    for line_words in ordered_lines:
        line_text = _clean(" ".join(w.get("text", "") for w in line_words))
        if not line_text:
            continue

        header_word_count = sum(1 for w in line_words if generic_word_is_header(w, median_size))
        line_is_header = len(line_words) > 0 and header_word_count / len(line_words) >= 0.6

        if not line_is_header and not any(w.get("size") for w in line_words):
            line_is_header = generic_looks_like_label(line_text)

        if line_is_header:
            flush()
            current_header = line_text.rstrip(":")
        elif ":" in line_text and not current_header:
            label, _, value = line_text.partition(":")
            if generic_looks_like_label(label):
                sections.append({"header": label.strip(), "value": value.strip()})
                continue
            current_value_parts.append(line_text)
        else:
            inline_pairs = re.findall(r'([A-Za-z][A-Za-z0-9 ]{1,30}):\s*([^\|;]+)', line_text)
            if inline_pairs and not current_header:
                for label, value in inline_pairs:
                    sections.append({"header": label.strip(), "value": value.strip()})
            else:
                current_value_parts.append(line_text)

    flush()
    return sections


def generic_is_header_row(row) -> bool:
    non_empty = [c for c in row if c.strip()]
    if not non_empty:
        return False
    label_count = sum(
        1 for c in non_empty
        if len(c.split()) <= 5 and not re.search(r'\d{4,}', c)
        and (c[0].isupper() or generic_is_mostly_upper(c))
    )
    return label_count / len(non_empty) >= 0.6


def generic_extract_page_tables(page):
    bboxes = []
    records = []
    try:
        found_tables = page.find_tables()
    except Exception:
        found_tables = []

    for tbl_obj in (found_tables or []):
        try:
            bb = tbl_obj.bbox
            bboxes.append(bb)
        except AttributeError:
            pass

        try:
            raw = tbl_obj.extract()
        except Exception:
            continue
        if not raw:
            continue

        cleaned_rows = []
        for row in raw:
            clean_row = [_clean(str(cell or "")) for cell in row]
            if any(clean_row):
                cleaned_rows.append(clean_row)
        if not cleaned_rows:
            continue

        if len(cleaned_rows) >= 2 and generic_is_header_row(cleaned_rows[0]):
            headers = cleaned_rows[0]
            data_rows = cleaned_rows[1:]
        else:
            n_cols = max(len(r) for r in cleaned_rows)
            headers = [f"Column_{i+1}" for i in range(n_cols)]
            data_rows = cleaned_rows

        n = len(headers)
        data_rows = [r + [""] * (n - len(r)) for r in data_rows]
        records.append({"headers": headers, "rows": data_rows})

    return bboxes, records


def generic_normalise_value(raw, type_label: str) -> str:
    if raw is None:
        return "No" if type_label == "Checkbox / Radio" else ""
    s = str(raw).strip()
    if type_label == "Checkbox / Radio":
        off_values = {"", "off", "/off", "no", "/no", "false", "/false"}
        return "No" if s.lower() in off_values else "Yes"
    if s.startswith("/"):
        s = s[1:]
    return _clean(s)


def generic_resolve(obj):
    return obj.get_object() if hasattr(obj, "get_object") else obj


def generic_page_idnum_map(reader):
    result = {}
    for i, page in enumerate(reader.pages, start=1):
        try:
            result[page.indirect_reference.idnum] = i
        except Exception:
            pass
    return result


def generic_get_widget_annotations(reader):
    page_map = generic_page_idnum_map(reader)
    widgets = []

    def _traverse(field_ref, inherited):
        try:
            field = generic_resolve(field_ref)
        except Exception:
            return

        attrs = dict(inherited)
        for key in ("/FT", "/TU", "/V", "/DV"):
            try:
                val = field.get(key)
                if val is not None:
                    attrs[key] = val
            except Exception:
                pass
        try:
            t_val = field.get("/T")
            if t_val is not None:
                attrs["_T"] = _clean(str(t_val))
        except Exception:
            pass

        try:
            kids = field.get("/Kids")
            if kids:
                for kid_ref in kids:
                    _traverse(kid_ref, attrs)
        except Exception:
            pass

        try:
            rect = field.get("/Rect")
            if rect is None:
                return
            rect_vals = [float(x) for x in generic_resolve(rect)]
        except Exception:
            return

        page_num = None
        try:
            p_ref = field.get("/P")
            if p_ref is not None:
                page_num = page_map.get(generic_resolve(p_ref).indirect_reference.idnum)
        except Exception:
            pass
        if page_num is None:
            return

        widgets.append({
            "page_num": page_num, "rect": rect_vals,
            "field_type": str(attrs.get("/FT") or ""),
            "field_name": attrs.get("_T") or "",
            "alt_name": _clean(str(attrs.get("/TU") or "")),
            "raw_value": attrs.get("/V"),
        })

    try:
        root = generic_resolve(reader.trailer["/Root"])
        acroform = generic_resolve(root["/AcroForm"])
        fields = acroform.get("/Fields") or []
        for f_ref in fields:
            _traverse(f_ref, {})
    except Exception:
        pass

    if widgets:
        return widgets

    for page_num, page in enumerate(reader.pages, start=1):
        try:
            annots_obj = page.get("/Annots")
            if annots_obj is None:
                continue
            annots_obj = generic_resolve(annots_obj)
        except Exception:
            continue

        for ref in annots_obj:
            try:
                annot = ref.get_object() if hasattr(ref, "get_object") else ref
            except Exception:
                continue
            if annot.get("/Subtype") != "/Widget":
                continue
            rect = annot.get("/Rect")
            if not rect:
                continue

            def _inh(key, node=annot):
                val = node.get(key)
                if val is not None:
                    return val
                try:
                    return _inh(key, generic_resolve(node["/Parent"]))
                except Exception:
                    return None

            try:
                rect_vals = [float(x) for x in rect]
            except Exception:
                continue

            widgets.append({
                "page_num": page_num, "rect": rect_vals,
                "field_type": str(_inh("/FT") or ""),
                "field_name": _clean(str(_inh("/T") or "")),
                "alt_name": _clean(str(_inh("/TU") or "")),
                "raw_value": _inh("/V"),
            })

    return widgets


def generic_text_lines_by_page(pdf_path: Path):
    result = {}
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            lines_by_y = defaultdict(list)
            for w in words:
                y_key = round(float(w.get("top", 0)) / 3) * 3
                lines_by_y[y_key].append(w)

            lines = []
            for y_key in sorted(lines_by_y):
                grp = sorted(lines_by_y[y_key], key=lambda w: float(w.get("x0", 0)))
                text = _clean(" ".join(w.get("text", "") for w in grp))
                if text:
                    lines.append({
                        "text": text, "x0": float(min(w.get("x0", 0) for w in grp)),
                        "x1": float(max(w.get("x1", 0) for w in grp)),
                        "top": float(min(w.get("top", 0) for w in grp)),
                        "bottom": float(max(w.get("bottom", 0) for w in grp)),
                    })
            result[page_num] = lines
    return result


def generic_find_label(field_rect_plumber, page_lines):
    fx0, ftop, fx1, fbottom = field_rect_plumber
    fy_mid = (ftop + fbottom) / 2
    fx_mid = (fx0 + fx1) / 2
    fwidth = fx1 - fx0

    candidates = []
    for line in page_lines:
        lx0, lx1 = line["x0"], line["x1"]
        ltop, lbottom = line["top"], line["bottom"]
        ly_mid = (ltop + lbottom) / 2
        lx_mid = (lx0 + lx1) / 2
        text = line["text"].rstrip(":").strip()
        if not text:
            continue

        if lx1 <= fx0 + 5 and abs(ly_mid - fy_mid) <= 14:
            dist = (fx0 - lx1) + abs(ly_mid - fy_mid) * 0.5
            candidates.append((dist, text))
        elif lbottom <= ftop + 4 and lx0 < fx1 + 10 and lx1 > fx0 - 10:
            dist = (ftop - lbottom) + abs(lx_mid - fx_mid) * 0.3
            candidates.append((dist, text))
        elif lbottom <= ftop + 4 and abs(lx_mid - fx_mid) <= fwidth * 0.9:
            dist = (ftop - lbottom) * 1.5 + abs(lx_mid - fx_mid)
            candidates.append((dist, text))

    if not candidates:
        return ""
    candidates.sort(key=lambda c: c[0])
    return candidates[0][1]


def generic_extract_form_fields(pdf_path: Path):
    if not HAS_PYPDF:
        return []
    source = pdf_path.name
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return []

    widgets = generic_get_widget_annotations(reader)
    if not widgets:
        return []

    page_heights = {}
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for i, p in enumerate(pdf.pages, start=1):
                page_heights[i] = float(p.height)
    except Exception:
        pass

    try:
        lines_by_page = generic_text_lines_by_page(pdf_path)
    except Exception:
        lines_by_page = {}

    results = []
    for w in widgets:
        page_num = w["page_num"]
        type_label = GENERIC_FT_LABELS.get(w["field_type"], "Field")
        value = generic_normalise_value(w["raw_value"], type_label)

        ph = page_heights.get(page_num, 792.0)
        llx, lly, urx, ury = w["rect"]
        field_plumber = (llx, ph - ury, urx, ph - lly)

        label = generic_find_label(field_plumber, lines_by_page.get(page_num, []))
        if not label:
            label = w["alt_name"] or w["field_name"]

        results.append({
            "source": source, "page": str(page_num), "field_name": label,
            "field_type": type_label, "value": value,
        })

    return results


def generic_extract_pdf(pdf_path: Path):
    source = pdf_path.name
    all_sections = []
    all_tables = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            bboxes, tbl_records = generic_extract_page_tables(page)
            page_sections = generic_extract_page_sections(page, bboxes)
            for sec in page_sections:
                all_sections.append({"source": source, "page": page_num, "header": sec["header"], "value": sec["value"]})
            for t_idx, rec in enumerate(tbl_records, start=1):
                all_tables.append({"source": source, "page": page_num, "table_num": t_idx,
                                    "headers": rec["headers"], "rows": rec["rows"]})

    form_fields = generic_extract_form_fields(pdf_path)
    return all_sections, all_tables, form_fields


def generic_build_unified_rows(all_sections, all_tables, all_form_fields):
    rows = []
    form_field_pages = {(ff.get("source", ""), str(ff.get("page", ""))) for ff in all_form_fields}

    for ff in all_form_fields:
        try:
            page_key = int(ff.get("page") or 0)
        except (ValueError, TypeError):
            page_key = 0
        data = [_csv_safe(ff.get("source", "")), _csv_safe(ff.get("page", "")), "Form Field",
                _csv_safe(ff.get("field_name", "")), _csv_safe(ff.get("value", ""))]
        rows.append((page_key, 0, data))

    for sec in all_sections:
        page_key = int(sec.get("page") or 0)
        value = sec.get("value", "")
        source = sec.get("source", "")
        page_str = str(sec.get("page", ""))
        if not value and (source, page_str) in form_field_pages:
            continue
        data = [_csv_safe(source), _csv_safe(page_str), "Section", _csv_safe(sec.get("header", "")), _csv_safe(value)]
        rows.append((page_key, 1, data))

    for tbl in all_tables:
        page_key = int(tbl.get("page") or 0)
        source = tbl.get("source", "")
        page = tbl.get("page", "")
        headers = tbl.get("headers", [])
        for r_idx, row in enumerate(tbl.get("rows", []), start=1):
            pairs = " | ".join(f"{h}: {v}" for h, v in zip(headers, row) if h or v)
            label = f"Row {r_idx} (Table {tbl.get('table_num', '')})"
            data = [_csv_safe(source), _csv_safe(page), "Table", label, _csv_safe(pairs)]
            rows.append((page_key, 2, data))

    rows.sort(key=lambda x: (x[0], x[1]))
    return [r[2] for r in rows]


def generic_write_workbook(out_path: Path, all_sections, all_tables, all_form_fields):
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Extracted Data"
    ws.append(GENERIC_OUTPUT_COLS)
    ws.row_dimensions[1].height = 18
    _style_header_row(ws, 1, len(GENERIC_OUTPUT_COLS))

    unified = generic_build_unified_rows(all_sections, all_tables, all_form_fields)
    alt_fill = PatternFill("solid", fgColor="DCE6F1")
    for i, row in enumerate(unified, start=2):
        ws.append(row)
        if i % 2 == 0:
            for col in range(1, len(GENERIC_OUTPUT_COLS) + 1):
                ws.cell(row=i, column=col).fill = alt_fill

    if not unified:
        ws.cell(row=2, column=1, value="No data extracted from the processed PDF(s).")

    _autofit(ws)
    wb.save(tmp_path)
    wb.close()
    _atomic_replace(tmp_path, out_path)


def cmd_generic(args) -> int:
    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required. Run: pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. Run: pip install openpyxl")
    if not HAS_PYPDF:
        print("WARNING: pypdf not installed -- form field values will be skipped. Run: pip install pypdf")

    target_str = args.target
    if not target_str:
        target_str = input("Enter path to PDF file or folder: ").strip().strip('"').strip("'")
    if not target_str:
        sys.exit("ERROR: no path provided.")

    target = Path(target_str)
    if not target.exists():
        sys.exit(f"ERROR: path not found -- {target}")

    files = list(_iter_pdfs(target, args.recursive))
    if not files:
        sys.exit(f"ERROR: no .pdf files found at {target}")

    print(f"Found {len(files)} PDF file(s). Extracting...")

    for pdf_path in files:
        print(f"  Processing: {pdf_path.name}")
        try:
            sections, tables, form_fields = generic_extract_pdf(pdf_path)
            out_path = _extracted_path(pdf_path, args.output_dir)
            generic_write_workbook(out_path, sections, tables, form_fields)
            table_rows = sum(len(t['rows']) for t in tables)
            print(f"    {len(sections)} section(s), {len(tables)} table(s) ({table_rows} row(s)), "
                  f"{len(form_fields)} form field(s) -> {out_path.name}")
        except Exception as exc:
            print(f"    ERROR: {exc}")

    return 0


# ===========================================================================
# bucket11 — Employee Payroll Changes from ADP "Employee Payroll Changes" PDFs
# Ported from bucket_11.py (pdfplumber, 3-column table + fuzzy header mapping).
# ===========================================================================

import difflib as _difflib

B11_OUTPUT_HEADERS = [
    "DOCID", "Name", "Associate ID", "Position ID", "Home Department",
    "Home Cost Number", "Changed Field", "Changed From", "Changed To",
]

B11_LINE_TOLERANCE = 3.0
B11_NAME_SUFFIXES = {"jr", "sr", "ii", "iii", "iv", "v"}
B11_DEFAULT_COL2_FRAC = 0.38
B11_DEFAULT_COL3_FRAC = 0.67
B11_NO_RECORDS = (
    "No employee records found -- no 'Associate ID:' header rows detected in "
    "this PDF. Run --diagnose to inspect the raw text."
)

B11_BUCKET11_MAP = {
    "Additional Earnings Amount":       (None, "Work-Related Information", "Salary or Compensation Information"),
    "Address - City":                   ("City", "Contact Information", "Home address"),
    "Address - Country":                ("Country of Residence", "Contact Information", "Home address"),
    "Address - Line 1":                 ("Residential Address", "Contact Information", "Home address"),
    "Address - Line 2":                 ("Residential Address", "Contact Information", "Home address"),
    "Address - Line 2 No":              ("Residential Address", "Contact Information", "Home address"),
    "Address - Line 3":                 ("Residential Address", "Contact Information", "Home address"),
    "Address - State":                  ("State of Residence (if US)", "Contact Information", "Home address"),
    "Address - State No":               ("State of Residence (if US)", "Contact Information", "Home address"),
    "Address - Zip / Postal Code":      ("Zip Code", "Contact Information", "Home address"),
    "Address - Zip / Postal Code No":   ("Zip Code", "Contact Information", "Home address"),
    "Basis of Pay":                     (None, "Work-Related Information", "Salary or Compensation Information"),
    "Basis of Pay No":                  (None, "Work-Related Information", "Salary or Compensation Information"),
    "Defer Social Security tax":        (None, "Government-Issued Identification", "Taxpayer Identification Number (TIN)"),
    "Dependents No":                    (None, "Family Information", None),
    "Direct Deposit - Account Number":  (None, "Financial Account Information", "Financial account number"),
    "Employee Name - First":            ("First Name", None, None),
    "Employee Name - First No":         ("First Name", None, None),
    "Employee Name - Last":             ("Last Name", None, None),
    "Employee Name - Last No":          ("Last Name", None, None),
    "Employee Name - Middle":           ("Middle Name", None, None),
    "Employee Name - Preferred":        ("PI Notes", None, None),
    "Employee Name - Salutation":       ("Suffix", None, None),
    "Ethnicity/Race":                   (None, "Demographic Information", "Race/ Ethnicity"),
    "Ethnicity/Race No":                (None, "Demographic Information", "Race/ Ethnicity"),
    "Gender for Insurance Coverage No": (None, "Health Related Information", "Health Insurance Information"),
    "Home Phone":                       ("Phone Number", "Contact Information", "Personal phone number (home)"),
    "Job Title No":                     (None, "Work-Related Information", "Employment Application Information"),
    "Lien Dependent Medical Insurance": (None, "Health Related Information", "Health Insurance Information"),
    "Other Income":                     (None, "Work-Related Information", "Salary or Compensation Information"),
    "Payee":                            (None, "Financial Account Information", "Financial account number"),
    "Payroll Name - First":             ("PI Notes", None, None),
    "Payroll Name - First No":          ("PI Notes", None, None),
    "Payroll Name - Last":              ("PI Notes", None, None),
    "Payroll Name - Last No":           ("PI Notes", None, None),
    "Personal E-mail":                  ("Email Address - Personal", "Contact Information", "Personal email address"),
    "Personal E-mail No":               ("Email Address - Personal", "Contact Information", "Personal email address"),
    "Personal Mobile":                  ("Phone Number", "Contact Information", "Personal phone number (mobile)"),
    "Personal Mobile No":               ("Phone Number", "Contact Information", "Personal phone number (mobile)"),
    "Rate 1":                           (None, "Work-Related Information", "Salary or Compensation Information"),
    "Rate 1 No":                        (None, "Work-Related Information", "Salary or Compensation Information"),
    "Rate 2":                           (None, "Work-Related Information", "Salary or Compensation Information"),
    "Rate 6":                           (None, "Work-Related Information", "Salary or Compensation Information"),
    "Rate 8":                           (None, "Work-Related Information", "Salary or Compensation Information"),
    "Social Security Number":           ("Social Security Number", "Government-Issued Identification", "Social Security Number (SSN)"),
    "Status":                           (None, "Work-Related Information", "Employment Application Information"),
    "Tax ID Type":                      (None, "Government-Issued Identification", "Taxpayer Identification Number (TIN)"),
    "Termination Date":                 (None, "Work-Related Information", "Employment Application Information"),
    "Termination Reason":               (None, "Work-Related Information", "Disciplinary Record or Report"),
}

B11_ALIASES = {
    "ssn": "Social Security Number", "ss number": "Social Security Number",
    "social security": "Social Security Number", "soc sec": "Social Security Number",
    "ss no": "Social Security Number",
    "first name": "Employee Name - First", "last name": "Employee Name - Last",
    "middle name": "Employee Name - Middle", "preferred name": "Employee Name - Preferred",
    "salutation": "Employee Name - Salutation", "payroll first name": "Payroll Name - First",
    "payroll last name": "Payroll Name - Last",
    "email": "Personal E-mail", "e mail": "Personal E-mail", "personal email": "Personal E-mail",
    "work email": "Personal E-mail", "mobile": "Personal Mobile", "cell": "Personal Mobile",
    "cell phone": "Personal Mobile", "phone": "Home Phone", "home phone": "Home Phone",
    "telephone": "Home Phone",
    "city": "Address - City", "state": "Address - State", "zip": "Address - Zip / Postal Code",
    "zip code": "Address - Zip / Postal Code", "postal code": "Address - Zip / Postal Code",
    "postcode": "Address - Zip / Postal Code", "country": "Address - Country",
    "address": "Address - Line 1", "address line 1": "Address - Line 1",
    "address line 2": "Address - Line 2", "address line 3": "Address - Line 3", "street": "Address - Line 1",
    "pay rate": "Rate 1", "hourly rate": "Rate 1", "rate": "Rate 1", "wage": "Rate 1", "wages": "Rate 1",
    "hourly wage": "Rate 1", "salary rate": "Rate 1", "base pay": "Basis of Pay", "salary": "Basis of Pay",
    "annual salary": "Basis of Pay", "base salary": "Basis of Pay", "gross pay": "Basis of Pay",
    "pay basis": "Basis of Pay", "earnings": "Additional Earnings Amount",
    "additional pay": "Additional Earnings Amount", "other pay": "Other Income", "income": "Other Income",
    "other income": "Other Income", "gross income": "Other Income", "annual income": "Other Income",
    "net income": "Other Income",
    "gender": "Gender for Insurance Coverage No", "sex": "Gender for Insurance Coverage No",
    "race": "Ethnicity/Race", "ethnicity": "Ethnicity/Race", "race ethnicity": "Ethnicity/Race",
    "job title": "Job Title No", "title": "Job Title No", "status": "Status",
    "employment status": "Status", "term date": "Termination Date", "termination": "Termination Date",
    "separation date": "Termination Date", "term reason": "Termination Reason",
    "separation reason": "Termination Reason", "reason for termination": "Termination Reason",
    "direct deposit": "Direct Deposit - Account Number", "bank account": "Direct Deposit - Account Number",
    "account number": "Direct Deposit - Account Number", "dd account": "Direct Deposit - Account Number",
    "payee": "Payee",
    "tax id": "Tax ID Type", "tin": "Tax ID Type", "taxpayer id": "Tax ID Type",
    "defer ss": "Defer Social Security tax", "defer soc sec": "Defer Social Security tax",
    "dependents": "Dependents No", "dependent count": "Dependents No",
    "medical insurance": "Lien Dependent Medical Insurance", "lien": "Lien Dependent Medical Insurance",
}


def b11_normalise(text: str) -> str:
    return " ".join(str(text or "").split()).lower()


def b11_parse_name(cell: str) -> dict:
    blank = {"Last Name": "", "First Name": "", "Middle Name": "", "Suffix": ""}
    s = _clean(cell)
    if not s or "," not in s:
        return blank

    last_part, _, rest = s.partition(",")
    last_toks = last_part.split()
    rest_toks = rest.split()

    suffix = ""
    if rest_toks and rest_toks[-1].strip(".").lower() in B11_NAME_SUFFIXES:
        suffix = rest_toks[-1]
        rest_toks = rest_toks[:-1]
    elif last_toks and last_toks[-1].strip(".").lower() in B11_NAME_SUFFIXES:
        suffix = last_toks[-1]
        last_toks = last_toks[:-1]

    return {
        "Last Name": " ".join(last_toks),
        "First Name": rest_toks[0] if rest_toks else "",
        "Middle Name": " ".join(rest_toks[1:]) if len(rest_toks) > 1 else "",
        "Suffix": suffix,
    }


def b11_group_lines(words):
    if not words:
        return []
    ordered = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))
    lines = []
    current = [ordered[0]]
    current_top = ordered[0]["top"]
    for w in ordered[1:]:
        if abs(w["top"] - current_top) <= B11_LINE_TOLERANCE:
            current.append(w)
        else:
            lines.append(sorted(current, key=lambda x: x["x0"]))
            current = [w]
            current_top = w["top"]
    lines.append(sorted(current, key=lambda x: x["x0"]))
    return lines


def b11_line_text(line) -> str:
    return " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]))


def b11_detect_boundaries(lines, page_width: float):
    for line in lines:
        norm = b11_normalise(b11_line_text(line))
        if "changed" in norm and "from" in norm and "to" in norm:
            words = sorted(line, key=lambda w: w["x0"])
            from_x = to_x = None
            for i, w in enumerate(words):
                t = w["text"].lower()
                if t == "changed" and i + 1 < len(words):
                    nxt = words[i + 1]["text"].lower()
                    if nxt == "from":
                        from_x = w["x0"]
                    elif nxt == "to":
                        to_x = w["x0"]
                elif t == "from" and from_x is None:
                    from_x = w["x0"]
                elif t == "to" and to_x is None:
                    to_x = w["x0"]
            if from_x and to_x and from_x < to_x:
                return from_x, to_x
    return page_width * B11_DEFAULT_COL2_FRAC, page_width * B11_DEFAULT_COL3_FRAC


def b11_is_table_header(line) -> bool:
    norm = b11_normalise(b11_line_text(line))
    return "changed" in norm and "from" in norm and ("field" in norm or "to" in norm)


def b11_is_employee_header(line) -> bool:
    norm = b11_normalise(b11_line_text(line))
    return ("associate id" in norm or "position id" in norm
            or "home department" in norm or "home cost number" in norm)


def b11_is_page_chrome(line) -> bool:
    words = [w["text"].lower() for w in line]
    norm = " ".join(words)
    if re.match(r"page\s+\d+\s+of\s+\d+", norm):
        return True
    if norm.strip() in {"employee payroll changes", "payroll changes"}:
        return True
    return False


def b11_is_name_line(line, col2_x: float) -> bool:
    if b11_is_table_header(line) or b11_is_page_chrome(line) or b11_is_employee_header(line):
        return False
    col1_text = " ".join(w["text"] for w in sorted(line, key=lambda w: w["x0"]) if w["x0"] < col2_x)
    return bool(re.search(r"[A-Za-z][A-Za-z'\-\.]*\s*,\s*[A-Za-z]", col1_text))


def b11_parse_employee_header(line_text: str) -> dict:
    result = {
        "Last Name": "", "First Name": "", "Middle Name": "",
        "Associate ID": "", "Position ID": "", "Home Department": "", "Home Cost Number": "",
    }
    text = _clean(line_text)

    assoc_m = re.search(r"Associate\s+ID\s*:\s*(\S+)", text, re.I)
    pos_m = re.search(r"Position\s+ID\s*:\s*(\S+)", text, re.I)
    dept_m = re.search(r"Home\s+Department\s*:\s*(.+?)(?=\s{2,}Home\s+Cost|\s*Home\s+Cost|$)", text, re.I)
    cost_m = re.search(r"Home\s+Cost\s+Number\s*:\s*(\S*)", text, re.I)

    if assoc_m:
        result["Associate ID"] = assoc_m.group(1).strip()
    if pos_m:
        result["Position ID"] = pos_m.group(1).strip()
    if dept_m:
        result["Home Department"] = _clean(dept_m.group(1))
    if cost_m:
        result["Home Cost Number"] = cost_m.group(1).strip()

    header_labels = re.compile(r"Associate\s+ID|Position\s+ID|Home\s+Department|Home\s+Cost", re.I)
    split_m = header_labels.search(text)
    before = text[:split_m.start()].strip() if split_m else text

    search_text = before if "," in before else text
    if "," in search_text:
        comma_pos = search_text.index(",")
        raw_last = search_text[:comma_pos].strip()
        raw_rest = search_text[comma_pos + 1:].strip()

        last_tokens = raw_last.split()
        while last_tokens and re.match(r"^[A-Z0-9]{1,8}$", last_tokens[0]):
            last_tokens.pop(0)

        if last_tokens:
            rest_tokens = raw_rest.split()
            if rest_tokens and rest_tokens[-1].strip(".").lower() in B11_NAME_SUFFIXES:
                rest_tokens.pop()
            result["Last Name"] = " ".join(last_tokens)
            result["First Name"] = rest_tokens[0] if rest_tokens else ""
            result["Middle Name"] = " ".join(rest_tokens[1:])

    return result


def b11_split_three_cols(line, col2_x: float, col3_x: float):
    col1, col2, col3 = [], [], []
    for w in sorted(line, key=lambda w: w["x0"]):
        if w["x0"] < col2_x:
            col1.append(w["text"])
        elif w["x0"] < col3_x:
            col2.append(w["text"])
        else:
            col3.append(w["text"])
    return " ".join(col1).strip(), " ".join(col2).strip(), " ".join(col3).strip()


def b11_extract_changes(pdf_path: Path, debug: bool = False):
    change_rows = []
    page_count = 0
    debug_count = 0

    col2_x = None
    col3_x = None
    cur_emp = None
    acc_emp_lines = []

    def _emit(field, frm, to):
        if cur_emp is None or not field:
            return
        row = {
            "Last Name": cur_emp.get("Last Name", ""), "First Name": cur_emp.get("First Name", ""),
            "Middle Name": cur_emp.get("Middle Name", ""), "Associate ID": cur_emp.get("Associate ID", ""),
            "Position ID": cur_emp.get("Position ID", ""), "Home Department": cur_emp.get("Home Department", ""),
            "Home Cost Number": cur_emp.get("Home Cost Number", ""),
            "Changed Field": field, "Changed From": frm, "Changed To": to,
        }
        change_rows.append(row)
        nonlocal debug_count
        if debug and debug_count < 6:
            print(f"  [{cur_emp.get('Last Name')} {cur_emp.get('First Name')}] {field!r} | {frm!r} -> {to!r}")
            debug_count += 1

    def _flush_emp_acc():
        nonlocal cur_emp, debug_count
        if acc_emp_lines:
            cur_emp = b11_parse_employee_header("  ".join(acc_emp_lines))
            acc_emp_lines.clear()
            if debug and debug_count < 20:
                print(f"  EMP: {cur_emp.get('Last Name')}, {cur_emp.get('First Name')}  assocID={cur_emp.get('Associate ID')!r}")
                debug_count += 1

    with pdfplumber.open(str(pdf_path)) as pdf:
        page_count = len(pdf.pages)

        for page_no, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            pw = float(page.width)
            lines = b11_group_lines(words)

            if col2_x is None:
                col2_x, col3_x = b11_detect_boundaries(lines, pw)

            for line in lines:
                if not line:
                    continue
                if b11_is_page_chrome(line):
                    continue
                if b11_is_table_header(line):
                    continue

                if b11_is_employee_header(line):
                    acc_emp_lines.append(b11_line_text(line))
                    continue

                if b11_is_name_line(line, col2_x):
                    if acc_emp_lines:
                        _flush_emp_acc()
                    acc_emp_lines.append(b11_line_text(line))
                    continue

                _flush_emp_acc()
                field, frm, to_ = b11_split_three_cols(line, col2_x, col3_x)
                _emit(field, frm, to_)

            _flush_emp_acc()

    warning = "" if change_rows else B11_NO_RECORDS
    return change_rows, warning, page_count


def b11_diagnose(pdf_path: Path, max_pages: int = 3) -> None:
    print(f"\n=== DIAGNOSE: {pdf_path.name} ===\n")
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            if page_no > max_pages:
                print(f"  (capped at {max_pages} pages -- use --diagnose-pages N)")
                break
            words = page.extract_words()
            pw = float(page.width)
            lines = b11_group_lines(words)
            c2, c3 = b11_detect_boundaries(lines, pw)
            print(f"--- PAGE {page_no}  col2_x={c2:.1f}  col3_x={c3:.1f} (page_width={pw:.1f}) ---")
            for li, line in enumerate(lines, start=1):
                raw = b11_line_text(line)
                tag = ""
                if b11_is_page_chrome(line):
                    tag = "  [CHROME]"
                elif b11_is_table_header(line):
                    tag = "  [TABLE HEADER]"
                elif b11_is_employee_header(line):
                    emp = b11_parse_employee_header(raw)
                    tag = (f"  [EMP-ID: {emp['Last Name']}, {emp['First Name']}"
                           f"  assocID={emp['Associate ID']!r}  posID={emp['Position ID']!r}]")
                elif b11_is_name_line(line, c2):
                    tag = "  [EMP-NAME]"
                else:
                    field, frm, to_ = b11_split_three_cols(line, c2, c3)
                    if field:
                        tag = f"  [CHG: {field!r} | {frm!r} -> {to_!r}]"
                print(f"  L{li:03d}: {raw}{tag}")
            print()


def b11_format_name(record: dict) -> str:
    last = record.get("Last Name", "").strip()
    first = record.get("First Name", "").strip()
    middle = record.get("Middle Name", "").strip()
    parts = " ".join(filter(None, [first, middle]))
    return f"{last}, {parts}" if last and parts else last or parts


def b11_build_row(record: dict, docid: str) -> list:
    row = []
    for col in B11_OUTPUT_HEADERS:
        if col == "DOCID":
            row.append(docid)
        elif col == "Name":
            row.append(b11_format_name(record))
        else:
            row.append(record.get(col, ""))
    return [_csv_safe(v) for v in row]


def b11_write_per_pdf_output(out_path: Path, change_rows, warning, template_headers=None, template_tags=None, bucket11_headers=None):
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    stats = {"template_employees": 0, "matched_fields": 0, "unmatched_fields": 0}
    try:
        ws1 = wb.active
        ws1.title = "Native extracted"
        ws1.append(B11_OUTPUT_HEADERS)
        _style_header_row(ws1, 1, len(B11_OUTPUT_HEADERS))
        name_idx = B11_OUTPUT_HEADERS.index("Name")
        last_name = ""
        for rec in change_rows:
            row = b11_build_row(rec, rec.get("DOCID", ""))
            if not row[name_idx]:
                row[name_idx] = last_name
            else:
                last_name = row[name_idx]
            ws1.append(row)
        _autofit(ws1)

        if template_headers and bucket11_headers:
            tmpl_rows, matched, unmatched = b11_pivot_to_template_rows(change_rows, "", template_headers, bucket11_headers)
            stats["template_employees"] = len(tmpl_rows)
            stats["matched_fields"] = matched
            stats["unmatched_fields"] = unmatched
            ws2 = wb.create_sheet("Import Template")
            ws2.append(template_headers)
            if template_tags:
                ws2.append(template_tags)
            for row_dict in tmpl_rows:
                ws2.append([_csv_safe(row_dict.get(h, "")) for h in template_headers])

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


def b11_write_processing_report(report_path: Path, rows):
    tmp_path = report_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()
    headers = ["File Name", "Change Rows", "Pages", "Records Found", "Warnings"]
    try:
        ws = wb.active
        ws.title = "Processing_Report"
        ws.append(headers)
        _style_header_row(ws, 1, len(headers))
        for r in rows:
            ws.append([_csv_safe(r["file"]), r["records"], r["pages"], r["records_found"], r["warning"]])
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


def b11_load_bucket11_headers(path: Path):
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active
    headers = [str(row[0]).strip() for i, row in enumerate(ws.iter_rows(values_only=True), 1) if i > 1 and row[0]]
    wb.close()
    return headers


def b11_load_template(path: Path):
    wb = openpyxl.load_workbook(str(path))
    ws = wb.active
    rows = list(ws.iter_rows(min_row=1, max_row=2, values_only=True))
    headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    tags = list(rows[1]) if len(rows) > 1 else []
    while len(tags) < len(headers):
        tags.append(None)
    wb.close()
    return headers, tags


def b11_norm_col(name: str) -> str:
    return re.sub(r"[\s–—\-]+", " ", name or "").lower().strip()


def b11_build_col_lookup(headers):
    return {b11_norm_col(h): h for h in headers if h}


def b11_fuzzy_match_header(field: str, candidates, cutoff: float = 0.65):
    field_norm = b11_norm_col(field)
    alias_hit = B11_ALIASES.get(field_norm)
    if alias_hit:
        return alias_hit
    cand_norms = [b11_norm_col(c) for c in candidates]
    matches = _difflib.get_close_matches(field_norm, cand_norms, n=1, cutoff=cutoff)
    if not matches:
        return None
    return candidates[cand_norms.index(matches[0])]


B11_BLANK_VALUES = {"", "no value", "n/a", "none", "null", "-", "–", "—"}


def b11_select_value(changed_from: str, changed_to: str) -> str:
    to_clean = changed_to.strip()
    return changed_from.strip() if to_clean.lower() in B11_BLANK_VALUES else to_clean


def b11_pivot_to_template_rows(change_rows, docid, template_headers, bucket11_headers):
    col_lookup = b11_build_col_lookup(template_headers)

    def _find_col(name):
        if not name:
            return None
        return col_lookup.get(b11_norm_col(name))

    emp_rows = {}
    cat_acc = {}
    matched = 0
    unmatched = 0

    for rec in change_rows:
        emp_key = (rec.get("Associate ID", ""), rec.get("Last Name", ""), rec.get("First Name", ""))

        if emp_key not in emp_rows:
            row = {h: "" for h in template_headers}
            row[_find_col("DOCID") or "DOCID"] = rec.get("DOCID", docid)
            row[_find_col("Last Name") or "Last Name"] = rec.get("Last Name", "")
            row[_find_col("First Name") or "First Name"] = rec.get("First Name", "")
            row[_find_col("Middle Name") or "Middle Name"] = rec.get("Middle Name", "")
            row[_find_col("Data Subject Type") or "Data Subject Type"] = "Employee"
            eic = _find_col("Employee Identification Number")
            if eic:
                row[eic] = rec.get("Associate ID", "")
            emp_rows[emp_key] = row
            cat_acc[emp_key] = {}

        row = emp_rows[emp_key]
        cats = cat_acc[emp_key]

        field = rec.get("Changed Field", "")
        value = b11_select_value(rec.get("Changed From", ""), rec.get("Changed To", ""))

        b11_hdr = b11_fuzzy_match_header(field, bucket11_headers)
        if b11_hdr is None:
            unmatched += 1
            continue
        mapping = B11_BUCKET11_MAP.get(b11_hdr)
        if mapping is None:
            unmatched += 1
            continue

        matched += 1
        value_col_name, cat_col_name, cat_subitem = mapping

        vcol = _find_col(value_col_name)
        if vcol and value:
            existing = row.get(vcol, "")
            if not existing:
                row[vcol] = value
            elif value not in existing:
                row[vcol] = existing + "; " + value

        ccol = _find_col(cat_col_name)
        if ccol and cat_subitem:
            items = cats.setdefault(ccol, [])
            if cat_subitem not in items:
                items.append(cat_subitem)

    for emp_key, row in emp_rows.items():
        for ccol, items in cat_acc[emp_key].items():
            if items:
                existing = row.get(ccol, "")
                for item in items:
                    if item not in existing:
                        existing = (existing + ";" + item) if existing else item
                row[ccol] = existing

    return list(emp_rows.values()), matched, unmatched


def b11_selftest() -> int:
    ok = True
    name_cases = [
        ("Smith, Jonathan", ("Smith", "Jonathan", "", "")),
        ("Torres, Efrain", ("Torres", "Efrain", "", "")),
        ("Vaughan, Antione D.", ("Vaughan", "Antione", "D.", "")),
        ("Strunk, William", ("Strunk", "William", "", "")),
        ("", ("", "", "", "")),
    ]
    for raw, exp in name_cases:
        got = b11_parse_name(raw)
        tup = (got["Last Name"], got["First Name"], got["Middle Name"], got["Suffix"])
        status = "OK " if tup == exp else "FAIL"
        if tup != exp:
            ok = False
        print(f"  [{status}] {raw!r:30} -> {tup}")

    hdr = ("7TG  Smith, Jonathan  Associate ID: A12345  Position ID: P67890  "
           "Home Department: B04103-LM Bellingham Warehouse  Home Cost Number: 99")
    emp = b11_parse_employee_header(hdr)
    emp_ok = (emp["Last Name"] == "Smith" and emp["First Name"] == "Jonathan"
              and emp["Associate ID"] == "A12345" and emp["Position ID"] == "P67890"
              and "Bellingham" in emp["Home Department"])
    print(f"  [{'OK ' if emp_ok else 'FAIL'}] b11_parse_employee_header -> "
          f"{emp['Last Name']}, {emp['First Name']}  assoc={emp['Associate ID']}  pos={emp['Position ID']}")
    if not emp_ok:
        ok = False

    def _w(text, x0, x1):
        return {"text": text, "x0": x0, "x1": x1, "top": 0.0, "bottom": 10.0}

    col_line = [
        _w("Deduction", 10, 80), _w("Code", 83, 120), _w("(W)", 123, 145),
        _w("Active", 240, 290), _w("Inactive", 450, 510),
    ]
    f, frm, to_ = b11_split_three_cols(col_line, col2_x=230.0, col3_x=440.0)
    split_ok = (f == "Deduction Code (W)" and frm == "Active" and to_ == "Inactive")
    print(f"  [{'OK ' if split_ok else 'FAIL'}] b11_split_three_cols -> field={f!r}  from={frm!r}  to={to_!r}")
    if not split_ok:
        ok = False

    hdr_line = [_w("7TG", 10, 30), _w("Smith,", 33, 70), _w("Jonathan", 73, 120),
                _w("Associate", 130, 185), _w("ID:", 188, 205), _w("X999", 208, 240)]
    non_hdr = [_w("Deduction", 10, 80), _w("Code", 83, 120)]
    hdr_ok = b11_is_employee_header(hdr_line) and not b11_is_employee_header(non_hdr)
    print(f"  [{'OK ' if hdr_ok else 'FAIL'}] b11_is_employee_header")
    if not hdr_ok:
        ok = False

    tbl_header_line = [
        _w("CHANGED", 10, 80), _w("FIELD", 83, 120),
        _w("CHANGED", 240, 300), _w("FROM", 303, 340),
        _w("CHANGED", 450, 510), _w("TO", 513, 530),
    ]
    c2, c3 = b11_detect_boundaries([tbl_header_line], page_width=612.0)
    bounds_ok = (230 < c2 < 260 and 440 < c3 < 460)
    print(f"  [{'OK ' if bounds_ok else 'FAIL'}] b11_detect_boundaries -> col2_x={c2:.1f}  col3_x={c3:.1f}")
    if not bounds_ok:
        ok = False

    print("Self-test:", "PASSED" if ok else "FAILED")
    return 0 if ok else 1


def b11_process_file(pdf_path: Path, debug: bool, log=print):
    log(f"Processing: {pdf_path.name}")
    change_rows, warning, pages = b11_extract_changes(pdf_path, debug=debug)
    if warning:
        log(f"  WARNING: {warning}")
    log(f"  Extracted {len(change_rows)} change row(s).")
    for rec in change_rows:
        rec["DOCID"] = pdf_path.name
    return (
        {"file": pdf_path.name, "records": len(change_rows), "pages": pages,
         "records_found": "Yes" if change_rows else "No", "warning": warning},
        change_rows,
    )


def cmd_bucket11(args) -> int:
    if args.selftest:
        return b11_selftest()

    if not args.target:
        print("ERROR: a PDF file or folder is required (or use --selftest)")
        return 1
    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required. pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required. pip install openpyxl")

    target = Path(args.target)
    if not target.exists():
        sys.exit(f"ERROR: path not found -- {target}")

    pdfs = list(_iter_pdfs(target, args.recursive))
    if not pdfs:
        sys.exit(f"ERROR: no PDF files found at {target}")

    if args.diagnose:
        for pdf_path in pdfs[:5]:
            b11_diagnose(pdf_path, max_pages=args.diagnose_pages)
        return 0

    script_dir = Path(__file__).parent
    tmpl_path = Path(args.template_file) if args.template_file else script_dir / "Latest Template.xlsx"
    bl_path = Path(args.bucket_list) if args.bucket_list else script_dir / "Bucket_11_List.xlsx"

    template_headers = template_tags = bucket11_headers = None
    if tmpl_path.exists() and bl_path.exists():
        template_headers, template_tags = b11_load_template(tmpl_path)
        bucket11_headers = b11_load_bucket11_headers(bl_path)
        print(f"Template mode: {len(template_headers)} template cols, {len(bucket11_headers)} Bucket-11 headers loaded.")
    else:
        if not tmpl_path.exists():
            print(f"WARNING: template file not found -- {tmpl_path} (Import Template sheet skipped)")
        if not bl_path.exists():
            print(f"WARNING: bucket-list not found -- {bl_path} (Import Template sheet skipped)")

    report_rows = []
    all_warnings = []
    output_folder = target if target.is_dir() else target.parent

    for pdf_path in tqdm(pdfs, desc="Processing", unit="file"):
        try:
            rpt, rows = b11_process_file(pdf_path, args.debug, print)
            report_rows.append(rpt)

            if not args.debug:
                out_path = _extracted_path(pdf_path, output_folder)
                stats = b11_write_per_pdf_output(
                    out_path, rows, rpt["warning"],
                    template_headers=template_headers, template_tags=template_tags,
                    bucket11_headers=bucket11_headers,
                )
                sheets = "Native extracted"
                if template_headers and bucket11_headers:
                    sheets += (f" + Import Template ({stats['template_employees']} employee(s), "
                               f"{stats['matched_fields']} field(s) mapped")
                    if stats["unmatched_fields"]:
                        sheets += f", {stats['unmatched_fields']} unmatched"
                    sheets += ")"
                elif not template_headers:
                    sheets += "  [Import Template skipped -- template file not loaded]"
                print(f"  Wrote: {out_path.name}  [{sheets}]  ({len(rows)} change row(s))")

            if rpt["warning"]:
                all_warnings.append(f"{pdf_path.name}: {rpt['warning']}")
        except Exception as exc:
            print(f"  ERROR processing {pdf_path.name}: {exc}")
            report_rows.append({
                "file": pdf_path.name, "records": 0, "pages": 0,
                "records_found": "No", "warning": f"Processing failed: {exc}",
            })
            all_warnings.append(f"{pdf_path.name}: Processing failed: {exc}")

    if not args.debug and report_rows:
        report_path = output_folder / "Processing_Report_Changes.xlsx"
        b11_write_processing_report(report_path, report_rows)
        print(f"Wrote: {report_path.name} ({len(report_rows)} file(s))")

    return 0


# ===========================================================================
# CLI dispatch
# ===========================================================================

def _add_common_args(p, recursive=True, output=True):
    if recursive:
        p.add_argument("--recursive", action="store_true",
                        help="Recurse into subfolders when target is a folder")
    if output:
        p.add_argument("-o", "--output-dir", default=None,
                        help="Directory to write outputs into (default: alongside each source PDF)")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="universal_pdf_extractor",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p = sub.add_parser("w2", help="Extract Employee Name/Address/SSN (+ optional wage boxes) from W-2 PDFs")
    p.add_argument("input", help="PDF file or folder of PDFs")
    _add_common_args(p)
    p.add_argument("--wages", action="store_true", help="Also extract Box 1-6 wage/tax amounts")
    p.add_argument("--split-fraction", type=float, default=0.5,
                   help="Fraction of page width where a two-up page is cut into left/right halves (default 0.5)")
    p.add_argument("--combined", action="store_true", help="Also write a combined_extracted.xlsx across all inputs")
    p.add_argument("--combined-name", default="combined_extracted.xlsx", metavar="FILENAME")
    p.add_argument("--debug", action="store_true", help="Print a masked line-by-line layout of one page and exit")
    p.add_argument("--debug-page", type=int, default=1, metavar="N")
    p.set_defaults(func=cmd_w2)

    p = sub.add_parser("1095c", help="Extract employee identity/address fields from IRS Form 1095-C PDFs")
    p.add_argument("input", help="PDF file or folder of PDFs")
    p.add_argument("template", help="Path to BDE Import Template.xlsx (read only, never modified)")
    _add_common_args(p, recursive=False)
    p.add_argument("--debug", action="store_true",
                   help="Also write <pdf>_debug.txt (line-numbered page text) to the output dir")
    p.set_defaults(func=cmd_1095c)

    p = sub.add_parser("claims", help="Extract patient/claim identity fields from remittance PDFs")
    p.add_argument("input", help="PDF file or folder of PDFs")
    _add_common_args(p, recursive=False)
    p.add_argument("--debug", action="store_true", help="Print every line of text as this script reads it")
    p.add_argument("--page", type=int, default=None, help="Only read/extract this page number (1-indexed)")
    p.set_defaults(func=cmd_claims)

    p = sub.add_parser("patient-info", help="Extract patient info from scanned/digital remittance PDFs (with OCR)")
    p.add_argument("input", help="PDF file or folder of PDFs")
    _add_common_args(p, recursive=False)
    p.set_defaults(func=cmd_patient_info)

    p = sub.add_parser("gross-pay", help="Extract Employee Gross-To-Net report data from PDFs")
    p.add_argument("input", help="PDF file or folder of PDFs")
    _add_common_args(p, recursive=False)
    p.add_argument("--template", default=None,
                   help="Path to BDE Import Template.xlsx (auto-detected in the input folder if omitted)")
    p.set_defaults(func=cmd_gross_pay)

    p = sub.add_parser("401k", help="Extract employee 401k data into the standard PII template format")
    p.add_argument("target", nargs="?", help="PDF file or folder (prompted if omitted)")
    _add_common_args(p)
    p.add_argument("--template", default=None, help="Path to Latest Template.xlsx (auto-detected if omitted)")
    p.set_defaults(func=cmd_401k)

    p = sub.add_parser("creditor-list", help="Extract creditor names/addresses from bankruptcy mailing-matrix PDFs")
    p.add_argument("target", nargs="?", help="PDF file or folder (prompted if omitted)")
    _add_common_args(p)
    p.set_defaults(func=cmd_creditor_list)

    p = sub.add_parser("generic", help="Dynamically extract headers/tables/form-fields from any PDF")
    p.add_argument("target", nargs="?", help="PDF file or folder (prompted if omitted)")
    _add_common_args(p)
    p.set_defaults(func=cmd_generic)

    p = sub.add_parser("bucket11", help="Extract Employee Payroll Changes from ADP 'Employee Payroll Changes' PDFs")
    p.add_argument("target", nargs="?", help="PDF file or folder of PDFs")
    _add_common_args(p, output=False)
    p.add_argument("--debug", action="store_true", help="Print sample detections; write no output files")
    p.add_argument("--diagnose", action="store_true", help="Dump raw PDF text and column detection to stdout")
    p.add_argument("--diagnose-pages", type=int, default=3, metavar="N")
    p.add_argument("--selftest", action="store_true", help="Run built-in parser tests and exit")
    p.add_argument("--template-file", metavar="XLSX",
                   help="Latest Template xlsx; when provided, also write the Import Template sheet")
    p.add_argument("--bucket-list", metavar="XLSX",
                   help="Bucket-11 header list xlsx (default: Bucket_11_List.xlsx beside this script)")
    p.set_defaults(func=cmd_bucket11)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
