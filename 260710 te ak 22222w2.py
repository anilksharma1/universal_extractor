"""
Extract Employee Name/Address/SSN from Form W-2 PDFs, with name and address
each split into separate CSV columns, and write:
  - ONE CSV PER INPUT FILE: "<stem>_w2_extracted.csv" (Page, First Name,
    Last Name, Suffix, Street Address, City, State, Zip Code, SSN), next to
    the input (or under -o/--output-dir if given).
  - ONE COMBINED CSV across every input file: "combined_w2_extracted.csv"
    (Document ID, then the same columns) -- pass --no-combined to skip it,
    or --combined-name to change its filename. "Document ID" is each PDF's
    filename stem; duplicate-copy removal (Copy A/B/C of the same employee)
    is always scoped to within that Document ID, never across documents,
    even though they land in the same combined file.

Name/address column splitting:
  - Standard IRS layout (separate "e"/"f" captions): First Name, Last Name,
    and Suffix are sliced by x-position using box e's own header row (the
    "Last name" and "Suff" sub-caption positions), not by guessing from
    whitespace -- so a multi-word first *or* last name still lands in the
    right column, and a blank Suffix column is never mistaken for part of
    the last name.
  - ADP-style combined "e/f" box (single name line, no sub-captions to
    anchor on): split by whitespace instead -- last token is Last Name
    (or a recognized suffix, e.g. "Jr"/"III", is peeled off first), the
    rest is First Name.
  - City/State/Zip Code always come from the city/state/zip-shaped line
    (regex-matched); everything above it in the address block is Street
    Address.

This is a per-file-output variant of w2_ssn_name_address_csv.py -- same
block-detection engine.

Layout handling:
  - SSN is read from the "a Employee's SSN" / "a Employee's social security
    number" box and validated/normalized to XXX-XX-XXXX.
  - Name is read from the "e Employee's first name and initial / Last name /
    Suff" box.
  - Address is read from the "f Employee's address and ZIP code" box.
  - Some layouts combine e/f into one "Employee's name, address, and ZIP
    code" box -- both caption forms are recognized.
  - A page can hold 1, 2, or 4 employees (side by side, stacked, or a 2x2
    grid) -- however many "SSN" captions land on each distinct row band
    decides the split, so no layout has to be assumed up front.

These PDFs are OCR'd (made searchable) via ABBYY rather than carrying native
text, so two extra tolerances are built in:
  - normalize_text() folds curly quotes/dashes and stray whitespace ABBYY
    tends to introduce, and the caption regexes allow an optional (rather
    than required) apostrophe character so a dropped apostrophe ("Employees
    SSN") still matches.
  - find_ssn() falls back to correcting single-character digit/letter OCR
    confusions (O<->0, I/l<->1, S<->5, etc.) within the SSN value window if
    the strict digit-only pattern finds nothing there.

Usage:
    python w2_ssn_name_address_per_file.py <pdf_file_or_folder> [-o output_dir] [--split-fraction 0.5]
    python w2_ssn_name_address_per_file.py <pdf_file_or_folder> --debug [--debug-page 1]

Requires:
    pip install pymupdf tqdm
"""

import re
import csv
import argparse
import unicodedata
from pathlib import Path

import pymupdf as fitz
from tqdm import tqdm

MIN_TEXT_CHARS_PER_PAGE = 20
COLUMN_SPLIT_FRACTION = 0.5

CSV_COLUMNS = ["Page", "First Name", "Last Name", "Suffix",
               "Street Address", "City", "State", "Zip Code", "SSN"]
COMBINED_CSV_COLUMNS = ["Document ID"] + CSV_COLUMNS

# Recognized name suffixes, checked against the last whitespace-split token
# so it isn't folded into Last Name (used only for the ADP-style combined
# box, which has no per-column header to slice by position instead).
NAME_SUFFIXES = {"JR", "SR", "II", "III", "IV", "V"}

# ".?" (not ".") for the apostrophe so a dropped/misread apostrophe -- common
# when ABBYY OCRs a scanned page -- still matches ("Employees SSN").
SSN_CAPTION_RE = re.compile(r"Employee.?s\s+(?:social security number|SSN)\b", re.IGNORECASE)

# Two distinct box layouts show up in practice:
#   - Combined ("ADP-style"): one "Employee's name, address, and ZIP code" box
#     holds name + street + city/state/zip together, stacked in that order.
#   - Separate (standard IRS paper form, e.g. the screenshot this script was
#     built against): box "e" holds only the name ("Employee's first name and
#     initial / Last name / Suff") and a later, separate box "f" ("Employee's
#     address and ZIP code") holds the address. Box e/f is one tall box that
#     runs alongside boxes 11-14 on the right, so box e's data row lands on
#     the same OCR text line as box 11/12a's printed headers/values, and the
#     real address doesn't start until *after* the "f" caption, well past
#     those unrelated box 11-14 rows.
NAME_ONLY_CAPTION_RE = re.compile(r"Employee.?s\s+first name and initial\b", re.IGNORECASE)
COMBINED_NAME_ADDR_CAPTION_RE = re.compile(r"Employee.?s\s+name,\s*address", re.IGNORECASE)
ADDRESS_CAPTION_RE = re.compile(r"\bf\s+Employee.?s address\b", re.IGNORECASE)

# Union of all three, used only for the debug-mode "is this a name/address
# caption line" check and the caption-less shape fallback -- the real
# extraction logic below always distinguishes which specific one matched.
NAME_CAPTION_RE = re.compile(
    r"Employee.?s\s+(?:first name and initial|name,\s*address)|\bf\s+Employee.?s address\b",
    re.IGNORECASE,
)

STOP_LABEL_RE = re.compile(
    r"^(?:\d{1,2}\s+State\b|Employer.s state ID|Local income tax|Locality name|"
    r"Form\s*W-?2\b|Wage\s*(?:&|and)\s*Tax Statement|Copy\s+[A-Z0-9]|"
    r"For (?:Official|Privacy)|Department of the Treasury|20\d{2}$)",
    re.IGNORECASE,
)

# A box-11/12a-d/13/14 numeric label ("11", "12a", "13", ...) sharing an OCR
# text row with box e's or f's data -- its x-position marks where that
# unrelated right-hand column starts, so it can be excluded instead of
# getting appended onto the name/address text.
BOX_NUMERIC_LABEL_RE = re.compile(r"^1[1-4][a-d]?$")

# Accepts dash, space, or no separator (9 digits run together) so a valid
# SSN isn't missed just because the PDF's text layer dropped the dashes;
# normalize_ssn() below reformats whatever is found to XXX-XX-XXXX.
SSN_VALUE_RE = re.compile(r"\b(\d{3})[-\s]?(\d{2})[-\s]?(\d{4})\b")

CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b")

# "b Employer ID no. (EIN)" value, e.g. "12-3456789" -- 2-7 grouping, distinct
# from the SSN's 3-2-4 grouping. Used only as an anchor in the shape-based
# fallback below, to tell the employer's address block (box c, which always
# prints directly after the EIN) apart from the employee's (box e, which
# always prints later, after box d).
EIN_VALUE_RE = re.compile(r"\b\d{2}-\d{7}\b")

# A name/address block starting within this many lines of the EIN is almost
# certainly box c (Employer's name/address) riding right below box b, not
# box e (Employee's) -- box e always comes later, after box c and box d.
EMPLOYER_BLOCK_MAX_GAP_LINES = 5

# Letters ABBYY commonly confuses with digits when the source image is a bit
# noisy (thin strokes, low DPI, etc.). Only applied as a fallback within the
# few lines around an already-found SSN caption -- not to the whole page --
# so it can't turn unrelated text elsewhere into a false SSN match.
OCR_DIGIT_FIX = str.maketrans({
    "O": "0", "o": "0", "Q": "0", "D": "0",
    "I": "1", "l": "1", "i": "1", "|": "1",
    "Z": "2", "z": "2",
    "S": "5", "s": "5",
    "G": "6", "b": "6",
    "T": "7",
    "B": "8",
    "g": "9", "q": "9",
})


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def normalize_text(s):
    """Fold ABBYY's curly quotes/dashes and odd whitespace to plain ASCII
    equivalents before any regex runs against OCR'd text."""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate({0x2018: "'", 0x2019: "'", 0x00B4: "'", 0x0060: "'",
                      0x201C: '"', 0x201D: '"', 0x00A0: " ",
                      0x2013: "-", 0x2014: "-"})
    return re.sub(r"[ \t]+", " ", s).strip()


def normalize_ssn(match):
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"


def group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def find_ssn(plain_lines, page_num, column_label):
    for i, line in enumerate(plain_lines):
        if SSN_CAPTION_RE.search(line):
            lo, hi = max(0, i - 2), min(len(plain_lines), i + 3)
            window_lines = plain_lines[lo:hi]
            window_text = " ".join(window_lines)
            m = SSN_VALUE_RE.search(window_text)
            if m:
                return normalize_ssn(m)
            # Strict digit match failed -- retry after mapping OCR digit-lookalike
            # letters (O/I/S/etc.) back to digits, in case ABBYY misread one.
            m = SSN_VALUE_RE.search(window_text.translate(OCR_DIGIT_FIX))
            if m:
                print(f"  page {page_num} ({column_label}): SSN recovered via OCR digit-lookalike "
                      f"correction (a letter ABBYY likely misread as a digit was substituted back)")
                return normalize_ssn(m)
            shape = " | ".join(l if SSN_CAPTION_RE.search(l) else mask_shape(l) for l in window_lines)
            print(f"  page {page_num} ({column_label}): SSN caption found but no SSN-shaped value nearby "
                  f"-- nearby line shapes (digits/letters masked, safe to share): {shape}")
            return ""
    print(f"  page {page_num} ({column_label}): SSN caption not found")
    return ""


def find_right_column_boundary(word_line):
    """Return the x0 of a box-11/12a-d/13/14 numeric label sharing this OCR
    text row (e.g. the row is really "e Employee's first name...  11
    Nonqualified plans  12a ..." all run together) -- marks where the
    unrelated right-hand column starts. Returns None if no such label is on
    this row, meaning there's nothing to exclude."""
    for x0, x1, text in word_line:
        if BOX_NUMERIC_LABEL_RE.match(text.strip(".:,")):
            return x0
    return None


def row_text_left_of(word_line, max_x):
    words = word_line if max_x is None else [w for w in word_line if w[0] < max_x]
    return normalize_text(" ".join(t for _, _, t in words))


def find_e_box_dimensions(caption_word_line):
    """Locate box e's own sub-column x-positions from its header row --
    "Employee's first name and initial | Last name | Suff" -- so the name
    DATA row below it can be sliced to fit those same dimensions instead of
    guessing from whitespace. Also returns a fallback right-edge boundary
    (past "Suff") for rows where find_right_column_boundary() finds no
    box-11-14 numeric label of its own (e.g. ABBYY split "12b" into a bare
    "12" + "b" token, so the per-row numeric-label match misses it)."""
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


def split_name_row(word_line, last_x0, suff_x0, max_x):
    """Slice one name-row's words into (first, last, suffix) by x-position
    using the box-e sub-column boundaries from find_e_box_dimensions()."""
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


def split_name_by_whitespace(name):
    """Fallback name splitter with no positional header to anchor on (the
    ADP-style combined name/address box): peel off a trailing suffix token
    if present, treat the last remaining token as Last Name, everything
    before it as First Name."""
    tokens = name.split()
    suffix = ""
    if tokens and tokens[-1].strip(".").upper() in NAME_SUFFIXES:
        suffix = tokens[-1]
        tokens = tokens[:-1]
    if not tokens:
        return "", "", suffix
    if len(tokens) == 1:
        return tokens[0], "", suffix
    return " ".join(tokens[:-1]), tokens[-1], suffix


def empty_name_address_fields():
    return {"First Name": "", "Last Name": "", "Suffix": "",
            "Street Address": "", "City": "", "State": "", "Zip Code": ""}


def split_address_block(block, start_idx, page_num, column_label):
    """Given a list of plain-text lines (block[start_idx:] is the address
    portion, with block[:start_idx] being the name row(s) already sliced
    out by the caller), split it into Street Address / City / State / Zip
    Code using the city/state/zip-shaped line as the anchor."""
    fields = {"Street Address": "", "City": "", "State": "", "Zip Code": ""}
    addr_lines = block[start_idx:]
    city_idx = next((idx for idx, l in enumerate(addr_lines) if CITY_STATE_ZIP_RE.match(l)), None)
    if city_idx is not None:
        m = CITY_STATE_ZIP_RE.match(addr_lines[city_idx])
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


def find_name_address(lines, page_num, column_label):
    """lines: list of per-row word tuples (x0, x1, text), as returned by
    group_words_into_lines -- word-level x-positions are needed (not just
    joined line strings) to strip box 11-14 content that shares a row with
    box e/f's data (see BOX_NUMERIC_LABEL_RE above), and to slice box e's
    First Name / Last Name / Suffix sub-columns by their real dimensions."""
    plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]

    name_idx = next((i for i, l in enumerate(plain_lines) if NAME_ONLY_CAPTION_RE.search(l)), None)
    if name_idx is not None:
        return _find_name_address_separate_boxes(lines, plain_lines, name_idx, page_num, column_label)

    for i, line in enumerate(plain_lines):
        if COMBINED_NAME_ADDR_CAPTION_RE.search(line):
            return _find_name_address_combined_box(plain_lines, i, page_num, column_label)

    print(f"  page {page_num} ({column_label}): name/address caption not found")
    return empty_name_address_fields()


def _find_name_address_combined_box(plain_lines, caption_idx, page_num, column_label):
    """ADP-style layout: one caption, then name, then street, then
    city/state/zip, all stacked directly below it. No per-column header to
    slice First/Last/Suffix by position, so split_name_by_whitespace() is
    used instead."""
    raw_block = plain_lines[caption_idx + 1:min(caption_idx + 9, len(plain_lines))]
    block = []
    for candidate in raw_block:
        candidate = candidate.strip()
        if not candidate:
            continue
        if STOP_LABEL_RE.search(candidate) or ADDRESS_CAPTION_RE.search(candidate):
            break
        block.append(candidate)
    if not block:
        shape = " | ".join(mask_shape(l) for l in raw_block)
        print(f"  page {page_num} ({column_label}): name/address caption found but block below it is "
              f"empty (line shapes below caption, safe to share): {shape}")
        return empty_name_address_fields()

    first, last, suffix = split_name_by_whitespace(block[0])
    fields = split_address_block(block, 1, page_num, column_label)
    fields.update({"First Name": first, "Last Name": last, "Suffix": suffix})
    return fields


def _find_name_address_separate_boxes(lines, plain_lines, name_idx, page_num, column_label):
    """Standard IRS layout: box e (name) and box f (address) print as one
    tall box running alongside boxes 11-14, so their data rows land on the
    same OCR text lines as those unrelated boxes' labels/values. The "f"
    caption itself can land anywhere in that range depending on the source
    document (sometimes right before the address, sometimes after it, per
    the actual scanned layout) -- so rather than treating it as a hard
    boundary, this treats box e/f as one continuous region from the "e"
    caption down to the state-tax stop-label, strips box 11-14 content off
    each row by x-position, and just skips the "f" caption line itself as a
    label with no data of its own.

    The right-edge boundary for that stripping is found two ways per row:
    first by looking for a box-11-14 numeric label token on that exact row
    (find_right_column_boundary); if none is found there -- e.g. ABBYY split
    "12b" into a bare "12" + "b" token, so no single token matches -- it
    falls back to the static boundary derived once from box e's own header
    ("Suff" column's right edge), fitting every row to that box's actual
    printed dimensions rather than leaving it unbounded.

    Two more tolerances, needed for forms with a taller box 11-14 region:
      - The search window is generous (40 rows) rather than a tight guess,
        since how many filler rows separate the name from the real address
        varies by document (a short box 11-14 region vs. one with several
        populated 12a-d/13/14 rows) -- STOP_LABEL_RE is what actually ends
        the scan, not this cap; it's just a backstop against a missing/
        unmatched stop label running away to the end of the page.
      - A row that survives x-position stripping but is just a stray
        single-character artifact (e.g. ABBYY misreading the vertical
        divider between box e/f and boxes 11-14 as a lone "X" on an
        otherwise-blank template row) is discarded rather than mistaken for
        the name or spliced into the address -- real content is either
        city/state/zip-shaped or has at least a few letters."""
    last_x0, suff_x0, static_boundary = find_e_box_dimensions(lines[name_idx])

    first = last = suffix = ""
    plain_block = []
    name_row_used = False
    for j in range(name_idx + 1, min(name_idx + 40, len(plain_lines))):
        if STOP_LABEL_RE.search(plain_lines[j]):
            break
        if ADDRESS_CAPTION_RE.search(plain_lines[j]):
            continue
        row_boundary = find_right_column_boundary(lines[j]) or static_boundary
        text = row_text_left_of(lines[j], row_boundary)
        if not text:
            continue
        if not (is_letterish(text) or CITY_STATE_ZIP_RE.match(text)):
            continue
        if not name_row_used:
            first, last, suffix = split_name_row(lines[j], last_x0, suff_x0, row_boundary)
            name_row_used = True
        else:
            plain_block.append(text)

    if not name_row_used:
        print(f"  page {page_num} ({column_label}): 'e' name caption found but box e/f's own column is "
              f"empty all the way down to the next stop label -- name/address left blank")
        return empty_name_address_fields()

    fields = split_address_block(plain_block, 0, page_num, column_label)
    fields.update({"First Name": first, "Last Name": last, "Suffix": suffix})
    return fields


def extract_employee(lines, plain_lines, page_num, column_label):
    ssn = find_ssn(plain_lines, page_num, column_label)
    fields = find_name_address(lines, page_num, column_label)
    return {"Page": page_num, "SSN": ssn, **fields}


def is_letterish(line):
    """True for lines that look like a name/street (letters outnumber
    digits) -- used to tell a wage-amount line apart from an address line
    when there's no caption to anchor on."""
    letters = sum(c.isalpha() for c in line)
    digits = sum(c.isdigit() for c in line)
    return letters >= 3 and letters > digits


def find_ssn_by_shape(plain_lines, page_num, column_label):
    """Caption-less fallback: find an SSN-shaped value directly, without an
    'Employee's SSN' caption to anchor on (used when that caption isn't in
    the OCR text at all -- e.g. it's part of a background template image
    ABBYY didn't recognize as text, while the payroll-filled values are)."""
    for line in plain_lines:
        m = SSN_VALUE_RE.search(line)
        if m:
            return normalize_ssn(m)
        m = SSN_VALUE_RE.search(line.translate(OCR_DIGIT_FIX))
        if m:
            print(f"  page {page_num} ({column_label}): [shape-fallback] SSN recovered via OCR "
                  f"digit-lookalike correction")
            return normalize_ssn(m)
    print(f"  page {page_num} ({column_label}): [shape-fallback] no SSN-shaped value found in this cell")
    return ""


def find_name_address_by_shape(plain_lines, page_num, column_label):
    """Caption-less fallback for name/address. A W-2 box stack always prints
    box c (Employer's name/address/ZIP) directly under box b (EIN), then box
    d (Control number), then box e (Employee's name/address/ZIP) -- so when
    there's no caption to anchor on, collect every name+street+city/state/zip
    -shaped block and use the EIN's position to tell an employer-adjacent
    block apart from the employee's, rather than assuming "the only block
    found" or "the last block" is automatically the employee's (either one
    can silently return the employer's info instead if the employee block
    wasn't detected for some reason)."""
    csz_indices = [i for i, l in enumerate(plain_lines) if CITY_STATE_ZIP_RE.match(l)]
    if not csz_indices:
        print(f"  page {page_num} ({column_label}): [shape-fallback] no city/state/zip-shaped line found "
              f"-- cannot locate name/address")
        return empty_name_address_fields()

    blocks = []
    for csz_idx in csz_indices:
        block = [csz_idx]
        j = csz_idx - 1
        while j >= 0 and len(block) < 4:
            line = plain_lines[j].strip()
            if not line or CITY_STATE_ZIP_RE.match(line) or not is_letterish(line):
                break
            block.insert(0, j)
            j -= 1
        if len(block) >= 2:
            blocks.append(block)

    if not blocks:
        print(f"  page {page_num} ({column_label}): [shape-fallback] city/state/zip-shaped line(s) found "
              f"but no preceding name/street line(s)")
        return empty_name_address_fields()

    ein_idx = next((i for i, l in enumerate(plain_lines) if EIN_VALUE_RE.search(l)), None)

    if ein_idx is None:
        # No EIN anchor to disambiguate with. Only safe to act when there's
        # more than one block (last = employee, per box order); a lone block
        # is genuinely ambiguous -- guessing it's the employee's is exactly
        # the bug this replaced (it can just as easily be the employer's).
        if len(blocks) > 1:
            print(f"  page {page_num} ({column_label}): [shape-fallback] no EIN found to anchor on, but "
                  f"{len(blocks)} name/address blocks were -- using the last one (box 'e' prints after box "
                  f"'c') as the employee's; spot-check this record")
            chosen = blocks[-1]
        else:
            print(f"  page {page_num} ({column_label}): [shape-fallback] only one name/address block found "
                  f"and no EIN to confirm whether it's box 'c' (employer) or box 'e' (employee) -- skipping "
                  f"rather than risk recording the employer's info as the employee's")
            return empty_name_address_fields()
    else:
        employee_like = [b for b in blocks if b[0] - ein_idx > EMPLOYER_BLOCK_MAX_GAP_LINES]
        if not employee_like:
            print(f"  page {page_num} ({column_label}): [shape-fallback] every name/address block found sits "
                  f"within {EMPLOYER_BLOCK_MAX_GAP_LINES} lines of the EIN -- that's box 'c' (employer's), "
                  f"not box 'e'; employee name/address not found in this cell")
            return empty_name_address_fields()
        if len(employee_like) > 1:
            print(f"  page {page_num} ({column_label}): [shape-fallback] {len(employee_like)} candidate "
                  f"employee blocks found past the EIN -- using the last one; spot-check this record")
        chosen = employee_like[-1]

    name = plain_lines[chosen[0]].strip()
    csz_line = plain_lines[chosen[-1]].strip()
    m = CITY_STATE_ZIP_RE.match(csz_line)
    first, last, suffix = split_name_by_whitespace(name)
    fields = {
        "First Name": first, "Last Name": last, "Suffix": suffix,
        "Street Address": " ".join(plain_lines[i].strip() for i in chosen[1:-1]),
        "City": m.group("city").rstrip(","), "State": m.group("state"), "Zip Code": m.group("zip"),
    }
    return fields


def extract_employee_by_shape(plain_lines, page_num, column_label):
    ssn = find_ssn_by_shape(plain_lines, page_num, column_label)
    fields = find_name_address_by_shape(plain_lines, page_num, column_label)
    return {"Page": page_num, "SSN": ssn, **fields}


def caption_line_groups(words, caption_re, y_tol=3):
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


def build_grid_cells(page, words, groups, split_fraction, page_num):
    row_ys = [y for y, _ in groups]
    row_counts = [c for _, c in groups]
    num_rows = len(row_ys)

    if num_rows <= 1:
        row_bounds = [(float("-inf"), float("inf"))]
    else:
        # Equal-height bands across the FULL PAGE, not midpoints between the
        # SSN-caption y-positions. Box "a" sits near the TOP of each stacked
        # copy (it's always the first box on the form), so a midpoint
        # between two captions falls in the middle of the FIRST copy's own
        # content -- chopping its address block in half and merging the
        # bottom half into the next employee's cell instead. Multi-part W-2
        # forms print each stacked copy at a fixed, equal height, so
        # dividing the page height evenly (mirroring how the left/right
        # split below already uses a fixed fraction of page WIDTH, not
        # caption x-positions) is the layout-accurate boundary.
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


def process_page(page, page_num, split_fraction):
    text = page.get_text()
    if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
        print(f"  page {page_num}: only {len(text.strip())} chars of text -- scanned/image-only page, skipping "
              f"(this script expects a real text layer; OCR it first if needed)")
        return []

    norm_text = normalize_text(text)
    caption_count = len(SSN_CAPTION_RE.findall(norm_text))
    shape_fallback = False
    marker_re = SSN_CAPTION_RE
    marker_count = caption_count

    if caption_count == 0:
        # The caption itself isn't in the OCR text at all -- happens when the
        # box labels are part of a background template image ABBYY didn't
        # recognize as text, while the payroll-filled values on top of it
        # were. If SSN-shaped values are on the page anyway, fall back to
        # locating employees by value shape/position instead of caption text.
        value_count = len(SSN_VALUE_RE.findall(norm_text))
        if value_count == 0:
            print(f"  page {page_num}: no 'Employee's SSN' caption and no SSN-shaped value found -- not a "
                  f"W-2 employee page, skipping")
            return []
        print(f"  page {page_num}: no 'Employee's SSN' caption in the OCR text (likely baked into a "
              f"background image) -- falling back to locating employees by SSN value shape/position")
        shape_fallback = True
        marker_re = SSN_VALUE_RE
        marker_count = value_count

    words = page.get_text("words")
    records = []

    if marker_count == 1:
        cells = [("Single", words)]
    else:
        groups = caption_line_groups(words, marker_re)
        if not groups:
            print(f"  page {page_num}: {marker_count} SSN {'values' if shape_fallback else 'captions'} found "
                  f"on the page but none could be grouped into row bands -- falling back to treating the "
                  f"whole page as one employee")
            cells = [("Single", words)]
        else:
            cells = build_grid_cells(page, words, groups, split_fraction, page_num)

    for column_label, column_words in cells:
        lines = group_words_into_lines(column_words)
        plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]

        if shape_fallback:
            has_ssn_here = any(SSN_VALUE_RE.search(l) or SSN_VALUE_RE.search(l.translate(OCR_DIGIT_FIX))
                                for l in plain_lines)
            if not has_ssn_here:
                print(f"  page {page_num} ({column_label}): [shape-fallback] no SSN-shaped value in this "
                      f"cell -- split-fraction may be off for this file, try --split-fraction")
                continue
            records.append(extract_employee_by_shape(plain_lines, page_num, column_label))
        else:
            if not any(SSN_CAPTION_RE.search(l) for l in plain_lines):
                print(f"  page {page_num} ({column_label}): no SSN caption on this cell -- split-fraction "
                      f"may be off for this file, try --split-fraction")
                continue
            records.append(extract_employee(lines, plain_lines, page_num, column_label))

    return records


def process_pdf(path: Path, split_fraction):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        records.extend(process_page(page, i, split_fraction))
    doc.close()
    return records


def dedupe_records(records):
    """Each employee typically appears on more than one identical W-2 copy
    (Copy B, Copy C, Copy 2, ...) within the same PDF, producing one record
    per copy with the same name/address/SSN but a different Page --
    collapse those down to a single row per employee, keeping the first
    occurrence's Page. Dedupes on SSN when present (the reliable identity
    key); falls back to the full name + address fields when SSN extraction
    failed."""
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
            # Nothing usable was extracted -- keep every such record rather
            # than collapsing unrelated blank rows into one.
            deduped.append(rec)
            continue
        if key in seen:
            continue
        seen.add(key)
        deduped.append(rec)
    return deduped


def write_csv(records, output_path, columns=CSV_COLUMNS):
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        for rec in records:
            writer.writerow({col: rec.get(col, "") for col in columns})


def debug_page(path: Path, page_num: int):
    """Print a masked, line-by-line view of one page's OCR text layer so a
    failed extraction can be diagnosed without exposing PII: caption lines
    print in full (they're just form labels), every other line is reduced to
    its digit/letter shape via mask_shape() -- safe to paste back here."""
    doc = fitz.open(path)
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return
    page = doc[page_num - 1]

    text = normalize_text(page.get_text())
    print(f"--- {path.name} page {page_num} of {len(doc)} ---")
    print(f"text layer: {len(text.strip())} chars "
          f"({'OK' if len(text.strip()) >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN_TEXT_CHARS_PER_PAGE -- treated as image-only'})")
    ssn_caption_count = len(SSN_CAPTION_RE.findall(text))
    name_caption_count = len(NAME_CAPTION_RE.findall(text))
    ssn_value_count = len(SSN_VALUE_RE.findall(text))
    print(f"SSN captions found on page: {ssn_caption_count}")
    print(f"Name/address captions found on page: {name_caption_count}")
    print(f"SSN-shaped values found on page: {ssn_value_count}")
    if ssn_caption_count == 0 and ssn_value_count > 0:
        print("  -> no caption text in the OCR layer; extraction will use the shape-based fallback "
              "(SSN located by value shape, name/address by position -- box 'e' assumed to be the LAST "
              "name/street/city-state-zip-shaped block in each cell)")

    words = page.get_text("words")
    lines = group_words_into_lines(words)
    plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]
    print(f"lines detected: {len(plain_lines)}")
    for i, line in enumerate(plain_lines):
        if SSN_CAPTION_RE.search(line):
            tag = "  <-- SSN caption"
            shown = line
        elif NAME_ONLY_CAPTION_RE.search(line):
            tag = "  <-- name caption (box e)"
            shown = line
        elif ADDRESS_CAPTION_RE.search(line):
            tag = "  <-- address caption (box f)"
            shown = line
        elif COMBINED_NAME_ADDR_CAPTION_RE.search(line):
            tag = "  <-- combined name/address caption"
            shown = line
        elif STOP_LABEL_RE.search(line):
            tag = "  <-- stop label"
            shown = line
        else:
            tag = ""
            shown = mask_shape(line)
        print(f"  [{i:>3}] {shown}{tag}")
    doc.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("-o", "--output-dir", default=None,
                         help="Folder to write each file's '<stem>_w2_extracted.csv' into, and the combined "
                              "CSV (default: same folder as each input PDF, or the input folder itself when "
                              "a folder of PDFs is given)")
    parser.add_argument("--combined-name", default="combined_w2_extracted.csv", metavar="FILENAME",
                         help="Filename for the combined CSV across all input files, written alongside the "
                              "per-file CSVs (default: combined_w2_extracted.csv). Adds a 'Document ID' column.")
    parser.add_argument("--no-combined", action="store_true",
                         help="Skip writing the combined CSV -- write only the per-file CSVs")
    parser.add_argument("--split-fraction", type=float, default=COLUMN_SPLIT_FRACTION,
                         help="Fraction of page width where a two-up page is cut into left/right halves "
                              f"(default {COLUMN_SPLIT_FRACTION})")
    parser.add_argument("--debug", action="store_true",
                         help="Print a masked line-by-line layout of one page per input file and exit -- "
                              "caption lines print in full, everything else is masked to digit/letter shape "
                              "(safe to paste back for troubleshooting). Use --debug-page to pick the page.")
    parser.add_argument("--debug-page", type=int, default=1, metavar="N",
                         help="Page number to debug when --debug is set (default: 1)")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    if args.debug:
        for pdf in pdf_files:
            debug_page(pdf, args.debug_page)
        return

    output_dir = Path(args.output_dir) if args.output_dir else None
    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    total_written = 0
    combined_records = []
    progress = tqdm(pdf_files, desc="Extracting", unit="file")
    for pdf in progress:
        progress.set_postfix_str(pdf.name)
        records = process_pdf(pdf, args.split_fraction)
        if not records:
            print(f"  no records extracted from {pdf.name} -- no CSV written for this file")
            continue

        # Scoped to this one document -- a duplicate-copy match (Copy A/B/C
        # of the same employee) never collapses records from a different
        # Document ID, even after they're merged into the combined CSV below.
        deduped_records = dedupe_records(records)
        removed = len(records) - len(deduped_records)
        if removed:
            print(f"  removed {removed} duplicate record(s) (same employee found on multiple copies/pages)")

        dest_dir = output_dir if output_dir else pdf.parent
        out_path = dest_dir / f"{pdf.stem}_w2_extracted.csv"
        write_csv(deduped_records, out_path)
        total_written += len(deduped_records)
        print(f"  -> wrote {len(deduped_records)} record(s) to {out_path}")

        for rec in deduped_records:
            combined_records.append({"Document ID": pdf.stem, **rec})

    if not args.no_combined and combined_records:
        combined_dir = output_dir if output_dir else (input_path if input_path.is_dir() else input_path.parent)
        combined_path = combined_dir / args.combined_name
        write_csv(combined_records, combined_path, columns=COMBINED_CSV_COLUMNS)
        print(f"\nWrote combined CSV: {len(combined_records)} record(s) across {len(pdf_files)} file(s) "
              f"-> {combined_path}")

    print(f"\nDone. {total_written} record(s) written across {len(pdf_files)} file(s).")


if __name__ == "__main__":
    main()
