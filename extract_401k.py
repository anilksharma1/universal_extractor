r"""
extract_401k.py — Extract employee data from 401k contribution PDFs
and write into the standard 38-column PII format.

Column mapping:
  A  DOCID                          ← PDF filename stem
  B  Last Name                      ← parsed from name
  C  First Name                     ← parsed from name
  D  Middle Name                    ← parsed from name or separate field
  E  Suffix                         ← parsed from name or separate field
  F  Data Subject Type              ← "Employee"
  G–Q blank (address / contact)
  R  Government-Issued Identification ← "Social Security Number (SSN)" tag
  S  Social Security Number         ← SSN value
  T–AL blank (other PII)
  AK Work-Related Information       ← salary & contribution amounts
  AL Employee Identification Number ← EE ID if present

Strategy (in order):
  1. AcroForm field extraction  — fillable PDFs
  2. Table extraction           — structured tables via pdfplumber
  3. Regex text scan            — fallback for unstructured / image-heavy PDFs

Output — one {docId}_extracted.xlsx per PDF, header in row 1,
         employee data from row 2 onward.

USAGE
    python extract_401k.py "C:\path\to\folder"
    python extract_401k.py "C:\path\to\file.pdf"
    python extract_401k.py "C:\path\to\folder" --recursive
    python extract_401k.py "C:\path\to\file.pdf" --output "C:\output"

DEPENDENCIES
    pip install pdfplumber openpyxl pypdf
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
    from tqdm import tqdm as _tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False

try:
    from pypdf import PdfReader
    HAS_PYPDF = True
except ImportError:
    HAS_PYPDF = False

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


# ---------------------------------------------------------------------------
# Column positions (1-based)
# ---------------------------------------------------------------------------

COL_DOCID          = 1
COL_LAST_NAME      = 2
COL_FIRST_NAME     = 3
COL_MIDDLE_NAME    = 4
COL_SUFFIX         = 5
COL_SUBJECT_TYPE   = 6
COL_GOV_ID_TAG     = 18   # R — Government-Issued Identification (tag)
COL_SSN            = 19   # S — Social Security Number (value)
COL_WORK_INFO      = 37   # AK — Work-Related Information
COL_EE_ID          = 38   # AL — Employee Identification Number
TEMPLATE_COL_COUNT = 38


# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

_SSN_RE = re.compile(r'\b(\d{3})-(\d{2})-(\d{4})\b')
# SSN where last group is masked/partial: 123-45-67XX, 123-45-****, 123-45-67**
# Uses (?<!\d) instead of \b so it also matches when the SSN immediately follows
# a letter (e.g. pdfplumber merges "Adeyemi, Chukwuemeke 611-56-73" into one string).
_SSN_RE_XPAD = re.compile(r'(?<!\d)(\d{3})-(\d{2})-([\dxX*]{1,4})(?![\dxX*])', re.I)
_AMOUNT_RE = re.compile(r'\$?\s*([\d]{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)')
# Title-case names: "John Smith", "Mary Ann Jones"
_NAME_RE = re.compile(r'\b([A-Z][a-z]+(?:[ \t]+[A-Z][a-z]*\.?){1,4})\b')
# ALL-CAPS names: "JOHN SMITH", "SMITH MARY ANN"
_NAME_RE_UPPER = re.compile(r'\b([A-Z]{2,}(?:[ \t]+[A-Z]{2,}){1,4})\b')
# "Last, First [Middle]" with comma
_NAME_RE_LASTFIRST = re.compile(r'\b([A-Za-z]{2,}),[ \t]+([A-Za-z]{2,}(?:[ \t]+[A-Za-z]{1,}\.?)?)\b')

# Partial SSN patterns checked when no full SSN is found in a line.
# All patterns assume digits are known from the LEFT; X pads missing tail.
_SSN_PARTIAL_PATTERNS: list[tuple[re.Pattern, str]] = [
    # DDD-DD-DDD  or  DDD-DD-DD  or  DDD-DD-D  (short last group, 1-3 digits)
    (re.compile(r'\b(\d{3})-(\d{2})-(\d{1,3})\b(?!\d)'), "short_last"),
    # DDD-DD  (first 5 digits, last group entirely missing)
    (re.compile(r'\b(\d{3})-(\d{2})\b(?!\s*-?\s*\d)'), "front5"),
    # XXX-XX-DDDD  or  ***-**-DDDD  (only last 4 known — keep as-is)
    (re.compile(r'[xX*]{3}[- ][xX*]{2}[- ](\d{4})\b'), "last4"),
    # DDD-XX-DDDD  (middle group masked)
    (re.compile(r'\b(\d{3})[- ][xX*]{2}[- ](\d{4})\b'), "mid_masked"),
]

_SUFFIXES = {
    "jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v",
    "esq", "esq.", "phd", "md", "dds", "cpa", "rn",
}


# ---------------------------------------------------------------------------
# Column-name → canonical field mapping
# ---------------------------------------------------------------------------

_COL_MAP: dict[str, str] = {}

def _add(canonical: str, *aliases: str) -> None:
    for a in aliases:
        _COL_MAP[a.lower()] = canonical

_add("first_name",
     "first name", "first", "fname", "given name", "employee first name",
     "participant first name")
_add("last_name",
     "last name", "last", "lname", "surname", "family name",
     "employee last name", "participant last name")
_add("middle_name",
     "middle name", "middle", "middle initial", "mi", "mname",
     "employee middle name")
_add("suffix",
     "suffix", "name suffix", "generational suffix", "jr", "sr")
_add("name",
     "name", "employee name", "emp name", "employee", "participant name",
     "participant", "last, first", "last/first", "full name")
_add("ssn",
     "ssn", "social security", "social security number", "soc sec",
     "ss number", "ss#", "ssn#", "tax id", "tin", "employee ssn",
     "participant ssn", "employee id / ssn")
_add("gross_salary",
     "gross salary", "gross pay", "gross wages", "gross compensation",
     "annual salary", "annual compensation", "compensation", "wages",
     "salary", "ytd gross", "ytd wages", "total compensation")
_add("ee_contribution",
     "ee contribution", "employee contribution", "ee deferral",
     "employee deferral", "employee 401k", "elective deferral",
     "pre-tax contribution", "pre tax contribution", "roth contribution",
     "employee amount", "ee amount", "employee contrib")
_add("er_match",
     "er match", "employer match", "employer contribution",
     "company match", "company contribution", "er contribution",
     "matching contribution", "employer match amount")
_add("ytd_contribution",
     "ytd contribution", "ytd total", "year to date", "ytd",
     "ytd deferral", "total ytd")
_add("ee_id",
     "employee id", "ee id", "emp id", "employee number", "emp number",
     "employee identification", "employee identification number", "ee#")


def _map_column(raw_header: str) -> str | None:
    normalised = re.sub(r'\s+', ' ', (raw_header or "").strip().lower().rstrip(":"))
    if normalised in _COL_MAP:
        return _COL_MAP[normalised]
    best, best_len = None, 0
    for alias, canonical in _COL_MAP.items():
        if alias in normalised and len(alias) > best_len:
            best, best_len = canonical, len(alias)
    return best


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean(text) -> str:
    return " ".join(str(text or "").split()).strip()


def _fmt_ssn(raw: str) -> str:
    """
    Format a raw SSN value as XXX-XX-XXXX (11 chars).

    Reads digits left-to-right (beginning of SSN first).
    Any missing digits are always padded with X at the END of the
    incomplete group, so the known prefix is always preserved exactly.

    Group sizes: [3]-[2]-[4]
    """
    raw = _clean(raw)
    if not raw:
        return raw

    # Already final format (may contain X placeholders already)
    if re.match(r'^[\dX]{3}-[\dX]{2}-[\dX]{4}$', raw, re.I):
        return raw.upper()

    # ── Structured partials: dashes reveal which groups are present ──────────

    # Full: DDD-DD-DDDD
    m = re.match(r'^(\d{3})-(\d{2})-(\d{4})$', raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # Short last group: DDD-DD-DDD / DDD-DD-DD / DDD-DD-D  (1–3 digits in last group)
    m = re.match(r'^(\d{3})-(\d{2})-(\d{1,3})$', raw)
    if m:
        serial = m.group(3).ljust(4, 'X')    # pad end of last group
        return f"{m.group(1)}-{m.group(2)}-{serial}"

    # Missing last group entirely: DDD-DD
    m = re.match(r'^(\d{3})-(\d{2})$', raw)
    if m:
        return f"{m.group(1)}-{m.group(2)}-XXXX"

    # Middle group masked: DDD-XX-DDDD
    m = re.match(r'^(\d{3})-[xX*]+-(\d{4})$', raw)
    if m:
        return f"{m.group(1)}-XX-{m.group(2)}"

    # First group masked: XXX-DD-DDDD
    m = re.match(r'^[xX*]+-(\d{2})-(\d{4})$', raw)
    if m:
        return f"XXX-{m.group(1)}-{m.group(2)}"

    # Only last group: XXX-XX-DDDD
    m = re.match(r'^[xX*]+-[xX*]+-(\d{4})$', raw)
    if m:
        return f"XXX-XX-{m.group(1)}"

    # ── Plain digit string: fill groups left-to-right, pad end with X ────────
    d = re.sub(r'\D', '', raw)
    n = len(d)
    if n >= 9:
        return f"{d[:3]}-{d[3:5]}-{d[5:9]}"
    if n == 8:   # 1 missing from end of last group
        return f"{d[0:3]}-{d[3:5]}-{d[5:8]}X"
    if n == 7:   # 2 missing from end of last group
        return f"{d[0:3]}-{d[3:5]}-{d[5:7]}XX"
    if n == 6:   # 3 missing from end of last group
        return f"{d[0:3]}-{d[3:5]}-{d[5]}XXX"
    if n == 5:   # last group entirely missing
        return f"{d[0:3]}-{d[3:5]}-XXXX"
    if n == 4:   # 1 digit into second group, rest missing
        return f"{d[0:3]}-{d[3]}X-XXXX"
    if n == 3:   # only first group (area code)
        return f"{d}-XX-XXXX"

    return raw  # too short / ambiguous


def _parse_amount(text: str) -> str:
    m = _AMOUNT_RE.search(_clean(text))
    return m.group(1).replace(',', '') if m else ""


# Label words stripped from text BEFORE name matching so they don't get
# captured as part of a name (e.g. "EMPLOYEE JOHN SMITH" → "JOHN SMITH").
_LABEL_STRIP_RE = re.compile(
    r'\b(?:EMPLOYEE|EMPLOYER|PARTICIPANT|MEMBER|CLAIMANT|'
    r'SALARY|COMPENSATION|CONTRIBUTION|DEFERRAL|MATCH|'
    r'GROSS|NET|TOTAL|AMOUNT|NUMBER|ID|SSN|TIN|TAX|'
    r'PLAN|COMPANY|PAGE|DATE|ADDRESS|STATE|CITY|ZIP|'
    r'PHONE|EMAIL|FAX|FOR|AND|THE|INC|LLC|CORP|CO|LTD)\b',
    re.I,
)


def _find_name_in_text(text: str) -> str:
    """
    Search text for a person name using three patterns tried in order:
      1. "Last, First [Middle]"  — comma separator makes structure explicit
      2. ALL CAPS words          — after stripping known label tokens
      3. Title Case words        — "John Smith", "Mary Ann Jones"
    Returns the best candidate as a string, or '' if nothing found.
    """
    text = _clean(text)
    if not text or _AMOUNT_RE.fullmatch(text) or _SSN_RE.search(text):
        return ''

    def _ok(s: str) -> bool:
        """Accept a candidate: ≥2 words, no digits."""
        words = s.split()
        return len(words) >= 2 and not any(c.isdigit() for c in s)

    # Pattern 1: "Last, First [Middle]" — preserve the comma so _parse_name_parts
    # can split correctly into Last / First components.
    m = _NAME_RE_LASTFIRST.search(text)
    if m:
        candidate = f"{m.group(1)}, {m.group(2)}"
        if _ok(candidate.replace(",", "")):   # _ok check on comma-stripped form
            return candidate.title()

    # Pattern 2: ALL CAPS — strip label words first so they don't pollute the match
    stripped = _LABEL_STRIP_RE.sub('', text)
    stripped = re.sub(r'\b\d[\d,.$%-]*\b', '', stripped)   # remove numbers
    stripped = re.sub(r'\s+', ' ', stripped).strip()
    m = _NAME_RE_UPPER.search(stripped)
    if m:
        candidate = m.group(1)
        if _ok(candidate):
            return candidate.title()

    # Pattern 3: Title Case
    m = _NAME_RE.search(text)
    if m:
        candidate = m.group(1)
        if _ok(candidate):
            return candidate

    return ''


def _ssn_from_cell(cell_text: str) -> str:
    """Extract and format an SSN from a table cell (full or partial). Returns '' if none."""
    cell_text = _clean(cell_text)
    if not cell_text:
        return ''
    # Full digit SSN: DDD-DD-DDDD or DDDDDDDDD
    m = _SSN_RE.search(cell_text)
    if m:
        return _fmt_ssn(m.group(0))
    # X-padded SSN: DDD-DD-DDXX, DDD-DD-DXXX, DDD-DD-XXXX
    m = _SSN_RE_XPAD.search(cell_text)
    if m:
        return _fmt_ssn(m.group(0))
    # Partial SSN patterns
    for pat, kind in _SSN_PARTIAL_PATTERNS:
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
            return _fmt_ssn(partial)
    return ''


def _ssn_real_digits(ssn: str) -> int:
    """Count actual digits (non-X) in a formatted SSN."""
    return sum(1 for c in ssn if c.isdigit())




def _drop_redacted_duplicates(records: list) -> list:
    """
    Narrow dedup: if two records share the same first-5 SSN digits (DDD-DD)
    and one has all 9 real digits while another has X-padding in the last group,
    drop the X-padded record.  No field merging; no name-based matching.
    """
    if len(records) <= 1:
        return records

    prefix_groups: dict[str, list[int]] = defaultdict(list)
    for i, rec in enumerate(records):
        if not rec.ssn:
            continue
        m = re.match(r'^(\d{3}-\d{2})-', rec.ssn)
        if m:
            prefix_groups[m.group(1)].append(i)

    drop: set[int] = set()
    for prefix, idxs in prefix_groups.items():
        if len(idxs) < 2:
            continue
        full_idxs = [i for i in idxs if _ssn_real_digits(records[i].ssn) == 9]
        if not full_idxs:
            continue
        for i in idxs:
            if i not in full_idxs:
                drop.add(i)

    return [r for i, r in enumerate(records) if i not in drop]


def _drop_exact_duplicates(records: list) -> list:
    """
    Remove records where (first_name, last_name, ssn) are all identical,
    keeping the first occurrence.  Records missing all three keys are kept.
    """
    seen: set[tuple] = set()
    out: list = []
    for rec in records:
        fn  = rec.first_name.strip().lower()
        ln  = rec.last_name.strip().lower()
        ssn = rec.ssn.strip()
        key = (fn, ln, ssn)
        if ssn or (fn and ln):
            if key in seen:
                continue
            seen.add(key)
        out.append(rec)
    return out



def _first_name_prefix_match(a: str, b: str) -> bool:
    """
    Return True only when the entire shorter first name is an exact
    character-by-character prefix of the longer one (case-insensitive,
    leading/trailing spaces ignored).

    Examples:
      "Michale"  vs "Michales" → True  ("Michale" == "Michales"[:7])
      "Michael"  vs "Michale"  → False (differ at index 5)
      "Jame"     vs "James"    → True
      "Anil"     vs "Aeel"     → False
    """
    a = a.strip().lower()
    b = b.strip().lower()
    if not a or not b:
        return False
    shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
    return longer.startswith(shorter)


def _merge_exact_duplicates(records: list) -> list:
    """
    Merge records where last name and SSN are exactly the same AND the first
    names share a common prefix (entire string, one starts with the other,
    case-insensitive).  The longer full first name is kept — last name is
    never modified.  Middle name and suffix are carried over from whichever
    record has them.  Records missing last name, SSN, or first name are untouched.
    """
    if len(records) <= 1:
        return records

    out: list = []

    for rec in records:
        fn  = rec.first_name.strip()
        ln  = rec.last_name.strip().lower()
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
            if not _first_name_prefix_match(fn, primary.first_name):
                continue
            # Keep the longer full first name — never touch last_name.
            # When the incoming record wins on length, its middle name and
            # suffix take priority (carry fields from the selected record).
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


def _deduplicate_extracted(records: list) -> list:
    """
    Deduplication rules for the Extracted sheet:

    A. Different SSN              → keep both.
    B. Same SSN, different last   → keep both.
    C. Same SSN, same last, first names share NO prefix → keep both.
    D. Same SSN, same last, first names share a prefix  → keep only the
       longer-first-name record; carry its middle name and suffix forward.
    """
    if len(records) <= 1:
        return list(records)

    out: list = []

    for rec in records:
        fn  = rec.first_name.strip()
        ln  = rec.last_name.strip().lower()
        ssn = rec.ssn.strip()

        matched = False
        for primary in out:
            # Rule A
            if primary.ssn.strip() != ssn:
                continue
            # Rule B
            if primary.last_name.strip().lower() != ln:
                continue
            # Rule C — no prefix overlap → keep both; try next primary
            if not _first_name_prefix_match(fn, primary.first_name.strip()):
                continue

            # Rule D — prefix match: keep the longer full first name
            if len(fn.strip()) > len(primary.first_name.strip()):
                primary.first_name  = fn.strip()
                primary.middle_name = rec.middle_name or primary.middle_name
                primary.suffix      = rec.suffix      or primary.suffix
            else:
                # Primary has longer/equal name — fill any missing fields from incoming
                if not primary.middle_name and rec.middle_name:
                    primary.middle_name = rec.middle_name
                if not primary.suffix and rec.suffix:
                    primary.suffix = rec.suffix

            matched = True
            break

        if not matched:
            out.append(rec)

    return out


def _parse_name_parts(full_name: str) -> tuple[str, str, str, str]:
    """Return (first, middle, last, suffix) from a full name string.

    Comma is the ONLY structural delimiter recognised:
      With comma:    "Smith, John A Jr" → last=Smith  first=John  middle=A  suffix=Jr
      Without comma: "XXXXXXX"          → last=XXXXXXX first=''   middle='' suffix=''

    When no comma is present the entire value is stored as Last Name and the
    remaining parts are left blank.  This handles masked / unstructured source
    data where space-based splitting would produce incorrect results.
    """
    name = _clean(full_name)
    if not name:
        return ("", "", "", "")

    if "," not in name:
        # No comma → whole value is Last Name; cannot reliably parse further
        return ("", "", name.title(), "")

    # "Last, First [Middle...] [Suffix]"
    last_part, _, rest_part = name.partition(",")
    last = last_part.strip().title()
    rest = rest_part.strip().split()

    suffix = ""
    if rest and rest[-1].lower().rstrip(".") in _SUFFIXES:
        suffix = rest.pop().title()

    first  = rest[0].title() if rest else ""
    middle = " ".join(w.title() for w in rest[1:]) if len(rest) > 1 else ""
    return (first, middle, last, suffix)


# ---------------------------------------------------------------------------
# Employee record
# ---------------------------------------------------------------------------

class EmployeeRecord:
    __slots__ = (
        "source", "page", "doc_id",
        "first_name", "middle_name", "last_name", "suffix",
        "_full_name",
        "ssn", "gross_salary", "ee_contribution",
        "er_match", "ytd_contribution", "ee_id",
    )

    def __init__(self, source: str = "", page: str = "", doc_id: str = ""):
        self.source          = source
        self.page            = page
        self.doc_id          = doc_id
        self.first_name      = ""
        self.middle_name     = ""
        self.last_name       = ""
        self.suffix          = ""
        self._full_name      = ""
        self.ssn             = ""
        self.gross_salary    = ""
        self.ee_contribution = ""
        self.er_match        = ""
        self.ytd_contribution = ""
        self.ee_id           = ""

    def set_field(self, canonical: str, value: str) -> None:
        value = _clean(value)
        if not value:
            return
        if canonical == "name" and not self._full_name:
            self._full_name = value
            first, middle, last, suffix = _parse_name_parts(value)
            if not self.first_name:  self.first_name  = first
            if not self.middle_name: self.middle_name = middle
            if not self.last_name:   self.last_name   = last
            if not self.suffix:      self.suffix      = suffix
        elif canonical == "first_name"  and not self.first_name:
            self.first_name = value.title()
        elif canonical == "middle_name" and not self.middle_name:
            self.middle_name = value.title()
        elif canonical == "last_name"   and not self.last_name:
            self.last_name = value.title()
        elif canonical == "suffix"      and not self.suffix:
            self.suffix = value.title()
        elif canonical == "ssn"         and not self.ssn:
            self.ssn = _fmt_ssn(value)
        elif canonical == "gross_salary"    and not self.gross_salary:
            self.gross_salary = _parse_amount(value) or value
        elif canonical == "ee_contribution" and not self.ee_contribution:
            self.ee_contribution = _parse_amount(value) or value
        elif canonical == "er_match"        and not self.er_match:
            self.er_match = _parse_amount(value) or value
        elif canonical == "ytd_contribution" and not self.ytd_contribution:
            self.ytd_contribution = _parse_amount(value) or value
        elif canonical == "ee_id"           and not self.ee_id:
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
        """Return a list of TEMPLATE_COL_COUNT values matching the template columns."""
        row = [""] * TEMPLATE_COL_COUNT

        row[COL_DOCID        - 1] = self.doc_id
        row[COL_LAST_NAME    - 1] = self.last_name
        row[COL_FIRST_NAME   - 1] = self.first_name
        row[COL_MIDDLE_NAME  - 1] = self.middle_name
        row[COL_SUFFIX       - 1] = self.suffix
        row[COL_SUBJECT_TYPE - 1] = "Employee"

        if self.ssn:
            row[COL_GOV_ID_TAG - 1] = "Social Security Number (SSN)"
            row[COL_SSN        - 1] = self.ssn

        work = self._work_info()
        if work:
            row[COL_WORK_INFO - 1] = work

        if self.ee_id:
            row[COL_EE_ID - 1] = self.ee_id

        return row


# ---------------------------------------------------------------------------
# Strategy 1 — AcroForm field extraction
# ---------------------------------------------------------------------------

def _resolve(obj):
    return obj.get_object() if hasattr(obj, "get_object") else obj


def _extract_acroform(pdf_path: Path, doc_id: str) -> list[EmployeeRecord]:
    if not HAS_PYPDF:
        return []
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return []
    try:
        root     = _resolve(reader.trailer["/Root"])
        acroform = _resolve(root["/AcroForm"])
        fields   = acroform.get("/Fields") or []
    except Exception:
        return []

    raw_fields: list[tuple[str, str]] = []

    def _traverse(ref):
        try:
            f = _resolve(ref)
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

    rec = EmployeeRecord(source=pdf_path.name, page="1", doc_id=doc_id)
    for fname, fval in raw_fields:
        canonical = _map_column(fname)
        if canonical:
            rec.set_field(canonical, fval)
        else:
            m = _SSN_RE.search(fval) or _SSN_RE_XPAD.search(fval)
            if m:
                rec.set_field("ssn", m.group(0))

    return [rec] if rec.is_meaningful() else []


# ---------------------------------------------------------------------------
# Strategy 2 — Table extraction  (full-table SSN scan)
# ---------------------------------------------------------------------------

def _extract_tables(pdf_path: Path, doc_id: str) -> list[EmployeeRecord]:
    """
    Scans EVERY cell of EVERY table on every page for SSN patterns.
    Does NOT require a recognised header row — works even when SSN data
    is buried mid-page or preceded by unrelated content.

    Column roles (name, salary, etc.) are detected opportunistically from
    any header-like row found above the SSN row. If no header exists, the
    script falls back to scanning the same row for a name candidate.
    """
    if not HAS_PDFPLUMBER:
        return []

    records: list[EmployeeRecord] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            try:
                tables = page.extract_tables()
            except Exception:
                continue

            for tbl in (tables or []):
                if not tbl:
                    continue

                # ── Pre-scan: detect column roles from any header-like row ──
                col_roles: dict[int, str] = {}  # col_idx → canonical field
                for row_i, row in enumerate(tbl):
                    tentative: dict[int, str] = {}
                    for col_i, cell in enumerate(row):
                        canonical = _map_column(str(cell or ""))
                        if canonical:
                            tentative[col_i] = canonical
                    if len(tentative) >= 2:
                        col_roles = tentative
                        break

                # ── Main scan: look at every row for an SSN ─────────────────
                for row_idx, row in enumerate(tbl):
                    if not row or all(not _clean(str(c)) for c in row):
                        continue

                    ssn_val = ''
                    ssn_col = -1

                    # Find SSN in any cell of this row
                    for col_i, cell in enumerate(row):
                        found = _ssn_from_cell(str(cell or ""))
                        if found:
                            ssn_val = found
                            ssn_col = col_i
                            break

                    if not ssn_val:
                        continue

                    rec = EmployeeRecord(
                        source=pdf_path.name,
                        page=str(page_num),
                        doc_id=doc_id,
                    )
                    rec.ssn = ssn_val

                    # ── Apply known column roles ─────────────────────────────
                    for col_i, canonical in col_roles.items():
                        if col_i == ssn_col or col_i >= len(row):
                            continue
                        rec.set_field(canonical, str(row[col_i] or ""))

                    # ── Name fallback: check SSN cell for embedded name ─────
                    # pdfplumber sometimes merges Name+SSN into one cell,
                    # e.g. "Adeyemi, Chukwuemeke 611-56-73". Strip the SSN
                    # from that cell and look for a name in the remainder.
                    if not rec._full_name and not rec.first_name:
                        ssn_cell_text = _clean(str(row[ssn_col] or ""))
                        name_part = re.sub(r'\d{3}-\d{2}-[\dxX*]{1,4}', '',
                                           ssn_cell_text, flags=re.I).strip()
                        name = _find_name_in_text(name_part)
                        if name:
                            rec.set_field("name", name)

                    # ── Name fallback: scan every non-SSN cell in same row ───
                    if not rec._full_name and not rec.first_name:
                        for col_i, cell in enumerate(row):
                            if col_i == ssn_col:
                                continue
                            cell_text = _clean(str(cell or ""))
                            if not cell_text:
                                continue
                            name = _find_name_in_text(cell_text)
                            if name:
                                rec.set_field("name", name)
                                break

                    # ── Name fallback: look at up to 3 rows above then 1 below ─
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
                                name = _find_name_in_text(cell_text)
                                if name:
                                    rec.set_field("name", name)
                                    break

                    if rec.is_meaningful():
                        records.append(rec)

    return records


# ---------------------------------------------------------------------------
# Strategy 3 — Full-page text scan (fallback / supplement)
# ---------------------------------------------------------------------------

_LABEL_PATS: dict[str, re.Pattern] = {
    "name":             re.compile(r'(?:employee|participant|emp(?:loyee)?)\s+name\s*[:\-]?\s*([A-Za-z][A-Za-z ,\.\'\-]+)', re.I),
    "first_name":       re.compile(r'first\s+name\s*[:\-]?\s*(.+)', re.I),
    "last_name":        re.compile(r'last\s+name\s*[:\-]?\s*(.+)', re.I),
    "middle_name":      re.compile(r'middle\s+(?:name|initial)\s*[:\-]?\s*(.+)', re.I),
    "suffix":           re.compile(r'(?:name\s+)?suffix\s*[:\-]?\s*(.+)', re.I),
    "ssn":              re.compile(r'(?:ssn|ss#|social\s+security(?:\s+number)?|tax\s+id)\s*[:\-#]?\s*([\d\- ]+)', re.I),
    "gross_salary":     re.compile(r'(?:gross\s+(?:salary|pay|wages?|comp(?:ensation)?)|annual\s+salary|salary)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "ee_contribution":  re.compile(r'(?:ee\s+(?:contribution|deferral)|employee\s+(?:contribution|deferral|401k)|elective\s+deferral)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "er_match":         re.compile(r'(?:er\s+match|employer\s+(?:match|contribution)|company\s+match)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "ytd_contribution": re.compile(r'(?:ytd\s+(?:contribution|total|deferral)|year\s+to\s+date)\s*[:\-]?\s*([\$\d,\.]+)', re.I),
    "ee_id":            re.compile(r'(?:employee\s+(?:id|number|identification)|ee\s+id|emp\s+(?:id|number))\s*[:\-#]?\s*([A-Z0-9\-]+)', re.I),
}


def _collect_ssn_positions(lines: list[str]) -> list[tuple[int, str]]:
    """Return (line_index, formatted_ssn) for every SSN found in the line list."""
    positions: list[tuple[int, str]] = []
    for line_i, line in enumerate(lines):
        found_on_line = False
        # Full digit SSN: DDD-DD-DDDD or DDDDDDDDD
        for m in _SSN_RE.finditer(line):
            positions.append((line_i, _fmt_ssn(m.group(0))))
            found_on_line = True
        # X-padded SSN: DDD-DD-DDXX, DDD-DD-DXXX, DDD-DD-XXXX (only when full not found)
        if not found_on_line:
            for m in _SSN_RE_XPAD.finditer(line):
                positions.append((line_i, _fmt_ssn(m.group(0))))
                found_on_line = True
        if not found_on_line:
            for pat, kind in _SSN_PARTIAL_PATTERNS:
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
                    positions.append((line_i, _fmt_ssn(partial)))
                    break
    return positions


def _extract_regex(pdf_path: Path, doc_id: str) -> list[EmployeeRecord]:
    """
    Scans every line of every page for SSN patterns — works regardless of
    where on the page the SSN appears.  For each SSN found:
      1. Checks text on the SAME line (before and after the SSN) for a name.
      2. Checks a ±20-line window using labelled patterns.
      3. Checks adjacent lines (±5) using broad name detection.
    """
    if not HAS_PDFPLUMBER:
        return []

    records: list[EmployeeRecord] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text  = page.extract_text() or ""
            lines = text.splitlines()

            ssn_positions = _collect_ssn_positions(lines)

            if not ssn_positions:
                # No SSNs found — try labelled-only extraction
                rec = EmployeeRecord(source=pdf_path.name, page=str(page_num), doc_id=doc_id)
                for canonical, pat in _LABEL_PATS.items():
                    for line in lines:
                        m = pat.search(line)
                        if m:
                            rec.set_field(canonical, m.group(1))
                            break
                if rec.is_meaningful():
                    records.append(rec)
                continue

            for ssn_line_i, ssn_val in ssn_positions:
                rec = EmployeeRecord(source=pdf_path.name, page=str(page_num), doc_id=doc_id)
                rec.ssn = ssn_val

                # ── 1. Search the SAME line for a name ───────────────────────
                raw_line = lines[ssn_line_i]
                # Remove the SSN itself, then look at what remains
                cleaned_line = re.sub(r'\d{3}-\d{2}-[\dxX*]{1,4}', '', raw_line, flags=re.I).strip()
                name = _find_name_in_text(cleaned_line)
                if name:
                    rec.set_field("name", name)

                # ── 2. Labelled fields in ±20-line window ────────────────────
                w_start = max(0, ssn_line_i - 20)
                w_end   = min(len(lines), ssn_line_i + 20)
                window  = "\n".join(lines[w_start:w_end])

                for canonical, pat in _LABEL_PATS.items():
                    if canonical == "ssn":
                        continue
                    m2 = pat.search(window)
                    if m2:
                        candidate = _clean(m2.group(1))
                        if not _AMOUNT_RE.fullmatch(candidate) and not _SSN_RE.search(candidate):
                            rec.set_field(canonical, candidate[:80])

                # ── 3. Broad name search in ±10 adjacent lines ───────────────
                if not rec._full_name and not rec.first_name:
                    for offset in (-1, 1, -2, 2, -3, 3, -4, 4, -5, 5,
                                   -6, 6, -7, 7, -8, 8, -9, 9, -10, 10):
                        check_i = ssn_line_i + offset
                        if not (0 <= check_i < len(lines)):
                            continue
                        check_line = re.sub(r'\d{3}-\d{2}-[\dxX*]{1,4}', '',
                                            lines[check_i], flags=re.I).strip()
                        name = _find_name_in_text(check_line)
                        if name:
                            rec.set_field("name", name)
                            break

                if rec.is_meaningful():
                    records.append(rec)

    return records


# ---------------------------------------------------------------------------
# Main extraction orchestrator
# ---------------------------------------------------------------------------

def extract_401k(pdf_path: Path, doc_id: str) -> list[EmployeeRecord]:
    """
    Extraction pipeline:
      1. AcroForm  — fillable PDFs (if found, used exclusively)
      2. Table scan — scans every cell of every table for SSNs
      3. Text scan  — scans every text line for SSNs (runs independently of table scan)
    All records are returned as-is with no merging or deduplication.
    """
    # AcroForm: structured fillable PDF — complete and authoritative
    records = _extract_acroform(pdf_path, doc_id)
    if records:
        records = _drop_redacted_duplicates(records)
        records = _drop_exact_duplicates(records)
        records = _merge_exact_duplicates(records)
        print(f"    [AcroForm] {len(records)} employee record(s) found.")
        return records

    # Table + text scan run fully independently — no shared seen_ssns
    tbl_records  = _extract_tables(pdf_path, doc_id)
    text_records = _extract_regex(pdf_path, doc_id)

    # Drop X-padded SSN duplicates, then exact duplicates, then merge prefix-name matches
    records = _merge_exact_duplicates(_drop_exact_duplicates(_drop_redacted_duplicates(tbl_records + text_records)))

    method_parts = []
    if tbl_records:
        method_parts.append(f"Table:{len(tbl_records)}")
    if text_records:
        method_parts.append(f"Text:{len(text_records)}")

    if records:
        method_str = " + ".join(method_parts) if method_parts else "no-header scan"
        print(f"    [{method_str}] {len(records)} employee record(s) found.")
    else:
        print(f"    [!] No employee records found — check PDF layout.")

    return records


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------

def _csv_safe(val) -> str:
    s = str(val or "")
    return ("'" + s) if s[:1] in ("=", "+", "-", "@") else s


_HEADER = [
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


def _write_sheet(ws, header: list, records: list) -> None:
    """Write header + data rows to a worksheet."""
    ws.append(header)
    for rec in records:
        ws.append([_csv_safe(v) for v in rec.to_template_row()])
    if not records:
        ws.append(["No employee records extracted."])


def write_workbook(
    out_path: Path,
    records: list[EmployeeRecord],
) -> None:
    tmp = out_path.with_suffix(".tmp.xlsx")
    wb  = openpyxl.Workbook()

    hdr = _HEADER

    # ── "Employees" sheet — all extracted records ─────────────────────────────
    ws = wb.active
    ws.title = "Employees"
    _write_sheet(ws, hdr, records)

    # ── "Extracted" sheet — deduplicated per SSN + name rules ────────────────
    ws_ext = wb.create_sheet(title="Extracted")
    _write_sheet(ws_ext, hdr, _deduplicate_extracted(records))

    wb.save(tmp)
    wb.close()

    for attempt in range(5):
        try:
            os.replace(tmp, out_path)
            break
        except PermissionError:
            if attempt < 4:
                time.sleep(1)
            else:
                print(f"  WARNING: could not write {out_path.name} — file locked.")
                try:
                    tmp.unlink(missing_ok=True)
                except OSError:
                    pass


# ---------------------------------------------------------------------------
# Progress bar
# ---------------------------------------------------------------------------

class _SimpleProgress:
    """Inline progress bar used when tqdm is not installed."""

    _BAR = 35

    def __init__(self, total: int) -> None:
        self.total     = total
        self.processed = 0
        self._print_header()

    def _print_header(self) -> None:
        print(f"  {'Total':<10}{'Processed':<12}{'Remaining':<12}  File")
        print("  " + "-" * 70)

    def update(self, filename: str, extra: str = "") -> None:
        self.processed += 1
        remaining = self.total - self.processed
        filled    = int(self._BAR * self.processed / max(self.total, 1))
        bar       = "#" * filled + "-" * (self._BAR - filled)
        name      = filename[:34]
        line = (
            f"\r  [{bar}] "
            f"{self.total:<6} "
            f"{self.processed:<12}"
            f"{remaining:<12}  "
            f"{name:<34}"
            f"  {extra}"
        )
        sys.stdout.write(line)
        sys.stdout.flush()

    def close(self) -> None:
        sys.stdout.write("\n")
        sys.stdout.flush()


# ---------------------------------------------------------------------------
# File iteration
# ---------------------------------------------------------------------------

def _iter_pdf(target: Path, recursive: bool):
    if target.is_file():
        yield target
        return
    pattern = "**/*.pdf" if recursive else "*.pdf"
    yield from sorted(target.glob(pattern))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract employee data from 401k PDFs into the PII format."
    )
    parser.add_argument("target", nargs="?",
                        help="PDF file or folder (prompted if omitted)")
    parser.add_argument("--recursive", action="store_true",
                        help="Recurse into subfolders")
    parser.add_argument("--output", default=None,
                        help="Output folder for extracted files (default: same folder as PDFs)")
    args = parser.parse_args()

    target_str = args.target
    if not target_str:
        target_str = input("Enter path to PDF file or folder: ").strip().strip('"').strip("'")
    if not target_str:
        sys.exit("ERROR: no path provided.")

    missing = [pkg for pkg, ok in [("pdfplumber", HAS_PDFPLUMBER), ("openpyxl", HAS_OPENPYXL)]
               if not ok]
    if missing:
        sys.exit(f"ERROR: missing packages — run:  pip install {' '.join(missing)}")
    if not HAS_PYPDF:
        print("WARNING: pypdf not installed — AcroForm extraction skipped.\n"
              "         Run:  pip install pypdf")

    target = Path(target_str)
    if not target.exists():
        sys.exit(f"ERROR: path not found — {target}")

    files = list(_iter_pdf(target, args.recursive))
    if not files:
        sys.exit(f"ERROR: no .pdf files found at {target}")

    total     = len(files)
    output_dir = Path(args.output) if args.output else (
        target if target.is_dir() else target.parent
    )

    print(f"\nFound {total} PDF file(s). Starting extraction...\n")

    total_records = 0
    errors: list[str] = []

    if HAS_TQDM:
        bar = _tqdm(
            files,
            desc="Processing",
            unit="doc",
            ncols=90,
            bar_format=(
                "  {l_bar}{bar}| "
                "{n_fmt}/{total_fmt} docs  "
                "Remaining: {remaining_s}  "
                "[{elapsed}]"
            ),
        )
    else:
        bar = None
        progress = _SimpleProgress(total)

    for pdf_path in (bar if HAS_TQDM else files):
        doc_id = pdf_path.stem

        if HAS_TQDM:
            bar.set_description(f"Processing  {pdf_path.name[:35]:<35}")

        try:
            records   = extract_401k(pdf_path, doc_id)
            out_path  = output_dir / f"{doc_id}_extracted.xlsx"
            write_workbook(out_path, records)
            total_records += len(records)
            msg = f"{len(records)} record(s)"
        except Exception as exc:
            errors.append(f"{pdf_path.name}: {exc}")
            msg = "ERROR"

        if HAS_TQDM:
            processed  = bar.n + 1          # +1 because tqdm updates after the loop body
            remaining  = total - processed
            bar.set_postfix({"records": total_records, "remaining": remaining}, refresh=True)
        else:
            progress.update(pdf_path.name, msg)

    if HAS_TQDM:
        bar.close()
    else:
        progress.close()

    if errors:
        print("\nErrors:")
        for e in errors:
            print(f"  [!] {e}")

    print(f"\nDone.  Total docs: {total}  |  Processed: {total - len(errors)}  |  "
          f"Failed: {len(errors)}  |  Employee records: {total_records}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
