"""
260716 pd ds unknown merge.py

Merge PII/PHI person records from an Excel export into one row per
confirmed person, for the notification report. Ten rules, built step by
step:

  Rule 1 (SSN Exists): rows with the same (non-blank) SSN are merged, AS
                        LONG AS both rows ALSO have a real (non-blank) DOB
                        and it's the SAME value - a shared SSN is NOT
                        trusted enough on its own; a blank DOB on either
                        side no longer counts as "no conflict" (unlike
                        Rules 4-8/9-10 below) - DOB must genuinely agree.
                        UNLESS:
                          - First, Middle, or Last Name genuinely disagrees
                            between the two rows (name_conflict()) - blank
                            on either side is fine, and an incomplete/
                            truncated entry that's a prefix of the fuller
                            one (e.g. "Did" vs "Didar") is NOT treated as a
                            conflict, but a real difference is; or
                          - Suffix is present on both sides and differs
                            (e.g. "Jr" vs "Sr").
                        In any of these cases the two rows are kept separate
                        rather than merged.
  Rule 2 (Exact Name, DOB): rows with a real (non-blank) First Name + Last
                        Name on both sides that are COMPATIBLE - not
                        required to be byte-exact, an incomplete/truncated
                        entry (e.g. an initial like "J" for "Jeffrey", or
                        "Did" for "Didar") counts as compatible, not a
                        mismatch (see name_prefix_compat()) - AND the same
                        DOB are merged, as long as their SSNs don't conflict
                        (blank on either/both sides is fine) and Middle
                        Name/Suffix don't actively disagree.
  (Rule 3 - a Name-Only, no-SSN/no-DOB fallback - previously existed here
  and has been removed. Numbering below is left as-is, with a gap at 3,
  rather than renumbering Rules 4-9.)
  Rule 4 (Employee ID, Name): a STRONG identifier, same trust tier as SSN.
                        Rows sharing at least one common Employee ID (a cell
                        can already contain multiple semicolon-joined IDs)
                        are merged as long as Name does NOT actively
                        CONFLICT (see name_conflict() - blank/placeholder on
                        either side, e.g. "Unknown"/"UNK", is NOT a conflict,
                        and an incomplete/truncated entry like "J" for
                        "Jeffrey" isn't either) and SSN/DOB are each either
                        matching or blank on both sides (a real, differing
                        value on both sides blocks it). A real name is NOT
                        required on both sides - a shared Employee ID alone
                        merges a placeholder-name row into the same person,
                        the same way Rule 1 already treats a blank/
                        placeholder name as no conflict for a shared SSN.
  Rule 5 (Driver's License, Name): same as Rule 4 (STRONG identifier, no
                        real name required on both sides), but matched on
                        Driver's License Number instead of Employee ID.
  Rule 6 (Passport Number, Name): same as Rule 4, but matched on Passport
                        Number instead of Employee ID.
  Rule 7 (Phone Number, Name): a WEAK identifier (can legitimately be shared
                        by different people, e.g. a household phone) -
                        unlike Rules 4-6/10, this REQUIRES a compatible,
                        REAL First+Last Name on both sides (see
                        name_compat_match() - an incomplete/truncated entry
                        still counts as compatible, not a mismatch), matched
                        on Phone Number instead of Employee ID.
  Rule 8 (Full Address, Name): Residential Address, City, State, Zip, and
                        Province are each checked individually - blank on
                        either side is fine, but a real, differing value in
                        ANY of them blocks the merge, and at least one field
                        must genuinely match on both sides. Unlike SSN,
                        address is NEVER enough to override a name
                        difference on its own - this also always requires
                        a compatible, REAL Name match (name_compat_match())
                        and SSN/DOB to be matching-or-blank, same as the
                        WEAK-identifier Rules 7, 9.
  Rule 9 (Email, Name): a WEAK identifier, same as Rule 7 (requires a
                        compatible, REAL Name on both sides) but matched on
                        Email Address - Personal instead of Phone Number.
  Rule 10 (Government-Issued ID Number, Name): a STRONG identifier, same as
                        Rule 4 (no real name required on both sides), but
                        matched on Government-Issued ID Number instead of
                        Employee ID.
  Rule 11 (Email, All PII Compatible): a last-resort bridge. Rows sharing an
                        Email Address token are merged even when only ONE
                        side has a real Name (e.g. the other side is
                        "[Unknown]" or truly blank - name_conflict() treats
                        that as no conflict, same tolerance as Rules 1/4-6/
                        10), AS LONG AS every other piece of PII - SSN, DOB,
                        Employee ID, Driver's License, Passport Number,
                        Government ID, Phone Number, and the Address fields
                        - is each either matching or blank on both sides (a
                        real, differing value in ANY of them blocks it; see
                        _all_pii_compatible()). Unlike Rules 4-6/10, Email
                        alone is still a WEAK identifier (a shared household
                        email), so this rule only trusts it to bridge an
                        Unknown/blank-named row once literally nothing else
                        contradicts - it does not get the address-override
                        trust that a STRONG match gets (see below).
  Rule 12 (Unknown/Partial Name, PII Overlap): a second last-resort bridge,
                        for the case where only PART of a Name is Unknown -
                        e.g. First is a real name but Last is "[Unknown]",
                        or vice versa - which Rules 4-6/10 do NOT bridge
                        (they only bridge a name that's fully blank or
                        fully "[Unknown]"/"[Unknown]", see
                        Rec.name_bridge_ok) and Rules 7-9 do not bridge
                        either (they require a REAL name on both sides).
                        Rows are merged when Name does not actively CONFLICT
                        (name_conflict() - so a partial-Unknown/blank Name
                        part on either side is fine, same tolerance as
                        Rule 1; only a genuine difference between two REAL
                        values in the SAME part blocks it), AND at least one
                        OTHER identifying field genuinely OVERLAPS (a shared
                        Employee ID/DL/Passport/Gov ID/Phone/Email token, or
                        a matching Address field - see _weak_overlap() -
                        deliberately NOT a bare SSN or DOB match alone, so
                        this can never reopen Rule 1's guard against
                        trusting a reused/junk SSN without a genuinely
                        agreeing DOB), AND every other PII field (including
                        SSN/DOB) is matching-or-blank (see
                        _all_pii_compatible()). Requiring genuine overlap
                        somewhere - never merging two rows purely because
                        neither one has enough data to contradict the other
                        - is what keeps this rule from collapsing every
                        near-empty row in the file into one person.

Either rule matching is enough to merge, and the match is transitive (if
A matches B and B matches C, all three end up in one merged row, even if
A and C don't directly match each other) - EXCEPT that a merge is refused
whenever it would combine 2+ DIFFERENT known SSNs, or 2+ DIFFERENT known
DOBs, into one group (this can happen via a blank-SSN/blank-DOB "bridge"
row that matches two otherwise-unrelated people). Rather than un-merging
the whole cluster, only the specific union that would cross real SSNs/DOBs
is refused - so e.g. 5 rows sharing one SSN and 3 rows sharing a different
SSN, bridged by a blank-SSN row, still end up as 2 clean groups instead of
8 separate rows (see group_ssn/group_dob/try_union() in main()).

A conflicting real ADDRESS gets the same transitive-bridge protection, but
only for the WEAK-evidence Rules 8 (Full Address, Name), 11 (Email, All PII
Compatible), and 12 (Unknown/Partial Name, PII Overlap) themselves - a
STRONG match (Rule 1 SSN, Rule 2 Name+DOB, or Rules 4-7/9-10 ID+Name) is
trusted enough to override a conflicting address on its own, since the same
person legitimately having two different addresses on file (e.g. they
moved) is far more likely than two different people sharing an SSN/DOB/ID
(see group_addr/try_union() in main()).

INPUT  : an Excel workbook with the columns listed in EXPECTED_COLS below.
OUTPUT : a new Excel workbook with two sheets:
         - "Merged Notification Data": ONE ROW PER CONFIRMED PERSON.
           - First Name, Middle Name, Last Name, and SSN: the single
             fullest/most complete value among the merged rows (placeholder
             values like "[Unknown]" already in the input are never picked
             as that value). If EVERY merged row's First/Last Name is blank
             or a placeholder, First Name and Last Name are each set to the
             literal "[Unknown]" instead of being left blank.
           - Suffix: the single fullest non-blank value.
           - DOB: the one real date shared by every row in the group,
             always displayed as a clean "MM/DD/YYYY" string - compared by
             the NORMALIZED date (see norm_dob()), never raw cell text, so
             the same date typed in different formats across rows (e.g.
             "01/01/1990" vs "1990-01-01") is treated as one date, not two,
             and a raw Excel SERIAL date number (e.g. "20037" - which the
             pyxlsb engine for .xlsb files hands back as-is for a
             date-formatted cell, unlike openpyxl for .xlsx) never leaks
             into the output (see dob_merge()). A group where 2+ rows have
             GENUINELY DIFFERENT real DOBs (possible via a blank-DOB bridge
             row - see group_dob/try_union()) is not merged at all here -
             every one of its original rows goes to "DOB Conflict Review"
             instead (see below).
           - Employee ID, Driver's License, and Passport Number: every
             distinct ID TOKEN seen (cells can already contain multiple
             semicolon-joined IDs - see parse_id_tokens()), deduplicated
             and joined with "; ".
           - Every OTHER column (DOCIDs, Gov ID, etc.): every distinct
             value seen, joined with "; ". If a group's merged DOCID list
             would exceed Excel's 32,767-char cell limit, it spills into
             extra "DOCIDs 2", "DOCIDs 3", ... columns, splitting only at
             "; " boundaries (see split_docid_chunks()).
           - Address fields (Residential Address, City, State, Province,
             Zip, Country) are kept TOGETHER as one unit: the most common
             address (by total row count) stays in those columns - with any
             field left blank on one row filled in from another row's fuller
             copy of that SAME address, so e.g. a blank Zip on one row never
             by itself creates a spurious extra address - and every OTHER,
             genuinely different address goes into a new "Other Address"
             column as one combined string per address, semicolon-joined
             (see address_key_conflict()).
         - "DOB Conflict Review": any candidate group where 2+ rows have
           genuinely different real DOBs is NOT merged at all - every
           original row from such a group is listed here as-is (unmerged),
           tagged with a "Candidate Group ID" (rows sharing one ID were
           clustered together) and "Candidate Group Size", for manual
           review. These rows do NOT appear in "Merged Notification Data".

Every group is merged regardless of size - a group of 50+ rows sharing one
value (e.g. a reused/junk SSN) is merged the same as a group of 2, with no
separate size-based review step.

This script does not touch the input file. Save the output only to the
secured/authorized folder for this data (never a desktop) - it contains
SSN, DOB, and other PII/PHI.

Designed for large row counts (uses "blocking" - only compares rows that
already share an exact SSN, an exact First+Last Name+DOB, a Last Name, or
an Employee ID + Name - instead of comparing every row to every other row).

Install once:
    pip install pandas openpyxl pyxlsb

Run:
    python "260716 pd ds unknown merge.py"
        (no argument - opens a native file-picker dialog; accepts a .csv,
        .xlsx, .xls, or .xlsb file, chosen by its extension - see
        pick_input_file()/load_input_file())
    python "260716 pd ds unknown merge.py" input_file.csv
        (skips the dialog, reads that file directly)
    python "260716 pd ds unknown merge.py" input_file.csv output.xlsx
        (also overrides the output filename)
"""

import sys
import os
import re
import itertools
import unicodedata
from collections import defaultdict
from multiprocessing import Pool
from tkinter import Tk, filedialog

import pandas as pd

# Worker processes for the pairwise clustering step (the slowest phase on
# large files). Plain `threading` would NOT help here - is_match() is pure
# CPU-bound Python, and the GIL serializes threads so they'd just take turns
# instead of running in parallel. Separate processes actually parallelize it.
PARALLEL_WORKERS = 4
# Below this many candidate pairs, run single-process instead - spinning up
# worker processes has real overhead (~tens of ms each) that isn't worth it
# for a small/quick run.
PARALLEL_THRESHOLD = 20_000

# ------------------------------------------------------------
# Progress bar - one line, updated in place, no per-item explanation.
# ------------------------------------------------------------
_last_pct = {}


def progress(label: str, current: int, total: int, extra: str = "") -> None:
    """Prints '[label]  42% |########------------|  420,000/1,000,000  extra'
    on a single line, overwriting itself. Only redraws when the whole percent
    value changes, so it's cheap to call on every loop iteration."""
    pct = 100 if total <= 0 else min(100, current * 100 // total)
    if _last_pct.get(label) == pct and current != total:
        return
    _last_pct[label] = pct
    bar = "#" * (pct // 5) + "-" * (20 - pct // 5)
    tail = f"  {extra}" if extra else ""
    end = "\n" if current >= total else ""
    print(f"\r  [{label}] {pct:3d}% |{bar}| {current:,}/{total:,}{tail}   ",
          end=end, flush=True)


# ------------------------------------------------------------
# 1) CONFIG - edit these to match your workbook
# ------------------------------------------------------------
# Input file is no longer a fixed path here - see pick_input_file()/
# load_input_file() in main() (file-picker dialog, or a CLI argument).
# Only used when the input is Excel with multiple sheets (ignored for .csv).
INPUT_SHEET = 0
OUTPUT_XLSX = "260716 re ds unknown merge output.xlsx"

COL_DOCID  = "Control ID"
COL_FIRST  = "First Name"
COL_LAST   = "Last Name"
COL_MIDDLE = "Middle Name"
COL_SUFFIX = "Suffix"
COL_DOB    = "Full Date of Birth (MM/DD/YYYY)"
COL_SSN    = "Social Security Number"

# COL_EMPID/COL_DL/COL_PASSPORT/COL_PHONE/COL_EMAIL/COL_GOVID are used for
# matching (Rules 4-7, 9-10) AND kept as plain semicolon-merged columns in
# OTHER_MERGE_COLS below (matching doesn't remove a column from display).
COL_DL       = "Driver's License Number"
COL_PASSPORT = "Passport Number"
COL_GOVID    = "Government-Issued ID Number"
COL_EMPID    = "Employee Identification Number"
COL_PHONE    = "Phone Number"
COL_EMAIL    = "Email Address - Personal"

# Address fields are handled specially (see ADDRESS_COLS below), not as
# plain semicolon-merged columns - they need to stay together as one unit
# per address, not be shuffled independently per field.
COL_ADDR    = "Residential Address"
COL_CITY    = "City"
COL_STATE   = "State of Residence (if US)"
COL_PROVINCE = "Province of Residence (if Canada)"
COL_ZIP     = "Zip Code"
COL_COUNTRY = "Country of Residence"

# The full set of fields that make up "one address". The MAJORITY (most
# common) address among a merged person's rows is kept in these columns as-is;
# every OTHER distinct address goes into the "Other Address" output column
# as one combined string per address, semicolon-joined.
ADDRESS_COLS = [COL_ADDR, COL_CITY, COL_STATE, COL_PROVINCE, COL_ZIP, COL_COUNTRY]

# Output-only column holding every non-majority address (see split_addresses()
# below). Not part of ADDRESS_COLS/EXPECTED_COLS - but if the INPUT already
# has it (e.g. it's the output of a prior merge), its content is carried
# forward and merged in too, rather than being silently overwritten by this
# run's own freshly-computed "Other Address" value.
COL_OTHER_ADDR = "Other Address"

# Every other column in the sheet - these get semicolon-merged as-is. Must
# match the INPUT file's real header text EXACTLY (a mismatch here - e.g.
# this list says "PI Notes" but the file's actual header is "PII Notes" -
# means the real column is never read for merging AND the header shows up
# in the wrong position in the output, since resolve_extra_merge_cols() in
# main() only catches a column this list doesn't know about AT ALL, not a
# near-miss spelling of one it does). Edit this list if your real headers
# differ - or, for a one-off unrecognized column, rely on
# resolve_extra_merge_cols() to auto-detect and merge it in unchanged.
OTHER_MERGE_COLS = [
    "Unique ID",
    "Data Sub Type",
    "Birth Information",
    "Address Comments",
    "Other Addresses",
    "PI Type",
    "Contact Information",
    "Access Credentials (Non-Financial Account)",
    COL_EMAIL,
    COL_PHONE,
    COL_DL,
    "DL Issuing Country",
    "DL Issuing Province (if Canada)",
    "DL Issuing State (if US)",
    "Passport Country",
    COL_PASSPORT,
    "Government ID Issuing Country",
    "Government- Issued Identification",
    COL_GOVID,
    "Health Related Information",
    COL_EMPID,
    "Tax Identification Number",
    "Work-Related Information",
    "Family Information",
    "Financial Account Information",
    "Student-Related Information",
    "Demographic Information",
    "Biometric Data",
    "PII Notes",
    "Partial Social Security Number",
    "Director identification Number (DIN)",
    "Indian Tax Identification PAN",
    "Digital Signature",
    "International Banking Account Number",
    "National identification number",
    "Username",
    "Password",
    "U.K. National Insurance Number",
    "Financial Account Number",
    "Digital Signature - Non-financial",
    "Credit/Debit Card Expiration",
    "Religion / Philosophical Beliefs",
    "Credit Card Number",
]

EXPECTED_COLS = (
    [COL_DOCID, COL_FIRST, COL_LAST, COL_MIDDLE, COL_SUFFIX, COL_DOB, COL_SSN]
    + ADDRESS_COLS + OTHER_MERGE_COLS
)

# Some source exports have varied the exact spelling of these headers from
# run to run (an apostrophe present/absent; periods vs. plain spaces; an
# abbreviated "Unq ID" instead of the full "Unique ID") - rather than
# hard-failing whenever a different variant shows up, each group below lists
# every spelling seen so far for ONE logical column. Whichever spelling is
# actually present in a given input file gets renamed to the group's
# canonical (first) name - the one used everywhere else in this script -
# right after reading the file (see resolve_column_aliases()/main()), so the
# rest of the script never needs to know an alternate spelling exists.
COLUMN_ALIASES = [
    [COL_DL, "Drivers License Number"],
    ["U.K. National Insurance Number", "U K  National Insurance Number"],
    ["Unique ID", "Unq ID"],
]

MERGE_SEP = "; "
EXCEL_MAX_ROWS = 1_048_576

# Placeholder name values treated as blank (never match/conflict on their
# own; a real name always supersedes these). Checked after stripping
# brackets/parens/periods, so "[Unknown]", "(unknown)", "N/A" all match.
NAME_PLACEHOLDERS = {
    "UNKNOWN", "UNK", "UNKN", "NA", "NONE", "NULL", "NIL",
    "XXX", "XX", "X", "NMN", "NONAME", "NOTGIVEN", "NOTPROVIDED",
}

# ------------------------------------------------------------
# 2) Normalization helpers
# ------------------------------------------------------------
# Zero-width space (U+200B), zero-width non-joiner (U+200C), zero-width
# joiner (U+200D), word joiner (U+2060), BOM (U+FEFF), non-breaking space
# (U+00A0), soft hyphen (U+00AD) - common invisible characters in Excel/CSV
# exports that make two values which LOOK identical fail an exact-string
# comparison. Built from chr() codes (not literal characters pasted into
# this file), so the source itself never contains an actual invisible char.
_INVISIBLE_CODEPOINTS = (0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF, 0x00A0, 0x00AD)
_INVISIBLE_CHARS = re.compile("[" + "".join(chr(c) for c in _INVISIBLE_CODEPOINTS) + "]")


def _numeric_cell_to_str(v) -> str:
    """Coerces a raw cell value to text without letting pandas' float parsing
    corrupt a whole-number ID (SSN, Employee ID, Driver's License, Passport,
    Phone). A purely-numeric column (no dashes, no leading zero) gets read as
    float64 by pandas whenever that column has even one blank cell elsewhere
    - so a valid value like 123456789 silently becomes 123456789.0, and
    str()'ing that adds a spurious extra digit ('123456789.0' -> 10 digits
    after stripping the dot), making an otherwise-identical ID fail to match
    another row's clean text/int version of the same number. A whole-number
    float is converted via int() first to avoid this."""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def norm_text(v) -> str:
    """Used as the DEDUP KEY everywhere (semicolon_merge, address_key, name
    matching). Normalizes unicode (NFKC - so visually-identical characters
    with different encodings compare equal), strips invisible characters
    (zero-width space/joiner, BOM, soft hyphen, non-breaking space) that are
    common in Excel/CSV exports and otherwise make two values that LOOK the
    same fail an exact-string dedup check, then collapses whitespace and
    upper-cases."""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    s = _INVISIBLE_CHARS.sub(" ", s)
    s = re.sub(r"\s+", " ", s).strip().upper()
    return "" if s.lower() in ("nan", "none", "null") else s


# Street-address abbreviation map: each variant -> one canonical token, so a
# street written with a full word and one written with the USPS abbreviation
# ("123 West Lane" vs "123 W Ln") normalize to the same value. Covers the
# common directionals and street-type suffixes; anything not listed is left
# unchanged. Only applied to the Residential Address (street) field.
_STREET_TOKEN_MAP = {
    # Directionals
    "NORTH": "N", "N": "N",
    "SOUTH": "S", "S": "S",
    "EAST": "E", "E": "E",
    "WEST": "W", "W": "W",
    "NORTHEAST": "NE", "NE": "NE",
    "NORTHWEST": "NW", "NW": "NW",
    "SOUTHEAST": "SE", "SE": "SE",
    "SOUTHWEST": "SW", "SW": "SW",
    # Street-type suffixes
    "STREET": "ST", "ST": "ST",
    "AVENUE": "AVE", "AVE": "AVE", "AV": "AVE",
    "BOULEVARD": "BLVD", "BLVD": "BLVD",
    "ROAD": "RD", "RD": "RD",
    "LANE": "LN", "LN": "LN",
    "DRIVE": "DR", "DR": "DR",
    "COURT": "CT", "CT": "CT",
    "CIRCLE": "CIR", "CIR": "CIR",
    "PLACE": "PL", "PL": "PL",
    "TERRACE": "TER", "TERR": "TER", "TER": "TER",
    "PARKWAY": "PKWY", "PKWY": "PKWY",
    "HIGHWAY": "HWY", "HWY": "HWY",
    "SQUARE": "SQ", "SQ": "SQ",
    "TRAIL": "TRL", "TRL": "TRL",
    "WAY": "WAY",
    "LOOP": "LOOP",
    "COVE": "CV", "CV": "CV",
    "POINT": "PT", "PT": "PT",
    "CROSSING": "XING", "XING": "XING",
    "PLAZA": "PLZ", "PLZ": "PLZ",
    "EXPRESSWAY": "EXPY", "EXPY": "EXPY",
    "FREEWAY": "FWY", "FWY": "FWY",
    "ROUTE": "RTE", "RTE": "RTE",
    "JUNCTION": "JCT", "JCT": "JCT",
    "MOUNT": "MT", "MT": "MT",
    "MOUNTAIN": "MTN", "MTN": "MTN",
    # Unit / secondary designators
    "APARTMENT": "APT", "APT": "APT",
    "SUITE": "STE", "STE": "STE",
    "BUILDING": "BLDG", "BLDG": "BLDG",
    "FLOOR": "FL", "FL": "FL",
    "UNIT": "UNIT",
}

_UNIT_DESIGNATORS = {"APT", "STE", "UNIT", "BLDG", "FL"}
# Directional abbreviations - used to tell a real directional ("203 W Shore
# Rd") apart from a bare unit letter loosely attached to the house number
# ("203 A West Shore Rd" - 'A' isn't a direction, so it must be a unit).
_DIRECTIONAL_TOKENS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}
_HOUSE_UNIT_RE = re.compile(r"^(\d+)([A-Z])$")


def norm_street(v) -> str:
    """Canonicalizes a street address so common formatting/abbreviation
    differences don't look like different addresses:
    - Directionals (West/W, North/N, ...) and street-type suffixes
      (Lane/Ln, Street/St, Avenue/Ave, ...) are each mapped to one standard
      token, and a trailing '.' on any token is dropped ('St.' -> 'ST').
      So '123 West Lane', '123 W Ln', and '123 W Ln.' all normalize to
      '123 W LN'.
    - A unit letter loosely attached to the house number - either jammed
      against it ('203A ...') or as its own token right after it
      ('203 A ...', where 'A' isn't a directional like 'W' or 'NE') - is
      pulled out and moved to the end as a normal 'UNIT <letter>' suffix.
      So '203A West Shore Rd', '203 A West Shore Rd', and '203 W Shore Rd
      Apt A' all normalize to the same base + unit (see split_street_unit(),
      street_compat())."""
    s = norm_text(v)
    if not s:
        return ""
    tokens = s.split(" ")
    unit_letter = ""
    m = _HOUSE_UNIT_RE.match(tokens[0])
    if m:
        tokens[0] = m.group(1)
        unit_letter = m.group(2)
    elif (len(tokens) > 1 and tokens[0].isdigit()
          and len(tokens[1]) == 1 and tokens[1] not in _DIRECTIONAL_TOKENS):
        unit_letter = tokens.pop(1)

    out = [_STREET_TOKEN_MAP.get(tok.rstrip("."), tok.rstrip(".")) for tok in tokens]

    if unit_letter and not any(t in _UNIT_DESIGNATORS for t in out):
        out += ["UNIT", unit_letter]

    return " ".join(out)


def norm_name(v) -> str:
    """Upper/trim; placeholder values ('[Unknown]', 'N/A', ...) become ''
    so they never out-compete or conflict with a real name."""
    s = norm_text(v)
    core = re.sub(r"[^A-Z0-9]", "", s)
    if not core or core in NAME_PLACEHOLDERS:
        return ""
    return s


def is_unknown_name(v) -> bool:
    """True ONLY for the literal '[Unknown]' placeholder (bracket/paren/
    period/case-insensitive - '[Unknown]', '(unknown)', 'Unknown.' all
    count), NOT a truly blank/empty cell and NOT any other placeholder like
    'UNK', 'N/A', 'NONE' (see NAME_PLACEHOLDERS - norm_name() treats all of
    those as equally blank, but the ID-merge bridge rules need to tell an
    explicit '[Unknown]' entry apart from a cell with nothing typed in it
    at all)."""
    if v is None:
        return False
    core = re.sub(r"[^A-Z0-9]", "", norm_text(v))
    return core == "UNKNOWN"


SSN_MIN_KNOWN_OVERLAP = 4   # min matching KNOWN digits to trust a masked SSN


def norm_ssn(v) -> str:
    """Return a 9-character pattern of digits and 'X' (X = redacted digit),
    or '' if unusable. Mask characters *, #, ? are treated as X, so
        123-45-6789 -> '123456789'
        123-45-XXXX -> '12345XXXX'
        123-45-6XXX -> '123456XXX'
    Rejected (-> ''): not 9 characters, or a masked SSN with fewer than
    SSN_MIN_KNOWN_OVERLAP known digits (too little information to trust).
    A fully-known 9-digit value is always accepted, even a placeholder-
    looking one (e.g. all-same-digit) - this data has no junk SSNs, so no
    such filtering is applied."""
    if v is None:
        return ""
    s = _numeric_cell_to_str(v).upper().replace("*", "X").replace("#", "X").replace("?", "X")
    kept = re.sub(r"[^0-9X]", "", s)
    if len(kept) != 9:
        return ""
    if "X" not in kept:                                    # fully known
        return kept
    known = sum(c != "X" for c in kept)                     # masked
    return kept if known >= SSN_MIN_KNOWN_OVERLAP else ""


# Excel's date epoch: day 1 = 1900-01-01, but Excel treats 1900 as a leap
# year (it wasn't) - using 1899-12-30 as day 0 reproduces that quirk, so a
# serial number converts to the SAME date Excel itself displays.
_EXCEL_SERIAL_EPOCH = pd.Timestamp("1899-12-30")
_EXCEL_SERIAL_RE = re.compile(r"\d{1,6}(\.\d+)?")   # bare serial, e.g. '20037', '20037.0', or '20037.5' (date + time-of-day)


def norm_dob(v) -> str:
    """Parses a DOB cell into 'YYYYMMDD', or '' if unparseable/blank.

    Also handles a raw Excel SERIAL date number (e.g. '20037') showing up
    as the cell's text instead of an actual date: unlike openpyxl (used for
    .xlsx), the pyxlsb engine (used for .xlsb) doesn't auto-convert a date-
    formatted numeric cell into a real date - it hands back the underlying
    serial number as-is, which pandas then stringifies verbatim. A bare
    integer (<=6 digits, so an 8-digit literal 'YYYYMMDD' date is never
    mistaken for one) is treated as such a serial and converted via Excel's
    date epoch; anything else falls through to normal date-text parsing.

    A serial can carry a fractional time-of-day component (e.g. '20037.5')
    even for a nominally date-only field - int() truncates that fraction,
    keeping just the date. Requiring an exact '.0' here previously fell
    through to pd.to_datetime(), which can't parse a bare serial number as
    text and silently returned '' - treating a real, known DOB as blank and
    letting it bridge two otherwise-unrelated people during merging."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "null"):
        return ""
    if _EXCEL_SERIAL_RE.fullmatch(s):
        serial = int(float(s))
        if 1 <= serial <= 60000:   # sane range: years ~1900-2064
            ts = _EXCEL_SERIAL_EPOCH + pd.Timedelta(days=serial)
            return ts.strftime("%Y%m%d")
    ts = pd.to_datetime(s, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y%m%d")


def parse_id_tokens(v) -> frozenset:
    """Splits an Employee ID / Driver's License / Passport / Phone cell into
    individual ID tokens. Cells may already contain multiple semicolon-
    joined IDs from an earlier merge (e.g. '12345; 12346' OR '12345;12346'
    - space after the ';' is optional/inconsistent in the source data), so
    this splits on ';' regardless of spacing and normalizes each token.
    Uses _numeric_cell_to_str() (not a plain str()) so a purely-numeric ID
    that pandas read as a float (e.g. 12345.0) doesn't pick up a spurious
    extra digit and silently fail to match another row's clean version of
    the same ID."""
    if v is None:
        return frozenset()
    return frozenset(norm_text(p) for p in _numeric_cell_to_str(v).split(";") if norm_text(p))


# ------------------------------------------------------------
# 3) Record type - __slots__ for fast attribute access at scale
#    (this loop runs millions of times, so dict-key lookups add up)
# ------------------------------------------------------------
class Rec:
    __slots__ = ("idx", "first", "last", "mid", "suffix", "dob", "ssn",
                 "empids", "dl_ids", "passport_ids", "phones", "emails",
                 "govids", "addr", "city", "state", "zip", "province",
                 "name_bridge_ok")

    def __init__(self, idx):
        self.idx = idx


def build_records(df: pd.DataFrame):
    """df MUST already have a 0..n-1 RangeIndex (see main()) - idx doubles
    as the record's position, so no separate index->position map is needed.
    Uses positional numpy-array access (df.values) rather than itertuples(),
    since itertuples() mangles column names with spaces/punctuation."""
    recs = []
    col_pos = {c: p for p, c in enumerate(df.columns)}
    values = df.values  # numpy object array, fast positional access
    fi, la, mi, sf, do, ss, ei, dl, pp, ph, em, gv, ad, ci, st, zp, pv = (
        col_pos[COL_FIRST], col_pos[COL_LAST], col_pos[COL_MIDDLE], col_pos[COL_SUFFIX],
        col_pos[COL_DOB], col_pos[COL_SSN], col_pos[COL_EMPID], col_pos[COL_DL],
        col_pos[COL_PASSPORT], col_pos[COL_PHONE], col_pos[COL_EMAIL], col_pos[COL_GOVID],
        col_pos[COL_ADDR], col_pos[COL_CITY], col_pos[COL_STATE], col_pos[COL_ZIP],
        col_pos[COL_PROVINCE],
    )
    for i in range(len(df)):
        row = values[i]
        r = Rec(i)
        r.first = norm_name(row[fi])
        r.last = norm_name(row[la])
        # Eligible to act as the "no real name" side of an ID-merge bridge
        # (see _id_name_match_strong()) only if it has a real name, OR both
        # First and Last are the literal "[Unknown]" placeholder - NOT if
        # the cell is truly blank/empty (nothing typed at all).
        r.name_bridge_ok = (
            bool(r.first) and bool(r.last)
        ) or (is_unknown_name(row[fi]) and is_unknown_name(row[la]))
        r.mid = norm_name(row[mi])
        r.suffix = norm_name(row[sf])
        r.dob = norm_dob(row[do])
        r.ssn = norm_ssn(row[ss])
        r.empids = parse_id_tokens(row[ei])
        r.dl_ids = parse_id_tokens(row[dl])
        r.passport_ids = parse_id_tokens(row[pp])
        r.phones = parse_id_tokens(row[ph])
        r.emails = parse_id_tokens(row[em])
        r.govids = parse_id_tokens(row[gv])
        r.addr = norm_street(row[ad])
        r.city = norm_text(row[ci])
        r.state = norm_text(row[st])
        r.zip = norm_text(row[zp])
        r.province = norm_text(row[pv])
        recs.append(r)
    return recs


# ------------------------------------------------------------
# 4) Pairwise matching rules - built STEP BY STEP.
#    Add each newly confirmed rule here as its own small function, then
#    call it from is_match() below.
# ------------------------------------------------------------
def name_prefix_compat(a: str, b: str) -> bool:
    """True if two (already normalized) name values are compatible: blank on
    either side, exactly equal, or one is a PREFIX of the other - an
    incomplete/truncated entry (e.g. 'DID' vs 'DIDAR') is not treated as a
    real difference. A genuine difference (neither a prefix of the other)
    returns False."""
    if not a or not b or a == b:
        return True
    return a.startswith(b) or b.startswith(a)


def name_conflict(r1: Rec, r2: Rec) -> bool:
    """True when First, Middle, or Last Name genuinely disagrees between two
    rows (see name_prefix_compat() - blank on either side, or one being a
    truncated/incomplete version of the other, is NOT a conflict), or when
    Suffix is present on both sides and differs (e.g. 'Jr' vs 'Sr' - no
    prefix tolerance for Suffix, since it isn't an abbreviation-of-the-same-
    value situation). Blocks Step 1 (SSN Exists) - a matching SSN is not
    trusted enough to override a genuine name difference."""
    if not name_prefix_compat(r1.first, r2.first):
        return True
    if not name_prefix_compat(r1.mid, r2.mid):
        return True
    if not name_prefix_compat(r1.last, r2.last):
        return True
    return bool(r1.suffix) and bool(r2.suffix) and r1.suffix != r2.suffix


def name_compat_match(r1: Rec, r2: Rec) -> bool:
    """True if First and Last Name are each real (non-blank) on both sides
    and COMPATIBLE - not required to be byte-exact (see
    name_prefix_compat(): an incomplete/truncated entry, e.g. an initial
    like 'J' for 'Jeffrey', or 'Did' for 'Didar', counts as compatible, not
    a mismatch) - and Middle Name/Suffix don't actively disagree (see
    name_conflict()). This is the shared POSITIVE name-match requirement for
    Rules 2, 4-9 (Rule 1 instead uses name_conflict() alone, as a blocking
    guard on top of an already-strong SSN match, not a positive
    requirement)."""
    if name_conflict(r1, r2):
        return False
    return bool(r1.first) and bool(r2.first) and bool(r1.last) and bool(r2.last)


def ssn_exists_match(r1: Rec, r2: Rec) -> bool:
    """Step 1 - 'SSN Exists': both rows have a usable (non-blank) SSN and
    it's the same value, AND both rows have a real (non-blank) DOB that's
    the SAME value - unlike Rules 4-8/9-10, a blank DOB on either side does
    NOT count as compatible here; DOB must genuinely agree for the SSN
    alone to be trusted. Also blocked by a genuinely differing Name/Suffix
    (name_conflict()). An incomplete/truncated name (e.g. 'Did' vs 'Didar')
    is NOT treated as a differing name here - see name_prefix_compat()."""
    if name_conflict(r1, r2):
        return False
    if not (r1.dob and r2.dob and r1.dob == r2.dob):
        return False
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn == r2.ssn


def ssn_conflict(r1: Rec, r2: Rec) -> bool:
    """True when BOTH rows have a usable, known SSN and it genuinely
    disagrees. Used to block Step 2 (Name+DOB) from merging two different
    real people who happen to share a name and DOB - without this guard,
    Union-Find's transitive clustering can chain unrelated SSNs together
    through a shared-name-and-DOB row (e.g. A<->B via SSN, B<->C via
    Name+DOB even though A and C have different SSNs), contaminating a
    single merged record with several different people's SSNs."""
    return bool(r1.ssn) and bool(r2.ssn) and r1.ssn != r2.ssn


def exact_name_dob_match(r1: Rec, r2: Rec) -> bool:
    """Step 2 - 'Exact Name, DOB': First and Last Name are each real
    (non-blank) on both sides and COMPATIBLE - not required to be byte-exact
    (see name_prefix_compat(): an incomplete/truncated entry, e.g. an
    initial like "J" for "Jeffrey", or "Did" for "Didar", counts as
    compatible, not a mismatch) - AND the same DOB, AND their SSNs don't
    conflict (blank on either/both sides is fine, but two different known
    SSNs block the merge). Middle Name and Suffix are only checked for an
    ACTIVE conflict (see name_conflict()) - they don't need to be present or
    match, just not actively disagree."""
    if ssn_conflict(r1, r2):
        return False
    if not name_compat_match(r1, r2):
        return False
    return bool(r1.dob) and r1.dob == r2.dob


def zip5(v: str) -> str:
    """First 5 digits of a ZIP code, ignoring hyphens/spaces and any ZIP+4
    suffix - so '62701' and '62701-1234' compare as the SAME base ZIP
    instead of a conflict. '' if fewer than 5 digits (not ZIP-shaped)."""
    digits = re.sub(r"[^0-9]", "", v)
    return digits[:5] if len(digits) >= 5 else ""


def split_street_unit(s: str) -> tuple:
    """Splits a normalized street (see norm_street()) into (base, unit):
    'base' is everything before a unit/apartment designator token (APT,
    STE, UNIT, BLDG, FL, or a bare '#123'-style token), 'unit' is just the
    VALUE after that designator - the designator WORD itself is dropped, so
    'APT A' and 'UNIT A' both give unit 'A' (same physical unit, different
    label word - norm_street() already funnels a bare unit letter through
    to a 'UNIT <letter>' suffix, so this keeps that consistent with an
    explicit 'Apt A' in the source data). If no designator token appears,
    base is the whole string and unit is ''."""
    tokens = s.split(" ") if s else []
    for i, tok in enumerate(tokens):
        if tok in _UNIT_DESIGNATORS:
            return " ".join(tokens[:i]), " ".join(tokens[i + 1:])
        if tok.startswith("#") and len(tok) > 1:
            return " ".join(tokens[:i]), " ".join([tok[1:]] + tokens[i + 1:])
    return s, ""


def street_compat(a: str, b: str) -> bool:
    """True if two normalized streets are equal, blank on either side, or
    the SAME base street with a unit/apartment suffix present on only ONE
    side (e.g. '123 ABC LN' vs '123 ABC LN APT 1' - one row is simply
    missing the unit detail, not a different address). If BOTH sides have
    a real unit and it disagrees (e.g. 'APT 1' vs 'APT 2'), that IS treated
    as a genuine conflict - different specific units at the same street,
    not the same address."""
    if not a or not b or a == b:
        return True
    base_a, unit_a = split_street_unit(a)
    base_b, unit_b = split_street_unit(b)
    if not base_a or base_a != base_b:
        return False
    return not (unit_a and unit_b and unit_a != unit_b)


def address_conflict(r1: Rec, r2: Rec) -> bool:
    """True when Residential Address, City, State, Zip, or Province has a
    real, DIFFERING value on both sides (blank on either side is never a
    conflict here). Zip is compared by its 5-digit prefix (zip5()), so a
    plain 5-digit ZIP and its ZIP+4 form are never treated as conflicting.
    Street is compared via street_compat(), so a bare street and the same
    street with an added apartment/unit suffix are never treated as
    conflicting either - only a genuinely different base street, or two
    real but DIFFERENT unit numbers, count as a conflict. Used to block
    Rule 8 from merging two same-named people whose addresses actively
    disagree - matching name alone isn't enough when the address itself
    contradicts it."""
    if not street_compat(r1.addr, r2.addr):
        return True
    fields1 = (r1.city, r1.state, zip5(r1.zip), r1.province)
    fields2 = (r2.city, r2.state, zip5(r2.zip), r2.province)
    return any(a and b and a != b for a, b in zip(fields1, fields2))


def compatible(a: str, b: str) -> bool:
    """True if either side is blank, or both sides are equal. Used for
    Rules 4-6's SSN/DOB check - blank never conflicts, matching is fine,
    but a real, differing value on both sides is a conflict."""
    return not a or not b or a == b


def _id_name_match(ids1: frozenset, ids2: frozenset, r1: Rec, r2: Rec) -> bool:
    """Shared logic for Rules 7, 9 (Phone, Email - WEAK identifiers that can
    legitimately be shared by different people, e.g. a household phone or a
    shared family email): rows share at least one common ID token (from
    whichever ID field is being checked - see parse_id_tokens(), a cell can
    already contain multiple semicolon-joined IDs) AND have a compatible
    First+Last Name (see name_compat_match() - an incomplete/truncated
    entry, e.g. an initial like "J" for "Jeffrey", counts as compatible,
    not a mismatch, but BOTH sides must have a real, non-blank name - a
    shared phone/email alone, with no name confirmation, is not trusted
    here), AND SSN and DOB are each either matching or blank on both sides
    (a real, differing value on either blocks it)."""
    if not (ids1 and ids2 and not ids1.isdisjoint(ids2)):
        return False
    if not name_compat_match(r1, r2):
        return False
    return compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)


def _id_name_match_strong(ids1: frozenset, ids2: frozenset, r1: Rec, r2: Rec) -> bool:
    """Shared logic for Rules 4-6, 10 (Employee ID, Driver's License,
    Passport, Government ID - STRONG/unique identifiers, same tier of trust
    as SSN): rows share at least one common ID token AND Name does not
    actively CONFLICT (see name_conflict() - blank/placeholder on either
    side, e.g. a row with no name on file yet, is NOT a conflict, and an
    incomplete/truncated entry like "Did" vs "Didar" isn't either), AND SSN
    and DOB are each either matching or blank on both sides.

    Unlike _id_name_match() (Phone/Email), a REAL name is not required on
    BOTH sides here - a shared Employee ID/DL/Passport/Gov ID is trusted
    enough on its own to merge a row explicitly labeled "[Unknown]" into
    the same person as another row sharing that ID. This is intentionally
    ONE-SIDED: it lets an "[Unknown]" row bridge to a real name, or to
    another "[Unknown]" row, but two rows that EACH have a real, genuinely
    differing name are NEVER merged here, no matter what else matches
    (SSN/DOB agreeing is not trusted enough to override an actual Name
    difference for this rule - only an explicit "[Unknown]" on at least one
    side is).

    Critically, a row with NOTHING typed in First/Last at all (truly blank,
    not the "[Unknown]" placeholder) does NOT get this bridging treatment -
    see Rec.name_bridge_ok in build_records() / is_unknown_name() - it
    never merges with anyone via this rule, even if it shares the ID."""
    if not (ids1 and ids2 and not ids1.isdisjoint(ids2)):
        return False
    if not (r1.name_bridge_ok and r2.name_bridge_ok):
        return False
    if name_conflict(r1, r2):
        return False
    return compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)


def empid_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 4 - 'Employee ID, Name' (see _id_name_match_strong())."""
    return _id_name_match_strong(r1.empids, r2.empids, r1, r2)


def dl_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 5 - 'Driver's License, Name' (see _id_name_match_strong())."""
    return _id_name_match_strong(r1.dl_ids, r2.dl_ids, r1, r2)


def passport_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 6 - 'Passport Number, Name' (see _id_name_match_strong())."""
    return _id_name_match_strong(r1.passport_ids, r2.passport_ids, r1, r2)


def phone_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 7 - 'Phone Number, Name' (see _id_name_match())."""
    return _id_name_match(r1.phones, r2.phones, r1, r2)


def email_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 9 - 'Email, Name' (see _id_name_match())."""
    return _id_name_match(r1.emails, r2.emails, r1, r2)


def govid_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 10 - 'Government-Issued ID Number, Name' (see _id_name_match_strong())."""
    return _id_name_match_strong(r1.govids, r2.govids, r1, r2)


def address_name_match(r1: Rec, r2: Rec) -> bool:
    """Step 8 - 'Full Address, Name': Residential Address, City, State,
    Zip, and Province are each checked individually - blank on either side
    is fine, but a real, differing value in ANY of them is a conflict that
    blocks the merge. At least one of these fields must have a genuine
    matching value on BOTH sides (two addresses that are entirely blank on
    one side never 'match' just because nothing conflicts).

    Unlike SSN (Rule 1), address is NEVER enough on its own to override a
    name difference - a compatible Name match (see name_compat_match() - an
    incomplete/truncated entry, e.g. an initial like "J" for "Jeffrey",
    counts as compatible, not a mismatch) is always required here too,
    plus SSN/DOB each being matching-or-blank (compatible()), same as
    Rules 4-7, 9-10. Zip is compared by its 5-digit prefix (zip5()), so a plain
    5-digit ZIP and its ZIP+4 form count as the same value, not a conflict
    and not a missed overlap. Street is compared via street_compat(), so
    '123 Abc Lane' vs '123 Abc Lane Apt 1' (unit detail on only one side)
    counts as a genuine overlap too, not a conflict and not a miss."""
    if address_conflict(r1, r2):
        return False
    # address_conflict() already confirmed the street is compatible (equal,
    # blank on one side, or same base street) - so both sides having ANY
    # real street text is itself genuine overlap evidence, even when the
    # exact text differs (e.g. one side adds "APT 1").
    street_overlap = bool(r1.addr) and bool(r2.addr)
    other_fields1 = (r1.city, r1.state, zip5(r1.zip), r1.province)
    other_fields2 = (r2.city, r2.state, zip5(r2.zip), r2.province)
    other_overlap = any(a and b and a == b for a, b in zip(other_fields1, other_fields2))
    if not (street_overlap or other_overlap):
        return False   # nothing real actually overlaps - not a match
    if not name_compat_match(r1, r2):
        return False
    return compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)


def _all_pii_compatible(r1: Rec, r2: Rec) -> bool:
    """True if SSN, DOB, and every other ID field (Employee ID, Driver's
    License, Passport, Government ID, Phone) are each either matching or
    blank on both sides (a real, differing value in any of them blocks it -
    see compatible()), and Residential Address/City/State/Zip/Province
    don't conflict either (see address_conflict() - blank-tolerant per
    field, same rules as Rule 8). This is a pure all-clear check - it does
    NOT require any field to actually overlap, unlike address_name_match()'s
    Rule 8. Used by email_bridge_match() (Rule 11) as the extra evidence
    that lets a shared Email override a blank/Unknown Name on one side: a
    shared identifier plus zero contradictions anywhere else in the row is
    stronger evidence than the shared Email alone."""
    if not (compatible(r1.ssn, r2.ssn) and compatible(r1.dob, r2.dob)):
        return False
    for ids1, ids2 in (
        (r1.empids, r2.empids), (r1.dl_ids, r2.dl_ids),
        (r1.passport_ids, r2.passport_ids), (r1.govids, r2.govids),
        (r1.phones, r2.phones),
    ):
        if ids1 and ids2 and ids1.isdisjoint(ids2):
            return False
    return not address_conflict(r1, r2)


def email_bridge_match(r1: Rec, r2: Rec) -> bool:
    """Step 11 - 'Email, All PII Compatible' (last-resort bridge): rows
    share at least one Email Address token, Name does not actively CONFLICT
    (see name_conflict() - blank/placeholder on either side, e.g. one row's
    Name is "[Unknown]" or truly blank and the other has a real name, is
    NOT a conflict; only a genuine difference between two REAL names blocks
    it), AND every other piece of PII - SSN, DOB, Employee ID, DL, Passport,
    Gov ID, Phone, Address - is either matching or blank on both sides (see
    _all_pii_compatible()).

    Unlike email_name_match() (Rule 9), a REAL name is NOT required on both
    sides here - this is the one case where a shared Email is trusted to
    bridge an Unknown/blank-named row into a real person's record, but ONLY
    once every other field has also been checked and found to agree or
    simply be absent. Email alone is still a WEAK identifier (e.g. a shared
    household email), so this rule is intentionally NOT added to
    try_union()'s strong_match list in main() - it does not get to override
    a conflicting address elsewhere in the transitive group, same as
    Rule 8."""
    if not (r1.emails and r2.emails and not r1.emails.isdisjoint(r2.emails)):
        return False
    if name_conflict(r1, r2):
        return False
    return _all_pii_compatible(r1, r2)


def _weak_overlap(r1: Rec, r2: Rec) -> bool:
    """True if at least one identifying field genuinely OVERLAPS between
    the two rows: a shared Employee ID/DL/Passport/Gov ID/Phone/Email
    token, or a genuinely matching Address field (street/city/state/zip/
    province - same overlap test as address_name_match()'s Rule 8).
    Deliberately excludes SSN and DOB matching alone - Rule 1 already
    treats a shared SSN as NOT trusted enough on its own without a
    genuinely agreeing DOB, and this helper must not reopen that door for
    an Unknown/partial-named bridge row with a blank DOB. Used by
    unknown_name_bridge_match() (Rule 12) as the "there IS real evidence,
    not just an absence of contradiction" requirement - same shape as the
    overlap already required by Rule 8, generalized across every other PII
    field."""
    for ids1, ids2 in (
        (r1.empids, r2.empids), (r1.dl_ids, r2.dl_ids),
        (r1.passport_ids, r2.passport_ids), (r1.govids, r2.govids),
        (r1.phones, r2.phones), (r1.emails, r2.emails),
    ):
        if ids1 and ids2 and not ids1.isdisjoint(ids2):
            return True
    if bool(r1.addr) and bool(r2.addr) and street_compat(r1.addr, r2.addr):
        return True
    other_fields1 = (r1.city, r1.state, zip5(r1.zip), r1.province)
    other_fields2 = (r2.city, r2.state, zip5(r2.zip), r2.province)
    return any(a and b and a == b for a, b in zip(other_fields1, other_fields2))


def unknown_name_bridge_match(r1: Rec, r2: Rec) -> bool:
    """Step 12 - 'Unknown/Partial Name, PII Overlap': Name does not
    actively CONFLICT (see name_conflict() - a blank/Unknown First OR Last
    on either side is NOT a conflict, INCLUDING a partial case where only
    one of First/Last is Unknown and the other is a real name, e.g.
    (Unknown, Smith) vs (Unknown, Unknown), or (Unknown, Smith) vs (John,
    Unknown) - only a genuine difference between two REAL values in the
    SAME position blocks it), AND at least one OTHER identifying field
    genuinely overlaps (see _weak_overlap()), AND every other PII field -
    SSN, DOB, and all of the above - is matching-or-blank (see
    _all_pii_compatible()).

    This exists because the STRONG-ID rules (4-6, 10) only bridge a Name
    that's fully blank or fully "[Unknown]" (see Rec.name_bridge_ok in
    build_records() - a row with ONE real Name part and ONE Unknown part
    does not qualify there), and the WEAK rules (7-9) require a REAL name
    on BOTH sides. This rule closes that specific gap - a partially-Unknown
    Name, as long as there's genuine overlap elsewhere and nothing else
    contradicts - without loosening those other rules' own, more targeted
    guarantees."""
    if name_conflict(r1, r2):
        return False
    if not _weak_overlap(r1, r2):
        return False
    return _all_pii_compatible(r1, r2)


def is_match(r1: Rec, r2: Rec) -> bool:
    return (
        ssn_exists_match(r1, r2)
        or exact_name_dob_match(r1, r2)
        or empid_name_match(r1, r2)
        or dl_name_match(r1, r2)
        or passport_name_match(r1, r2)
        or phone_name_match(r1, r2)
        or address_name_match(r1, r2)
        or email_name_match(r1, r2)
        or govid_name_match(r1, r2)
        or email_bridge_match(r1, r2)
        or unknown_name_bridge_match(r1, r2)
    )


# ------------------------------------------------------------
# 4b) Multiprocessing workers for the pairwise clustering step. Must be
#     module-level functions (not closures) so they can be pickled and sent
#     to worker processes. _worker_recs is set once per worker via the Pool
#     initializer, instead of re-sending the full `recs` list with every
#     chunk.
# ------------------------------------------------------------
_worker_recs = None


def _init_worker(recs):
    global _worker_recs
    _worker_recs = recs


def _match_chunk(chunk):
    """Runs in a worker process: tests every pair in this chunk, returns only
    the ones that matched."""
    return [(a, b) for a, b in chunk if is_match(_worker_recs[a], _worker_recs[b])]


def _chunk_pairs(pairs_list, target_chunks=200, min_chunk_size=200):
    """Splits pairs_list into roughly target_chunks pieces (never smaller
    than min_chunk_size), so progress can update ~target_chunks times
    regardless of how many pairs there are in total."""
    n = len(pairs_list)
    if n == 0:
        return []
    chunk_size = max(min_chunk_size, -(-n // target_chunks))   # ceil division
    return [pairs_list[i:i + chunk_size] for i in range(0, n, chunk_size)]


# ------------------------------------------------------------
# 5) Union-Find (disjoint set) for transitive clustering
# ------------------------------------------------------------
class UnionFind:
    def __init__(self, n):
        self.parent = list(range(n))

    def find(self, x):
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[max(ra, rb)] = min(ra, rb)


# ------------------------------------------------------------
# 6) Blocking - only test pairs that already share something exact, so we
#    never do a full N x N comparison. Candidate pairs come from one bucket
#    per active rule; is_match() then applies the real rules within each.
# ------------------------------------------------------------


def bucket_candidate_pairs(recs):
    buckets = defaultdict(list)
    for r in recs:
        if r.ssn:                                    # Rule 1: SSN Exists
            buckets[("ssn", r.ssn)].append(r.idx)
        if r.first and r.last and r.dob:              # Rule 2: Exact Name, DOB
            buckets[("namedob", r.first, r.last, r.dob)].append(r.idx)
        # Rules 4-6, 10 (Employee ID/DL/Passport/Gov ID - STRONG identifiers)
        # bucket by the ID token ALONE, not name - a row with a blank/
        # placeholder name must still land in the same bucket as a
        # real-name row sharing that ID, so is_match() (via
        # _id_name_match_strong()) actually gets a chance to test the pair.
        if r.empids:                                   # Rule 4: Employee ID, Name
            for tok in r.empids:
                buckets[("empidname", tok)].append(r.idx)
        if r.dl_ids:                                    # Rule 5: Driver's License, Name
            for tok in r.dl_ids:
                buckets[("dlname", tok)].append(r.idx)
        if r.passport_ids:                              # Rule 6: Passport Number, Name
            for tok in r.passport_ids:
                buckets[("ppname", tok)].append(r.idx)
        if r.govids:                                    # Rule 10: Government ID, Name
            for tok in r.govids:
                buckets[("govidname", tok)].append(r.idx)
        # Rules 7, 9 (Phone/Email - WEAK identifiers) still require a real
        # name on both sides to match (see _id_name_match()), so bucketing
        # by name too keeps these buckets from ballooning on a widely-shared
        # phone/email with no name confirmation possible anyway.
        if r.phones and r.first and r.last:            # Rule 7: Phone Number, Name
            for tok in r.phones:
                buckets[("phonename", tok, r.first, r.last)].append(r.idx)
        # Rule 12 (Unknown/Partial Name, PII Overlap) can match on a shared
        # Phone even when Name is only partially known - same reasoning as
        # the "emailbridge" bucket above, so this buckets by the phone
        # token ALONE too.
        if r.phones:                                     # Rule 12: Unknown/Partial Name, PII Overlap
            for tok in r.phones:
                buckets[("phonebridge", tok)].append(r.idx)
        if r.emails and r.first and r.last:             # Rule 9: Email, Name
            for tok in r.emails:
                buckets[("emailname", tok, r.first, r.last)].append(r.idx)
        # Rule 11 (Email, All PII Compatible) can match even when only ONE
        # side has a real name, so - unlike the "emailname" bucket above -
        # this buckets by the email token ALONE, same as the strong-ID
        # buckets (empidname/dlname/ppname/govidname), so an Unknown/blank-
        # named row still lands in the same bucket as a real-named row
        # sharing that email and actually gets tested by is_match().
        if r.emails:                                     # Rule 11: Email, All PII Compatible
            for tok in r.emails:
                buckets[("emailbridge", tok)].append(r.idx)
        # Rule 8: Full Address, Name - bucket by each address FIELD
        # individually (not the whole address as one key), since
        # address_name_match() only needs ONE field to genuinely overlap.
        # is_match() then does the real per-field compatibility check.
        # State/Province deliberately excluded here (though still checked
        # for conflicts by address_name_match() itself) - they're too
        # low-cardinality for large files (e.g. "CA" alone can be shared by
        # tens of thousands of unrelated rows), which would make that one
        # bucket's pairwise comparison extremely slow. A real matching
        # address almost always also shares City or Zip, so this loses
        # very little real coverage.
        if r.addr:
            buckets[("addrfield", "addr", r.addr)].append(r.idx)
        if r.city:
            buckets[("addrfield", "city", r.city)].append(r.idx)
        if r.zip:
            buckets[("addrfield", "zip", r.zip)].append(r.idx)

    pairs = set()
    total = len(buckets)
    for n, (key, idxs) in enumerate(buckets.items(), 1):
        progress("Bucketing", n, total)
        if len(idxs) < 2:
            continue
        for a, b in itertools.combinations(sorted(idxs), 2):
            pairs.add((a, b))
    return pairs


# ------------------------------------------------------------
# 7) Merge helpers for building the output
# ------------------------------------------------------------
def semicolon_merge(values) -> str:
    """Distinct, non-blank values joined with '; ', first-seen order,
    original casing preserved; dedup key is upper/trimmed. Splits every
    cell on ';' FIRST and dedupes at that token level (not just whole-cell
    strings) - source cells can already contain multiple semicolon-joined
    sub-values themselves (e.g. '12345; 12346', or '12345;12346' with
    inconsistent spacing, or a literal repeat like '123;456;123'), and
    whole-cell-only dedup would leave those repeats sitting in the output
    untouched. Applies uniformly to every semicolon-merged column, not just
    the ID fields, since this repeated-token pattern can turn up in any
    column depending on how the source data was entered."""
    seen = set()
    out = []
    for v in values:
        if v is None:
            continue
        for tok in str(v).split(";"):
            raw = tok.strip()
            key = norm_text(raw)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(raw)
    return MERGE_SEP.join(out)


def dob_merge(values) -> str:
    """Like semicolon_merge(), but dedupes by the NORMALIZED date
    (norm_dob()) instead of raw text, and always DISPLAYS that normalized
    date as a clean 'MM/DD/YYYY' string - never the row's original raw text.
    This matters for two reasons: (1) the same date typed in different
    formats across rows (e.g. '01/01/1990', '1990-01-01', '1/1/90') collapses
    into ONE consistently-formatted entry instead of showing as multiple
    different-looking DOBs, and (2) a raw Excel SERIAL date number (e.g.
    '20037' - see norm_dob()) never leaks into the output as-is. Only a
    genuinely different real date produces a second entry - and the group-
    level safety net in main() (group_dob/try_union()) already refuses to
    merge two rows with different real DOBs in the first place, so this is
    purely a display fix, not a new matching rule."""
    seen = set()
    out = []
    for v in values:
        if v is None:
            continue
        for tok in str(v).split(";"):
            raw = tok.strip()
            key = norm_dob(raw)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(f"{key[4:6]}/{key[6:8]}/{key[0:4]}")
    return MERGE_SEP.join(out)


DOCID_CHUNK_SIZE = 20_000   # keep well under Excel's 32,767-char cell limit
OTHER_ADDR_CHUNK_SIZE = 25_000   # keep well under Excel's 32,767-char cell limit


def split_docid_chunks(docid_str, max_chars=DOCID_CHUNK_SIZE):
    """Splits an already-merged DOCID string ('DOC001; DOC002; ...') into
    chunks no longer than max_chars, breaking ONLY at '; ' boundaries (never
    mid-DOCID). A group merging tens of thousands of rows can produce a
    DOCID string longer than Excel's 32,767-character cell limit, which
    Excel silently truncates - splitting across extra 'DOCIDs 2', 'DOCIDs
    3', ... columns instead keeps every DOCID intact and visible."""
    if len(docid_str) <= max_chars:
        return [docid_str]
    parts = docid_str.split(MERGE_SEP)
    chunks = []
    current = []
    current_len = 0
    for part in parts:
        added_len = len(part) + (len(MERGE_SEP) if current else 0)
        if current and current_len + added_len > max_chars:
            chunks.append(MERGE_SEP.join(current))
            current = [part]
            current_len = len(part)
        else:
            current.append(part)
            current_len += added_len
    if current:
        chunks.append(MERGE_SEP.join(current))
    return chunks


def fullest_value(raw_values, norm_values) -> str:
    """Longest raw value whose norm form is non-blank (placeholders/blanks
    are skipped); '' if every value is blank/placeholder. Used for First/
    Middle/Last Name, Suffix, and SSN - these keep ONE final value per
    merged person, not every variant."""
    best, best_len = "", -1
    for raw, norm in zip(raw_values, norm_values):
        if not norm:
            continue
        raw = "" if raw is None else str(raw).strip()
        if len(raw) > best_len:
            best, best_len = raw, len(raw)
    return best


def has_variation(raw_values) -> bool:
    """True if the group had 2+ distinct real values for this field, even
    though only one (the fullest) was kept in the output - used for the
    'Names Differ' review flag."""
    return len({norm_text(v) for v in raw_values if norm_text(v)}) > 1


def zip_key(v) -> str:
    """Normalized comparison key for a ZIP/postal code: the 5-digit prefix
    for a US-style ZIP (so '12345' and '12345-1234' compare equal), or the
    plain normalized text for anything else (e.g. non-US postal codes that
    don't have 5+ digits to extract a prefix from)."""
    z5 = zip5(norm_text(v))
    return z5 if z5 else norm_text(v)


def address_key(values) -> tuple:
    """Normalized tuple used to tell whether two rows have the SAME address
    (all fields blank-insensitive) - used to find the majority address.
    ADDRESS_COLS order is Residential Address, City, State, Province, Zip,
    Country - Zip is compared via zip_key() (5-digit prefix) and the street
    via norm_street() (abbreviation-canonicalized), so a plain 5-digit ZIP
    vs its ZIP+4 form, and '123 West Lane' vs '123 W Ln', each count as the
    SAME address here - matching how they're already treated during merging
    (see zip5() / norm_street())."""
    def _key(col, v):
        if col == COL_ZIP:
            return zip_key(v)
        if col == COL_ADDR:
            return norm_street(v)
        return norm_text(v)
    return tuple(_key(col, v) for col, v in zip(ADDRESS_COLS, values))


def address_key_conflict(k1: tuple, k2: tuple) -> bool:
    """True if two normalized address keys (see address_key(), each a
    Street/City/State/Province/Zip/Country tuple) genuinely disagree - blank
    on either side is never a conflict, but a real, differing value is.
    Street is compared via street_compat() (a unit/apt suffix present on
    only one side isn't a conflict, matching how Rule 8 treats it); Zip is
    already the zip5-prefix-or-raw key from address_key(), so a plain
    compare is enough; City/State/Province/Country are a plain blank-
    tolerant equality check. Used by split_addresses() so a row that's
    missing just one field (e.g. blank Zip) of an otherwise-identical
    address is folded into that SAME address instead of being listed as a
    separate 'Other Address' - mirrors how address_conflict() already treats
    blanks when deciding whether two rows are the same person in the first
    place, so the output doesn't contradict the matching decision."""
    addr1, city1, state1, prov1, zip1, country1 = k1
    addr2, city2, state2, prov2, zip2, country2 = k2
    if not street_compat(addr1, addr2):
        return True
    pairs = ((city1, city2), (state1, state2), (prov1, prov2), (zip1, zip2), (country1, country2))
    return any(a and b and a != b for a, b in pairs)


def format_full_address(values) -> str:
    """One address, all its parts combined into a single readable string,
    e.g. '123 ABC Ln, Springfield, IL, 62701, USA'. Blank parts are skipped."""
    parts = []
    for v in values:
        s = "" if v is None else str(v).strip()
        if s and s.lower() not in ("nan", "none", "null"):
            parts.append(s)
    return ", ".join(parts)


def split_addresses(df, group_idxs):
    """Returns (majority_values, other_address_string) for one merged group.

    Rows are first bucketed by their exact normalized address_key() (blank-
    insensitive per field, abbreviation-canonicalized), then those DISTINCT
    keys are clustered together whenever they don't genuinely conflict (see
    address_key_conflict()) - e.g. the same street/city with Zip blank on
    one row and present on another are one cluster, one address, not two.
    majority_values: the fullest non-blank value per field (see
    fullest_value()) across the winning cluster - the one with the most
    rows in it - so a row missing a field borrows it from another row's
    fuller copy of the same address, instead of leaving it blank or
    spinning off a spurious 'Other Address' entry. These go into the normal
    Residential Address/City/State/Zip/Country columns.
    other_address_string: every OTHER, genuinely different address cluster,
    combined into one string per address and semicolon-joined, for the
    'Other Address' column. A row with no address at all doesn't count as
    a "real" address unless it's the only kind of address in the group."""
    key_order = []       # distinct keys, first-seen order
    key_count = {}       # key -> row count
    key_raw = {}         # key -> first-seen raw ADDRESS_COLS values
    for idx in group_idxs:
        raw = tuple(df.at[idx, c] for c in ADDRESS_COLS)
        key = address_key(raw)
        if key not in key_count:
            key_count[key] = 0
            key_raw[key] = raw
            key_order.append(key)
        key_count[key] += 1

    # Cluster the DISTINCT keys (typically a handful per person, even in a
    # large merged group) that don't conflict with each other - plain
    # pairwise since there are far fewer distinct addresses than rows.
    parent = list(range(len(key_order)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    for i, j in itertools.combinations(range(len(key_order)), 2):
        if not address_key_conflict(key_order[i], key_order[j]):
            union(i, j)

    clusters = defaultdict(list)
    for i in range(len(key_order)):
        clusters[find(i)].append(i)

    def cluster_weight(positions):
        return sum(key_count[key_order[i]] for i in positions)

    def cluster_values(positions):
        """Fullest non-blank raw value per field across every distinct key
        in this cluster - fills a gap like a blank Zip from another row's
        fuller address rather than leaving it blank."""
        out = []
        for col_i in range(len(ADDRESS_COLS)):
            raws = [key_raw[key_order[i]][col_i] for i in positions]
            norms = [key_order[i][col_i] for i in positions]
            out.append(fullest_value(raws, norms))
        return tuple(out)

    non_blank_clusters = [c for c in clusters.values()
                          if any(any(key_order[i]) for i in c)]
    candidates = non_blank_clusters or list(clusters.values())
    majority_cluster = max(candidates, key=lambda c: (cluster_weight(c), -min(c)))

    other_strings = []
    for c in clusters.values():
        if c is majority_cluster or not any(any(key_order[i]) for i in c):
            continue
        other_strings.append(format_full_address(cluster_values(c)))

    return cluster_values(majority_cluster), MERGE_SEP.join(other_strings)


def pick_input_file() -> str:
    """Opens a native file-picker dialog for the input file, returns the
    chosen path ('' if the user cancels). Accepts any format
    load_input_file() knows how to read - CSV or Excel (.xlsx/.xls/.xlsb)."""
    root = Tk()
    root.withdraw()
    path = filedialog.askopenfilename(
        title="Select input file",
        filetypes=[
            ("CSV and Excel files", "*.csv *.xlsx *.xls *.xlsb"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx *.xls *.xlsb"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    return path


def load_input_file(path: str) -> pd.DataFrame:
    """Reads path into a DataFrame with every column as plain text
    (dtype=str, same as the rest of this script requires), picking the
    reader/engine from its file extension: .csv via pd.read_csv(), .xlsb via
    pd.read_excel(engine='pyxlsb') (pandas can't infer that engine the way
    it does for .xlsx/.xls - see norm_dob()'s docstring for the same
    underlying quirk), anything else via plain pd.read_excel()."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str)
    engine = "pyxlsb" if ext == ".xlsb" else None
    return pd.read_excel(path, sheet_name=INPUT_SHEET, dtype=str, engine=engine)


def resolve_column_aliases(df: pd.DataFrame, alias_groups) -> pd.DataFrame:
    """For each group in alias_groups (see COLUMN_ALIASES), renames whichever
    spelling is actually present in df's columns to that group's canonical
    (first) name - so an input file using an alternate spelling of the same
    logical column doesn't fail the EXPECTED_COLS check below. If the
    canonical name is already present, or none of the spellings are present
    at all, df is returned unchanged (a genuinely missing column still
    surfaces via the normal EXPECTED_COLS check, under its canonical name)."""
    rename = {}
    for group in alias_groups:
        canonical = group[0]
        if canonical in df.columns:
            continue
        for alias in group[1:]:
            if alias in df.columns:
                rename[alias] = canonical
                break
    return df.rename(columns=rename) if rename else df


def resolve_extra_merge_cols(df: pd.DataFrame, other_merge_cols) -> list:
    """Returns other_merge_cols plus any column ACTUALLY present in df that
    isn't a core matching field (Name/DOB/SSN/Address), a DOCID/Other
    Address overflow column, or already listed in other_merge_cols.

    This is the safety net for a real header the configured OTHER_MERGE_COLS
    list doesn't know about AT ALL - as opposed to a near-miss spelling of
    one it DOES know about (e.g. the list says "PI Notes" but the file's
    actual header is "PII Notes" - that case isn't caught here; it just
    silently fails to match either name and needs the list itself fixed).
    Without this, a column the list has never heard of would be silently
    dropped from the output entirely, instead of being merged in with a
    semicolon like every other passthrough column. Detected extras are
    appended in the ORIGINAL order they appear in df, so the "Column order"
    logic in main() still places them at their real position in the input
    header, not at the end."""
    docid_extra_re = re.compile(rf"^{re.escape(COL_DOCID)} \d+$")
    other_addr_extra_re = re.compile(rf"^{re.escape(COL_OTHER_ADDR)} \d+$")
    core_cols = ({COL_DOCID, COL_FIRST, COL_LAST, COL_MIDDLE, COL_SUFFIX,
                  COL_DOB, COL_SSN, COL_OTHER_ADDR} | set(ADDRESS_COLS))
    known = core_cols | set(other_merge_cols)
    extras = [c for c in df.columns
              if c not in known
              and not docid_extra_re.match(c)
              and not other_addr_extra_re.match(c)]
    if extras:
        print(f"  NOTE: input has {len(extras)} column(s) not listed in "
              f"OTHER_MERGE_COLS - merging them in too, unchanged: {extras}")
    return other_merge_cols + extras


# ------------------------------------------------------------
# 8) Main
# ------------------------------------------------------------
def main() -> None:
    # No CLI argument -> a native file-picker dialog (CSV or Excel, any of
    # .csv/.xlsx/.xls/.xlsb - see pick_input_file()/load_input_file()).
    # A CLI argument skips the dialog and reads that file directly.
    input_path = sys.argv[1] if len(sys.argv) > 1 else pick_input_file()
    if not input_path:
        print("No input file selected - exiting.")
        return
    print(f"Reading {input_path} ...")
    df = load_input_file(input_path)
    df.columns = [str(c).strip() for c in df.columns]
    df = resolve_column_aliases(df, COLUMN_ALIASES)
    # other_merge_cols = OTHER_MERGE_COLS plus any column the input has that
    # the configured list doesn't know about at all (see
    # resolve_extra_merge_cols()) - both are purely append-only pass-through
    # data, never used for matching (see is_match()), so a CONFIGURED name
    # missing from the input entirely (a newer/aspirational field the real
    # export doesn't have yet) shouldn't hard-fail the whole run; it just
    # comes through blank in the output instead of being real evidence one
    # way or another.
    other_merge_cols = resolve_extra_merge_cols(df, OTHER_MERGE_COLS)
    for c in other_merge_cols:
        if c not in df.columns:
            df[c] = ""
    df = df.reset_index(drop=True)   # guarantees row position == record idx

    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise SystemExit(
            f"These expected columns were not found in {input_path}:\n"
            f"  {missing}\nColumns present:\n  {list(df.columns)}\n"
            "Fix the COL_*/OTHER_MERGE_COLS names in the CONFIG block."
        )
    print(f"  {len(df):,} rows read.")

    # "DOCIDs" is always required (see EXPECTED_COLS above), but the input
    # may ALSO already have "DOCIDs 2"/"DOCIDs 3"/... - e.g. it's the output
    # of a prior merge that itself overflowed into those columns. Whichever
    # of those are present get read and merged in too (see SEMICOLON_COLS
    # handling below), so a value that only lives in one row's "DOCIDs 3"
    # isn't silently dropped.
    _docid_extra_re = re.compile(rf"^{re.escape(COL_DOCID)} (\d+)$")
    docid_input_cols = [COL_DOCID] + sorted(
        (c for c in df.columns if _docid_extra_re.match(c)),
        key=lambda c: int(_docid_extra_re.match(c).group(1)),
    )
    if len(docid_input_cols) > 1:
        print(f"  Input already has overflow DOCID columns "
              f"({', '.join(docid_input_cols[1:])}) - merging them in too.")

    # "Other Address" is optional (not in EXPECTED_COLS - a fresh source file
    # won't have it), but if the input already has it - plus any
    # "Other Address 2"/"Other Address 3"/... from a prior merge that itself
    # overflowed - all of those are read and merged in too.
    _other_addr_extra_re = re.compile(rf"^{re.escape(COL_OTHER_ADDR)} (\d+)$")
    other_addr_input_cols = ([COL_OTHER_ADDR] if COL_OTHER_ADDR in df.columns else []) + sorted(
        (c for c in df.columns if _other_addr_extra_re.match(c)),
        key=lambda c: int(_other_addr_extra_re.match(c).group(1)),
    )
    if len(other_addr_input_cols) > 1:
        print(f"  Input already has overflow Other Address columns "
              f"({', '.join(other_addr_input_cols[1:])}) - merging them in too.")

    recs = build_records(df)   # recs[i].idx == i == df row position

    print("Clustering (blocked comparison) ...")
    pairs = bucket_candidate_pairs(recs)
    print(f"  {len(pairs):,} candidate pairs to test.")

    uf = UnionFind(len(recs))
    # group_ssn[root] = every DISTINCT known SSN currently inside that root's
    # merged group (not just the two rows being compared). Guards against
    # transitive contamination: a Step 2/3 match can bridge through a row
    # with a BLANK SSN (blank never conflicts pairwise), which would
    # otherwise let two groups with different real SSNs merge into one via
    # that bridge row even though neither pairwise SSN conflicts directly.
    # Refusing the union as soon as it would combine 2 different known SSNs
    # - rather than merging everything first and un-merging afterward -
    # means the SAFE parts of a would-be conflicting cluster still merge
    # normally: e.g. 5 rows sharing SSN=111 and 3 rows sharing SSN=999,
    # bridged by one blank-SSN row, end up as 2 clean groups (not 8 separate
    # rows) - the bridge row itself lands in whichever group it's processed
    # against first (order can vary; which side "wins" isn't predictable,
    # but two different real SSNs never end up in the same group either way).
    group_ssn = [({r.ssn} if r.ssn else set()) for r in recs]
    # group_dob[root] = every DISTINCT known DOB currently inside that
    # root's group. Same bridging risk as group_ssn: a blank-DOB row can
    # pairwise-match two OTHER rows (via Rules 4-8/9-10, which still treat a
    # blank DOB as compatible - unlike Rule 1, which now requires DOB to
    # genuinely agree) that have different real DOBs from each other,
    # transitively merging two genuinely different DOBs into one group.
    group_dob = [({r.dob} if r.dob else set()) for r in recs]
    # group_addr[root] = 5 sets (one per address field: Street, City, State,
    # Zip, Province), each holding every DISTINCT known value currently in
    # that root's group. Same bridging risk again: Rule 8 needs a genuine
    # value match on at least one field, but a row can overlap with two
    # OTHERS via a DIFFERENT field (e.g. matching City, blank Street) even
    # though those other two rows have a genuinely conflicting Street with
    # each other - transitively combining them. Checked in try_union() ONLY
    # for weak-evidence Rule 8 matches - a STRONG match (SSN, Name+DOB, or
    # any ID+Name rule) is trusted enough to override a conflicting address
    # on its own, so it skips this check (see the strong_match guard in
    # try_union()).
    group_addr = [
        (({split_street_unit(r.addr)[0]} if r.addr else set()),   # base street, so a unit/apt suffix doesn't falsely conflict
         ({r.city} if r.city else set()),
         ({r.state} if r.state else set()),
         ({zip5(r.zip)} if zip5(r.zip) else set()),   # 5-digit prefix, so ZIP/ZIP+4 don't falsely conflict
         ({r.province} if r.province else set()))
        for r in recs
    ]
    # group_first/group_last[root] = every DISTINCT known First/Last Name
    # currently inside that root's group. Same bridging risk as group_ssn/
    # group_dob/group_addr: since Rules 4-6/10 now allow a blank/placeholder
    # name to match a shared strong ID (see _id_name_match_strong()), a
    # blank-name row can pairwise-match two OTHER rows that have genuinely
    # DIFFERENT real names from each other - e.g. "Alice Johnson" and "Mike
    # Green" both sharing one Employee ID with an "Unknown" row in between -
    # transitively merging two different real people into one group even
    # though the direct Alice-vs-Mike pairwise test itself correctly refuses
    # (name_conflict()). Checked for EVERY union (not just strong-ID
    # matches), since this bridging risk applies to any rule that tolerates
    # a blank name on either side.
    group_first = [({r.first} if r.first else set()) for r in recs]
    group_last = [({r.last} if r.last else set()) for r in recs]
    refused_ssn = 0
    refused_dob = 0
    refused_addr = 0
    refused_name = 0

    def try_union(a_idx, b_idx):
        nonlocal refused_ssn, refused_dob, refused_addr, refused_name
        ra, rb = uf.find(a_idx), uf.find(b_idx)
        if ra == rb:
            return
        sa, sb = group_ssn[ra], group_ssn[rb]
        if sa and sb and sa.isdisjoint(sb):
            refused_ssn += 1
            return   # would combine two different real SSNs - refused
        da, db = group_dob[ra], group_dob[rb]
        if da and db and da.isdisjoint(db):
            refused_dob += 1
            return   # would combine two different real DOBs - refused
        r1, r2 = recs[a_idx], recs[b_idx]
        # A mismatch on EITHER First OR Last (both known, and disjoint) is
        # enough to refuse - a shared Last Name alone (e.g. two different
        # "Holland"s) is NOT enough to treat two genuinely different First
        # names as the same person. Only an Unknown/blank-name row is meant
        # to bridge here (into a real name, or into another Unknown row) -
        # two rows/groups that EACH have a real, differing name must never
        # merge through such a bridge, no exceptions (not even a matching
        # SSN/DOB - see _id_name_match_strong()).
        fa1, fa2 = group_first[ra], group_first[rb]
        la1, la2 = group_last[ra], group_last[rb]
        first_conflict = bool(fa1) and bool(fa2) and fa1.isdisjoint(fa2)
        last_conflict = bool(la1) and bool(la2) and la1.isdisjoint(la2)
        if first_conflict or last_conflict:
            refused_name += 1
            return   # would combine two rows/groups with genuinely different real names - refused
        # A STRONG match (SSN, Name+DOB, or any ID+Name rule) is trusted
        # enough to override a conflicting address on its own - e.g. the
        # same SSN turning up with two different addresses on file is far
        # more likely to mean "this person moved" than "two different
        # people happen to share this SSN". Only the WEAK-evidence Rule 8
        # itself (which IS the address rule, so it can't be used to
        # override its own conflict check) still gets blocked by a
        # conflicting address.
        strong_match = (
            ssn_exists_match(r1, r2) or exact_name_dob_match(r1, r2)
            or empid_name_match(r1, r2) or dl_name_match(r1, r2)
            or passport_name_match(r1, r2) or phone_name_match(r1, r2)
            or email_name_match(r1, r2) or govid_name_match(r1, r2)
        )
        aa, ab = group_addr[ra], group_addr[rb]
        if not strong_match:
            if any(fa and fb and fa.isdisjoint(fb) for fa, fb in zip(aa, ab)):
                refused_addr += 1
                return   # would combine two conflicting real addresses - refused
        uf.union(a_idx, b_idx)
        merged_root = min(ra, rb)
        group_ssn[merged_root] = sa | sb
        group_dob[merged_root] = da | db
        group_addr[merged_root] = tuple(fa | fb for fa, fb in zip(aa, ab))
        group_first[merged_root] = fa1 | fa2
        group_last[merged_root] = la1 | la2

    pairs_list = list(pairs)
    total_pairs = len(pairs_list)

    if total_pairs < PARALLEL_THRESHOLD:
        # Small run - not worth the process-startup overhead, just test
        # in-process.
        for tested, (a_idx, b_idx) in enumerate(pairs_list, 1):
            if is_match(recs[a_idx], recs[b_idx]):
                try_union(a_idx, b_idx)
            progress("Clustering", tested, total_pairs)
    else:
        # Large run - spread the pairwise is_match() tests across
        # PARALLEL_WORKERS separate processes (real parallelism; a thread
        # pool would not help here, see PARALLEL_WORKERS comment above).
        # The actual union calls stay single-process afterward, since
        # Union-Find (and group_ssn) is shared, mutable state that can't be
        # split safely.
        chunks = _chunk_pairs(pairs_list)
        done = 0
        with Pool(processes=PARALLEL_WORKERS, initializer=_init_worker, initargs=(recs,)) as pool:
            for matched in pool.imap_unordered(_match_chunk, chunks):
                for a_idx, b_idx in matched:
                    try_union(a_idx, b_idx)
                done += 1
                progress("Clustering", done, len(chunks))

    groups = defaultdict(list)
    for r in recs:
        groups[uf.find(r.idx)].append(r.idx)
    groups = list(groups.values())

    print(f"  {len(df):,} rows -> {len(groups):,} merged people "
          f"({len(df) - len(groups):,} rows collapsed by a match).")
    if refused_ssn:
        print(f"  {refused_ssn:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real SSNs into one group.")
    if refused_dob:
        print(f"  {refused_dob:,} candidate merge(s) were refused - would have "
              f"combined 2+ different real DOBs into one group.")
    if refused_addr:
        print(f"  {refused_addr:,} candidate merge(s) were refused - would have "
              f"combined 2+ conflicting real addresses into one group.")
    if refused_name:
        print(f"  {refused_name:,} candidate merge(s) were refused - would have "
              f"combined 2+ genuinely different real Names into one group.")

    print("Building merged output ...")
    SEMICOLON_COLS = [COL_DOCID] + other_merge_cols
    total_groups = len(groups)
    out_rows = []
    dob_conflict_rows = []
    docid_overflow_groups = 0
    max_docid_cols = 1
    other_addr_overflow_groups = 0
    max_other_addr_cols = 1
    for n, group_idxs in enumerate(groups, 1):
        progress("Building output", n, total_groups)
        sub = df.iloc[group_idxs]           # O(group size), not O(n)
        sub_recs = [recs[i] for i in group_idxs]

        merged_dob = dob_merge(sub[COL_DOB])
        if MERGE_SEP in merged_dob:
            # 2+ genuinely different real DOBs ended up in one group (via a
            # blank-DOB bridge row) - don't merge this group at all. Every
            # original row goes into "DOB Conflict Review" as-is, tagged
            # with which candidate group they'd have landed in, for manual
            # review instead of silently combining conflicting birth dates
            # into one person.
            for idx in group_idxs:
                conflict_row = df.iloc[idx].to_dict()
                conflict_row["Candidate Group ID"] = n
                conflict_row["Candidate Group Size"] = len(group_idxs)
                dob_conflict_rows.append(conflict_row)
            continue

        row = {c: semicolon_merge(sub[c]) for c in SEMICOLON_COLS if c != COL_DOCID}
        row[COL_DOCID] = semicolon_merge(
            itertools.chain.from_iterable(sub[c] for c in docid_input_cols)
        )
        row[COL_DOB] = merged_dob

        # A group merging enough rows can produce a DOCID string longer than
        # Excel's 32,767-char cell limit (silently truncated otherwise) -
        # split it across 'DOCIDs', 'DOCIDs 2', 'DOCIDs 3', ... columns,
        # breaking only at '; ' boundaries so no DOCID is ever cut in half.
        docid_chunks = split_docid_chunks(row[COL_DOCID])
        row[COL_DOCID] = docid_chunks[0]
        for extra_i, chunk in enumerate(docid_chunks[1:], start=2):
            row[f"{COL_DOCID} {extra_i}"] = chunk
        if len(docid_chunks) > 1:
            docid_overflow_groups += 1
            max_docid_cols = max(max_docid_cols, len(docid_chunks))

        row[COL_FIRST] = fullest_value(sub[COL_FIRST], [r.first for r in sub_recs]) or "[Unknown]"
        row[COL_LAST] = fullest_value(sub[COL_LAST], [r.last for r in sub_recs]) or "[Unknown]"
        row[COL_MIDDLE] = fullest_value(sub[COL_MIDDLE], [r.mid for r in sub_recs])
        row[COL_SUFFIX] = fullest_value(sub[COL_SUFFIX], [norm_text(v) for v in sub[COL_SUFFIX]])
        row[COL_SSN] = fullest_value(sub[COL_SSN], [r.ssn for r in sub_recs])

        majority_addr_values, other_address = split_addresses(df, group_idxs)
        for c, v in zip(ADDRESS_COLS, majority_addr_values):
            row[c] = v
        if other_addr_input_cols:
            # Preserve whatever the input already had here (e.g. from a
            # prior merge pass) instead of overwriting it with only this
            # run's own newly-computed value.
            other_address = semicolon_merge(itertools.chain(
                itertools.chain.from_iterable(sub[c] for c in other_addr_input_cols),
                [other_address],
            ))

        # Same overflow protection as DOCIDs above - split across
        # 'Other Address', 'Other Address 2', 'Other Address 3', ...
        other_addr_chunks = split_docid_chunks(other_address, max_chars=OTHER_ADDR_CHUNK_SIZE)
        row[COL_OTHER_ADDR] = other_addr_chunks[0]
        for extra_i, chunk in enumerate(other_addr_chunks[1:], start=2):
            row[f"{COL_OTHER_ADDR} {extra_i}"] = chunk
        if len(other_addr_chunks) > 1:
            other_addr_overflow_groups += 1
            max_other_addr_cols = max(max_other_addr_cols, len(other_addr_chunks))

        row["Rows Merged"] = len(group_idxs)
        row["Names Differ"] = has_variation(sub[COL_FIRST]) or has_variation(sub[COL_LAST])
        out_rows.append(row)

    df_out = pd.DataFrame(out_rows)

    # Column order: follow the INPUT file's own header sequence (not a
    # fixed order), so the output layout matches the source workbook the
    # user already knows - any 'DOCIDs 2', 'DOCIDs 3', ... overflow columns
    # go right after 'DOCIDs', and any 'Other Address 2', 'Other Address 3',
    # ... overflow columns go right after 'Other Address' (regardless of
    # whether the input already had those overflow columns itself - they're
    # always placed via this fixed logic, never via their own raw input
    # position too, so a chained run's input never ends up with the same
    # overflow column listed twice). Columns with no input counterpart
    # ('Other Address' on a fresh source file, 'Rows Merged', 'Names Differ')
    # go at the end.
    docid_extra_cols = [f"{COL_DOCID} {i}" for i in range(2, max_docid_cols + 1)]
    other_addr_extra_cols = [f"{COL_OTHER_ADDR} {i}" for i in range(2, max_other_addr_cols + 1)]
    docid_extra_names = {f"{COL_DOCID} {i}" for i in range(2, max_docid_cols + 1)}
    other_addr_extra_names = {f"{COL_OTHER_ADDR} {i}" for i in range(2, max_other_addr_cols + 1)}
    input_order = list(df.columns)
    extra_cols = [c for c in df_out.columns
                  if c not in input_order
                  and c not in docid_extra_names
                  and c not in other_addr_extra_names]
    new_order = []
    for c in input_order:
        if c in docid_extra_names or c in other_addr_extra_names:
            continue
        if c not in df_out.columns:
            continue
        new_order.append(c)
        if c == COL_DOCID:
            new_order.extend(docid_extra_cols)
        if c == COL_OTHER_ADDR:
            new_order.extend(other_addr_extra_cols)
    for c in extra_cols:
        new_order.append(c)
        if c == COL_OTHER_ADDR:
            new_order.extend(other_addr_extra_cols)
    df_out = df_out[new_order]

    df_out = df_out.sort_values(["Rows Merged"], ascending=False).reset_index(drop=True)

    if docid_overflow_groups:
        print(f"  {docid_overflow_groups:,} group(s) had a DOCID list too long for one "
              f"cell (> {DOCID_CHUNK_SIZE:,} chars) - split across up to "
              f"{max_docid_cols} '{COL_DOCID}' columns.")
    if other_addr_overflow_groups:
        print(f"  {other_addr_overflow_groups:,} group(s) had an Other Address list too "
              f"long for one cell (> {OTHER_ADDR_CHUNK_SIZE:,} chars) - split across up "
              f"to {max_other_addr_cols} '{COL_OTHER_ADDR}' columns.")

    n_multi = (df_out["Rows Merged"] > 1).sum()
    print(f"  {n_multi:,} merged groups combine 2+ original rows.")
    biggest = df_out["Rows Merged"].max()
    print(f"  Largest merged group: {biggest:,} rows.")

    df_dob_conflict = pd.DataFrame(dob_conflict_rows)
    if len(df_dob_conflict):
        df_dob_conflict = df_dob_conflict.sort_values(
            ["Candidate Group Size", "Candidate Group ID"], ascending=[False, True]
        ).reset_index(drop=True)
        n_conflict_groups = df_dob_conflict["Candidate Group ID"].nunique()
        print(f"  WARNING: {n_conflict_groups:,} candidate group(s) ({len(df_dob_conflict):,} rows) "
              f"had 2+ genuinely different DOBs and were held back from merging entirely "
              f"- see the 'DOB Conflict Review' sheet for manual review.")

    print(f"Writing {OUTPUT_XLSX} ...")
    _write_workbook(OUTPUT_XLSX, {
        "Merged Notification Data": df_out,
        "DOB Conflict Review": df_dob_conflict,
    })
    print(f"Done -> {OUTPUT_XLSX}")
    print("Reminder: save the output only to the secured/authorized folder for "
          "this data - never a desktop or personal drive. It contains SSN, "
          "DOB, and other PII/PHI.")


def _write_workbook(path, sheets: dict):
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for sheet, df in sheets.items():
            if len(df) > EXCEL_MAX_ROWS:
                csv_name = path.replace(".xlsx", f" {sheet}.csv")
                df.to_csv(csv_name, index=False, encoding="utf-8-sig")
                print(f"  {sheet}: {len(df):,} rows exceed Excel limit -> {csv_name}")
                pd.DataFrame({"note": [f"{sheet} exported to {csv_name} (too large)"]}
                             ).to_excel(xl, sheet_name=sheet, index=False)
            else:
                df.to_excel(xl, sheet_name=sheet, index=False)


if __name__ == "__main__":
    # Optional command-line overrides - see the module docstring's Run
    # section:
    #   python "260716 pd ds unknown merge.py"                     (opens the file-picker dialog, OUTPUT_XLSX above as-is)
    #   python "260716 pd ds unknown merge.py" input.csv           (skips the dialog, reads that file directly)
    #   python "260716 pd ds unknown merge.py" input.csv out.xlsx  (also overrides the output filename)
    if len(sys.argv) > 3:
        print(f"Usage: python {sys.argv[0]!r} [input_file] [output_file]", file=sys.stderr)
        sys.exit(1)
    if len(sys.argv) > 2:
        OUTPUT_XLSX = sys.argv[2]

    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

