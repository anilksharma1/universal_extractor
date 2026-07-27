"""
Extract patient/claim identity fields from ABBYY-searchable remittance PDFs.
Two known layouts are handled, auto-detected line by line (a given PDF is
expected to be entirely one layout or the other):

Format B -- "Combined Check": repeating label-per-record blocks --

    <Last, First>                       Payee: <provider>
    Health Plan ID: <id>                Claim: <claim #>
    Pat. Acct #<acct>                   Status: <status>

    Service Date  Revenue Code  CPT-  Billed Amount  ...
    <date>        <code>        <cpt> <amount>       ...
                                       <amount>       ...   (totals row, no date)

    Fields extracted, one row per block: Patient Name, Health Plan ID,
    Claim #, Pat. Acct #, Service Date, CPT Code. Anchored on "Health Plan
    ID:"; the patient name is the line directly above it (left of any
    "Payee:" text on that same line); Claim # is read off the "Claim:" label
    on that same "Health Plan ID:" line; Service Date/CPT come from the
    first data row under the header that carries a date (the totals row
    that follows has none, so it's skipped automatically).

Format A -- tabular "Remittance Advice": one or more claim-line-item rows
per member, each row carrying a Received Date/Service Date From/Service Date
To/Proc, ended by "Patient Acct. #" and "Member Totals:" lines --

    Extraction is anchored on the patient name: a block with no detectable
    name is skipped entirely rather than having its details attributed to
    the wrong record. Fields extracted, one row per member/claim block:
    Patient Name (the comma-formatted name printed on the block's first
    data row, under the "Patient Name" header), Claim # and Member # (read
    positionally -- Claim # leads the same line as the patient name and
    repeats on every line-item row under that claim; Member # sits alone on
    its own line directly above, printed once per member rather than once
    per claim -- so it's optional and left blank if that line isn't
    findable), Pat. Acct #, and Service Date -- the LATEST Service Date To
    among that block's rows, counting only rows that have a value in the
    Proc column (a row with no Proc value doesn't count towards "latest").
    Health Plan ID and CPT Code aren't part of this layout, so those columns
    are left blank for these
    rows.

Both extractors run over the same PDF; each is anchored on markers only its
own layout has ("Health Plan ID:" vs "Member Totals:"), so a PDF in one
format simply yields zero records from the other extractor.

This assumes the PDF already has a text layer (e.g. made searchable via
ABBYY, as in this project) -- no OCR is performed here.

This file contains patient-identifying data (PHI), so console output --
both --debug's line dump and the "missing field" warnings -- is REDACTED by
default: every letter becomes 'X' and every digit becomes '#', except for a
whitelist of known structural/label words (Health, Plan, ID, Service, Date,
CPT, Payee, Claim, Status, Pat, Acct, User, Comments, etc.), so line
structure, spacing, and punctuation are still visible for troubleshooting
without exposing real names/IDs/dates. The output .xlsx itself is never
redacted -- only what gets printed to the console.

Usage:
    python extract_claims_info.py <pdf_file_or_folder> [-o output_dir]

Each input PDF produces its own "<filename>_extracted.xlsx", written next to
the source file unless -o/--output-dir is given. After every PDF is
processed, a single "combined_extracted.xlsx" is also written alongside them
(or into -o/--output-dir), unioning every PDF's records onto one sheet with
an added "Source File" column so each row stays traceable back to its PDF.

Requires:
    pip install pymupdf openpyxl
"""

import re
import argparse
from datetime import datetime
from pathlib import Path

import pymupdf as fitz
import openpyxl
from openpyxl.utils import get_column_letter

# Last name may be multiple capitalized words ("GARCIA VARGAS, ..."); the gap
# after the comma is [^A-Za-z]* rather than \s* because OCR occasionally fuses a
# stray punctuation character (e.g. "!") onto the front of the first name -- that
# stray character sits outside both capture groups so it's dropped from the result
# (see extract_name, which reassembles the name from group(1)/group(2) rather than
# taking the raw matched span).
NAME_RE = re.compile(r"([A-Z][A-Za-z'\-]+(?:\s+[A-Z][A-Za-z'\-]+)*),[^A-Za-z]*([A-Za-z][A-Za-z'\-\. ]*)")
# Format B occasionally OCR-misreads the name-separating comma as a period (seen on
# some pages). Used only by extract_name, not Format A's find_name_match/tabular scan --
# extract_name's search is tightly bounded to 1-2 lines right around a confirmed
# "Health Plan ID:" anchor, whereas Format A scans much more free-form text (disclaimer
# prose, etc.) where treating an ordinary sentence-ending period as a name separator
# would risk far more false positives.
NAME_RE_LOOSE_SEP = re.compile(r"([A-Z][A-Za-z'\-]+(?:\s+[A-Z][A-Za-z'\-]+)*)[,.][^A-Za-z]*([A-Za-z][A-Za-z'\-\. ]*)")
DATE_RE = re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}")

# Format A's "Line of Business" value (e.g. "Medi-Cal") sits directly before the
# patient name with no separating label, and is itself a capitalized, hyphenated
# word -- exactly what the multi-word-surname part of NAME_RE also matches. Without
# this, "Medi-Cal VILLA, Daniel" would wrongly capture "Medi-Cal VILLA" as the
# surname. find_name_match skips a match whenever its surname's leading word is one
# of these known fillers and keeps searching further into the text instead.
#
# "pat" is here for a different reason: with NAME_RE_LOOSE_SEP (period accepted as a
# separator), the field label "Pat. Acct #" itself matches the pattern outright --
# "Pat" + "." + "Acct" looks exactly like "Lastname. Firstname" -- and was observed
# coming out as a literal "Pat, Acct" patient name.
NAME_PRECEDING_FILLER_WORDS = {"medi-cal", "medicare", "commercial", "medi-medi", "pat"}


def find_name_match(text, pos=0, pattern=NAME_RE):
    while True:
        m = pattern.search(text, pos)
        if not m:
            return None
        first_word = m.group(1).split()[0]
        if first_word.lower() not in NAME_PRECEDING_FILLER_WORDS:
            return m
        pos = m.start(1) + len(first_word)

# Words that carry no patient-identifying content -- kept as-is in console
# output so the layout/structure stays readable; everything else printed to
# the console gets its letters/digits masked (see redact_word/redact_value).
DEBUG_SAFE_WORDS = {
    "health", "plan", "id", "claim", "status", "paid", "pat", "acct", "payee",
    "service", "date", "revenue", "code", "cpt", "billed", "amount", "not",
    "allowed", "contract", "adjust", "interest", "penalty", "deduct", "coin",
    "copay", "net", "amt", "user", "comments", "member", "totals", "line",
    "ver", "received", "from", "to", "proc", "mod", "qty", "patient", "name",
    "provider", "business", "medi", "cal", "no", "pcp", "network", "reason",
    "of", "by", "program", "all", "programs", "combined", "check", "vendor",
    "group", "community", "clinics", "connect", "e", "p", "s", "t", "ra",
}


def redact_word(word):
    core = word.strip(":#.,()/-’'")
    if core.lower() in DEBUG_SAFE_WORDS:
        return word
    return "".join("#" if ch.isdigit() else ("X" if ch.isalpha() else ch) for ch in word)


def redact_line_for_debug(line):
    return " ".join(redact_word(t) for _, _, t in line)


def redact_value(value):
    return "".join("#" if ch.isdigit() else ("X" if ch.isalpha() else ch) for ch in str(value))

# ---- Format B: "Combined Check" (label-per-record) ----
HEALTH_PLAN_RE = re.compile(r"Health\s*Plan\s*ID\s*:\s*(\S+)")
HEALTH_PLAN_LABEL_RE = re.compile(r"Health\s*Plan\s*ID\s*:")
PAT_ACCT_RE = re.compile(r"Pat\.?\s*Acct\s*#\s*(\S+)")
# "Claim:" sits on the same visual line as "Health Plan ID:" (see module docstring
# layout diagram), so it's read straight off that line rather than via lookahead.
CLAIM_LABEL_RE = re.compile(r"Claim\s*:")
# A genuine continuation of an OCR-split claim number (see extract_claim) -- a short
# alnum fragment, unlike the next real field's value which starts with a recognizable
# multi-letter word.
CLAIM_CONTINUATION_RE = re.compile(r"^[A-Za-z0-9-]{1,3}\d+$")

# ---- Format A: tabular "Remittance Advice" ----
PATIENT_ACCT_RE = re.compile(r"Patient\s+Acct\.?\s*#\s*(\S+)")
# ":" is occasionally OCR-misread as ";" -- accept either so a record's closing
# "Member Totals" line isn't missed, leaving that record (and everything after it,
# since `current` never gets reset) merged into whatever follows.
MEMBER_TOTALS_RE = re.compile(r"Member Totals\s*[:;]")
# Proc is the code token after the row's last date (dates read left-to-right
# as Received Date, Service Date From, Service Date To -- the last one is
# Service Date To). re.search from that point skips over any non-matching
# text in between (e.g. a "Line of Business" value like "Medi-Cal" that only
# appears on a block's first row), landing on the actual Proc code either way.
PROC_TOKEN_RE = re.compile(r"[A-Za-z]?\d{3,5}")
DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y")
PROVIDER_NAME_HEADER_RE = re.compile(r"Provider\s*Name")


def find_provider_name_x0(lines):
    """Patient Name and Provider Name are adjacent columns with nothing
    separating their values on a data row -- NAME_RE's first-name group
    allows spaces, so without a hard stop it happily reads straight through
    into the provider's name too (e.g. "Christina Sai Wong"). Locate the
    "Provider Name" header's x0 once so name-matching can be restricted to
    words left of it."""
    for line in lines:
        if PROVIDER_NAME_HEADER_RE.search(line_text(line)):
            x0 = word_x0(line, re.compile(r"^Provider$"))
            if x0 is not None:
                return x0
    return None


def parse_date(date_str):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

Y_TOL = 5             # px gap tolerance for grouping words onto the same visual line
COL_X_TOL = 2         # px tolerance when assigning a word to a header column
LOOKBACK_NAME = 2      # lines above "Health Plan ID:" to search for the patient name
                       # before giving up -- deliberately tight (docstring's own "one
                       # extra stray line" rationale) so a record with no real name
                       # line above it (e.g. the first record on a page, right below
                       # the page header) can't wander into unrelated header text.
LOOKAHEAD_ACCT = 4     # lines after "Health Plan ID:" to search for "Pat. Acct #"
LOOKAHEAD_HEADER = 8   # lines after "Health Plan ID:" to search for the Service Date/CPT header
LOOKAHEAD_DATA_ROWS = 25  # lines after the header to search for the first dated data row --
                          # generous enough to span a page break: when a record's data rows
                          # are the very first thing on the NEXT page (the record's header
                          # printed at the bottom of this one), the search has to cross this
                          # page's footer/disclaimer block (~7 lines) plus the next page's own
                          # running header (~5 lines) before reaching real data.

COLUMNS = ["Patient Name", "Health Plan ID", "Member #", "Claim #", "Pat. Acct #", "Service Date", "CPT Code", "Page"]


def group_words_into_lines(words):
    """words: (x0, y0, x1, y1, text, block_no, line_no, word_no) tuples from
    page.get_text('words'). Returns lines (top-to-bottom), each a list of
    (x0, x1, text) sorted left-to-right.

    Real OCR'd text has enough per-word vertical jitter within one nominal
    printed row (differing character heights/baselines, e.g. "Payee:" vs a
    name next to it) that a fixed-grid y-rounding scheme can split a single
    row into two lines, or interleave it with a neighboring row, once that
    jitter crosses a grid boundary. Instead, cluster by GAP: sort every word
    by vertical center, then start a new line only when the gap to the
    previous word's vertical center exceeds Y_TOL -- this tolerates gradual
    per-word drift within a row while still separating genuinely different
    rows (which are spaced much further apart than the jitter)."""
    entries = sorted(
        ((x0, x1, text, (y0 + y1) / 2) for x0, y0, x1, y1, text, *_ in words),
        key=lambda e: (e[3], e[0]),
    )
    lines = []
    current = []
    last_y = None
    for x0, x1, text, ymid in entries:
        if last_y is not None and abs(ymid - last_y) > Y_TOL:
            lines.append(sorted(current, key=lambda t: t[0]))
            current = []
        current.append((x0, x1, text))
        last_y = ymid
    if current:
        lines.append(sorted(current, key=lambda t: t[0]))
    return lines


def line_text(line):
    return " ".join(t for _, _, t in line)


def build_lines_for_page(page, mirror):
    words = page.get_text("words")
    if mirror:
        page_width = page.rect.width
        words = [(page_width - x1, y0, page_width - x0, y1, text, *rest)
                  for x0, y0, x1, y1, text, *rest in words]
    return group_words_into_lines(words)


def page_looks_readable(lines):
    text = "\n".join(line_text(l) for l in lines)
    return bool(
        HEALTH_PLAN_RE.search(text) or MEMBER_TOTALS_RE.search(text)
        or PATIENT_ACCT_RE.search(text) or PAT_ACCT_RE.search(text)
    )


def score_complete_records(lines, line_pages):
    """How many fully non-blank records a quiet trial extraction finds for
    this candidate line order -- used to auto-pick the correct top-to-bottom
    line order for a page (see get_page_lines). Regex label matches (e.g.
    'Health Plan ID:') work regardless of line order since they're within a
    single line, so they can't detect a reversed line sequence; completeness
    of the surrounding-line lookups (name above, acct/date below) can."""
    b_rows = extract_records(lines, line_pages, quiet=True)
    a_rows = extract_tabular_records(lines, line_pages, quiet=True)
    complete_b = sum(
        1 for r in b_rows
        if all(r.get(c) for c in ("Patient Name", "Claim #", "Pat. Acct #", "Service Date", "CPT Code"))
    )
    complete_a = sum(
        1 for r in a_rows
        if all(r.get(c) for c in ("Patient Name", "Pat. Acct #", "Service Date"))
    )
    return complete_b + complete_a


def get_page_lines(page, debug=False, page_num=None):
    """Some ABBYY-produced text layers come out with word coordinates
    mirrored relative to the visible page (the rendered image still looks
    normal -- only the invisible OCR text run is off), which silently
    reverses reading order and breaks both label regexes and
    surrounding-line/column lookups. Two independent things can be flipped:

    - Left-right (within a line): detected via known anchor phrases, which
      only read correctly forward in the right orientation.
    - Top-to-bottom (line sequence): label regexes match either way (they're
      single-line), so this can't be detected via a simple presence check.
      score_complete_records used to be the tiebreaker, but extract_name's
      forward-search fallback (added for resilience on scrambled pages) also
      made it direction-INSENSITIVE: a record can now score "complete" in
      EITHER orientation, so a nonzero forward score no longer proves the
      orientation is correct. count_backward_names is the direction-
      sensitive replacement -- it counts names found the CONVENTIONAL
      (backward-only, no fallback) way, which should be common in the
      correct orientation and rare in the wrong one. Only switch to reversed
      when it clearly wins (strictly more backward-only names, and neither
      side wins nothing) -- ties or both-zero default to the normal
      orientation, same "don't be too eager" caution as before.
    """
    normal = build_lines_for_page(page, mirror=False)
    oriented = normal
    if not page_looks_readable(normal):
        mirrored = build_lines_for_page(page, mirror=True)
        if page_looks_readable(mirrored):
            oriented = mirrored
            if debug:
                print(f"  [debug] page {page_num}: text layer is mirrored left-right -- corrected automatically")

    page_nums = [page_num] * len(oriented)
    reversed_lines = list(reversed(oriented))

    # count_backward_names only has a signal on Format B pages (it's anchored on
    # "Health Plan ID:"); on Format A pages it's 0 either way, so fall back to the
    # original completeness-based heuristic there. Deciding via count_backward_names
    # FIRST (rather than gating on score_complete_records like before) matters: that
    # score requires EVERY field non-empty including CPT Code, which can be 0 in BOTH
    # orientations for reasons that have nothing to do with ordering (e.g. a garbled
    # CPT cell) -- gating on it would mask a real, fixable ordering problem.
    forward_strict = count_backward_names(oriented)
    reversed_strict = count_backward_names(reversed_lines)
    if forward_strict == 0 and reversed_strict == 0:
        forward_score = score_complete_records(oriented, page_nums)
        if forward_score > 0:
            return oriented
        reversed_score = score_complete_records(reversed_lines, page_nums)
        if reversed_score > 0:
            if debug:
                print(f"  [debug] page {page_num}: top-to-bottom line order was reversed -- corrected "
                      f"automatically ({reversed_score} complete record(s) found reversed, 0 found normally)")
            return reversed_lines
        return oriented

    if reversed_strict > forward_strict:
        if debug:
            print(f"  [debug] page {page_num}: top-to-bottom line order was reversed -- corrected automatically "
                  f"({reversed_strict} record(s) with a conventionally-positioned name reversed, "
                  f"{forward_strict} found normally)")
        return reversed_lines
    return oriented


def word_x0(line, token_re):
    for x0, x1, text in line:
        if token_re.match(text):
            return x0
    return None


def bucket_row(line, col_defs):
    """col_defs: [(label, x0), ...] sorted by x0 ascending. Assigns each word
    to the rightmost column whose start is still <= the word's x0."""
    buckets = {label: [] for label, _ in col_defs}
    for x0, x1, text in line:
        label = col_defs[0][0]
        for cand_label, start in col_defs:
            if x0 + COL_X_TOL >= start:
                label = cand_label
        buckets[label].append(text)
    return {label: " ".join(words).strip() for label, words in buckets.items()}


def extract_name(lines, hp_idx, allow_forward_fallback=True):
    """The patient name is normally the line directly above 'Health Plan
    ID:', but real OCR'd text can still land an extra stray line-group in
    between (residual row-grouping noise), or -- seen on some pages -- get
    merged onto the very same line as 'Health Plan ID:' itself (tight line
    spacing collapsing two visual rows into one). So the name line's own
    'Health Plan ID:' line is checked first (using only the text before
    that label), then the search reads backward line by line until a
    comma-formatted name is found -- stopping only at a natural record
    boundary (the previous block's own 'Health Plan ID:', or its 'User
    Comments:' trailer) so an overshoot can't wander into an earlier,
    unrelated block's name (e.g. the page header/address block, which can
    itself contain a stray comma that coincidentally matches NAME_RE).

    On some pages the local line order around a record comes out scrambled
    (seen: 'Pat. Acct .../Status:' printing BEFORE 'Health Plan ID:', with
    the name pushed to just after it instead of before) even though the
    page-level mirror/reversal auto-correction didn't trigger. If nothing
    turns up backward, also try a small forward window, bounded the same
    way, before giving up -- unless allow_forward_fallback is False, which
    count_backward_names (see get_page_lines) uses to get a direction-
    sensitive signal: with the fallback enabled, a name becomes findable
    regardless of which way the page actually reads, so it can no longer
    tell orientation-detection which direction is correct."""
    same_line = line_text(lines[hp_idx]).split("Health Plan ID:")[0]
    m = find_name_match(same_line.split("Payee:")[0], pattern=NAME_RE_LOOSE_SEP)
    if m:
        return f"{m.group(1).strip()}, {m.group(2).strip()}"

    lookback_stop = max(hp_idx - 1 - LOOKBACK_NAME, -1)
    for i in range(hp_idx - 1, lookback_stop, -1):
        text = line_text(lines[i])
        if HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        above = text.split("Payee:")[0]
        m = find_name_match(above, pattern=NAME_RE_LOOSE_SEP)
        if m:
            return f"{m.group(1).strip()}, {m.group(2).strip()}"

    if not allow_forward_fallback:
        return ""

    lookahead_stop = min(hp_idx + 1 + LOOKBACK_NAME, len(lines))
    for i in range(hp_idx + 1, lookahead_stop):
        text = line_text(lines[i])
        if HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        above = text.split("Payee:")[0]
        m = find_name_match(above, pattern=NAME_RE_LOOSE_SEP)
        if m:
            return f"{m.group(1).strip()}, {m.group(2).strip()}"
    return ""


def count_backward_names(lines):
    """How many 'Health Plan ID:' records find their name the CONVENTIONAL
    way (backward-only, no forward fallback) -- used only by get_page_lines
    to pick the correct page orientation. score_complete_records uses the
    fully resilient extractors, which (thanks to extract_name's forward
    fallback) can now find "complete" records in EITHER orientation, so a
    nonzero score there no longer proves the orientation is right. This
    stays direction-sensitive because it deliberately doesn't use that
    fallback."""
    return sum(
        1 for i, line in enumerate(lines)
        if HEALTH_PLAN_RE.search(line_text(line)) and extract_name(lines, i, allow_forward_fallback=False)
    )


def extract_health_plan_id(text):
    """The token right after 'Health Plan ID:' is occasionally a stray word
    that doesn't belong to the ID (seen on merged/scrambled lines), with the
    real ID -- always containing at least one digit -- one token further
    along. Skip non-digit tokens rather than blindly taking the first one."""
    label = HEALTH_PLAN_LABEL_RE.search(text)
    if not label:
        return ""
    for tok in text[label.end():].split():
        if any(ch.isdigit() for ch in tok):
            return tok
    return ""


def extract_claim(lines, hp_idx):
    text = line_text(lines[hp_idx])
    label = CLAIM_LABEL_RE.search(text)
    if not label:
        return ""
    tokens = text[label.end():].split()
    if not tokens:
        return ""
    claim = tokens[0]
    if len(tokens) > 1 and CLAIM_CONTINUATION_RE.match(tokens[1]):
        claim += tokens[1]
    return claim


def extract_acct(lines, hp_idx):
    for i in range(hp_idx, min(hp_idx + LOOKAHEAD_ACCT, len(lines))):
        m = PAT_ACCT_RE.search(line_text(lines[i]))
        if m:
            return m.group(1)
    # On some pages this record's own "Pat. Acct #"/"Status:" line prints BEFORE
    # "Health Plan ID:" instead of after (see find_service_header for the matching
    # header-position quirk) -- fall back to a small backward window, bounded by
    # the previous record's own boundary, before giving up.
    lookback_stop = max(hp_idx - 1 - LOOKAHEAD_ACCT, -1)
    for i in range(hp_idx - 1, lookback_stop, -1):
        text = line_text(lines[i])
        if HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        m = PAT_ACCT_RE.search(text)
        if m:
            return m.group(1)
    return ""


def find_service_header(lines, hp_idx):
    """Locate the Service Date/Revenue Code/CPT-/Billed Amount header line
    and return (line_index, col_defs) with col_defs sorted by x0.

    Anchored on "CPT-" alone -- not the literal words "Service Date" or
    "Billed" -- because different pages have been seen to OCR-garble
    different header cells (one page mangles "Service", another mangles
    "Billed", while the rest of the same header line reads cleanly).
    Requiring any one specific word beyond "CPT-" (which has held up across
    every page seen so far) makes the header undetectable on whichever page
    happens to garble that word. Service Date's and the amounts column's x0
    are taken positionally instead -- the header line's first word, and the
    word immediately following "CPT-" -- rather than by matching their
    literal text, since those positions hold regardless of what those
    cells actually OCR'd to.

    Tried backward first, then forward: on some pages this record's own
    header/Pat. Acct print BEFORE "Health Plan ID:" instead of after (the
    whole label block reads in a locally reversed order), so a forward-only
    search can walk right past this record's real header and grab the NEXT
    record's instead. Checking backward first is safe for normal pages too
    -- it's bounded by the previous record's own "Health Plan ID:"/"User
    Comments:" boundary, so it simply finds nothing there and falls through
    to the forward search unchanged."""
    def header_col_defs(line):
        cpt_idx = next((idx for idx, (_, _, text) in enumerate(line) if re.match(r"^CPT-?$", text)), None)
        if cpt_idx is None or cpt_idx + 1 >= len(line):
            return None
        svc_x0 = line[0][0]
        cpt_x0 = line[cpt_idx][0]
        billed_x0 = line[cpt_idx + 1][0]
        rev_x0 = word_x0(line, re.compile(r"^Revenue$"))
        return sorted(
            {"service_date": svc_x0, "revenue_code": rev_x0 if rev_x0 is not None else cpt_x0,
             "cpt": cpt_x0, "amounts": billed_x0}.items(),
            key=lambda kv: kv[1],
        )

    lookback_stop = max(hp_idx - 1 - LOOKAHEAD_HEADER, -1)
    for i in range(hp_idx - 1, lookback_stop, -1):
        text = line_text(lines[i])
        if HEALTH_PLAN_RE.search(text) or "User Comments:" in text:
            break
        col_defs = header_col_defs(lines[i])
        if col_defs:
            return i, col_defs

    for i in range(hp_idx, min(hp_idx + LOOKAHEAD_HEADER, len(lines))):
        col_defs = header_col_defs(lines[i])
        if col_defs:
            return i, col_defs
    return None, None


def extract_service_date_and_cpt(lines, header_idx, col_defs):
    """Date and CPT are searched independently (not required off the same
    row) because a claim's first data row is sometimes revenue-code-only --
    its CPT value only appears on the following row -- so pinning both to
    one row can find a date but miss a CPT that's one row down.

    LOOKAHEAD_DATA_ROWS is wide enough to cross a page break, so this stops
    early if it runs into the NEXT record's own 'Health Plan ID:' line --
    otherwise a record with genuinely no data of its own could pick up a
    later, unrelated record's date/CPT. That width can also reach into the
    CURRENT page's own footer/disclaimer text (which doesn't contain
    "Health Plan ID:", so the check above doesn't catch it) when a record's
    real CPT genuinely isn't nearby -- a stray disclaimer word can land in
    the CPT column's x-range and get grabbed as if it were a real code. A
    real CPT/procedure code always contains a digit ("99213", "G0439"), so
    requiring one rejects that kind of plain-word garbage.

    Service Date takes the LATEST date seen across the record's rows (a
    claim can span several dated line-items, or -- now that the lookahead
    crosses page breaks -- even continue past a page boundary), same
    rationale as Format A's "latest Service Date To". CPT still takes the
    first valid value found, since it isn't described as needing "latest"
    treatment the way a date range is."""
    latest_date, cpt = None, ""
    for i in range(header_idx + 1, min(header_idx + 1 + LOOKAHEAD_DATA_ROWS, len(lines))):
        if HEALTH_PLAN_RE.search(line_text(lines[i])):
            break
        bucket = bucket_row(lines[i], col_defs)
        date_match = DATE_RE.search(bucket.get("service_date", ""))
        if date_match:
            parsed = parse_date(date_match.group())
            if parsed and (latest_date is None or parsed > latest_date):
                latest_date = parsed
        cpt_candidate = bucket.get("cpt", "")
        if not cpt and cpt_candidate and any(ch.isdigit() for ch in cpt_candidate):
            cpt = cpt_candidate
    service_date = latest_date.strftime("%m/%d/%Y") if latest_date else ""
    return service_date, cpt


def extract_records(lines, line_pages, quiet=False):
    rows = []
    for i, line in enumerate(lines):
        hp_match = HEALTH_PLAN_RE.search(line_text(line))
        if not hp_match:
            continue

        health_plan_id = extract_health_plan_id(line_text(line))
        name = extract_name(lines, i)
        claim = extract_claim(lines, i)
        acct = extract_acct(lines, i)
        header_idx, col_defs = find_service_header(lines, i)
        service_date, cpt = ("", "")
        if header_idx is not None:
            # header_idx can be BEFORE hp_idx (see find_service_header) -- data rows
            # still belong to this record's own forward span past "Health Plan ID:",
            # not right after a backward-found header, so search from whichever of
            # the two comes later.
            service_date, cpt = extract_service_date_and_cpt(lines, max(header_idx, i), col_defs)

        missing = [
            label for label, val in [
                ("Patient Name", name), ("Health Plan ID", health_plan_id),
                ("Claim #", claim), ("Pat. Acct #", acct),
                ("Service Date", service_date), ("CPT Code", cpt),
            ] if not val
        ]
        if missing and not quiet:
            print(f"    record at Health Plan ID {redact_value(health_plan_id or hp_match.group(1))} "
                  f"(page {line_pages[i]}): missing {', '.join(missing)} -- check the layout around this block")

        rows.append({
            "Patient Name": name,
            "Health Plan ID": health_plan_id,
            "Member #": "",
            "Claim #": claim,
            "Pat. Acct #": acct,
            "Service Date": service_date,
            "CPT Code": cpt,
            "Page": line_pages[i],
        })
    return rows


def last_date_and_proc(text):
    dates = list(DATE_RE.finditer(text))
    if not dates:
        return None, ""
    proc_match = PROC_TOKEN_RE.search(text, dates[-1].end())
    return dates[-1].group(), (proc_match.group() if proc_match else "")


def extract_tabular_records(lines, line_pages, quiet=False):
    """Format A. Anchored on the patient name: everything else about a
    record (Member #, Claim #, Pat. Acct #, Service Date) is only extracted
    once a name is actually found for that block, and runs through to
    'Member Totals:'. If a block has no detectable name, `current` simply
    never gets set for it (see the `current is None` guard below), so its
    Patient Acct/date lines are skipped rather than mis-attributed to
    whichever record came before it.

    Claim # leads the same line as the patient name (per the "Claim#" header
    column -- it repeats on every line-item row under that claim, Line/Ver#
    incrementing alongside it). Member # sits alone on its own line directly
    above the name row, printed once per member rather than once per claim.
    A name match is only treated as a genuine record, not the report's own
    check/provider address block (which can contain a coincidental comma,
    e.g. "Some City, CA 90210" -- and has nothing resembling a claim number
    before it), when it's preceded on its line by a digit-bearing prefix."""
    rows = []
    skipped = []
    current = None
    provider_name_x0 = find_provider_name_x0(lines)

    for i, line in enumerate(lines):
        text = line_text(line)
        if provider_name_x0 is not None:
            name_search_text = line_text([w for w in line if w[0] < provider_name_x0 - COL_X_TOL])
        else:
            name_search_text = text

        name_match = find_name_match(name_search_text)
        prefix = name_search_text[:name_match.start(1)].strip() if name_match else ""
        if name_match and not any(ch.isdigit() for ch in prefix):
            name_match = None

        if name_match:
            name = f"{name_match.group(1).strip()}, {name_match.group(2).strip()}"
            prefix_tokens = prefix.split()
            claim = prefix_tokens[0] if prefix_tokens else ""
            member_id = ""
            if i > 0:
                prev_text = line_text(lines[i - 1]).strip()
                prev_tokens = prev_text.split()
                # Skip the previous line if it's a Totals/Acct line (this member has no
                # separate standalone Member # line, e.g. right after a "Member Totals"
                # boundary with nothing in between) or a header/label row (no digit --
                # a genuine Member # always has one).
                if (prev_tokens and "Totals" not in prev_text and "Acct" not in prev_text
                        and any(ch.isdigit() for ch in prev_tokens[0])):
                    member_id = prev_tokens[0]
            current = {
                "name": name, "member_id": member_id, "claim": claim,
                "acct": "", "latest_date": None, "page": line_pages[i],
            }

        if current is None:
            continue

        acct_match = PATIENT_ACCT_RE.search(text)
        if acct_match:
            current["acct"] = acct_match.group(1)
            continue

        if MEMBER_TOTALS_RE.search(text):
            # Pat. Acct #/Service Date are hard requirements (a record without them was
            # already unusable before Member #/Claim # existed). Member #/Claim # are
            # extracted when available but, being read positionally, aren't always
            # findable (e.g. a single-claim record has no separate Ver#/dates row to pull
            # Claim # from) -- their absence is reported but shouldn't drop an otherwise
            # good record that DOES have a name/acct/date.
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
                    print(f"    tabular record for '{redact_value(current['name'])}' (page {current['page']}): "
                          f"missing {', '.join(optional_missing)} -- kept, but check the layout around this block")
                rows.append({
                    "Patient Name": current["name"],
                    "Health Plan ID": "",
                    "Member #": current["member_id"],
                    "Claim #": current["claim"],
                    "Pat. Acct #": current["acct"],
                    "Service Date": current["latest_date"].strftime("%m/%d/%Y"),
                    "CPT Code": "",
                    "Page": current["page"],
                })
            current = None
            continue

        date_str, proc = last_date_and_proc(text)
        if date_str and proc:
            parsed = parse_date(date_str)
            if parsed and (current["latest_date"] is None or parsed > current["latest_date"]):
                current["latest_date"] = parsed

    if not quiet:
        for name, page, missing in skipped:
            print(f"    tabular record for '{redact_value(name)}' (page {page}): "
                  f"missing {', '.join(missing)} -- skipped")
    return rows


def merge_duplicate_rows(rows):
    """Same Health Plan ID + same Patient Name almost always means the same
    underlying claim recorded more than once (e.g. a payment and its
    reversal) rather than two distinct records -- collapse those into a
    single row. Service Date takes the LATEST date across the group (the
    most clinically relevant one, same rationale as Format A's "latest
    Service Date To" in extract_tabular_records); every other column is
    joined with ';' wherever the group's values differ, so nothing is
    silently dropped or overwritten. Rows with a blank Health Plan ID or
    Patient Name are left alone: there's nothing reliable to group them on."""
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
        for col in COLUMNS:
            if col == "Service Date":
                dates = [d for d in (parse_date(str(row.get(col, ""))) for row in group) if d]
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


def page_sort_key(row):
    """Rows merged by merge_duplicate_rows can carry a ';'-joined Page value
    (e.g. '27;28'); sort on the first (lowest) page number in that value."""
    try:
        return int(str(row.get("Page", "")).split(";")[0])
    except ValueError:
        return 0


def extract_pdf(path: Path, debug=False, only_page=None):
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
        page_lines = get_page_lines(page, debug=debug, page_num=page_num)
        all_lines.extend(page_lines)
        line_pages.extend([page_num] * len(page_lines))
    doc.close()

    if debug:
        print(f"  [debug] {len(all_lines)} line(s) read from the PDF's text layer "
              f"(names/IDs/dates masked with X/# below -- safe to share):")
        for i, line in enumerate(all_lines):
            print(f"  [debug] line {i} (page {line_pages[i]}): {redact_line_for_debug(line)!r}")

    rows = (extract_records(all_lines, line_pages, quiet=not debug)
            + extract_tabular_records(all_lines, line_pages, quiet=not debug))
    rows = merge_duplicate_rows(rows)
    rows.sort(key=page_sort_key)
    return rows


def build_workbook(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "extracted"
    ws.append(COLUMNS)
    for row in rows:
        ws.append([row.get(c, "") for c in COLUMNS])
    for i, header in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(16, len(header) + 2)
    return wb


COMBINED_COLUMNS = ["Source File"] + COLUMNS


def build_combined_workbook(source_and_rows):
    """source_and_rows: [(source_filename, row_dict), ...] -- one entry per
    record across every PDF processed this run. Each per-file xlsx is left
    exactly as before; this just additionally unions everything into one
    sheet with a "Source File" column so records stay traceable back to
    their originating PDF."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "combined"
    ws.append(COMBINED_COLUMNS)
    for source, row in source_and_rows:
        ws.append([source] + [row.get(c, "") for c in COLUMNS])
    for i, header in enumerate(COMBINED_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(16, len(header) + 2)
    return wb


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("-o", "--output-dir", default=None,
                         help="Directory to write outputs into (default: alongside each source PDF)")
    parser.add_argument("--debug", action="store_true",
                         help="Print every line of text as this script reads it from the PDF, "
                              "so mismatches against the expected layout are easy to spot")
    parser.add_argument("--page", type=int, default=None,
                         help="Only read/extract this page number (1-indexed) -- combine with "
                              "--debug to inspect one page at a time instead of the whole file")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    combined_rows = []
    for pdf in pdf_files:
        print(f"Processing {pdf.name} ...")
        rows = extract_pdf(pdf, debug=args.debug, only_page=args.page)
        if not rows:
            print(f"  no 'Health Plan ID:' records found -- check that the PDF's text is selectable "
                  f"and that the layout matches what this script expects")
            continue

        output_dir = Path(args.output_dir) if args.output_dir else pdf.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        output_xlsx = output_dir / f"{pdf.stem}_extracted.xlsx"
        try:
            build_workbook(rows).save(output_xlsx)
        except PermissionError:
            print(f"  could not write {output_xlsx.name} -- it's likely open in Excel or another program; "
                  f"close it and re-run")
            continue
        print(f"  -> {output_xlsx.name} ({len(rows)} record(s))")
        combined_rows.extend((pdf.name, row) for row in rows)

    if combined_rows:
        combined_output_dir = Path(args.output_dir) if args.output_dir else pdf_files[0].parent
        combined_output_dir.mkdir(parents=True, exist_ok=True)
        combined_xlsx = combined_output_dir / "combined_extracted.xlsx"
        try:
            build_combined_workbook(combined_rows).save(combined_xlsx)
        except PermissionError:
            print(f"could not write {combined_xlsx.name} -- it's likely open in Excel or another program; "
                  f"close it and re-run")
        else:
            print(f"-> {combined_xlsx.name} ({len(combined_rows)} record(s) across "
                  f"{len({s for s, _ in combined_rows})} file(s))")


if __name__ == "__main__":
    main()
