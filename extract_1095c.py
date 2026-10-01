"""
Extract employee identity/address fields from IRS Form 1095-C ("Employer-
Provided Health Insurance Offer and Coverage") PDFs -- one employee record
per page, across two known layouts:

  Format A - "substitute statement" style (seen on a 2019-dated render):
             a single combined caption "EMPLOYEE's name, address, ZIP/postal
             code & country" followed by 2-3 physical lines (Name, Street,
             "City, ST ZIP[, Country]"), terminated by the boilerplate line
             "Do not attach to your tax return...". SSN sits in its own
             "EMPLOYEE's social security number (SSN)" field further down
             the same page.

  Format B - official numbered-box layout (seen on a 2015-dated render):
             "Part I Employee", fields 1-6 ("1 Name of employee", "2 Social
             security number (SSN)", "3 Street address ...", "4 City or
             town", "5 State or province", "6 Country and ZIP or foreign
             postal code"), each caption followed by its value either on
             the same line or the line(s) below, up to the next numbered
             caption. Field 7 ("Name of employer") is used only as the end
             boundary for field 6, never extracted.

Only Part I (employee identity/address) fields are extracted -- Part II
(monthly coverage offer codes) and Part III (covered individuals) are
intentionally never captured.

This is a first-pass "raw" best guess built from the two documented layouts,
not from real extracted text -- run it against a real (non-sensitive) file
and tune FORMAT_A_* / FORMAT_B_FIELDS below if a page's fields come back
blank. The console output below only ever prints filenames/page numbers/
field *names* that failed to match -- never a field's actual value -- so it
is safe to run and read its output without exposing PHI in a terminal/log
that might get shared.

One workbook is produced per input PDF:
    <filename>_1095c_extracted.xlsx - every field listed above, one row per
                                        employee page

Employee name is assumed to be exactly "First Middle Last" (3 space-
separated words); 2-word or 4+-word names still split (2 -> First/Last with
Middle blank, 4+ -> first word/last word/everything between as Middle) but
are flagged in the console output for manual review since the 3-word
assumption doesn't hold for them.

Usage:
    python extract_1095c.py <pdf_file_or_folder> [-o output_dir]

Requires:
    pip install pymupdf openpyxl
"""

import re
import argparse
from pathlib import Path

import pymupdf as fitz
import openpyxl
from openpyxl.utils import get_column_letter

MIN_TEXT_CHARS_PER_PAGE = 20  # below this, a page is treated as scanned/image-only

OUTPUT_COLUMNS = [
    "Page", "Format", "First Name", "Middle Name", "Last Name",
    "Street Address", "City", "State", "Zip Code", "Country", "SSN",
]

# "City, ST 12345" or "City, ST 12345-6789", optionally with a trailing
# country name/word that isn't part of the ZIP.
CITY_STATE_ZIP_RE = re.compile(
    r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b\s*(?P<country>.*)$"
)
SSN_VALUE_RE = re.compile(r"[\dXx\*\-]{9,}")

FORMAT_A_NAME_LABEL_RE = re.compile(r"EMPLOYEE.s name,\s*address,\s*ZIP/?postal code\s*&?\s*country", re.IGNORECASE)
FORMAT_A_SSN_LABEL_RE = re.compile(r"EMPLOYEE.s social security number\s*\(SSN\)", re.IGNORECASE)
FORMAT_A_NAME_BLOCK_END_RE = re.compile(r"Do not attach to your tax return", re.IGNORECASE)
FORMAT_A_SSN_BLOCK_END_RE = re.compile(
    r"Department of the Treasury|Covered Individuals|APPLICABLE LARGE EMPLOYER", re.IGNORECASE
)

# Ordered numbered-box captions for Format B, Part I Employee (field 7 is the
# end boundary for field 6 and is never itself extracted).
FORMAT_B_FIELDS = [
    ("employee_name", re.compile(r"^\s*1\s+Name of employee", re.IGNORECASE)),
    ("ssn", re.compile(r"^\s*2\s+Social security number\s*\(SSN\)", re.IGNORECASE)),
    ("street", re.compile(r"^\s*3\s+Street address(?:\s*\(including apartment no\.?\))?", re.IGNORECASE)),
    ("city", re.compile(r"^\s*4\s+City or town", re.IGNORECASE)),
    ("state", re.compile(r"^\s*5\s+State or province", re.IGNORECASE)),
    ("zip", re.compile(r"^\s*6\s+Country and ZIP(?:\s+or foreign postal code)?", re.IGNORECASE)),
    ("_end", re.compile(r"^\s*7\s+Name of employer", re.IGNORECASE)),
]

# Matches a leading, not-fully-consumed remainder of field 6's caption tail
# ("or foreign postal code") when its line wrap splits that phrase across
# two lines -- see the zip/country split in extract_format_b.
ZIP_CAPTION_LEFTOVER_RE = re.compile(r"^(?:\b(?:or|foreign|postal|code)\b\s*)+", re.IGNORECASE)


def group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def split_employee_name(full_name):
    full_name = re.sub(r"\s+", " ", (full_name or "").strip())
    if not full_name:
        return "", "", "", False
    parts = full_name.split(" ")
    if len(parts) == 2:
        return parts[0], "", parts[1], False
    if len(parts) == 1:
        return parts[0], "", "", False
    # 3 words is the assumed case; 4+ still splits (first/last word, rest as
    # Middle) but is flagged as non-standard for manual review.
    return parts[0], " ".join(parts[1:-1]), parts[-1], len(parts) != 3


def find_label_span(plain_lines, label_re, start=0, max_window=3):
    """Return (start_idx, end_idx, trailing_text) for the first place
    label_re matches, checking windows of 1..max_window consecutive lines
    joined together -- handles a caption that wraps onto a second (or
    third) physical line, which a single-line search would silently miss.
    trailing_text is whatever remains after the match within that window
    (usually empty when the caption itself takes the whole window).

    Uses .match (anchored at the window's first line), not .search: with
    .search, a wider window starting a few lines EARLY can still contain
    the label text further inside it, so a short window at the label's
    true start (tried later) never gets a chance to win -- and worse, the
    call site then treats those earlier, actually-value lines as part of
    the label instead of data. Anchoring at the window start means a
    window only matches when the label genuinely begins at that line."""
    for i in range(start, len(plain_lines)):
        for span in range(1, max_window + 1):
            if i + span > len(plain_lines):
                break
            joined = " ".join(plain_lines[i:i + span])
            m = label_re.match(joined)
            if m:
                return i, i + span - 1, joined[m.end():].strip()
    return None, None, None


def value_after_span(plain_lines, span_end_idx, trailing_text, boundary_start_idx):
    """Trailing text left over on the label's own span, PLUS the non-empty
    line(s) between the label span and `boundary_start_idx` (exclusive) --
    combined, not either/or: a caption regex that only partially consumes
    a wrapped caption (e.g. matches "...ZIP" but not the "or foreign postal
    code" that continues on the next line) leaves that leftover caption
    text in `trailing_text`, and the real value still follows on the lines
    after it -- short-circuiting on non-empty trailing would silently drop
    that value."""
    between = [l.strip() for l in plain_lines[span_end_idx + 1:boundary_start_idx] if l.strip()]
    combined = ([trailing_text] if trailing_text else []) + between
    return " ".join(combined)


# Generic 1095-C boilerplate (present on every real form, never personal
# data) used only to diagnose *why* neither format matched -- e.g. "the form
# text extracted fine but our caption wording is off" vs "almost no text
# came out of this page at all". Printed as yes/no + counts, never content.
DIAGNOSTIC_TERMS = ["1095-C", "employee", "employer", "social security", "coverage"]


def log_no_match_diagnostics(plain_lines, page_num):
    joined = " ".join(plain_lines).lower()
    hits = {term: joined.count(term.lower()) for term in DIAGNOSTIC_TERMS}
    print(f"  page {page_num}: neither Format A nor Format B employee block found -- not a 1095-C "
          f"employee page, skipping")
    print(f"    diagnostic (counts only, no PHI): {len(plain_lines)} line(s) extracted; "
          + ", ".join(f"'{term}': {count}" for term, count in hits.items()))
    if len(plain_lines) < 5:
        print("    -> very few lines extracted overall: this page may be scanned/image-only, or the PDF "
              "stores field values as fillable form widgets rather than page content (run --debug and "
              "check if the dump is nearly empty for this page)")
    elif hits["employee"] == 0 and hits["1095-C"] == 0:
        print("    -> no 'employee'/'1095-C' boilerplate found at all: this page's text likely isn't a "
              "1095-C employee block, or is a continuation/cover page")
    else:
        print("    -> boilerplate terms ARE present but our caption regex didn't match: run --debug and "
              "tell me the exact label line text (the caption only, e.g. the line starting with "
              "'EMPLOYEE' or '1 Name of employee') -- that line is generic form text, not PHI, safe to paste")


def extract_format_a(plain_lines, page_num):
    name_start, name_end, name_trailing = find_label_span(plain_lines, FORMAT_A_NAME_LABEL_RE)
    if name_start is None:
        return None

    end_start, _, _ = find_label_span(plain_lines, FORMAT_A_NAME_BLOCK_END_RE, name_end + 1)
    end_idx = end_start if end_start is not None else min(name_end + 6, len(plain_lines))

    block = [l.strip() for l in [name_trailing] + plain_lines[name_end + 1:end_idx] if l.strip()]

    if not block:
        print(f"  page {page_num} (Format A): name/address block empty after label -- skipping")
        return None

    full_name = block[0]
    street = block[1] if len(block) > 1 else ""
    city = state = zip_code = country = ""
    if len(block) > 2:
        m = CITY_STATE_ZIP_RE.match(block[2])
        if m:
            city, state, zip_code, country = m.group("city").rstrip(","), m.group("state"), m.group("zip"), m.group("country").strip()
        else:
            city = block[2]
            print(f"  page {page_num} (Format A): could not split city/state/zip line -- placing raw text in City")

    ssn = ""
    ssn_start, ssn_end, ssn_trailing = find_label_span(plain_lines, FORMAT_A_SSN_LABEL_RE, end_idx)
    if ssn_start is not None:
        ssn_boundary_start, _, _ = find_label_span(plain_lines, FORMAT_A_SSN_BLOCK_END_RE, ssn_end + 1)
        if ssn_boundary_start is None:
            ssn_boundary_start = min(ssn_end + 4, len(plain_lines))
        ssn_text = value_after_span(plain_lines, ssn_end, ssn_trailing, ssn_boundary_start)
        ssn_match = SSN_VALUE_RE.search(ssn_text)
        ssn = ssn_match.group() if ssn_match else ""
        if not ssn:
            print(f"  page {page_num} (Format A): SSN label found but no SSN-shaped value nearby")
    else:
        print(f"  page {page_num} (Format A): SSN label not found")

    return {
        "Page": page_num, "Format": "A", "full_name": full_name,
        "Street Address": street, "City": city, "State": state,
        "Zip Code": zip_code, "Country": country, "SSN": ssn,
    }


def extract_format_b(plain_lines, page_num):
    spans = {}
    search_from = 0
    for key, label_re in FORMAT_B_FIELDS:
        start, end, trailing = find_label_span(plain_lines, label_re, search_from)
        if start is None:
            if key == "employee_name":
                return None  # not a Format B page
            continue
        spans[key] = (start, end, trailing)
        search_from = end + 1

    if "employee_name" not in spans:
        return None

    ordered_keys = [k for k, _ in FORMAT_B_FIELDS]
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
        values[key] = value_after_span(plain_lines, span_end, trailing, next_start)

    missing = [k for k in ("employee_name", "ssn", "street", "city", "state", "zip") if not values.get(k)]
    if missing:
        print(f"  page {page_num} (Format B): fields not matched: {', '.join(missing)}")

    # The "6 Country and ZIP..." caption's optional tail ("or foreign postal
    # code") is only consumed by FORMAT_B_FIELDS' regex when it lands whole
    # on one line; if the wrap point splits it, the leftover words end up
    # prepended to this field's value (via value_after_span) and must be
    # stripped before what's left is treated as a country name.
    zip_text = ZIP_CAPTION_LEFTOVER_RE.sub("", values.get("zip", "")).strip()
    zip_match = re.search(r"\d{5}(?:-\d{4})?", zip_text)
    zip_code = zip_match.group() if zip_match else ""
    country = zip_text[:zip_match.start()].strip() if zip_match else zip_text.strip()

    ssn_match = SSN_VALUE_RE.search(values.get("ssn", ""))

    return {
        "Page": page_num, "Format": "B", "full_name": values.get("employee_name", ""),
        "Street Address": values.get("street", ""), "City": values.get("city", ""),
        "State": values.get("state", ""), "Zip Code": zip_code, "Country": country,
        "SSN": ssn_match.group() if ssn_match else "",
    }


def dump_debug_lines(plain_lines, pdf_stem, page_num, output_dir):
    """Write this page's line-grouped text to a local, gitignored file so it
    can be inspected on the machine that has the real data -- never printed
    to the console and never sent anywhere. Caller reviews the file itself;
    only the *structural* mismatch (which line number holds what) needs to
    be described back for a regex fix, not its contents."""
    debug_path = output_dir / f"{pdf_stem}_debug.txt"
    with open(debug_path, "a", encoding="utf-8") as f:
        f.write(f"\n===== page {page_num} =====\n")
        for i, line in enumerate(plain_lines):
            f.write(f"[{i}] {line}\n")
    return debug_path


def process_page(page, page_num, debug=False, pdf_stem=None, output_dir=None):
    text_len = len(page.get_text().strip())
    if text_len < MIN_TEXT_CHARS_PER_PAGE:
        print(f"  page {page_num}: only {text_len} chars of text -- scanned/image-only page, skipping "
              f"(these PDFs are expected to be searchable already)")
        return None

    words = page.get_text("words")
    lines = group_words_into_lines(words)
    plain_lines = [" ".join(t for _, _, t in ln) for ln in lines]

    if debug:
        dump_debug_lines(plain_lines, pdf_stem, page_num, output_dir)

    record = extract_format_a(plain_lines, page_num) or extract_format_b(plain_lines, page_num)
    if record is None:
        log_no_match_diagnostics(plain_lines, page_num)
        return None

    first, middle, last, nonstandard = split_employee_name(record.pop("full_name"))
    record["First Name"], record["Middle Name"], record["Last Name"] = first, middle, last
    if nonstandard:
        print(f"  page {page_num}: employee name has more than 3 words -- split as first/last word with "
              f"everything between as Middle Name; review this row")
    return record


def process_pdf(path: Path, debug=False, output_dir=None):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        record = process_page(page, i, debug=debug, pdf_stem=path.stem, output_dir=output_dir)
        if record:
            records.append(record)
    doc.close()
    return records


def build_workbook(records):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "extracted"
    ws.append(OUTPUT_COLUMNS)
    for rec in records:
        ws.append([rec.get(c, "") for c in OUTPUT_COLUMNS])
    for i, header in enumerate(OUTPUT_COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(header) + 2)
    return wb


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("-o", "--output-dir", default=".", help="Directory to write outputs into")
    parser.add_argument("--debug", action="store_true",
                         help="Also write <pdf>_debug.txt (line-numbered page text) to the output dir "
                              "for inspecting layout mismatches locally -- never printed to the console")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    for pdf in pdf_files:
        print(f"Processing {pdf.name} ...")
        records = process_pdf(pdf, debug=args.debug, output_dir=output_dir)
        if not records:
            print(f"  no records extracted from {pdf.name}")
            continue

        extracted_path = output_dir / f"{pdf.stem}_1095c_extracted.xlsx"

        build_workbook(records).save(extracted_path)

        print(f"  -> {extracted_path.name} ({len(records)} record(s))")


if __name__ == "__main__":
    main()
