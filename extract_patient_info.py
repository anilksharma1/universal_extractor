"""
Extract patient-information fields (name, health plan ID, claim #, account #,
service date) from remittance-advice / EOB PDFs, handling both scanned
(image-only) and digitally generated pages, across two known layouts:

  Format A - tabular "Remittance Advice" layout: repeating blocks keyed by
             Member #, with Claim #, Received Date, Service Date, Proc/Mod/Qty
             and Patient Name, terminated by a "Patient Acct. #" line and
             "Claim Totals:" / "Member Totals:" rows.

  Format B - label-per-record "Combined Check" layout: a patient name line
             followed by "Health Plan ID:", "Claim:", "Pat. Acct #", a
             Service Date / Revenue Code / CPT- header, and a data row.

Billing/payment columns (billed/allowed/paid amounts, adjustments, etc.) are
intentionally NOT extracted -- only patient-identifying/claim-reference
fields are kept.

Runs in two passes over the whole input batch, each with its own progress
bar, so every file is made searchable before any extraction is attempted:

  Pass 1 (searchable): for every input PDF, any page without a usable text
     layer is treated as scanned -- it is rasterized, its rotation is
     corrected using Tesseract's orientation detection (OSD), and it is
     OCR'd. A new, searchable version of the PDF is written alongside the
     original as '<filename>_searchable.pdf' (pages that already have a
     text layer are copied through unchanged).

  Pass 2 (extraction): once every file has a searchable PDF on disk, each
     one is read back and scanned for both formats' patterns. Each file's
     matches are written to their own '<filename>_extracted_raw.xlsx',
     sheet "extracted" -- one workbook per input file, never combined.

The regex patterns below are a first-pass ("raw") best guess built from the
documented column layouts, not from real sample text. Run this against a
real file and tune FORMAT_A_BLOCK_RE / FORMAT_B_RE if a block fails to
match -- OCR line-wrapping and spacing can vary from what's assumed here.

Usage:
    python extract_patient_info.py <pdf_file_or_folder> [-o output_dir]

Requires:
    pip install pymupdf openpyxl pytesseract pillow tqdm
    Tesseract OCR engine installed (on PATH, or set TESSERACT_CMD below)
"""

import re
import argparse
import bisect
from datetime import datetime
from pathlib import Path

import pymupdf as fitz
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.utils import get_column_letter
from tqdm import tqdm

# If tesseract.exe isn't on PATH, uncomment and set the full path:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

DPI = 300
MIN_TEXT_CHARS_PER_PAGE = 20  # below this, a page is treated as scanned/image-only

# First/middle-name part only allows single spaces between words, so a multi-space
# column gap (e.g. into an adjacent "Provider Name" column) stops the match instead
# of being swallowed into it. Continuation words don't require a capital start and
# the fragment count is generous (up to 4) because OCR can split one word into
# several pieces (ink/highlighter artifacts on the scanned original) -- e.g.
# "CHRISTINA" coming back as "CHR ISTI NA" would otherwise get truncated.
NAME_RE = r"[A-Z][A-Za-z'\-]+,\s*[A-Z][A-Za-z'\-\.]*(?: [A-Za-z][A-Za-z'\-\.]*){0,4}"
DATE_RE = r"\d{1,2}/\d{1,2}/\d{2,4}"
NAME_SEARCH_RE = re.compile(NAME_RE)
DATE_SEARCH_RE = re.compile(DATE_RE)

# Fixed-size windows (chars) searched around each anchor below, instead of
# scanning the whole document -- keeps every regex O(window size), not
# O(document size), and avoids catastrophic backtracking on large OCR text.
WINDOW_BEFORE = 200
WINDOW_AFTER = 500

# ---- Format B: label-per-record ("Combined Check"), anchored on "Health Plan ID:" ----
HEALTH_PLAN_RE = re.compile(r"Health Plan ID:\s*(?P<planid>[\w\d]+)")
CLAIM_LABEL_RE = re.compile(r"Claim:\s*(?P<claim>[\w\d]+)")
PAT_ACCT_LABEL_RE = re.compile(r"Pat\.?\s*Acct\s*#\s*(?P<acct>[\w\d]+)")
SVC_HEADER_RE = re.compile(r"Service Date.{0,80}?CPT-?", re.DOTALL)

# ---- Format A: tabular ("Remittance Advice") ----
# Read line by line with running state instead of windowed/anchored regex blocks.
# A member's claims can be split across a PDF page boundary (its last claim row,
# Patient Acct #, and Member Totals: printed at the top of the NEXT page, with
# nothing page-local to anchor on). Since the full document is one continuous
# text stream (all pages concatenated in reading order), a stateful line reader
# naturally carries a patient's data across that boundary; a block regex bounded
# by character-offset windows cannot, because the "record" isn't a contiguous,
# boundable span in the text -- it's whatever lines arrive between two events
# regardless of which page they happen to land on.
#
# Patient Name sits on the top line of a record (same line as Member#/Claim#/
# "Medi-Cal"), before that record's own "Patient Acct. #" line(s). Provider
# names in this layout ("SAI WONG") aren't comma-formatted, so NAME_RE only
# ever matches the patient's "Last, First" name.
NAME_LINE_RE = NAME_SEARCH_RE
MEMBER_TOTALS_RE = re.compile(r"Member Totals\s*:")
PATIENT_ACCT_RE = re.compile(r"Patient Acct\.?\s*#\s*(?P<acct>[\w\d]+)")
MEMBER_RE = re.compile(r"(?P<member>\d{9,})")
CLAIM_LINE_RE = re.compile(r"(?P<claim>\d{6,})")

# Proc is the code token right after the last date on a row (a row can carry up
# to 3 dates -- Received Date, Service Date From, Service Date To -- the LAST
# being the actual/latest service date for that row).
PROC_TOKEN_RE = re.compile(r"[A-Za-z]?\d{3,5}")


def deskew_page_image(pil_img):
    """Correct 0/90/180/270 rotation using Tesseract's orientation detection."""
    try:
        osd = pytesseract.image_to_osd(pil_img)
        angle = int(re.search(r"Rotate: (\d+)", osd).group(1))
    except pytesseract.TesseractError:
        angle = 0
    if angle:
        pil_img = pil_img.rotate(-angle, expand=True)
    return pil_img


def make_searchable_pdf(input_path: Path, output_path: Path):
    """Write a searchable version of `input_path` to `output_path`."""
    doc = fitz.open(input_path)
    out_pdf = fitz.open()

    pages = tqdm(list(enumerate(doc)), desc=f"  {input_path.name}", unit="page", leave=False)
    for page_index, page in pages:
        text = page.get_text().strip()

        if len(text) >= MIN_TEXT_CHARS_PER_PAGE:
            out_pdf.insert_pdf(doc, from_page=page_index, to_page=page_index)
            continue

        pages.set_postfix_str("OCR'ing")
        try:
            pix = page.get_pixmap(dpi=DPI)
            pil_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            pil_img = deskew_page_image(pil_img)

            ocr_pdf_bytes = pytesseract.image_to_pdf_or_hocr(pil_img, extension="pdf")
            ocr_page = fitz.open("pdf", ocr_pdf_bytes)
            out_pdf.insert_pdf(ocr_page)
            ocr_page.close()
        except Exception as exc:
            # Never let one bad page (corrupt image data, a Tesseract crash, an
            # out-of-memory render at 300 DPI, etc.) kill the whole file's run and
            # silently truncate everything from that page onward. Fall back to the
            # original, unrecognized page so the page COUNT and ORDER stay intact --
            # it just won't be searchable/extractable -- and keep going.
            tqdm.write(f"    {input_path.name} page {page_index + 1}: OCR failed ({exc}), keeping original page as-is")
            out_pdf.insert_pdf(doc, from_page=page_index, to_page=page_index)

    if out_pdf.page_count != doc.page_count:
        tqdm.write(
            f"    WARNING: {input_path.name}: searchable PDF has {out_pdf.page_count} pages, "
            f"expected {doc.page_count} -- some content may be missing"
        )

    out_pdf.save(output_path)
    out_pdf.close()
    doc.close()


def extract_pdf_text(path: Path):
    """Return (full_text, page_offsets): page_offsets[i] is the character offset
    within full_text where page i+1 (1-indexed) begins, for page_number_at()."""
    doc = fitz.open(path)
    # sort=True orders spans top-to-bottom/left-to-right; multi-column layouts
    # (e.g. a name/claim block next to a separate "Network"/label column)
    # otherwise come out in the wrong reading order under the default mode,
    # silently breaking line-position-based matches (name-above-anchor, etc.).
    parts = [page.get_text(sort=True) for page in doc]
    doc.close()

    page_offsets = []
    offset = 0
    for part in parts:
        page_offsets.append(offset)
        offset += len(part) + 1  # +1 for the "\n" the join below inserts

    return "\n".join(parts), page_offsets


def page_number_at(char_offset, page_offsets):
    """1-indexed page number containing `char_offset` in the joined full text."""
    return bisect.bisect_right(page_offsets, char_offset)


def extract_format_b(text, page_offsets):
    rows = []
    skipped = []
    for i, hp_match in enumerate(HEALTH_PLAN_RE.finditer(text), start=1):
        before = text[max(0, hp_match.start() - WINDOW_BEFORE):hp_match.start()]
        after = text[hp_match.end():hp_match.end() + WINDOW_AFTER]

        # Patient name is the line directly above "Health Plan ID:". That line can
        # also carry "Payee: <provider name>" further right in the same comma
        # format, so only look left of "Payee:" to avoid grabbing the provider.
        name_line = before.rstrip("\n").rsplit("\n", 1)[-1]
        name_line = name_line.split("Payee:", 1)[0]
        name_match = NAME_SEARCH_RE.search(name_line)

        claim_match = CLAIM_LABEL_RE.search(after)
        acct_match = PAT_ACCT_LABEL_RE.search(after)
        svc_header_match = SVC_HEADER_RE.search(after)
        missing = [
            field for field, m in [
                ("Patient Name", name_match), ("Claim", claim_match),
                ("Pat. Acct #", acct_match), ("Service Date/CPT header", svc_header_match),
            ] if not m
        ]
        if missing:
            skipped.append((i, missing))
            continue

        date_match = DATE_SEARCH_RE.search(after, svc_header_match.end())
        if not date_match:
            skipped.append((i, ["Service Date value"]))
            continue

        # CPT value sits between the service date and the Billed Amount ($) on the
        # same data row. Column boundaries in OCR text aren't reliable, so this
        # captures the raw Revenue Code + CPT + modifier span rather than trying
        # to split them -- good enough for a first-pass "raw" extract.
        dollar_match = re.search(r"\$", after[date_match.end():])
        cpt_end = date_match.end() + dollar_match.start() if dollar_match else date_match.end() + 40
        cpt_value = re.sub(r"\s+", " ", after[date_match.end():cpt_end]).strip()

        rows.append({
            "Format": "B",
            "Patient Name": name_match.group().strip(),
            "Health Plan ID": hp_match.group("planid"),
            "Claim #": claim_match.group("claim"),
            "Member #": "",
            "Pat. Acct #": acct_match.group("acct"),
            "Service Date": date_match.group(),
            "CPT": cpt_value,
            "Page": page_number_at(hp_match.start(), page_offsets),
        })

    for i, missing in skipped:
        tqdm.write(f"    Format B record #{i}: skipped -- missing/unmatched: {', '.join(missing)}")
    return rows


def _line_latest_date_and_proc(line):
    """Last date on the line (Service Date To, if 3 dates are present) and the
    Proc code token right after it."""
    dates = list(DATE_SEARCH_RE.finditer(line))
    if not dates:
        return "", ""
    proc_match = PROC_TOKEN_RE.search(line, dates[-1].end())
    return dates[-1].group(), (proc_match.group() if proc_match else "")


def extract_format_a(text, page_offsets):
    rows = []
    skipped = []
    current = None  # in-progress patient record, or None between records
    record_count = 0

    offset = 0
    for raw_line in text.split("\n"):
        line_start = offset
        offset += len(raw_line) + 1  # +1 for the removed "\n"
        line = raw_line

        name_match = NAME_LINE_RE.search(line)
        if name_match:
            # New patient starts here -- Member#/Claim# and this row's own
            # date/proc come from this same line.
            record_count += 1
            member_match = MEMBER_RE.search(line)
            claim_match = CLAIM_LINE_RE.search(line)
            svc_date, proc_value = _line_latest_date_and_proc(line)
            current = {
                "name": name_match.group().strip(),
                "member": member_match.group("member") if member_match else "",
                "claim": claim_match.group("claim") if claim_match else "",
                "accts": [],
                "svc_date": svc_date,
                "proc": proc_value,
                "line_start": line_start,
            }
            continue

        if current is None:
            continue  # no patient context yet (e.g. document opens mid-record)

        acct_match = PATIENT_ACCT_RE.search(line)
        if acct_match:
            acct = acct_match.group("acct")
            if acct not in current["accts"]:
                current["accts"].append(acct)
            continue

        if MEMBER_TOTALS_RE.search(line):
            missing = [
                field for field, v in [
                    ("Member #/Claim #", current["member"] or current["claim"]),
                    ("Patient Acct #", current["accts"]),
                    ("Service Date", current["svc_date"]),
                ] if not v
            ]
            if missing:
                skipped.append((record_count, missing))
            else:
                rows.append({
                    "Format": "A",
                    "Patient Name": current["name"],
                    "Health Plan ID": "",
                    "Claim #": current["claim"],
                    "Member #": current["member"],
                    "Pat. Acct #": ":".join(current["accts"]),
                    "Service Date": current["svc_date"],
                    "Proc": current["proc"],
                    "Page": page_number_at(current["line_start"], page_offsets),
                })
            current = None
            continue

        # A continuation claim row for the same patient (e.g. a 2nd/3rd claim
        # under one Member Totals:) -- keep the latest date seen across all of
        # this patient's claim rows, not just the first (name) row.
        svc_date, _proc = _line_latest_date_and_proc(line)
        if svc_date and (parse_date(svc_date) or datetime.min) > (parse_date(current["svc_date"]) or datetime.min):
            current["svc_date"] = svc_date

    for i, missing in skipped:
        tqdm.write(f"    Format A record #{i}: skipped -- missing/unmatched: {', '.join(missing)}")
    return rows


DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y")


def parse_date(date_str):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


def merge_same_patient(rows):
    """Rows with the same Patient Name AND the same (non-blank) Health Plan ID
    are the same patient reported across multiple claims -- merge them into one
    row: combine their distinct Patient Acct #s with ':' and keep only the
    latest Service Date. Rows with no Health Plan ID (Format A) or a differing
    Health Plan ID/name are left as separate rows, per patient-identity rule:
    same name + different plan ID or account # = a different patient."""
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

        latest_row = max(group, key=lambda r: parse_date(r.get("Service Date", "")) or datetime.min)

        combined = dict(group[0])
        combined["Pat. Acct #"] = ":".join(accts)
        combined["Service Date"] = latest_row.get("Service Date", "")
        merged.append(combined)

    return passthrough + merged


def build_workbook(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "extracted"
    columns = ["Format", "Patient Name", "Health Plan ID", "Claim #", "Member #", "Pat. Acct #", "Service Date", "Proc", "CPT", "Page"]
    ws.append(columns)
    for row in rows:
        ws.append([row.get(c, "") for c in columns])
    for i, header in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(header) + 2)
    return wb


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("-o", "--output-dir", default=".", help="Directory to write outputs into")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Pass 1: make every file searchable (OCR + rotation fix) before any extraction runs.
    # Files that already have a _searchable.pdf on disk are skipped (e.g. re-running
    # after tuning extraction logic shouldn't force OCR to redo).
    searchable_paths = {}
    for pdf in tqdm(pdf_files, desc="Pass 1/2: making files searchable", unit="file"):
        searchable_path = output_dir / f"{pdf.stem}_searchable.pdf"
        if searchable_path.exists():
            expected_pages = fitz.open(pdf).page_count
            actual_pages = fitz.open(searchable_path).page_count
            if actual_pages != expected_pages:
                tqdm.write(
                    f"  {pdf.name}: existing searchable PDF has {actual_pages} pages, "
                    f"expected {expected_pages} -- redoing OCR"
                )
                make_searchable_pdf(pdf, searchable_path)
            else:
                tqdm.write(f"  {pdf.name}: searchable PDF already exists ({actual_pages} pages), skipping OCR")
        else:
            make_searchable_pdf(pdf, searchable_path)
        searchable_paths[pdf] = searchable_path

    # Pass 2: extract patient info from each searchable file, one workbook per input file.
    for pdf in tqdm(pdf_files, desc="Pass 2/2: extracting patient info", unit="file"):
        text, page_offsets = extract_pdf_text(searchable_paths[pdf])
        rows = extract_format_b(text, page_offsets) + extract_format_a(text, page_offsets)
        rows = merge_same_patient(rows)
        if not rows:
            tqdm.write(f"  {pdf.name}: no records matched either format -- check regex patterns against actual OCR text")

        output_xlsx = output_dir / f"{pdf.stem}_extracted_raw.xlsx"
        build_workbook(rows).save(output_xlsx)
        tqdm.write(f"  {pdf.name} -> {output_xlsx.name} ({len(rows)} record(s))")


if __name__ == "__main__":
    main()
