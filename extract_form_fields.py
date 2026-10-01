# extract_form_fields.py — Extract fillable form field values (AcroForm) and
# digital signatures from PDFs into a CSV.
#
# A plain text-layer reader reads the page's own text layer, so it misses values typed
# into form fields (e.g. an "HCC Academic Advisor" name box): those values are
# stored in the field itself, not in the page content. This script reads the
# fields directly with PyMuPDF — including read-only / locked fields that
# can't be edited in a PDF viewer.
#
# Signatures ("Digitally signed by <name>  Date: ...") are found three ways:
#   1. Signature field  — signer name/date read from the signature data itself
#                         (works whatever the box looks like).
#   2. Flattened, text  — the box was flattened into ordinary page text; found
#                         by searching the page text.
#   3. Flattened, image — the box is only a picture ("snip"); found by OCR of
#                         the page, only when --ocr is given.
# The "Source" column says which one each row came from. This is
# informational only — it does not validate signatures cryptographically.
#
# Usage:
#   python extract_form_fields.py form.pdf               # one row per file, fields as column headers
#   python extract_form_fields.py form.pdf --long        # one row per field instead
#   python extract_form_fields.py ./forms/ --output fields.csv
#   python extract_form_fields.py ./forms/ --include-empty
#   python extract_form_fields.py ./forms/ --ocr          # also image-only signatures
#   python extract_form_fields.py ./forms/ --label "HCC Academic Advisor" --ocr
#       -> one row per file: advisor name + signer name/date only
#
# Requires: pip install PyMuPDF pandas
#   --ocr also needs: pip install pytesseract Pillow, plus Tesseract installed
#   (path taken from DEFAULT_TESSERACT_CMD in this script).

import argparse
import os
import re
import sys

import fitz  # PyMuPDF
import pandas as pd


# Damaged PDFs make MuPDF print "MuPDF error: ..." to the console; the file is
# still repaired and read, so keep the output readable.
fitz.TOOLS.mupdf_display_errors(False)

DEFAULT_TESSERACT_CMD = r"C:\Tessaract\tesseract.exe"  # set to your tesseract.exe, or None if on PATH


def resolve_pdf_files(path: str) -> list:
    """Return list of absolute PDF paths from a file or folder."""
    if os.path.isfile(path):
        if not path.lower().endswith(".pdf"):
            print(f"Error: '{path}' is not a PDF file.")
            sys.exit(1)
        return [os.path.abspath(path)]
    elif os.path.isdir(path):
        files = [
            os.path.abspath(os.path.join(path, f))
            for f in os.listdir(path)
            if f.lower().endswith(".pdf")
        ]
        if not files:
            print(f"Error: No PDF files found in '{path}'.")
            sys.exit(1)
        return sorted(files)
    else:
        print(f"Error: Path does not exist: '{path}'")
        sys.exit(1)
if hasattr(fitz.TOOLS, "mupdf_display_warnings"):  # newer PyMuPDF only
    fitz.TOOLS.mupdf_display_warnings(False)

READ_ONLY_FLAG = 1  # PDF spec: field flag bit 1 = ReadOnly
SIGNATURE_COLUMNS = ["Signed", "Signer Name", "Signed Date", "Reason", "Location", "Contact Info", "Source"]

# "Digitally signed by Arturo Torres" — loose enough for common OCR slips
# ("Digitaly", "signed bv").
SIGNED_BY_RE = re.compile(r"Digital+y\s+sign\w*\s+b[yv]\s*:?\s*(.*)", re.IGNORECASE)
# "Date: 2024.12.09 08:36:48 -06'00'"
SIGNED_DATE_RE = re.compile(
    r"Date\s*:?\s*(\d{4}[.\-/]\d{2}[.\-/]\d{2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?(?:\s*[+-]\d{2}'?\s*\d{2}'?)?)",
    re.IGNORECASE,
)


def _pdf_string(doc, xref: int, key: str) -> str:
    """Read a string value from a PDF object, '' if missing."""
    kind, value = doc.xref_get_key(xref, key)
    if kind in ("string", "name", "text"):
        return value.lstrip("/")
    return ""


def _format_pdf_date(raw: str) -> str:
    """'D:20240315143000-05'00'' -> '2024-03-15 14:30:00 -05:00' (best effort)."""
    s = raw[2:] if raw.startswith("D:") else raw
    if len(s) < 8 or not s[:8].isdigit():
        return raw
    date = f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    time = ":".join(p for p in (s[8:10], s[10:12], s[12:14]) if p.isdigit())
    tz = s[14:].replace("'", "")
    if tz and tz[0] in "+-" and len(tz) >= 5:
        tz = f"{tz[:3]}:{tz[3:5]}"
    elif tz == "Z":
        tz = "UTC"
    return " ".join(p for p in (date, time, tz) if p)


def _format_visible_date(raw: str) -> str:
    """'2024.12.09 08:36:48 -06'00'' -> '2024-12-09 08:36:48 -06:00'."""
    s = re.sub(r"\s+", " ", raw.replace("'", "")).strip()
    s = re.sub(r"^(\d{4})[./](\d{2})[./](\d{2})", r"\1-\2-\3", s)
    return re.sub(r"\s*([+-]\d{2})\s*(\d{2})$", r" \1:\2", s)


def parse_signature_text(text: str) -> list[dict]:
    """
    Find every "Digitally signed by <name> ... Date: <date>" block in a piece
    of text (page text or OCR output). The name is the rest of its line; the
    date is looked for in the next few lines, since signing tools usually
    wrap the date (and its timezone) onto separate lines.
    """
    found = []
    lines = text.splitlines()
    for i, line in enumerate(lines):
        m = SIGNED_BY_RE.search(line)
        if not m:
            continue
        name = SIGNED_DATE_RE.split(m.group(1))[0].strip(" :,")
        following = " ".join(lines[i:i + 4])
        d = SIGNED_DATE_RE.search(following)
        found.append(
            {"Signer Name": name, "Signed Date": _format_visible_date(d.group(1)) if d else ""}
        )
    return found


def ocr_region(page, dpi: int, clip=None) -> str:
    """Render a page (or part of it, form fields included) and OCR it."""
    import pytesseract
    from PIL import Image

    if DEFAULT_TESSERACT_CMD and os.path.isfile(DEFAULT_TESSERACT_CMD):
        pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESSERACT_CMD
    pix = page.get_pixmap(dpi=dpi, clip=clip)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return pytesseract.image_to_string(img)


def signature_info(doc, widget) -> dict:
    """
    Details of a digital signature field, read from its signature
    dictionary (/V). An unsigned field has no /V. Signer name, date, reason
    and location are what the signing software recorded — this does NOT
    cryptographically validate the signature or its certificate.
    """
    xref = widget.xref
    kind, value = doc.xref_get_key(xref, "V")
    if kind == "null":
        # Widget may be a kid of the actual field — /V then lives on the parent.
        p_kind, p_value = doc.xref_get_key(xref, "Parent")
        if p_kind == "xref":
            xref = int(p_value.split()[0])
            kind, value = doc.xref_get_key(xref, "V")
    if kind != "xref":
        return {"Signed": "No", "Source": "signature field"}

    sig_xref = int(value.split()[0])
    return {
        "Signed": "Yes",
        "Signer Name": _pdf_string(doc, sig_xref, "Name"),
        "Signed Date": _format_pdf_date(_pdf_string(doc, sig_xref, "M")),
        "Reason": _pdf_string(doc, sig_xref, "Reason"),
        "Location": _pdf_string(doc, sig_xref, "Location"),
        "Contact Info": _pdf_string(doc, sig_xref, "ContactInfo"),
        "Source": "signature field",
    }


def extract_form_fields(
    pdf_path: str, include_empty: bool = False, ocr: bool = False, dpi: int = 300
) -> list[dict]:
    """Return one record per form field and per flattened signature in the
    PDF (page order). Signature fields are always included, signed or not."""
    pdf_file = os.path.basename(pdf_path)
    records = []
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"  Failed to open '{pdf_file}': {e}")
        return records

    def add(page, name="", label="", ftype="", value="", read_only="", sig=None):
        record = {
            "File Name": pdf_file,
            "Page Number": page.number + 1,
            "Field Name": name,
            "Field Label": label,
            "Field Type": ftype,
            "Value": value,
            "Read Only": read_only,
        }
        record.update({col: (sig or {}).get(col, "") for col in SIGNATURE_COLUMNS})
        records.append(record)

    with doc:
        for page in doc:
            # Names already reported from a signature field on this page, so
            # the same signature isn't listed again from the page text.
            field_signers = set()

            for widget in page.widgets() or []:
                read_only = "Yes" if (widget.field_flags or 0) & READ_ONLY_FLAG else "No"
                if widget.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
                    sig = signature_info(doc, widget)
                    if sig["Signed"] == "Yes" and ocr and not (sig["Signer Name"] and sig["Signed Date"]):
                        # Name/date missing from the signature data (some tools
                        # keep the name only in the certificate) — read the
                        # visible box instead.
                        visible = parse_signature_text(ocr_region(page, dpi, clip=widget.rect))
                        if visible:
                            sig["Signer Name"] = sig["Signer Name"] or visible[0]["Signer Name"]
                            sig["Signed Date"] = sig["Signed Date"] or visible[0]["Signed Date"]
                            sig["Source"] = "signature field + OCR of box"
                    if sig.get("Signer Name"):
                        field_signers.add(sig["Signer Name"].lower())
                    add(page, widget.field_name or "", widget.field_label or "",
                        widget.field_type_string, sig.get("Signer Name", ""), read_only, sig)
                    continue

                value = widget.field_value
                if isinstance(value, bool):
                    value = "Yes" if value else "Off"
                value = "" if value is None else str(value).strip()
                if not value and not include_empty:
                    continue
                add(page, widget.field_name or "", widget.field_label or "",
                    widget.field_type_string, value, read_only)

            # Flattened signatures: first the page's own text, then (with
            # --ocr, only if the text had none) an OCR pass for image-only ones.
            flattened = [(s, "page text (flattened)") for s in parse_signature_text(page.get_text())]
            if not flattened and ocr:
                flattened = [(s, "OCR (image only)") for s in parse_signature_text(ocr_region(page, dpi))]
            for s, source in flattened:
                if s["Signer Name"].lower() in field_signers:
                    continue
                add(page, ftype="Signature (flattened)", value=s["Signer Name"],
                    sig={**s, "Signed": "Yes", "Source": source})

    return records


LABEL_SEARCH_PT = 40  # how far above/below a label to look for its value box


def value_near_label(page, label: str, ocr: bool = False, dpi: int = 300) -> tuple[str, str]:
    """
    Value of the box belonging to a printed label such as "HCC Academic
    Advisor" — the label sits either just below the box or just above it
    (as a header bar). Tries, in order: a form field there, the page text
    there (flattened form), then OCR of that strip (--ocr). Returns
    (value, source).
    """
    for lr in page.search_for(label):
        above = fitz.Rect(lr.x0 - LABEL_SEARCH_PT, lr.y0 - LABEL_SEARCH_PT, lr.x1 + LABEL_SEARCH_PT, lr.y0)
        below = fitz.Rect(lr.x0 - LABEL_SEARCH_PT, lr.y1, lr.x1 + LABEL_SEARCH_PT, lr.y1 + LABEL_SEARCH_PT)

        # 1. Form field closest to the label, overlapping it horizontally.
        best = None
        for w in page.widgets() or []:
            if w.field_type == fitz.PDF_WIDGET_TYPE_SIGNATURE:
                continue
            value = "" if w.field_value is None else str(w.field_value).strip()
            if not value or not (w.rect.intersects(above) or w.rect.intersects(below)):
                continue
            overlap = min(w.rect.x1, lr.x1) - max(w.rect.x0, lr.x0)
            if overlap <= 0:
                continue
            gap = min(abs(lr.y0 - w.rect.y1), abs(w.rect.y0 - lr.y1))
            if best is None or gap < best[0]:
                best = (gap, value)
        if best:
            return best[1], "form field"

        # 2. / 3. Flattened: page text in the strip, then OCR of the strip.
        for band in (above, below):
            lines = [l.strip() for l in page.get_text("text", clip=band).splitlines()]
            lines = [l for l in lines if l and label.lower() not in l.lower()]
            if lines:
                return lines[-1] if band is above else lines[0], "page text (flattened)"
        if ocr:
            for band in (above, below):
                lines = [l.strip() for l in ocr_region(page, dpi, clip=band).splitlines()]
                lines = [l for l in lines if l and label.lower() not in l.lower()]
                if lines:
                    return lines[-1] if band is above else lines[0], "OCR (image only)"
    return "", ""


def extract_label_summary(pdf_path: str, labels: list[str], ocr: bool = False, dpi: int = 300) -> dict:
    """One row per file: the value for each label, plus the signature(s)."""
    row = {"File Name": os.path.basename(pdf_path)}
    for label in labels:
        row[label] = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            for label in labels:
                if not row[label]:
                    value, _ = value_near_label(page, label, ocr, dpi)
                    row[label] = value

    sigs = [r for r in extract_form_fields(pdf_path, ocr=ocr, dpi=dpi) if r["Signed"] == "Yes"]
    row["Signed"] = "Yes" if sigs else "No"
    row["Signer Name"] = "; ".join(s["Signer Name"] for s in sigs if s["Signer Name"])
    row["Signed Date"] = "; ".join(s["Signed Date"] for s in sigs if s["Signed Date"])
    row["Signature Source"] = "; ".join(dict.fromkeys(s["Source"] for s in sigs))
    return row


def records_to_wide_row(pdf_path: str, records: list[dict]) -> dict:
    """
    One row per file, header-wise: each form field becomes its own column
    (header = field label, else field name), and each signature gets
    "Signature N ..." columns. Duplicate headers get " (2)", " (3)", ...
    """
    row = {"File Name": os.path.basename(pdf_path)}
    sig_no = 0
    for r in records:
        if r["Signed"]:
            sig_no += 1
            prefix = f"Signature {sig_no}"
            row[f"{prefix} Signed"] = r["Signed"]
            row[f"{prefix} Signer Name"] = r["Signer Name"]
            row[f"{prefix} Date"] = r["Signed Date"]
            row[f"{prefix} Source"] = r["Source"]
            continue
        header = r["Field Label"] or r["Field Name"] or f"Field (page {r['Page Number']})"
        key, n = header, 1
        while key in row:
            n += 1
            key = f"{header} ({n})"
        row[key] = r["Value"]
    return row


def main():
    parser = argparse.ArgumentParser(
        description="Extract fillable form field values and digital signatures from PDF(s) into a CSV.",
    )
    parser.add_argument(
        "path",
        nargs="+",
        help="Path to a single PDF file or a folder containing PDFs. Spaces in path are handled automatically.",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output CSV file name (default: derived from input path, with '_fields' suffix).",
    )
    parser.add_argument(
        "--include-empty",
        action="store_true",
        help="Also list fields that have no value (default: only filled fields).",
    )
    parser.add_argument(
        "--ocr",
        action="store_true",
        help=(
            "Also OCR pages to find image-only (snipped) signatures, and "
            "signature boxes whose name/date aren't in the signature data. "
            "Slower — every page without a text signature is OCR'd."
        ),
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI used when rendering pages for --ocr (default: 300).",
    )
    parser.add_argument(
        "--label",
        action="append",
        default=None,
        help=(
            "Only output the value of the box next to this printed label, "
            "plus the signature — one row per file. Repeat for more labels, "
            "e.g. --label \"HCC Academic Advisor\" --label \"E-mail or Phone\"."
        ),
    )
    parser.add_argument(
        "--long",
        action="store_true",
        help=(
            "Output one row per field (File, Page, Field Name, Value, ...) "
            "instead of the default one row per file with each field as a column header."
        ),
    )
    args = parser.parse_args()
    input_path = " ".join(args.path)
    pdf_files = resolve_pdf_files(input_path)

    if args.ocr:
        try:
            import pytesseract  # noqa: F401
            from PIL import Image  # noqa: F401
        except ImportError:
            print("Error: --ocr requires: pip install pytesseract Pillow (and Tesseract installed).")
            sys.exit(1)

    if args.output:
        output_csv = args.output
    elif os.path.isfile(input_path):
        output_csv = os.path.splitext(os.path.basename(input_path))[0] + "_fields.csv"
    else:
        folder_name = os.path.basename(os.path.normpath(input_path)) or "output"
        output_csv = folder_name + "_fields.csv"

    print(f"\nFound {len(pdf_files)} PDF file(s)  |  OCR: {'on' if args.ocr else 'off'}  →  output: {output_csv}\n")

    if args.label:
        rows = []
        for idx, pdf_path in enumerate(pdf_files, start=1):
            try:
                row = extract_label_summary(pdf_path, args.label, args.ocr, args.dpi)
            except Exception as e:
                print(f"  [{idx}/{len(pdf_files)}] failed — {os.path.basename(pdf_path)}: {e}")
                continue
            rows.append(row)
            shown = "  |  ".join(f"{l}: {row[l] or '-'}" for l in args.label)
            print(f"  [{idx}/{len(pdf_files)}] {shown}  |  Signer: {row['Signer Name'] or '-'}  —  {row['File Name']}")
        if not rows:
            print("\nNothing extracted. CSV not written.")
            sys.exit(0)
        pd.DataFrame(rows).to_csv(output_csv, index=False, encoding="utf-8-sig")
        missing = sum(1 for r in rows if not all(r[l] for l in args.label) or r["Signed"] == "No")
        print(f"\nDone. {len(rows)} file(s) saved to: {output_csv}")
        if missing:
            hint = "" if args.ocr else " (try --ocr for image-only boxes/signatures)"
            print(f"  {missing} file(s) are missing a value or signature{hint}.")
        return

    all_records = []
    wide_rows = []
    no_field_files = 0
    for idx, pdf_path in enumerate(pdf_files, start=1):
        try:
            records = extract_form_fields(pdf_path, args.include_empty, args.ocr, args.dpi)
        except Exception as e:
            print(f"  [{idx}/{len(pdf_files)}] failed — {os.path.basename(pdf_path)}: {e}")
            continue
        if not records:
            no_field_files += 1
        all_records.extend(records)
        wide_rows.append(records_to_wide_row(pdf_path, records))
        print(f"  [{idx}/{len(pdf_files)}] {len(records)} field(s)  —  {os.path.basename(pdf_path)}")

    if not all_records:
        print("\nNo form field values or signatures found. CSV not written.")
        sys.exit(0)

    if args.long:
        pd.DataFrame(all_records).to_csv(output_csv, index=False, encoding="utf-8-sig")
        print(f"\nDone. {len(all_records)} row(s) from {len(pdf_files)} file(s) saved to: {output_csv}")
    else:
        pd.DataFrame(wide_rows).to_csv(output_csv, index=False, encoding="utf-8-sig")
        print(f"\nDone. {len(wide_rows)} file(s), one row each, saved to: {output_csv}")
    sig_rows = [r for r in all_records if r["Signed"]]
    if sig_rows:
        signed = sum(1 for r in sig_rows if r["Signed"] == "Yes")
        by_source = {}
        for r in sig_rows:
            by_source[r["Source"]] = by_source.get(r["Source"], 0) + 1
        detail = ", ".join(f"{k}: {v}" for k, v in by_source.items())
        print(f"  Signatures: {len(sig_rows)}  (signed: {signed}, unsigned: {len(sig_rows) - signed})  —  {detail}")
    if no_field_files:
        print(f"  {no_field_files} file(s) had no form fields or signatures.")


if __name__ == "__main__":
    main()
