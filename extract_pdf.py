r"""
extract_pdf.py — Dynamically extract data from PDF files into Excel.

Detects headers automatically (by font size, bold, ALL-CAPS, or "Label:" patterns)
and extracts the content beneath each header.  Also pulls any embedded tables.

Output — Excel workbook with two sheets:
  "Sections"  — Source File | Page | Header | Value  (one row per key-value pair)
  "Tables"    — Source File | Page | Table  | (columns discovered from the table)

USAGE
    python extract_pdf.py "C:\path\to\folder"
    python extract_pdf.py "C:\path\to\file.pdf"
    python extract_pdf.py "C:\path\to\folder" --recursive
    python extract_pdf.py                              # prompts for path

DEPENDENCIES
    pip install pdfplumber openpyxl
    (pdfplumber pulls in pdfminer.six automatically)
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
from pathlib import Path
from collections import defaultdict

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


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _clean(text: str) -> str:
    return " ".join(str(text or "").split()).strip()


def _norm(text: str) -> str:
    return _clean(text).lower()


def _is_mostly_upper(text: str) -> bool:
    letters = [c for c in text if c.isalpha()]
    if len(letters) < 3:
        return False
    return sum(1 for c in letters if c.isupper()) / len(letters) >= 0.75


def _looks_like_label(text: str) -> bool:
    """True for short lines that look like field labels."""
    stripped = text.strip().rstrip(":")
    if len(stripped.split()) > 8:
        return False
    if text.strip().endswith(":"):
        return True
    if _is_mostly_upper(stripped) and len(stripped.split()) <= 5:
        return True
    return False


# ---------------------------------------------------------------------------
# Header detection
# ---------------------------------------------------------------------------

_MEDIAN_SIZE_FALLBACK = 10.0   # used when font metadata is absent


def _median(values: list[float]) -> float:
    if not values:
        return _MEDIAN_SIZE_FALLBACK
    sorted_v = sorted(values)
    mid = len(sorted_v) // 2
    return sorted_v[mid]


def _word_font_sizes(words: list[dict]) -> list[float]:
    sizes = []
    for w in words:
        try:
            s = float(w.get("size") or 0)
            if s > 0:
                sizes.append(s)
        except (TypeError, ValueError):
            pass
    return sizes


def _is_bold(word: dict) -> bool:
    font_name = str(word.get("fontname") or "").lower()
    return "bold" in font_name or "bd" in font_name


def _word_is_header(word: dict, median_size: float) -> bool:
    try:
        size = float(word.get("size") or 0)
    except (TypeError, ValueError):
        size = 0.0
    if size > median_size * 1.15:
        return True
    if _is_bold(word):
        return True
    return False


# ---------------------------------------------------------------------------
# Per-page extraction
# ---------------------------------------------------------------------------

def _word_in_bbox(word: dict, bboxes: list[tuple]) -> bool:
    """True if the word's centre point falls inside any of the given bounding boxes."""
    try:
        wx = (float(word["x0"]) + float(word["x1"])) / 2
        wy = (float(word["top"]) + float(word["bottom"])) / 2
    except (KeyError, TypeError, ValueError):
        return False
    for (x0, top, x1, bottom) in bboxes:
        if x0 <= wx <= x1 and top <= wy <= bottom:
            return True
    return False


def _extract_page_sections(page, table_bboxes: list[tuple]) -> list[dict]:
    """
    Return a list of {"header": str, "value": str} dicts from one page.

    Words that fall inside a detected table bounding box are skipped so that
    table content never leaks into the Sections sheet.

    Strategy A: use word-level font metadata to separate headers from body text.
    Strategy B (fallback): line-by-line heuristics when font metadata is absent.
    """
    all_words = page.extract_words(extra_attrs=["fontname", "size"])
    # Exclude words that belong to a table region
    words = [w for w in all_words if not _word_in_bbox(w, table_bboxes)]

    sizes = _word_font_sizes(words)
    median_size = _median(sizes) if sizes else _MEDIAN_SIZE_FALLBACK

    # Group words into lines by their top-y coordinate (within 2pt tolerance)
    lines_by_y: dict[int, list[dict]] = defaultdict(list)
    for w in words:
        y_bucket = round(float(w.get("top") or 0) / 2) * 2
        lines_by_y[y_bucket].append(w)

    ordered_lines = [lines_by_y[y] for y in sorted(lines_by_y)]

    sections: list[dict] = []
    current_header = ""
    current_value_parts: list[str] = []

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

        # Check if the majority of words on this line are header-like
        header_word_count = sum(1 for w in line_words if _word_is_header(w, median_size))
        line_is_header = (
            len(line_words) > 0
            and header_word_count / len(line_words) >= 0.6
        )

        # Fallback for lines with no font metadata
        if not line_is_header and not any(w.get("size") for w in line_words):
            line_is_header = _looks_like_label(line_text)

        if line_is_header:
            flush()
            current_header = line_text.rstrip(":")
        elif ":" in line_text and not current_header:
            label, _, value = line_text.partition(":")
            if _looks_like_label(label):
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


def _is_header_row(row: list[str]) -> bool:
    """
    Heuristic: a row is a header row if most non-empty cells look like labels
    (short, no digits, title-case or ALL-CAPS).
    """
    non_empty = [c for c in row if c.strip()]
    if not non_empty:
        return False
    label_count = sum(
        1 for c in non_empty
        if len(c.split()) <= 5 and not re.search(r'\d{4,}', c)
        and (c[0].isupper() or _is_mostly_upper(c))
    )
    return label_count / len(non_empty) >= 0.6


def _extract_page_tables(page) -> tuple[list[tuple], list[dict]]:
    """
    Detect tables on the page using pdfplumber's structural finder.

    Returns:
        bboxes  — list of (x0, top, x1, bottom) for each detected table
        records — list of {"headers": [...], "rows": [[...], ...]}

    Header-row detection:
      1. If the first row looks like labels → use it as the header.
      2. Otherwise generate Column_1, Column_2, … as synthetic headers.
    """
    bboxes: list[tuple] = []
    records: list[dict] = []

    try:
        found_tables = page.find_tables()
    except Exception:
        found_tables = []

    for tbl_obj in (found_tables or []):
        # Bounding box for masking text extraction
        try:
            bb = tbl_obj.bbox          # (x0, top, x1, bottom)
            bboxes.append(bb)
        except AttributeError:
            pass

        # Extract cell data
        try:
            raw = tbl_obj.extract()
        except Exception:
            continue
        if not raw:
            continue

        # Clean every cell
        cleaned_rows = []
        for row in raw:
            clean_row = [_clean(str(cell or "")) for cell in row]
            if any(clean_row):
                cleaned_rows.append(clean_row)
        if not cleaned_rows:
            continue

        # Decide header row
        if len(cleaned_rows) >= 2 and _is_header_row(cleaned_rows[0]):
            headers = cleaned_rows[0]
            data_rows = cleaned_rows[1:]
        else:
            # Synthesise generic column names
            n_cols = max(len(r) for r in cleaned_rows)
            headers = [f"Column_{i+1}" for i in range(n_cols)]
            data_rows = cleaned_rows

        # Pad short rows to match header width
        n = len(headers)
        data_rows = [r + [""] * (n - len(r)) for r in data_rows]

        records.append({"headers": headers, "rows": data_rows})

    return bboxes, records


# ---------------------------------------------------------------------------
# AcroForm field extraction — with spatial label association
# ---------------------------------------------------------------------------

_FT_LABELS = {
    "/Tx":  "Text",
    "/Ch":  "Dropdown / List",
    "/Btn": "Checkbox / Radio",
    "/Sig": "Signature",
}


def _normalise_value(raw, type_label: str) -> str:
    if raw is None:
        return "No" if type_label == "Checkbox / Radio" else ""
    s = str(raw).strip()
    if type_label == "Checkbox / Radio":
        # PDF spec: unchecked = /Off (or empty); anything else = checked
        off_values = {"", "off", "/off", "no", "/no", "false", "/false"}
        return "No" if s.lower() in off_values else "Yes"
    if s.startswith("/"):
        s = s[1:]
    return _clean(s)


def _resolve(obj):
    return obj.get_object() if hasattr(obj, "get_object") else obj


def _page_idnum_map(reader) -> dict[int, int]:
    """Build {pdf_object_id → 1-based page number} for fast /P lookups."""
    result = {}
    for i, page in enumerate(reader.pages, start=1):
        try:
            result[page.indirect_reference.idnum] = i
        except Exception:
            pass
    return result


def _get_widget_annotations(reader) -> list[dict]:
    """
    Collect every Widget annotation (one per visible form field) from the PDF.

    Primary strategy — AcroForm field tree:
        Traverses /AcroForm/Fields recursively, inheriting /FT, /T, /TU, /V
        from parent nodes.  This finds fields in nested groups (repeating
        sections, sub-forms) that are NOT listed in per-page /Annots.

    Fallback strategy — per-page /Annots:
        Used when the AcroForm root is absent or yields nothing.

    Returns list of dicts:
        page_num  — 1-based
        rect      — [llx, lly, urx, ury] PDF native coords (y=0 at bottom)
        field_type, field_name, alt_name, raw_value
    """
    page_map = _page_idnum_map(reader)
    widgets: list[dict] = []

    # ------------------------------------------------------------------ #
    # Strategy 1 — AcroForm field tree                                    #
    # ------------------------------------------------------------------ #
    def _traverse(field_ref, inherited: dict):
        try:
            field = _resolve(field_ref)
        except Exception:
            return

        # Merge inheritable attributes downward
        attrs = dict(inherited)
        for key in ("/FT", "/TU", "/V", "/DV"):
            try:
                val = field.get(key)
                if val is not None:
                    attrs[key] = val
            except Exception:
                pass
        # /T is NOT inherited — it is always the leaf name
        try:
            t_val = field.get("/T")
            if t_val is not None:
                attrs["_T"] = _clean(str(t_val))
        except Exception:
            pass

        # Recurse into child nodes first
        try:
            kids = field.get("/Kids")
            if kids:
                for kid_ref in kids:
                    _traverse(kid_ref, attrs)
        except Exception:
            pass

        # If this node has a /Rect it is (also) a widget
        try:
            rect = field.get("/Rect")
            if rect is None:
                return
            rect_vals = [float(x) for x in _resolve(rect)]
        except Exception:
            return

        # Resolve page number via /P → object id
        page_num = None
        try:
            p_ref = field.get("/P")
            if p_ref is not None:
                page_num = page_map.get(_resolve(p_ref).indirect_reference.idnum)
        except Exception:
            pass
        if page_num is None:
            return

        widgets.append({
            "page_num":   page_num,
            "rect":       rect_vals,
            "field_type": str(attrs.get("/FT") or ""),
            "field_name": attrs.get("_T") or "",
            "alt_name":   _clean(str(attrs.get("/TU") or "")),
            "raw_value":  attrs.get("/V"),
        })

    try:
        root     = _resolve(reader.trailer["/Root"])
        acroform = _resolve(root["/AcroForm"])
        fields   = acroform.get("/Fields") or []
        for f_ref in fields:
            _traverse(f_ref, {})
    except Exception:
        pass

    if widgets:
        return widgets

    # ------------------------------------------------------------------ #
    # Strategy 2 — per-page /Annots fallback                              #
    # ------------------------------------------------------------------ #
    for page_num, page in enumerate(reader.pages, start=1):
        try:
            annots_obj = page.get("/Annots")
            if annots_obj is None:
                continue
            annots_obj = _resolve(annots_obj)
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
                    return _inh(key, _resolve(node["/Parent"]))
                except Exception:
                    return None

            try:
                rect_vals = [float(x) for x in rect]
            except Exception:
                continue

            widgets.append({
                "page_num":   page_num,
                "rect":       rect_vals,
                "field_type": str(_inh("/FT") or ""),
                "field_name": _clean(str(_inh("/T") or "")),
                "alt_name":   _clean(str(_inh("/TU") or "")),
                "raw_value":  _inh("/V"),
            })

    return widgets


def _text_lines_by_page(pdf_path: Path) -> dict[int, list[dict]]:
    """
    Use pdfplumber to extract text lines (grouped words) per page.
    Each line: {text, x0, x1, top, bottom}  — all in pdfplumber coords (y=0 at top).
    """
    result: dict[int, list[dict]] = {}
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            lines_by_y: dict[int, list] = defaultdict(list)
            for w in words:
                y_key = round(float(w.get("top", 0)) / 3) * 3
                lines_by_y[y_key].append(w)

            lines = []
            for y_key in sorted(lines_by_y):
                grp = sorted(lines_by_y[y_key], key=lambda w: float(w.get("x0", 0)))
                text = _clean(" ".join(w.get("text", "") for w in grp))
                if text:
                    lines.append({
                        "text":   text,
                        "x0":     float(min(w.get("x0", 0) for w in grp)),
                        "x1":     float(max(w.get("x1", 0) for w in grp)),
                        "top":    float(min(w.get("top", 0) for w in grp)),
                        "bottom": float(max(w.get("bottom", 0) for w in grp)),
                    })
            result[page_num] = lines
    return result


def _find_label(field_rect_plumber: tuple, page_lines: list[dict]) -> str:
    """
    Given a field's bounding box in pdfplumber coords (x0, top, x1, bottom),
    return the text of the nearest label to the LEFT or ABOVE the field.

    Search order / priority:
      1. Same-row text immediately to the LEFT  (tight vertical band ±14 pt)
      2. Text directly ABOVE, horizontally overlapping the field
      3. Column-header text anywhere above, whose horizontal centre is within
         the field's x-span (catches leftmost fields in a table-like layout
         where the label is in a header row above, not beside the field)
    """
    fx0, ftop, fx1, fbottom = field_rect_plumber
    fy_mid  = (ftop + fbottom) / 2
    fx_mid  = (fx0 + fx1) / 2
    fwidth  = fx1 - fx0

    candidates: list[tuple[float, str]] = []

    for line in page_lines:
        lx0, lx1 = line["x0"], line["x1"]
        ltop, lbottom = line["top"], line["bottom"]
        ly_mid  = (ltop + lbottom) / 2
        lx_mid  = (lx0 + lx1) / 2
        text    = line["text"].rstrip(":").strip()
        if not text:
            continue

        # 1. Same-row label to the LEFT
        if lx1 <= fx0 + 5 and abs(ly_mid - fy_mid) <= 14:
            dist = (fx0 - lx1) + abs(ly_mid - fy_mid) * 0.5
            candidates.append((dist, text))

        # 2. Text directly ABOVE, horizontally overlapping
        elif lbottom <= ftop + 4 and lx0 < fx1 + 10 and lx1 > fx0 - 10:
            dist = (ftop - lbottom) + abs(lx_mid - fx_mid) * 0.3
            candidates.append((dist, text))

        # 3. Column-header above: centre of label within field's x-span (±half field width)
        elif lbottom <= ftop + 4 and abs(lx_mid - fx_mid) <= fwidth * 0.9:
            dist = (ftop - lbottom) * 1.5 + abs(lx_mid - fx_mid)
            candidates.append((dist, text))

    if not candidates:
        return ""
    candidates.sort(key=lambda c: c[0])
    return candidates[0][1]


def extract_form_fields(pdf_path: Path) -> list[dict]:
    """
    Extract AcroForm fields from a fillable PDF.

    For each field widget, the visible label text nearest to it on the page
    is used as the field name — so output matches what the user sees on the form
    rather than internal technical field names.

    Each record: {"source", "page", "field_name", "field_type", "value"}
    """
    if not HAS_PYPDF:
        return []

    source = pdf_path.name

    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return []

    widgets = _get_widget_annotations(reader)
    if not widgets:
        return []

    # Page heights needed for coordinate conversion (PDF native → pdfplumber)
    page_heights: dict[int, float] = {}
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for i, p in enumerate(pdf.pages, start=1):
                page_heights[i] = float(p.height)
    except Exception:
        pass

    # Text lines per page for label association
    try:
        lines_by_page = _text_lines_by_page(pdf_path)
    except Exception:
        lines_by_page = {}

    results: list[dict] = []

    for w in widgets:
        page_num   = w["page_num"]
        type_label = _FT_LABELS.get(w["field_type"], "Field")
        value      = _normalise_value(w["raw_value"], type_label)

        # Convert PDF native rect → pdfplumber coords
        ph = page_heights.get(page_num, 792.0)
        llx, lly, urx, ury = w["rect"]
        field_plumber = (llx, ph - ury, urx, ph - lly)  # (x0, top, x1, bottom)

        # Try spatial label lookup first, fall back to /TU then /T
        label = _find_label(field_plumber, lines_by_page.get(page_num, []))
        if not label:
            label = w["alt_name"] or w["field_name"]

        results.append({
            "source":     source,
            "page":       str(page_num),
            "field_name": label,
            "field_type": type_label,
            "value":      value,
        })

    return results


# ---------------------------------------------------------------------------
# Full PDF extraction
# ---------------------------------------------------------------------------

def extract_pdf(pdf_path: Path) -> tuple[list[dict], list[dict], list[dict]]:
    """
    Returns:
        sections     — list of {"source", "page", "header", "value"}
        tables       — list of {"source", "page", "table_num", "headers", "rows"}
        form_fields  — list of {"source", "page", "field_name", "field_type", "value"}
    """
    source = pdf_path.name
    all_sections: list[dict] = []
    all_tables:   list[dict] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            # 1. Detect tables first so we know which areas to mask from text
            bboxes, tbl_records = _extract_page_tables(page)

            # 2. Extract sections, excluding table regions
            page_sections = _extract_page_sections(page, bboxes)
            for sec in page_sections:
                all_sections.append({
                    "source": source,
                    "page":   page_num,
                    "header": sec["header"],
                    "value":  sec["value"],
                })

            # 3. Store table records
            for t_idx, rec in enumerate(tbl_records, start=1):
                all_tables.append({
                    "source":    source,
                    "page":      page_num,
                    "table_num": t_idx,
                    "headers":   rec["headers"],
                    "rows":      rec["rows"],
                })

    # 4. AcroForm fields (fillable form values — dropdown, text inputs, checkboxes)
    form_fields = extract_form_fields(pdf_path)

    return all_sections, all_tables, form_fields


# ---------------------------------------------------------------------------
# Excel helpers
# ---------------------------------------------------------------------------

_HEADER_FILL   = PatternFill("solid", fgColor="1F3864")
_HEADER_FONT   = Font(bold=True, color="FFFFFF", size=11)
_ALT_FILL      = PatternFill("solid", fgColor="DCE6F1")
_SECTION_FONT  = Font(bold=True, size=10)

def _style_header_row(ws, row_idx: int, col_count: int) -> None:
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font   = _HEADER_FONT
        cell.fill   = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _autofit(ws) -> None:
    for col in ws.columns:
        max_len = max((len(str(cell.value or "")) for cell in col), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 80)


def _csv_safe(value) -> str:
    s = str(value or "")
    if s[:1] in ("=", "+", "-", "@"):
        return "'" + s
    return s


# ---------------------------------------------------------------------------
# Workbook writer
# ---------------------------------------------------------------------------

OUTPUT_COLS = ["Source File", "Page", "Category", "Field Name", "Value"]


def _build_unified_rows(
    all_sections: list[dict],
    all_tables: list[dict],
    all_form_fields: list[dict],
) -> list[list]:
    """
    Merge all three data sources into a single list of rows, ordered by page.

    Category values:
      "Form Field"  — AcroForm fillable field (text input, dropdown, checkbox)
      "Section"     — header/label with text content found in the document body
      "Table"       — one row per table data row; Value = "Col: Val | Col: Val …"

    For pages that already have Form Field data, blank-value Section rows are
    suppressed — they are just the static label text that was already captured
    as the Field Name of the corresponding Form Field row.
    """
    rows: list[tuple] = []   # (sort_key, priority, data_row)

    # Build set of (source, page) pairs that have form field data
    form_field_pages: set[tuple] = {
        (ff.get("source", ""), str(ff.get("page", "")))
        for ff in all_form_fields
    }

    for ff in all_form_fields:
        try:
            page_key = int(ff.get("page") or 0)
        except (ValueError, TypeError):
            page_key = 0
        data = [
            _csv_safe(ff.get("source", "")),
            _csv_safe(ff.get("page", "")),
            "Form Field",
            _csv_safe(ff.get("field_name", "")),
            _csv_safe(ff.get("value", "")),
        ]
        rows.append((page_key, 0, data))

    for sec in all_sections:
        page_key = int(sec.get("page") or 0)
        value    = sec.get("value", "")
        source   = sec.get("source", "")
        page_str = str(sec.get("page", ""))

        # Skip blank-value sections on pages that have form fields
        if not value and (source, page_str) in form_field_pages:
            continue

        data = [
            _csv_safe(source),
            _csv_safe(page_str),
            "Section",
            _csv_safe(sec.get("header", "")),
            _csv_safe(value),
        ]
        rows.append((page_key, 1, data))

    for tbl in all_tables:
        page_key  = int(tbl.get("page") or 0)
        source    = tbl.get("source", "")
        page      = tbl.get("page", "")
        headers   = tbl.get("headers", [])
        for r_idx, row in enumerate(tbl.get("rows", []), start=1):
            # Represent each table row as "Col: Val | Col: Val …"
            pairs = " | ".join(
                f"{h}: {v}" for h, v in zip(headers, row) if h or v
            )
            label = f"Row {r_idx} (Table {tbl.get('table_num', '')})"
            data = [
                _csv_safe(source),
                _csv_safe(page),
                "Table",
                label,
                _csv_safe(pairs),
            ]
            rows.append((page_key, 2, data))

    rows.sort(key=lambda x: (x[0], x[1]))
    return [r[2] for r in rows]


def write_workbook(
    out_path: Path,
    all_sections: list[dict],
    all_tables: list[dict],
    all_form_fields: list[dict],
) -> None:
    tmp_path = out_path.with_suffix(".tmp.xlsx")
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Extracted Data"
    ws.append(OUTPUT_COLS)
    ws.row_dimensions[1].height = 18
    _style_header_row(ws, 1, len(OUTPUT_COLS))

    unified = _build_unified_rows(all_sections, all_tables, all_form_fields)
    for i, row in enumerate(unified, start=2):
        ws.append(row)
        if i % 2 == 0:
            for col in range(1, len(OUTPUT_COLS) + 1):
                ws.cell(row=i, column=col).fill = _ALT_FILL

    if not unified:
        ws.cell(row=2, column=1, value="No data extracted from the processed PDF(s).")

    _autofit(ws)

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

def _iter_pdf(target: Path, recursive: bool):
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
        description="Dynamically extract headers and data from PDF files into Excel."
    )
    parser.add_argument(
        "target", nargs="?",
        help="PDF file or folder of PDF files (prompted if omitted)"
    )
    parser.add_argument(
        "--recursive", action="store_true",
        help="Recurse into subfolders when target is a folder"
    )
    parser.add_argument(
        "--output", default=None,
        help="Custom output .xlsx path (default: next to target)"
    )
    args = parser.parse_args()

    target_str = args.target
    if not target_str:
        target_str = input("Enter path to PDF file or folder: ").strip().strip('"').strip("'")
    if not target_str:
        sys.exit("ERROR: no path provided.")

    if not HAS_PDFPLUMBER:
        sys.exit("ERROR: pdfplumber is required.  Run:  pip install pdfplumber")
    if not HAS_OPENPYXL:
        sys.exit("ERROR: openpyxl is required.   Run:  pip install openpyxl")
    if not HAS_PYPDF:
        print("WARNING: pypdf not installed — form field values will be skipped.\n"
              "         Run:  pip install pypdf")

    target = Path(target_str)
    if not target.exists():
        sys.exit(f"ERROR: path not found — {target}")

    files = list(_iter_pdf(target, args.recursive))
    if not files:
        sys.exit(f"ERROR: no .pdf files found at {target}")

    print(f"Found {len(files)} PDF file(s).  Extracting...")

    all_sections:    list[dict] = []
    all_tables:      list[dict] = []
    all_form_fields: list[dict] = []

    for pdf_path in files:
        print(f"  Processing: {pdf_path.name}")
        try:
            sections, tables, form_fields = extract_pdf(pdf_path)
            all_sections.extend(sections)
            all_tables.extend(tables)
            all_form_fields.extend(form_fields)
            print(f"    {len(sections)} section(s), {len(tables)} table(s), "
                  f"{len(form_fields)} form field(s) found.")
        except Exception as exc:
            print(f"    ERROR: {exc}")

    if args.output:
        out_path = Path(args.output)
    else:
        output_folder = target if target.is_dir() else target.parent
        out_stem      = target.stem if target.is_file() else target.name
        out_path      = output_folder / f"{out_stem}_extracted.xlsx"

    write_workbook(out_path, all_sections, all_tables, all_form_fields)
    table_rows = sum(len(t['rows']) for t in all_tables)
    print(f"\nWrote: {out_path}")
    print(f"  Form fields : {len(all_form_fields)} field(s)")
    print(f"  Sections    : {len(all_sections)} row(s)")
    print(f"  Tables      : {table_rows} data row(s) across {len(all_tables)} table(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
