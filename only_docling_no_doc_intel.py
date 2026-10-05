"""
only_docling.py - table extraction from PDF documents into Excel workbooks
(one worksheet per detected table), using Docling's local "standard"
pipeline (OCR + layout analysis + the TableFormer table-structure model)
instead of Azure Document Intelligence.

This is a trimmed-down variant of ocr_docling_test.py: same Docling table
extraction, same reliability/confidence gates, same continuation-merging and
stacked-header reconstruction, same xlsx output format -- but with NO GLiNER2
PII gate at all. Docling's own OCR text is used only to find and parse
tables; nothing here filters which tables are kept based on their content.
Every table that clears the reliability/confidence gates below is written to
the output workbook. Docling itself runs entirely locally: no Azure
Document Intelligence call, no network dependency at all once Docling's
models are cached on first run.

There is no Azure Document Intelligence (or any other) fallback here at
all -- Docling's own local extraction is the only engine this script ever
runs, and it never makes a network call or incurs any cloud cost once its
models are cached (see --warm-models below). A file Docling itself flags as
low-confidence or finds no reliable table in is not deferred to any other
engine; a size-verified COPY of the native file is simply placed in
"manual review/" for a human to look at, and the native file itself is left
exactly where it was (see the "manual review/" entry under Output layout
below). This script is meant to run entirely on its own -- it does not
shell out to, or expect to be driven by, triage_pipeline.py, and it always
writes its own metrics workbook (see metrics_docling.xlsx below) rather
than relying on a caller to supply one.

What this script does:
    - Government/tax form exclusion (same marker regex, same "excludes the
      whole PDF" policy as doc_reader_v2.py).
    - The table reliability gate: empty-cell ratio, per-row sparsity,
      "ambiguous narrow/headerless table" check, and page-text coverage
      ratio (see _table_reliable).
    - The "Reviewer Note" sheet for pages whose OCR text isn't substantially
      captured inside any extracted table.
    - The one-workbook-per-PDF, one-sheet-per-table xlsx writer, output
      folder layout, and copy-verify-then-delete file relocation.
    - Continuation-table merging across pages (see
      _merge_continuation_tables): a page whose table has no Docling-declared
      header row and whose column count matches the immediately preceding
      accepted table is folded into it as a continuation, rather than
      written out as its own separate, header-less sheet.
    - Stacked-header/split-row reconstruction (see
      _unstack_docling_header_table): when TableFormer fuses several real,
      narrow columns into one detected grid column, each fused column's
      header still visually stacks the real columns' labels, and each
      record's row still visually spans several print-lines. For a
      born-digital PDF (native_text -- see _pdf_has_native_text), each header
      TableCell's own bbox is handed to pypdfium2's own text layer (see
      _PdfHeaderLineReader) to recover the PDF's true line breaks and
      reconstruct the real header split.
    - Docling's own per-page confidence scores (result.confidence.pages[n],
      see _read_page_confidence): a page whose OCR confidence is low gets
      flagged on the Reviewer Note ("low Docling confidence -- verify
      manually"). table_score is NOT used the same way, despite being part
      of the same confidence report -- no model in the installed Docling
      package ever actually writes it (see _read_page_confidence's own
      docstring), so there is no working table-confidence gate here.
      layout_score (the one per-page signal that IS populated) is reported
      in the Metrics sheet as a labeled proxy instead, for a human to read,
      not acted on by any rejection logic.

What this script deliberately does NOT do:
    - No GLiNER2 (or any other) PII/entity gate -- every reliable table is
      kept and written out, unfiltered, regardless of its content.
    - No LLM or Document Intelligence fallback of any kind -- Docling's
      local extraction above never calls out to any other engine, model, or
      cloud service. A file that fails the benchmarks below is simply
      copied to "manual review/" for a human, not handed off elsewhere.
    - No label:value entity pivoting/orientation detection or "staircase"
      artifact detection (doc_reader_v2.py subsystems hand-tuned to specific
      Document Intelligence Layout failure modes that don't apply here). A
      page whose table doesn't cleanly fit Docling's own header/data grid
      (e.g. a label:value form) is instead just rejected by the reliability
      gate above and left for the Reviewer Note to flag.

Output layout (inside --output-folder, default 'docling_output' next to the
input):
    tabular data/      one "<name>_intermediate.xlsx" per PDF where at least
                        one reliable table was detected (one worksheet per
                        such table), plus a COPY of the native PDF itself
                        placed alongside it under its original name (see
                        _copy_verified) -- e.g. "test1.pdf" in produces both
                        "test1.pdf" and "test1_intermediate.xlsx" here. The
                        "_intermediate" tag marks the xlsx as Docling's own
                        parse, not a final reviewed deliverable. The source
                        PDF at its original input location is left in place,
                        untouched, same as the copy placed here. If a file of
                        that name already exists here (a prior run on the
                        same input), the new one is written alongside it as
                        "<name>_intermediate_1.xlsx"/"<name>_1.pdf", etc.
                        instead of overwriting it.
    government forms/  PDFs identified as a government/tax form -- excluded
                        outright, no extraction attempted; relocated the
                        same copy-verify-delete way. This is a policy
                        exclusion, not a quality/confidence issue.
    manual review/      PDFs that either failed one of the benchmark gates
                        below or hit a genuine, unresolved processing
                        failure. Unlike "tabular data/"/"government forms/",
                        this is copy-only (see _copy_verified) -- the native
                        file is always left exactly where it was, and a
                        size-verified COPY is placed here so the outcome
                        isn't silently invisible. Covers three cases, each
                        distinguished in the metrics workbook's Status/Notes
                        (see METRICS_COLUMNS):
                          (1) No reliable table was found anywhere in the
                              document (Status "Manual Review (Not
                              Tabular)").
                          (2) A page that contributed an accepted table also
                              had low Docling OCR confidence -- same
                              MIN_OCR_CONFIDENCE_SCORE threshold used for the
                              Reviewer Note (Status "Manual Review (Low
                              Confidence)").
                          (3) A genuine, unresolved processing failure, e.g.
                              Docling's own conversion raised an exception
                              (a corrupt/unreadable PDF) (Status "Error").
                        There is no fallback engine to re-verify a rejected
                        table against, and no rerun mechanism of its own --
                        every one of these outcomes needs a human to look at
                        the native file directly.
    metrics_docling.xlsx  persistent run log for this output folder (or
                        wherever --metrics-path points). Runs are appended,
                        never overwritten, across invocations. Sheets:
                          "Docling Metrics" (--metrics-sheet-name) -- one row
                          per file processed this run. Its own "TableFormer
                          Check" column (Passed / Failed Reliability Check /
                          No Table Detected / N/A (Excluded) -- see
                          process_pdf's own table_check_status) tells apart a
                          structural table-detection failure from a low-OCR-
                          confidence outcome -- two independent reasons a
                          file can end up with 0 tables found; when it reads
                          "Failed Reliability Check", Notes names the exact
                          check (and page) that failed.
                          "Low Confidence (Not Parsed)" -- one row per file
                          copied to manual review/ above for low OCR
                          confidence specifically, so those documents are
                          easy to find/filter on their own.

Handling note: output workbooks and relocated PDFs may contain real PII/PHI.
Keep them in approved secure storage only; do not commit them to source
control.

Usage:
    python only_docling.py
        (no arguments: opens a file-picker dialog to select one or more PDFs)
    python only_docling.py path/to/file.pdf
    python only_docling.py path/to/folder_of_pdfs
    python only_docling.py path/to/folder_of_pdfs --output-folder ./docling_output
    python only_docling.py path/to/file.pdf --debug

Setup:
    pip install docling
    Model downloads are NOT done implicitly on first use -- HF_HUB_OFFLINE/
    TRANSFORMERS_OFFLINE are forced on below (see the top of this file) so a
    missing/uncached model fails loudly at startup instead of this script
    silently reaching Hugging Face/EasyOCR's GitHub releases mid-run, after
    a PDF's PHI is already in memory. Run this once, on a machine with no
    PHI present, before processing anything for real:
        python only_docling.py --warm-models
    That downloads Docling's layout + TableFormer + OCR (EasyOCR, English)
    models into Docling's local cache; every later run against real
    documents is then fully offline.

    No other setup is needed -- this script never calls out to Azure
    Document Intelligence, Azure OpenAI, or any other engine, and does not
    need any Document Intelligence credentials/.env to be present.
"""
from __future__ import annotations

import argparse
import contextlib
import json
import math
import os
import re
import shutil
import sys
import time
from pathlib import Path
from typing import List, Optional, Tuple

# Table content transcribed from OCR/native PDF text can contain non-ASCII
# characters; Windows' default console codepage (cp1252) can't encode all of
# them, which would otherwise crash a print() with a UnicodeEncodeError.
# Force UTF-8 for stdout/stderr up front -- harmless on platforms where it
# was already UTF-8.
if sys.platform == "win32":
    with contextlib.suppress(Exception):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")

# Force offline model loading so a missing/uncached Docling model fails
# loudly at startup instead of this process silently reaching Hugging Face
# mid-run -- i.e. after a PDF's PHI is already in memory. huggingface_hub
# and transformers each read their *_OFFLINE env var once, at import time,
# so this has to happen before docling (which pulls in both transitively)
# is imported below. The one sanctioned exception is --warm-models (see
# main()), meant to be run once on a machine with no PHI present to
# populate the local model cache; argparse hasn't run yet at this point,
# so it's detected directly off sys.argv instead.
WARM_MODELS = "--warm-models" in sys.argv
if not WARM_MODELS:
    os.environ.setdefault("HF_HUB_OFFLINE", "1")
    os.environ.setdefault("TRANSFORMERS_OFFLINE", "1")
os.environ.setdefault("HF_HUB_DISABLE_TELEMETRY", "1")

try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None

import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import IllegalCharacterError
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill, Font, Alignment

import pypdfium2 as pdfium

from docling.document_converter import DocumentConverter, PdfFormatOption
from docling.datamodel.base_models import InputFormat, ConversionStatus
from docling.datamodel.pipeline_options import (
    PdfPipelineOptions,
    TableFormerMode,
    EasyOcrOptions,
)
from docling.datamodel.settings import settings as docling_settings
from docling.utils.model_downloader import download_models

# Docling's layout/object-detection engine defaults to torch.compile()-ing its
# model for a speed boost, which needs an MSVC C++ compiler (cl.exe) on
# Windows. That's not installed here, so leave models un-compiled -- eager
# mode is slower per-page but works everywhere.
docling_settings.inference.compile_torch_models = False

DOCUMENT_TYPE = "PDF"

# Docling's own default (1.0, roughly a 72 DPI-equivalent render) is the
# resolution its layout/TableFormer models see the page AT, regardless of
# do_ocr -- too low to reliably resolve column boundaries on a dense, many-
# narrow-column table (observed: a 23-column surgical schedule got collapsed
# to 11 columns with rows split across multiple grid rows at scale=1.0).
# Raised here across the board since higher resolution never makes structure
# detection worse, only slower.
DEFAULT_IMAGES_SCALE = 2.0

# Below this many extracted characters across a sample of a PDF's first few
# pages, treat it as a scanned/image PDF with no usable native text layer --
# see _pdf_has_native_text.
MIN_NATIVE_TEXT_CHARS = 20
ENGINE_LABEL = "Docling (local, standard pipeline, TableFormer)"


# ----------------- Government/tax form exclusion -----------------
# Pure text-marker matching, not tied to any Document-Intelligence-specific
# object, so it applies just as well to Docling's plain-text output. A hit
# on ANY page excludes the WHOLE PDF; this repo has dedicated per-form
# extractor scripts instead.
_GOVERNMENT_FORM_MARKERS = [
    r"internal revenue service",
    r"department of the treasury",
    r"omb (?:no\.?|number|control number)\s*[:\-]?\s*\d",
    r"for paperwork reduction act notice",
    r"privacy act and paperwork reduction act notice",
    r"\bwage and tax statement\b",
    r"\bform\s+(?:w-?2|w-?4|w-?9|1099[-a-z]*|940|941|944|1040[-a-z]*|ss-4|ss-8|i-9|8850|8832)\b",
    r"employer identification number",
    r"social security administration",
    r"u\.s\.\s*citizenship and immigration services",
    r"department of homeland security",
    r"\bdepartment of revenue\b",
    r"\bcat\.\s*no\.\s*\d",
]
_GOVERNMENT_FORM_RE = re.compile("|".join(_GOVERNMENT_FORM_MARKERS), re.IGNORECASE)


def _detect_government_form(page_texts: dict) -> Optional[tuple]:
    """Return (page_number, matched phrase) for the first page (in page
    order) whose text matches a strong government/tax-form signal, or None."""
    for page_number in sorted(page_texts):
        match = _GOVERNMENT_FORM_RE.search(page_texts[page_number])
        if match:
            return page_number, match.group(0)
    return None


# ----------------- Table reliability gate -----------------
MIN_RELIABLE_TABLE_COLS = 2       # a single-column "table" is usually a mis-detected text block
MAX_EMPTY_CELL_RATIO = 0.5        # more than half blank cells suggests a broken/misaligned grid
AMBIGUOUS_HEADERLESS_MAX_COLS = 3
MIN_AMBIGUOUS_FORM_ROWS = 9
MIN_TEXT_COVERAGE_RATIO = 0.4     # table(s) should account for a reasonable share of the page's text
                                  # -- skipped when every table on the page is otherwise
                                  # structurally sound (real header, has data rows); see _table_reliable
MAX_ROW_EMPTY_RATIO = 0.6         # a single row this empty looks like a split-off fragment
MAX_SPARSE_ROW_FRACTION = 0.3     # too many fragment-like rows means the row grid itself is broken
PARSE_COVERAGE_WARNING_THRESHOLD = 0.80

# Docling's own confidence report grades each page's OCR quality on a 0-1
# scale (see _read_page_confidence). Below this, the page is called out on
# the Reviewer Note regardless of whether a table was found on it at all.
# There is deliberately no equivalent MIN_TABLE_CONFIDENCE_SCORE gate on
# table_score, the report's other field -- no model in the installed
# Docling package ever actually writes it (see _read_page_confidence's own
# docstring), so a constant/gate for it would just be dead code with a
# threshold nothing ever compares against.
#
# Deliberately biased toward parsing over deferring to manual review: 0.65
# rejected plenty of real-world scans (skewed, low-DPI, table-heavy pages)
# whose OCR text was still substantially usable. 0.45 only rejects pages
# where EasyOCR's own confidence suggests genuinely garbled text -- expect
# more incorrect cell values to make it into the output workbook unflagged
# as a result; there's no downstream engine to catch them.
MIN_OCR_CONFIDENCE_SCORE = 0.45


def _table_char_count(table: dict) -> int:
    headers = table.get("headers") or []
    rows = table.get("rows") or []
    total = sum(len(str(h)) for h in headers if h is not None)
    for row in rows:
        if isinstance(row, list):
            total += sum(len(str(c)) for c in row if c is not None)
    return total


def _cached_char_count(table: dict) -> int:
    if "_char_count" not in table:
        table["_char_count"] = _table_char_count(table)
    return table["_char_count"]


def _table_reliable(
    tables: list,
    page_text: str,
    min_cols: int = MIN_RELIABLE_TABLE_COLS,
    max_empty_ratio: float = MAX_EMPTY_CELL_RATIO,
    min_coverage_ratio: float = MIN_TEXT_COVERAGE_RATIO,
    max_row_empty_ratio: float = MAX_ROW_EMPTY_RATIO,
    max_sparse_row_fraction: float = MAX_SPARSE_ROW_FRACTION,
    max_ambiguous_headerless_cols: int = AMBIGUOUS_HEADERLESS_MAX_COLS,
    min_ambiguous_form_rows: int = MIN_AMBIGUOUS_FORM_ROWS,
    debug: bool = False,
    debug_label: str = "",
) -> tuple:
    """(False, reason) if any Docling-detected table on this page looks too
    thin, empty, fragmented, form-like, or incomplete to trust as-is; the
    reason string names the specific check that failed and the measured
    value, e.g. "table#0 empty_ratio=0.62 > max_empty_ratio=0.50" -- used by
    process_pdf to record WHY a page's table(s) were rejected on the
    "Docling Metrics" sheet's Notes column (see its own "TableFormer Check"
    column), not just print it under --debug. There is no LLM fallback to
    re-verify a rejected table against here, so a page that fails this
    simply ends up on the Reviewer Note as uncaptured text instead of being
    retried. Returns (True, "") when every table on the page passes.

    No table_score-based confidence gate here -- Docling's own per-page
    confidence report has a table_score field, but no model in the
    installed package ever actually writes to it (see
    _read_page_confidence's own docstring); a gate against it would never
    fire, so there's deliberately no MIN_TABLE_CONFIDENCE_SCORE constant or
    branch here to imply one does.

    The page-text coverage-ratio check (see MIN_TEXT_COVERAGE_RATIO) is
    skipped when every table on the page already has a real declared
    header, at least one data row, and header labels that read like labels
    rather than data (see _looks_like_real_header) -- a table that's
    structurally sound shouldn't be discarded just because the rest of the
    page also happens to carry a lot of non-tabular narrative text."""
    def _reject(reason: str) -> tuple:
        if debug:
            print(f"  [debug]{debug_label} table rejected: {reason}", file=sys.stderr)
        return False, reason

    # Tracks whether EVERY table on this page has strong-enough structural
    # signals (real declared header, at least one data row, header labels
    # that read like labels rather than data -- see _looks_like_real_header)
    # to skip the page-text coverage-ratio check below entirely. That check
    # assumes a table capturing only a small share of the page's text means
    # a broken/fragmentary detection -- true for a genuinely mis-detected
    # table, but a false positive for a small, perfectly legitimate table
    # that happens to share a page with a lot of unrelated narrative text
    # (observed: a 3x6 table on a page that also had several paragraphs of
    # prose got coverage_ratio=0.13, well under the 0.4 default, and was
    # discarded outright despite having a clean header and no empty/sparse
    # rows). A table with weaker signals (no declared header, or a lone
    # header row with no data at all) still needs the coverage check, since
    # those are exactly the shapes a genuinely fragmentary detection
    # produces.
    all_structurally_sound = True

    for idx, table in enumerate(tables):
        headers = table.get("headers") or []
        rows = table.get("rows") or []
        tag = f" table#{idx}"

        if len(headers) < min_cols:
            return _reject(f"{tag} has {len(headers)} column(s), fewer than min_cols={min_cols}")

        has_declared_header = table.get("has_declared_header", True)
        is_narrow = len(headers) <= max_ambiguous_headerless_cols
        total_rows = 1 + len(rows)
        if is_narrow and (not has_declared_header or total_rows >= min_ambiguous_form_rows):
            return _reject(
                f"{tag} is narrow ({len(headers)} column(s)) and either has no Docling-declared "
                f"header row (has_declared_header={has_declared_header}) or has {total_rows} total "
                f"row(s) (>= min_ambiguous_form_rows={min_ambiguous_form_rows}) -- ambiguous "
                "form-vs-table shape"
            )

        if not has_declared_header and not is_narrow and len(rows) >= 2:
            return _reject(
                f"{tag} is wide ({len(headers)} column(s)) with no Docling-declared header row and "
                f"{len(rows)} data row(s) -- row 0 may be real data mistaken for column titles"
            )

        cells = list(headers) + [c for row in rows if isinstance(row, list) for c in row]
        if not cells:
            return _reject(f"{tag} has no cells at all")
        empty_ratio = sum(1 for c in cells if not str(c or "").strip()) / len(cells)
        if empty_ratio > max_empty_ratio:
            return _reject(f"{tag} empty_ratio={empty_ratio:.2f} > max_empty_ratio={max_empty_ratio}")

        data_rows = [row for row in rows if isinstance(row, list) and row]
        if data_rows:
            sparse_rows = sum(
                1 for row in data_rows
                if sum(1 for c in row if not str(c or "").strip()) / len(row) > max_row_empty_ratio
            )
            sparse_fraction = sparse_rows / len(data_rows)
            if sparse_fraction > max_sparse_row_fraction:
                return _reject(
                    f"{tag} sparse_row_fraction={sparse_fraction:.2f} "
                    f"> max_sparse_row_fraction={max_sparse_row_fraction}"
                )

        if not (has_declared_header and data_rows and _looks_like_real_header(headers)):
            all_structurally_sound = False

    text_len = len(page_text.strip())
    if text_len > 0 and not all_structurally_sound:
        captured_chars = sum(_cached_char_count(t) for t in tables)
        coverage_ratio = captured_chars / text_len
        if coverage_ratio < min_coverage_ratio:
            return _reject(f"coverage_ratio={coverage_ratio:.2f} < min_coverage_ratio={min_coverage_ratio}")

    if debug:
        print(f"  [debug]{debug_label} table(s) accepted as reliable", file=sys.stderr)
    return True, ""


def _find_pages_with_uncaptured_text(
    page_texts: dict,
    table_chars_by_page: dict,
    threshold: float = PARSE_COVERAGE_WARNING_THRESHOLD,
) -> list:
    """Return [(page_number, coverage_ratio), ...] for pages whose text is
    only partly accounted for by the table(s) extracted from that page.
    Takes an already-built {page_number: char count} dict rather than
    deriving one from a list of table dicts, since _merge_continuation_tables
    collapses a multi-page table's several per-page entries into one table
    dict carrying a page RANGE (e.g. "1-3") for display -- deriving
    table_chars_by_page from that merged list would credit the whole range's
    char count to a single dict key that never matches an individual
    page_number below. Building this dict during the per-page loop in
    process_pdf, before any merging happens, keeps each page's own
    captured-char count correct regardless of how tables get combined
    afterward for the workbook."""
    flagged = []
    for page_number in sorted(page_texts):
        text = page_texts[page_number]
        text_len = len(text.strip())
        captured_chars = table_chars_by_page.get(page_number, 0)
        if text_len == 0:
            if captured_chars == 0:
                flagged.append((page_number, None))
            continue
        ratio = min(captured_chars / text_len, 1.0)
        if ratio < threshold:
            flagged.append((page_number, ratio))
    return flagged


# ----------------- xlsx writer -----------------
def _style_header_row(ws: Worksheet, row: int = 1) -> None:
    fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    align = Alignment(horizontal="center", vertical="center")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


def _autosize_columns_from_rows(ws: Worksheet, rows: list, preserve_existing: bool = False) -> None:
    """Sizes ws's columns from row VALUES already in hand, not by reading
    them back off the sheet -- ws.columns/iter_cols() inflates a sparse
    sheet to its full dense size."""
    max_lens: dict = {}
    for row in rows:
        for idx, value in enumerate(row):
            if value is None:
                continue
            length = len(str(value))
            if length > max_lens.get(idx, 0):
                max_lens[idx] = length

    for idx, max_len in max_lens.items():
        col_letter = get_column_letter(idx + 1)
        width = min(max(max_len + 2, 8), 60)
        if preserve_existing:
            existing = ws.column_dimensions[col_letter].width
            if existing:
                width = max(width, existing)
        ws.column_dimensions[col_letter].width = width


def _safe_sheet_name(name: str) -> str:
    for bad in ["\\", "/", "*", "[", "]", ":", "?"]:
        name = name.replace(bad, "_")
    return name[:31]


def _normalize_row(row: list, width: int) -> list:
    row = list(row)
    if width <= 0:
        return row
    if len(row) < width:
        return row + [None] * (width - len(row))
    return row[:width]


def _sanitize_cell_value(value):
    """Neutralize CSV/Excel formula injection in transcribed cell content,
    and strip XML-illegal control characters. Table content is transcribed
    verbatim from an untrusted PDF, so a crafted cell like =HYPERLINK(...)
    would otherwise become a live formula the moment the workbook is
    opened."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
    if isinstance(value, str) and value and value[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


def _write_table_to_sheet(wb: openpyxl.Workbook, table: dict, sheet_name: str, source: str, page: int) -> None:
    headers = table.get("headers") or []
    rows = table.get("rows") or []

    ws = wb.create_sheet(_safe_sheet_name(sheet_name))

    note_text = f"Detected via: {source}  |  {DOCUMENT_TYPE} | Page {page}"
    note_cell = ws.cell(row=1, column=1, value=note_text)
    note_cell.font = Font(italic=True, color="666666", size=9)

    written_rows = [[note_text]]

    if headers:
        header_row = [_sanitize_cell_value(str(h)) for h in headers]
        ws.append(header_row)
        written_rows.append(header_row)

    width = len(headers)
    for row_idx, row in enumerate(rows, start=1):
        if isinstance(row, list):
            if width > 0 and len(row) > width:
                print(
                    f"  Warning: '{sheet_name}' row {row_idx} has {len(row)} value(s) but only "
                    f"{width} header(s); discarding extra value(s) {row[width:]!r}.",
                    file=sys.stderr,
                )
            written_row = [_sanitize_cell_value(v) for v in _normalize_row(row, width)]
            ws.append(written_row)
            written_rows.append(written_row)

    if headers:
        _style_header_row(ws, row=2)
    ws.freeze_panes = "A3"
    _autosize_columns_from_rows(ws, written_rows)


def _write_reviewer_note_sheet(
    wb: openpyxl.Workbook,
    source_filename: str,
    flagged_pages: list,
    low_confidence_pages: Optional[list] = None,
) -> None:
    """Insert a warning sheet at the front of the workbook listing pages
    whose non-tabular text wasn't captured, plus pages Docling itself
    flagged as low-confidence OCR/table quality (see
    MIN_OCR_CONFIDENCE_SCORE)."""
    ws = wb.create_sheet("Reviewer Note", 0)

    title_cell = ws.cell(row=1, column=1, value="REVIEWER NOTE - PARSING MAY BE INCOMPLETE")
    title_cell.font = Font(bold=True, size=13, color="9C0006")
    ws.merge_cells("A1:C1")

    explanation = (
        f"This workbook contains only data captured inside tables Docling detected in "
        f"'{source_filename}'. The page(s) below may be missing data because they contain "
        "non-tabular text that was NOT extracted into this workbook, because Docling's own "
        "OCR/table confidence on that page was low, or because a candidate table failed a "
        "reliability check and was discarded rather than trusted as-is (this script has no LLM "
        "fallback to re-verify a rejected table). Review the native source file directly before "
        "relying on this workbook alone."
    )
    explanation_cell = ws.cell(row=3, column=1, value=explanation)
    explanation_cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A3:C3")
    ws.row_dimensions[3].height = 60

    next_row = 5
    if flagged_pages:
        header_row = next_row
        ws.cell(row=header_row, column=1, value="Page")
        ws.cell(row=header_row, column=2, value="Approx. % of Page Text Captured in Tables")
        _style_header_row(ws, row=header_row)

        for offset, (page_number, ratio) in enumerate(flagged_pages, start=1):
            ws.cell(row=header_row + offset, column=1, value=page_number)
            note = f"{ratio * 100:.0f}%" if ratio is not None else "No OCR text extracted -- verify manually"
            ws.cell(row=header_row + offset, column=2, value=note)

        next_row = header_row + len(flagged_pages) + 2

    if low_confidence_pages:
        lc_title_cell = ws.cell(
            row=next_row, column=1,
            value="Pages Docling itself flagged as low OCR/table confidence",
        )
        lc_title_cell.font = Font(bold=True, color="9C0006")
        next_row += 1

        header_row = next_row
        ws.cell(row=header_row, column=1, value="Page")
        ws.cell(row=header_row, column=2, value="Docling Confidence Score")
        _style_header_row(ws, row=header_row)
        for offset, (page_number, score) in enumerate(sorted(low_confidence_pages), start=1):
            ws.cell(row=header_row + offset, column=1, value=page_number)
            ws.cell(row=header_row + offset, column=2, value=f"{score:.2f}")

    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20


def _safe_save(wb: openpyxl.Workbook, path: Path) -> None:
    """Write wb to a .tmp sibling, then atomically rename to path."""
    tmp = Path(str(path) + ".tmp")
    try:
        wb.save(tmp)
        for attempt in range(5):
            try:
                os.replace(tmp, path)
                return
            except PermissionError:
                if attempt < 4:
                    time.sleep(0.5)
                else:
                    raise
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def _strip_illegal_characters(wb: openpyxl.Workbook) -> None:
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cleaned = ILLEGAL_CHARACTERS_RE.sub("", cell.value)
                    if cleaned != cell.value:
                        cell.value = cleaned


def _save_workbook_with_recovery(wb: openpyxl.Workbook, out_path: Path) -> None:
    try:
        _safe_save(wb, out_path)
    except IllegalCharacterError:
        print(
            "  Warning: workbook contained control character(s) (transcribed verbatim from the "
            "source PDF) that openpyxl can't write to XML -- stripping them and retrying the save "
            "once.",
            file=sys.stderr,
        )
        _strip_illegal_characters(wb)
        _safe_save(wb, out_path)


@contextlib.contextmanager
def _file_lock(path: Path, timeout: float = 30.0, poll_interval: float = 0.1):
    """Advisory lock (sidecar `<path>.lock` file) serializing access to
    `path` across concurrently running instances of this script."""
    lock_path = Path(str(path) + ".lock")
    deadline = time.monotonic() + timeout
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            break
        except FileExistsError:
            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"Timed out waiting for lock on {path} -- if no other run is "
                    f"active, delete the stale lock file at {lock_path}."
                )
            time.sleep(poll_interval)
    try:
        yield
    finally:
        lock_path.unlink(missing_ok=True)


# ----------------- Destination relocation -----------------
_NUMBERED_DEST_SUFFIX_RE = re.compile(r"^(.*)_(\d+)$")


def _highest_existing_dest_suffix(parent: Path, stem: str, suffix: str) -> int:
    highest = 1
    try:
        entries = os.listdir(parent)
    except OSError:
        return highest
    for entry_name in entries:
        entry_path = Path(entry_name)
        if entry_path.suffix != suffix:
            continue
        m = _NUMBERED_DEST_SUFFIX_RE.match(entry_path.stem)
        if m and m.group(1) == stem:
            try:
                n = int(m.group(2))
            except ValueError:
                continue
            if n > highest:
                highest = n
    return highest


def _reserve_unique_destination(dest_path: Path) -> Path:
    """Atomically claim a destination path that won't clobber an existing
    file, even across two concurrently running instances of this script."""
    stem, suffix, parent = dest_path.stem, dest_path.suffix, dest_path.parent
    candidate = dest_path
    n = None
    while True:
        try:
            fd = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return candidate
        except FileExistsError:
            if n is None:
                n = _highest_existing_dest_suffix(parent, stem, suffix) + 1
            else:
                n += 1
            candidate = parent / f"{stem}_{n}{suffix}"


def _relocate_verified(src: Path, dest_path: Path) -> Path:
    """Relocate src to a non-clobbering path near dest_path via copy, verify,
    then delete -- never a bare move/rename."""
    reserved = _reserve_unique_destination(dest_path)
    try:
        shutil.copy2(str(src), str(reserved))
        src_size = src.stat().st_size
        if reserved.stat().st_size != src_size:
            raise IOError(
                f"Copy verification failed relocating {src} -> {reserved}: "
                f"copied size {reserved.stat().st_size} != source size {src_size}"
            )
    except Exception:
        reserved.unlink(missing_ok=True)
        raise
    src.unlink()
    return reserved


def _copy_verified(src: Path, dest_path: Path) -> Path:
    """Like _relocate_verified, but leaves src exactly where it was -- for
    the "manual review/" destination (see process_pdf's benchmark-gate
    branches and main()'s top-level exception handling), where the native
    file must stay in its original location (it may still need a rerun, or
    a reviewer may need it in context) and only a verified COPY goes into
    the review folder."""
    reserved = _reserve_unique_destination(dest_path)
    try:
        shutil.copy2(str(src), str(reserved))
        src_size = src.stat().st_size
        if reserved.stat().st_size != src_size:
            raise IOError(
                f"Copy verification failed copying {src} -> {reserved}: "
                f"copied size {reserved.stat().st_size} != source size {src_size}"
            )
    except Exception:
        reserved.unlink(missing_ok=True)
        raise
    return reserved


# ----------------- Metrics log -----------------
METRICS_COLUMNS = [
    "Timestamp", "File", "Status", "Notes", "Page Count", "Tables Found",
    # "Passed" / "Failed Reliability Check" / "No Table Detected" / "N/A
    # (Excluded)" -- see process_pdf's own table_check_status comment. Makes
    # it possible to tell, at a glance, WHY a file has 0 tables found or was
    # deferred: TableFormer's structural checks failing (this column) is a
    # completely separate reason from low OCR confidence ("Pages Flagged
    # (Low Docling Confidence)" below) -- a file can fail one without the
    # other, or neither, or (rarely) both. When it reads "Failed Reliability
    # Check", the Notes column identifies exactly which check failed, on
    # which page.
    "TableFormer Check",
    # "Layout Confidence" is a labeled stand-in for table-detection confidence,
    # not Docling's own table_score field -- that field is never populated by
    # the installed Docling version (confirmed by inspecting its source; see
    # _read_page_confidence), so it would read "n/a" for every file/page.
    # layout_score is the closest actually-populated per-page signal to "how
    # confidently was this page's table region detected".
    "Layout Confidence Scores (by page) -- table_score proxy",
    "OCR/Text Confidence Scores (by page)",
    "Pages Flagged (Uncaptured Text)", "Pages Flagged (Low Docling Confidence)", "Run Time (s)",
]

# A second sheet in the same workbook, one row per file copied to "manual
# review/" for low OCR confidence specifically (see process_pdf's OCR
# confidence gate) -- every such file also gets an ordinary row in the
# "Metrics" sheet above via Status="Manual Review (Low Confidence)"; this
# sheet exists just to make those documents easy to find/filter on their
# own.
LOW_CONFIDENCE_SHEET_NAME = "Low Confidence (Not Parsed)"
LOW_CONFIDENCE_COLUMNS = ["Timestamp", "File", "Low Confidence Page(s)", "Notes"]

DOCLING_SHEET_NAME = "Docling Metrics"

_SHEET_SIDECAR_SAFE_RE = re.compile(r"[^A-Za-z0-9_-]+")


class MetricsWriter:
    """Appends rows to one or more named sheets in the same persistent
    workbook at self._path, never overwriting rows already there -- each
    sheet gets its own jsonl sidecar so a mid-batch crash loses nothing
    already appended to any of them, and a run against an existing
    metrics_docling.xlsx adds to it rather than starting over.

    Sheet name and columns are passed explicitly on each append()/save()
    call (rather than fixed at construction) so ONE writer -- and ONE
    workbook -- can serve every sheet a run produces: this script's own
    per-file row and the low-confidence-only sheet. See main()'s
    sheet_specs for the full list save() consolidates."""

    def __init__(self, path: Path):
        self._path = path

    def _sidecar_path(self, sheet_name: str) -> Path:
        safe = _SHEET_SIDECAR_SAFE_RE.sub("_", sheet_name)
        return Path(str(self._path) + f".{safe}.pending.jsonl")

    def append(self, sheet_name: str, columns: list, **fields) -> None:
        row = [fields.get(col, "") for col in columns]
        sidecar = self._sidecar_path(sheet_name)
        sidecar.parent.mkdir(parents=True, exist_ok=True)
        with _file_lock(sidecar):
            with open(sidecar, "a", encoding="utf-8") as f:
                f.write(json.dumps(row) + "\n")

    def _read_pending_rows(self, sidecar_path: Path) -> list:
        if not sidecar_path.exists():
            return []
        with open(sidecar_path, "r", encoding="utf-8") as f:
            return [json.loads(line) for line in f if line.strip()]

    def _get_or_create_sheet(self, wb: openpyxl.Workbook, sheet_name: str, columns: list) -> Worksheet:
        if sheet_name in wb.sheetnames:
            return wb[sheet_name]
        ws = wb.create_sheet(sheet_name)
        ws.append(columns)
        _style_header_row(ws)
        return ws

    def save(self, sheet_specs: list) -> None:
        """sheet_specs: [(sheet_name, columns), ...] -- every sheet this
        writer might hold pending rows for. Listing a sheet that never
        actually got an append() call this run is harmless: its sidecar
        just doesn't exist, _read_pending_rows returns [], and the loop
        below skips creating it."""
        sidecar_paths = {name: self._sidecar_path(name) for name, _ in sheet_specs}
        for sidecar in sidecar_paths.values():
            sidecar.parent.mkdir(parents=True, exist_ok=True)

        with contextlib.ExitStack() as stack:
            for sidecar in sidecar_paths.values():
                stack.enter_context(_file_lock(sidecar))
            pending_by_sheet = {name: self._read_pending_rows(sidecar_paths[name]) for name, _ in sheet_specs}
            if not any(pending_by_sheet.values()):
                return

            with _file_lock(self._path):
                if self._path.exists():
                    wb = openpyxl.load_workbook(self._path)
                else:
                    self._path.parent.mkdir(parents=True, exist_ok=True)
                    wb = openpyxl.Workbook()
                    wb.remove(wb.active)

                for name, columns in sheet_specs:
                    rows = pending_by_sheet[name]
                    if not rows:
                        continue
                    ws = self._get_or_create_sheet(wb, name, columns)
                    for row in rows:
                        ws.append(row)
                    _autosize_columns_from_rows(ws, rows, preserve_existing=True)

                _safe_save(wb, self._path)
                wb.close()

            for sidecar in sidecar_paths.values():
                sidecar.unlink(missing_ok=True)


# ----------------- Docling setup -----------------
def _pdf_has_native_text(
    pdf_path: Path, sample_pages: int = 3, min_chars: int = MIN_NATIVE_TEXT_CHARS
) -> Tuple[bool, Optional[str]]:
    """(has_native_text, probe_error). has_native_text is True if pdf_path
    already has a real embedded/programmatic text layer (born-digital) on
    its first few pages, as opposed to being a scanned image PDF with no
    text at all. Checked cheaply via pypdfium2 -- entirely independent of
    Docling's own (much more expensive) pipeline -- so the right
    do_ocr/force_backend_text combination can be chosen per file before
    conversion runs, rather than guessing one setting for every PDF
    regardless of how it was produced.

    probe_error is None on a normal probe (whether or not native text was
    found); it carries the exception message when pdfium itself couldn't
    even open the file (e.g. a transient lock, or a genuinely corrupt PDF)
    -- has_native_text is always False in that case, but that False is a
    fallback DEFAULT, not a real "this PDF is scanned" determination, and
    the caller (_converter_for) needs to tell the two apart: silently
    treating a probe failure the same as a confirmed scanned PDF means a
    born-digital file could go through OCR instead of its own exact
    embedded text with nothing in the logs or metrics explaining why."""
    try:
        pdf = pdfium.PdfDocument(str(pdf_path))
    except Exception as exc:
        print(
            f"  WARNING: could not open '{pdf_path.name}' to probe for native text ({exc}) -- "
            "defaulting to OCR mode for this file; see the metrics Notes column.",
            file=sys.stderr,
        )
        return False, str(exc)
    try:
        total_chars = 0
        for i in range(min(sample_pages, len(pdf))):
            page = pdf[i]
            try:
                tp = page.get_textpage()
                try:
                    total_chars += len(tp.get_text_range().strip())
                finally:
                    tp.close()
            finally:
                page.close()
        return total_chars >= min_chars, None
    finally:
        pdf.close()


def _build_converter(native_text: bool, images_scale: float = DEFAULT_IMAGES_SCALE) -> DocumentConverter:
    """Standard Docling pipeline: layout analysis + the TableFormer
    table-structure model. No remote services, no VLM.

    native_text (see _pdf_has_native_text) switches between two distinct
    text sources -- OCR is only useful/necessary when there's no better
    option:
        True  -- born-digital PDF: skip OCR entirely (do_ocr=False) and set
                 force_backend_text=True so Docling uses the PDF's own exact
                 embedded text instead of the layout model's predicted text.
                 Faster, and removes OCR misreads as a source of error for
                 text that's already exact.
        False -- scanned/image PDF: OCR (EasyOCR) is the only way to get any
                 text at all, so it stays on and force_backend_text stays
                 off (there's no reliable backend text to force to).

    Either way, table STRUCTURE (row/column boundaries) is still predicted
    visually from the rendered page image at images_scale -- native_text
    only changes where cell/text CONTENT comes from, not how column
    boundaries are found. images_scale defaults higher than Docling's own
    1.0 default (see DEFAULT_IMAGES_SCALE).

    TableFormerMode is left at FAST (not ACCURATE) -- on a CPU-only,
    2-core machine ACCURATE mode was the dominant runtime cost (measured:
    ~100s for a single page).
    """
    pipeline_options = PdfPipelineOptions()
    pipeline_options.images_scale = images_scale
    pipeline_options.do_table_structure = True
    pipeline_options.table_structure_options.mode = TableFormerMode.FAST
    # Without this, every model (layout, TableFormer, EasyOCR) falls back to
    # its own default lookup location instead of the folder --warm-models
    # actually wrote into (docling_settings.cache_dir / "models") -- layout/
    # TableFormer would each try a plain snapshot_download() against the
    # standard HF hub cache (empty, since download_models() wrote via
    # local_dir instead) and fail under HF_HUB_OFFLINE; EasyOCR would look in
    # its own default ~/.EasyOCR/model/ instead of .../models/EasyOcr/.
    #pipeline_options.artifacts_path = docling_settings.cache_dir / "models"
    if native_text:
        pipeline_options.do_ocr = False
        pipeline_options.force_backend_text = True
    else:
        pipeline_options.do_ocr = True
        # EasyOCR downloads its weights straight from GitHub via urllib, not
        # through huggingface_hub -- HF_HUB_OFFLINE above has no effect on
        # it. download_enabled=False is what actually keeps it from
        # reaching the network mid-run once warmed; --warm-models (see
        # main()) is the only path that should ever set WARM_MODELS True.
        pipeline_options.ocr_options = EasyOcrOptions(download_enabled=WARM_MODELS)
        pipeline_options.force_backend_text = False
    return DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)}
    )


def _docling_tables_by_page(doc, header_line_reader: Optional["_PdfHeaderLineReader"] = None) -> dict:
    """Convert Docling's native TableItem/TableData structures into a
    {"headers": [...], "rows": [...]} shape, keyed by 1-indexed page number.

    A cell's column_span is filled forward across the columns it visually
    covers (safe -- it only ever fills in a blank on the SAME row), but
    row_span is deliberately NOT filled forward into additional rows --
    duplicating a spanning cell's content into every row it covers can make
    a genuinely sparse/broken row grid look artificially dense, defeating
    the empty-ratio/row-sparsity checks in _table_reliable that exist to
    catch exactly that.

    header_line_reader, when given, is used to recover each single-row
    header's true per-column stacked labels (see _PdfHeaderLineReader and the
    module comment above _unstack_docling_header_table) into a parallel
    "header_lines" entry -- only for a table whose header is exactly one
    physical row, since a table that already declares more than one header
    row doesn't have anything left to unstack."""
    tables_by_page: dict = {}

    for table in doc.tables or []:
        data = table.data
        row_count = data.num_rows
        col_count = data.num_cols
        if row_count == 0 or col_count == 0:
            continue

        grid = [["" for _ in range(col_count)] for _ in range(row_count)]
        # Parallel to grid: True at a position a col_span fill-forward wrote
        # into, False at a cell's own starting (r, c) position -- see the
        # module comment above _unstack_docling_header_table for why the
        # unstack reconstruction needs to tell these apart (TableFormer, on
        # this file's own dense-schedule test case, occasionally reports a
        # col_span of 2 on a stacked value simply because its text visually
        # runs wide for its narrow real column, bleeding that value into the
        # NEXT fused header-group's column too -- fine for the general
        # empty-ratio/no-gaps purpose col_span-filling exists for elsewhere,
        # but it would otherwise read as a genuine extra stacked value on a
        # column it was never really part of).
        fill_mask = [[False for _ in range(col_count)] for _ in range(row_count)]
        header_row_indices = set()
        header_cell_bbox_by_col: dict = {}

        for cell in data.table_cells:
            r, c = cell.start_row_offset_idx, cell.start_col_offset_idx
            if not (0 <= r < row_count and 0 <= c < col_count):
                continue
            content = (cell.text or "").strip()
            col_span = max(cell.col_span or 1, 1)
            for cc in range(c, min(c + col_span, col_count)):
                grid[r][cc] = content
                if cc != c:
                    fill_mask[r][cc] = True
            if cell.column_header:
                header_row_indices.add(r)
                if r == 0:
                    header_cell_bbox_by_col[c] = cell.bbox

        # Contiguous run of header-tagged rows starting at row 0, so a
        # mid-table repeated header (a section restart) doesn't get
        # mistaken for the table's real header extending that far down.
        if header_row_indices:
            header_row_count = 0
            while header_row_count in header_row_indices:
                header_row_count += 1
        else:
            header_row_count = 1 if grid else 0

        headers = grid[0] if grid else []
        data_rows = grid[header_row_count:]

        prov = getattr(table, "prov", None) or []
        page_number = prov[0].page_no if prov else 1

        table_dict = {
            "headers": headers,
            "rows": data_rows,
            "has_declared_header": bool(header_row_indices),
            "_col_span_fill_mask": fill_mask[header_row_count:],
        }

        if header_line_reader is not None and header_row_count == 1 and headers:
            table_dict["header_lines"] = [
                header_line_reader.lines_for(page_number, header_cell_bbox_by_col.get(c))
                or ([headers[c]] if headers[c] else [""])
                for c in range(col_count)
            ]

        tables_by_page.setdefault(page_number, []).append(table_dict)

    return tables_by_page


# ----------------- Stacked-header reconstruction -----------------
# When TableFormer fuses several real, narrow columns into one detected grid
# column (observed on a dense surgical-schedule PDF), each fused column's
# header cell still visually stacks the real columns' labels (e.g.
# "IP/OP" above "Time"), and each record's row still visually spans several
# print-lines -- but Docling flattens a header cell's stacked labels with a
# plain SPACE ("IP/OP Time"), not a newline, and the correspondingly stacked
# DATA lines instead land as several separate, mostly-empty detected rows.
# Left alone this produces a table with too many (mostly empty) rows and too
# few (multi-label) columns.
#
# Splitting "IP/OP Time" back into ["IP/OP", "Time"] can't be done by
# whitespace alone -- indistinguishable from a genuine multi-word label like
# "Last Name". Docling has no per-OCR-line geometry, but for a born-digital
# PDF (native_text -- see _pdf_has_native_text) there is an equally
# authoritative source: pypdfium2's own text layer, bounded to the header
# TableCell's own bbox (Docling reports one in real PDF-point page
# coordinates regardless of do_ocr/force_backend_text) -- see
# _PdfHeaderLineReader. pypdfium2 inserts the PDF's own line breaks
# ("\r\n") between stacked labels, which a flattened, space-joined string
# can't distinguish on its own.
MIN_UNSTACK_STACK_DEPTH = 2  # at least one column must have 2+ stacked labels to bother


class _PdfHeaderLineReader:
    """Recovers a Docling header TableCell's true stacked labels, top-to-
    bottom, from pdf_path's own PDF text layer -- see the module comment
    above. Opens pdf_path lazily (only if a header cell actually needs
    looking up) and caches one PdfTextPage per page number queried, since a
    multi-table document may ask about several header cells on the same
    page.

    Returns [] for a page with no real PDF text layer (a scanned/image PDF,
    most notably) -- pypdfium2's own bounded-text extraction just comes back
    an empty string in that case, no exception involved, so a document this
    can't help falls straight through to the "no split" fallback in
    _docling_tables_by_page, exactly today's flattened-header behavior. A
    genuine lookup failure (a bad page index, a malformed bbox, a
    pypdfium2-internal error) is a different thing entirely -- caught below,
    but reported with a one-line stderr warning naming the page and cell
    rather than silently returning [] the same way -- indistinguishable from
    "this document just has no stacked headers" otherwise, with nothing to
    say a whole page's worth of reconstruction quietly stopped working.

    Caches only the single most-recently-used page/textpage rather than one
    per page number ever seen -- the calling loop already groups every
    header lookup for one table (and, in practice, one page) together
    before moving to the next, so this doesn't cause repeated re-opens; it
    just stops a many-page document from holding every page's handles open
    at once for the run's whole duration."""

    def __init__(self, pdf_path: Path):
        self._pdf_path = pdf_path
        self._pdf: Optional[pdfium.PdfDocument] = None
        self._page_cache: Optional[Tuple[int, "pdfium.PdfPage", "pdfium.PdfTextPage"]] = None

    def lines_for(self, page_number: int, bbox, pad: float = 1.0) -> list:
        if bbox is None:
            return []
        try:
            if self._pdf is None:
                self._pdf = pdfium.PdfDocument(str(self._pdf_path))
            if self._page_cache is None or self._page_cache[0] != page_number:
                if self._page_cache is not None:
                    _, old_page, old_textpage = self._page_cache
                    with contextlib.suppress(Exception):
                        old_textpage.close()
                    with contextlib.suppress(Exception):
                        old_page.close()
                page = self._pdf[page_number - 1]
                self._page_cache = (page_number, page, page.get_textpage())
            _, page, textpage = self._page_cache
            height = page.get_size()[1]
            bl = bbox.to_bottom_left_origin(height)
            text = textpage.get_text_bounded(
                left=bl.l - pad, bottom=bl.b - pad, right=bl.r + pad, top=bl.t + pad,
            )
        except Exception as e:
            print(
                f"  Warning: header-line lookup failed on page {page_number}, cell {bbox} -- "
                f"{type(e).__name__}: {e} -- header left unstacked for this cell.",
                file=sys.stderr,
            )
            return []
        return [p.strip() for p in text.replace("\r\n", "\n").split("\n") if p.strip()]

    def close(self) -> None:
        if self._page_cache is not None:
            _, page, textpage = self._page_cache
            with contextlib.suppress(Exception):
                textpage.close()
            with contextlib.suppress(Exception):
                page.close()
            self._page_cache = None
        if self._pdf is not None:
            with contextlib.suppress(Exception):
                self._pdf.close()
            self._pdf = None


def _unstack_docling_header_table(table: dict, page_number: Optional[int] = None) -> Optional[dict]:
    """If table["header_lines"] (see _docling_tables_by_page /
    _PdfHeaderLineReader) shows several real columns were merged into one,
    and the data-row count divides evenly into same-size record blocks,
    return a reconstructed table dict with one flattened header per real
    column and one row per real record. Returns None if the pattern doesn't
    apply (no header_lines available, no stacked header, or the row count
    doesn't cleanly divide) -- callers should keep using the original table
    dict in that case.

    Same positional (not row-offset) mapping of a column's non-blank values
    onto its stacked labels within a record's block of rows, and the same
    overflow-to-last-column handling when a narrower column's block still
    holds more values than it has labels for. One caveat: a record whose
    stacked sub-fields are blank anywhere but the END of a column's stack
    would have a later value shifted into the missing one's slot.

    A position table["_col_span_fill_mask"] (see _docling_tables_by_page)
    marks as a col_span fill-forward duplicate, not that cell's own content,
    is treated as blank here regardless of what grid value landed there --
    see the module comment above this function for why a value legitimately
    belonging to one stacked column can otherwise bleed into gathering the
    NEXT one's values instead."""
    headers = table.get("headers") or []
    raw_rows = table.get("rows") or []
    raw_mask = table.get("_col_span_fill_mask") or []
    rows: list = []
    fill_mask: list = []
    for i, r in enumerate(raw_rows):
        if isinstance(r, list):
            rows.append(r)
            fill_mask.append(raw_mask[i] if i < len(raw_mask) else [False] * len(r))
    if not headers or not rows:
        return None

    header_lines = table.get("header_lines")
    if not header_lines or len(header_lines) != len(headers):
        return None
    split_headers = [parts if parts else [""] for parts in header_lines]

    stack_depth = max(len(parts) for parts in split_headers)
    if stack_depth < MIN_UNSTACK_STACK_DEPTH:
        return None
    if len(rows) % stack_depth != 0:
        return None

    new_headers: list = []
    for parts in split_headers:
        new_headers.extend(parts)

    new_rows: list = []
    for block_start in range(0, len(rows), stack_depth):
        block = rows[block_start:block_start + stack_depth]
        mask_block = fill_mask[block_start:block_start + stack_depth]
        new_row: list = []
        for c, parts in enumerate(split_headers):
            values = [
                row[c] if c < len(row) and not (c < len(mask_block[i]) and mask_block[i][c]) else ""
                for i, row in enumerate(block)
            ]
            non_empty = [v for v in values if str(v or "").strip()]
            if len(non_empty) > len(parts):
                # This column has fewer stacked labels than stack_depth (the
                # MAX label count across every column) but the block still
                # holds one legitimate value per physical row -- keep this
                # column's own label count of slots and fold whatever's left
                # over into the last one, loudly, rather than silently
                # dropping real data past len(parts).
                keep = non_empty[:len(parts) - 1]
                overflow = non_empty[len(parts) - 1:]
                slots = keep + ["; ".join(overflow)]
                page_desc = f"page {page_number}" if page_number is not None else "unknown page"
                print(
                    f"  WARNING: stacked-header unstack on {page_desc}, column {c + 1} -- "
                    f"found {len(non_empty)} value(s) but this column only has {len(parts)} "
                    "stacked label(s) (narrower than the table's overall stack depth); "
                    "overflow value(s) joined into the last column instead of being dropped.",
                    file=sys.stderr,
                )
            else:
                slots = (non_empty + [""] * len(parts))[:len(parts)]
            new_row.extend(slots)
        new_rows.append(new_row)

    result = dict(table)
    result["headers"] = new_headers
    result["rows"] = new_rows
    result.pop("header_lines", None)
    result.pop("_col_span_fill_mask", None)
    return result


def _page_texts_by_number(doc) -> dict:
    """Build a per-page text blob to compare table content against for the
    coverage-ratio check and government-form detection: every free-text
    item's text, plus every table's own cell text, grouped by page number.

    Docling doesn't expose an equivalent "everything visible on this page,
    regardless of structure" text dump on DoclingDocument itself, only
    already-classified items. Using texts + table cell content together
    approximates that: how much of a page's real content sits inside vs.
    outside a table."""
    texts_by_page: dict = {}

    for item in doc.texts or []:
        for prov in (getattr(item, "prov", None) or []):
            texts_by_page.setdefault(prov.page_no, []).append(item.text or "")

    for table in doc.tables or []:
        cell_text = " ".join(
            (c.text or "").strip() for c in table.data.table_cells if (c.text or "").strip()
        )
        if not cell_text:
            continue
        for prov in (getattr(table, "prov", None) or []):
            texts_by_page.setdefault(prov.page_no, []).append(cell_text)

    page_count = doc.num_pages()
    return {p: "\n".join(texts_by_page.get(p, [])) for p in range(1, page_count + 1)}


def _read_page_confidence(result) -> dict:
    """{page_number: (table_score, ocr_score, layout_score)} from Docling's
    own confidence report, or {} if the installed Docling version/pipeline
    didn't populate one -- every call site treats a missing entry as "no
    signal" rather than failing.

    table_score is carried through here for completeness (it's part of
    Docling's own confidence report shape) but is not acted on anywhere --
    as of the Docling version this was written against, no model in the
    installed package ever actually writes a per-page table_score --
    confirmed by grep'ing site-packages/docling for every assignment site: it
    is declared with a permanent np.nan default and nothing overwrites it.
    _table_reliable has no table_score-based gate for exactly this reason
    (see its own docstring) -- a threshold compared against a value that's
    always NaN would never fire, so there is no such threshold. layout_score
    (the layout model's own per-page confidence, which DOES get populated
    every run) is read here too, for _format_page_confidence to report as a
    labeled stand-in in the Metrics sheet, since it's the closest actually-
    populated signal to "how confidently was this page's table region
    detected/classified" available from Docling's public API right now."""
    confidence = getattr(result, "confidence", None)
    if confidence is None:
        return {}
    pages = getattr(confidence, "pages", None) or {}
    out = {}
    for page_number, scores in pages.items():
        out[page_number] = (
            getattr(scores, "table_score", None),
            getattr(scores, "ocr_score", None),
            getattr(scores, "layout_score", None),
        )
    return out


def _format_page_confidence(confidence_by_page: dict, page_count: int) -> tuple:
    """Render confidence_by_page (see _read_page_confidence) as two
    "p1: 0.95 | p2: 0.88 | ..." strings covering every page 1..page_count --
    not just the pages that ended up contributing an accepted table -- so the
    Metrics sheet shows the same confidence numbers the reliability/OCR-
    confidence gates actually acted on (for ocr_score), plus a labeled proxy
    for table-detection confidence (layout_score, since Docling's own
    table_score field is unpopulated in this install -- see
    _read_page_confidence), per file, without having to reopen the PDF in
    --debug mode. A page missing from confidence_by_page, or whose score is
    NaN (e.g. ocr_score on a native-text page, where OCR never ran at all),
    prints as "n/a" rather than being silently dropped from the string."""
    def _fmt(score) -> str:
        if score is None or (isinstance(score, float) and math.isnan(score)):
            return "n/a"
        return f"{score:.2f}"

    layout_parts = []
    ocr_parts = []
    for page_number in range(1, page_count + 1):
        _table_score, ocr_score, layout_score = confidence_by_page.get(page_number, (None, None, None))
        layout_parts.append(f"p{page_number}: {_fmt(layout_score)}")
        ocr_parts.append(f"p{page_number}: {_fmt(ocr_score)}")
    return " | ".join(layout_parts), " | ".join(ocr_parts)


# Fraction of a candidate header row's non-empty cells that must be
# digit-free to be trusted as real column labels rather than a data row --
# see _looks_like_real_header for why has_declared_header can't be used for
# this instead.
MIN_LABEL_LIKE_FRACTION = 0.6


def _looks_like_real_header(values: list) -> bool:
    """True if values reads like column labels ("SSN", "Fname", "Wage
    Code") rather than a data row ("315-78-1207", "$8,810.00", "WC-105").

    Needed because Docling's TableFormer, observed against this script's own
    test PDFs, tags row 0 as column_header on very nearly every
    independently detected table region -- including a continuation page
    that has no real header at all, just more data -- so
    has_declared_header (see _docling_tables_by_page) is true almost
    unconditionally and can't tell a real header apart from a continuation
    page's first data row. Real column labels are overwhelmingly digit-free
    short words/phrases; a data row -- an SSN, a dollar amount, a code like
    "WC-104" -- is not. This is intentionally a narrow, single-purpose check
    (only used to decide whether to merge a table into the previous page's
    as a continuation), not a general label:value classifier."""
    non_empty = [str(v).strip() for v in values if v is not None and str(v).strip()]
    if not non_empty:
        return False
    label_like = sum(1 for v in non_empty if not any(ch.isdigit() for ch in v))
    return (label_like / len(non_empty)) >= MIN_LABEL_LIKE_FRACTION


def _merge_continuation_tables(tables_in_page_order: list) -> list:
    """Fold a multi-page table's later-page fragments back into its first
    page's entry, e.g. a 3-page roster whose header only prints once on
    page 1 -- ideal output is one continuous "Table_1 (p1-3)" sheet rather
    than three separate, page-1-headered-looking sheets.

    tables_in_page_order: process_pdf's all_tables, in the ascending page
    order they were appended in (one table per accepted page, in this
    script's simplified model -- see _docling_tables_by_page). A table is
    treated as the immediately preceding (anchor) table's continuation onto
    a later page when ALL of:
        - its column count matches the anchor's,
        - its page is exactly the page right after whatever page the
          anchor most recently absorbed (no gap-tolerance -- a page in
          between with no reliable table of its own breaks the chain here),
        - its "headers" (Docling always hands back SOME first row as
          headers -- see _docling_tables_by_page) does NOT look like a real
          header row (see _looks_like_real_header) -- i.e. it's actually
          more data, arbitrarily promoted, the shape an un-headered
          continuation page produces.
    That "headers" row is folded back in as an ordinary data row, appended
    onto the anchor, and the anchor's "page" field widens from a single int
    into a "first-last" range string for display (_write_table_to_sheet
    only ever uses "page" to build a label, so a string is safe there).

    Deliberately simple in one other way too: no verification that the
    anchor's real headers make sense as this continuation's headers. Good
    enough for the common "one header page, N data-only pages" shape."""
    merged: list = []
    for table in tables_in_page_order:
        prev = merged[-1] if merged else None
        prev_last_page = prev.get("_last_page") if prev else None
        is_continuation = (
            prev is not None
            and isinstance(prev_last_page, int)
            and table["page"] == prev_last_page + 1
            and len(table.get("headers") or []) == len(prev.get("headers") or [])
            and not _looks_like_real_header(table.get("headers") or [])
        )
        if is_continuation:
            prev["rows"].append(table["headers"])
            prev["rows"].extend(table.get("rows") or [])
            first_page = prev.get("_first_page", prev["page"])
            prev["_first_page"] = first_page
            prev["_last_page"] = table["page"]
            prev["page"] = f"{first_page}-{table['page']}"
            prev.pop("_char_count", None)
            continue
        table = dict(table)
        table["_last_page"] = table["page"]
        merged.append(table)
    return merged


# ----------------- Per-file processing -----------------
def process_pdf(
    converter: DocumentConverter, mode_label: str, pdf_path: Path, output_folder: Path,
    debug: bool = False,
) -> dict:
    print(f"Processing: {pdf_path.name} ({mode_label})")
    source_label = f"{ENGINE_LABEL}, {mode_label}"

    result = converter.convert(pdf_path)
    if result.status == ConversionStatus.FAILURE:
        errors = "; ".join(str(e) for e in (result.errors or [])) or "unknown error"
        raise RuntimeError(f"Docling conversion failed: {errors}")

    doc = result.document
    page_count = doc.num_pages()
    page_texts = _page_texts_by_number(doc)
    confidence_by_page = _read_page_confidence(result)
    # Per-page confidence, formatted once and attached to every return path
    # below (including the early relocation returns) so the Metrics sheet
    # always shows the same table/OCR confidence numbers the reliability and
    # OCR-confidence gates actually acted on for this file -- see
    # _format_page_confidence.
    layout_confidence_str, ocr_confidence_str = _format_page_confidence(confidence_by_page, page_count)

    gov_form_hit = _detect_government_form(page_texts)
    if gov_form_hit is not None:
        gov_page, gov_phrase = gov_form_hit
        gov_forms_dir = output_folder / "government forms"
        gov_forms_dir.mkdir(parents=True, exist_ok=True)
        dest_path = _relocate_verified(pdf_path, gov_forms_dir / pdf_path.name)
        print(
            f"  Detected as a government/tax form (page {gov_page} matched {gov_phrase!r}) -- "
            f"moved to {gov_forms_dir} without extraction."
        )
        return {
            "dest_path": dest_path, "page_count": page_count, "tables_found": 0,
            "pages_flagged": 0, "low_confidence_pages": 0,
            "layout_confidence": layout_confidence_str, "ocr_confidence": ocr_confidence_str,
            # Excluded outright before table detection was even attempted --
            # nothing for TableFormer/_table_reliable to have judged.
            "table_check": "N/A (Excluded)",
            "status": "Skipped (Government/Tax Form)", "notes": f"page {gov_page}: {gov_phrase!r}",
        }

    header_line_reader = _PdfHeaderLineReader(pdf_path)
    try:
        raw_tables_by_page = _docling_tables_by_page(doc, header_line_reader)
    finally:
        header_line_reader.close()

    # Reconstruct any stacked-header tables (see the module comment above
    # _unstack_docling_header_table) before the reliability gate or
    # continuation-merge logic below sees them -- a table that's still
    # fused/split-across-rows at that point would otherwise very plausibly
    # be rejected outright, or merged as a bogus "continuation" of the wrong
    # table.
    def _unstack_and_clean(t: dict, page_number: int) -> dict:
        result = _unstack_docling_header_table(t, page_number) or dict(t)
        result.pop("header_lines", None)
        result.pop("_col_span_fill_mask", None)
        return result

    for page_number, page_tables in raw_tables_by_page.items():
        raw_tables_by_page[page_number] = [
            _unstack_and_clean(t, page_number) for t in page_tables
        ]

    all_tables = []
    low_confidence_pages = []
    table_chars_by_page: dict = {}
    # (page_number, reason) for every page whose candidate table(s) TableFormer
    # actually detected were then rejected by _table_reliable -- distinct from
    # a page with no candidates at all (that page never reaches this list).
    # Used below to tell "TableFormer found nothing" apart from "TableFormer
    # found something but it failed a reliability check" -- see
    # table_check_status/table_check_detail.
    table_check_failures: list = []
    for page_number in range(1, page_count + 1):
        page_text = page_texts.get(page_number, "")
        _table_score, ocr_score, _layout_score = confidence_by_page.get(page_number, (None, None, None))

        if (
            ocr_score is not None
            and not (isinstance(ocr_score, float) and math.isnan(ocr_score))
            and ocr_score < MIN_OCR_CONFIDENCE_SCORE
        ):
            low_confidence_pages.append((page_number, ocr_score))

        page_tables = raw_tables_by_page.get(page_number)
        if not page_tables:
            continue

        reliable, reject_reason = _table_reliable(
            page_tables, page_text, debug=debug, debug_label=f" page {page_number}"
        )
        if not reliable:
            table_check_failures.append((page_number, reject_reason))
            continue

        for table in page_tables:
            all_tables.append({**table, "page": page_number, "source": source_label})
            # Captured BEFORE _merge_continuation_tables folds later pages'
            # tables into an earlier page's entry -- see that function's
            # docstring for why this dict has to be built per-page, now,
            # rather than derived from the (possibly merged) table list
            # afterward.
            table_chars_by_page[page_number] = (
                table_chars_by_page.get(page_number, 0) + _table_char_count(table)
            )

    all_tables = _merge_continuation_tables(all_tables)
    flagged_pages = _find_pages_with_uncaptured_text(page_texts, table_chars_by_page)

    # "TableFormer Check" (see the Docling Metrics sheet's own column of that
    # name): tells apart three distinct reasons a file might end up without a
    # parsed table, or deferred/skipped despite one -- see the module
    # docstring's "OCR confidence" vs "table reliability" distinction:
    #   "Passed"                  -- at least one page's candidate table(s)
    #                                cleared _table_reliable (all_tables is
    #                                non-empty) -- true whether the file goes
    #                                on to parse successfully OR gets
    #                                deferred/skipped for LOW CONFIDENCE
    #                                below; either way, TableFormer's own
    #                                structural checks were not the reason.
    #   "Failed Reliability Check" -- TableFormer found candidate(s)
    #                                somewhere, but every one was rejected by
    #                                _table_reliable (table_check_failures is
    #                                non-empty and all_tables ended up empty)
    #                                -- this IS why the file was deferred/
    #                                relocated as "Not Tabular" below.
    #   "No Table Detected"        -- TableFormer never proposed a table
    #                                candidate on any page at all -- nothing
    #                                for _table_reliable to have even judged.
    if all_tables:
        table_check_status = "Passed"
        table_check_detail = ""
    elif table_check_failures:
        table_check_status = "Failed Reliability Check"
        table_check_detail = "; ".join(f"page {p}: {reason}" for p, reason in table_check_failures)
    else:
        table_check_status = "No Table Detected"
        table_check_detail = ""

    # OCR confidence gate: a document where any page that contributed an
    # accepted table also had low Docling OCR confidence is treated as too
    # unreliable to trust at all -- extraction is skipped for the WHOLE
    # document (a copy goes to "manual review/" below, native file left in
    # place), since untrustworthy OCR text undermines the table content
    # itself. There's no fallback engine to re-verify a low-confidence table
    # against. Revisit once the OCR model itself is upgraded.
    low_confidence_table_pages = [
        (page_number, ocr_score) for page_number, ocr_score in low_confidence_pages
        if page_number in table_chars_by_page
    ]
    if low_confidence_table_pages:
        pages_str = ", ".join(f"p{p} (ocr={s:.2f})" for p, s in low_confidence_table_pages)
        manual_dir = output_folder / "manual review"
        manual_dir.mkdir(parents=True, exist_ok=True)
        dest_path = _copy_verified(pdf_path, manual_dir / pdf_path.name)
        print(
            f"  Low Docling OCR confidence on page(s) contributing table(s) ({pages_str}) -- "
            f"copied to {manual_dir} for manual review (original left in place)."
        )
        return {
            "dest_path": dest_path, "page_count": page_count, "tables_found": 0,
            "pages_flagged": len(flagged_pages), "low_confidence_pages": len(low_confidence_pages),
            "layout_confidence": layout_confidence_str, "ocr_confidence": ocr_confidence_str,
            # Always "Passed" here -- low_confidence_table_pages only ever
            # has entries for a page that already contributed an ACCEPTED
            # table (see table_chars_by_page above), so reaching this
            # branch means TableFormer's own checks were never the issue.
            "table_check": table_check_status,
            "status": "Manual Review (Low Confidence)",
            "notes": f"confidence score was low on {pages_str} -- not parsed, copy placed at {dest_path}",
            "low_confidence_skip": True, "low_confidence_pages_detail": pages_str,
        }

    if not all_tables:
        # table_check_status is "Failed Reliability Check" or "No Table
        # Detected" here -- never "Passed", since all_tables is empty. The
        # detail (which specific check, on which page) is folded into the
        # Notes below so it's visible without a --debug rerun.
        note_suffix = f" ({table_check_detail})" if table_check_detail else ""
        manual_dir = output_folder / "manual review"
        manual_dir.mkdir(parents=True, exist_ok=True)
        dest_path = _copy_verified(pdf_path, manual_dir / pdf_path.name)
        print(
            f"  No reliable table found ({table_check_status}) -- copied to {manual_dir} for manual "
            "review (original left in place)."
        )
        return {
            "dest_path": dest_path, "page_count": page_count, "tables_found": 0,
            "pages_flagged": len(flagged_pages), "low_confidence_pages": len(low_confidence_pages),
            "layout_confidence": layout_confidence_str, "ocr_confidence": ocr_confidence_str,
            "table_check": table_check_status,
            "status": "Manual Review (Not Tabular)",
            "notes": f"no reliable table found by Docling{note_suffix} -- copy placed at {dest_path}",
        }

    tabular_dir = output_folder / "tabular data"
    tabular_dir.mkdir(parents=True, exist_ok=True)
    # Never clobber a prior run's output for this same PDF -- claim
    # "<stem>_intermediate.xlsx", or "<stem>_intermediate_1.xlsx", "_2.xlsx",
    # etc. if it's already taken, atomically (safe even across two
    # concurrently running instances of this script). See
    # _reserve_unique_destination. The "_intermediate" tag marks this xlsx as
    # Docling's own parse of the table(s) -- not a final, reviewed deliverable
    # -- and distinguishes it from the native PDF copy placed alongside it
    # below.
    out_path = _reserve_unique_destination(tabular_dir / f"{pdf_path.stem}_intermediate.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if flagged_pages or low_confidence_pages:
        _write_reviewer_note_sheet(wb, pdf_path.name, flagged_pages, low_confidence_pages)
    for idx, table in enumerate(all_tables, start=1):
        sheet_name = f"Table {idx} (p{table['page']})"
        _write_table_to_sheet(wb, table, sheet_name, table["source"], table["page"])
    _save_workbook_with_recovery(wb, out_path)
    wb.close()

    # A copy of the native PDF sits alongside its parsed intermediate xlsx so
    # a reviewer can compare the two without hunting down the original in its
    # source folder -- the source PDF itself is left untouched, same as every
    # other successful-parse path here.
    pdf_copy_path = _copy_verified(pdf_path, tabular_dir / pdf_path.name)

    print(
        f"  -> {out_path} ({len(all_tables)} table(s) across {page_count} page(s)); "
        f"native copy -> {pdf_copy_path}"
    )
    return {
        "dest_path": out_path, "page_count": page_count, "tables_found": len(all_tables),
        "pages_flagged": len(flagged_pages), "low_confidence_pages": len(low_confidence_pages),
        "layout_confidence": layout_confidence_str, "ocr_confidence": ocr_confidence_str,
        "table_check": table_check_status,
        "status": "Success", "notes": "",
    }


# ----------------- CLI -----------------
def select_pdf_files() -> List[Path]:
    """Open a file-picker dialog so the user can select one or more PDFs."""
    if tk is None:
        sys.exit(
            "Error: tkinter is not available -- cannot open the file picker. "
            "Pass a file or folder path as a command-line argument instead."
        )
    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        paths = filedialog.askopenfilenames(
            title="Select PDF file(s) to process",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
    finally:
        if root is not None:
            root.destroy()
    return [Path(p) for p in paths]


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract tables from PDFs into Excel workbooks using Docling's local "
                    "standard pipeline (OCR + layout + TableFormer) only -- no GLiNER/PII "
                    "gate, no LLM fallback.",
    )
    parser.add_argument(
        "input",
        metavar="INPUT",
        nargs="?",
        default=None,
        help="Path to a single PDF file, or a folder containing PDFs. "
             "If omitted, a file-picker dialog opens to select one or more PDFs.",
    )
    parser.add_argument(
        "--output-folder",
        default="",
        help="Folder to write .xlsx output into. Defaults to 'docling_output' next to the input.",
    )
    parser.add_argument(
        "--metrics-path",
        default=None,
        help="Path to a single persistent metrics log shared across all output folders. "
             "Default: a metrics_docling.xlsx is written inside each output folder used.",
    )
    parser.add_argument(
        "--metrics-sheet-name",
        default=DOCLING_SHEET_NAME,
        help=f"Sheet name for this script's own per-file metrics row inside the metrics "
             f"workbook. Default: '{DOCLING_SHEET_NAME}'. Override when --metrics-path points "
             "at a workbook shared with another writer.",
    )
    parser.add_argument(
        "--images-scale",
        type=float,
        default=DEFAULT_IMAGES_SCALE,
        help=f"Resolution scale for the page images Docling's layout/TableFormer models see "
             f"(Docling's own default is 1.0, roughly 72 DPI-equivalent -- too low to reliably "
             f"resolve column boundaries on a dense, many-column table). Default: "
             f"{DEFAULT_IMAGES_SCALE}. Higher is slower but never less accurate.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print extra diagnostics to stderr about table detection and reliability-gate "
             "decisions per page. Never prints cell content.",
    )
    parser.add_argument(
        "--warm-models",
        action="store_true",
        help="Download and cache Docling's layout, TableFormer, and EasyOCR (English) models, "
             "then exit without processing any PDFs. Run this once, on a machine with no PHI "
             "present -- every other run enforces HF_HUB_OFFLINE (see the top of this file), "
             "so a missing/uncached model fails loudly at startup instead of this script "
             "silently reaching Hugging Face/GitHub mid-run.",
    )
    args = parser.parse_args()

    if args.warm_models:
        print("Downloading Docling models (layout, TableFormer, EasyOCR English)...")
        # This script's own pipeline (see _build_converter) only ever uses the
        # layout, TableFormer, and EasyOCR models -- picture_classifier,
        # code_formula, and rapidocr are download_models()'s own defaults, not
        # anything this script's pipeline calls, and rapidocr's default model
        # source (modelscope.cn) is commonly blocked by corporate firewalls
        # that otherwise allow huggingface.co/github.com. Disabled here so
        # --warm-models only needs those two hosts unblocked.
        download_models(
            with_easyocr=True,
            easyocr_languages=["en"],
            with_picture_classifier=False,
            with_code_formula=False,
            with_rapidocr=False,
            progress=True,
        )
        print("Done -- models cached locally. Safe to run without --warm-models from here on.")
        return

    if args.input is None:
        pdf_files = select_pdf_files()
        if not pdf_files:
            print("No files selected. Exiting.")
            return
    else:
        input_path = Path(args.input)
        if not input_path.exists():
            print(f"Error: path not found: {input_path}", file=sys.stderr)
            sys.exit(1)

        if input_path.is_dir():
            pdf_files = sorted(p for p in input_path.iterdir() if p.suffix.lower() == ".pdf")
        else:
            if input_path.suffix.lower() != ".pdf":
                print(f"Error: not a PDF file: {input_path}", file=sys.stderr)
                sys.exit(1)
            pdf_files = [input_path]

    if not pdf_files:
        print("No PDF files found.", file=sys.stderr)
        sys.exit(1)

    metrics_override = Path(args.metrics_path) if args.metrics_path else None
    metrics_logs: dict = {}

    def _metrics_for(output_folder: Path) -> MetricsWriter:
        key = metrics_override if metrics_override is not None else output_folder / "metrics_docling.xlsx"
        writer = metrics_logs.get(key)
        if writer is None:
            writer = MetricsWriter(key)
            metrics_logs[key] = writer
        return writer

    print("Building Docling converter(s) (first run downloads model weights)...")
    # Built lazily, one per native_text mode actually needed (a batch of all
    # born-digital or all-scanned PDFs never pays for the unused one) -- see
    # _build_converter for why native-text vs. OCR needs a different pipeline
    # configuration rather than one converter handling both.
    converters: dict = {}

    def _converter_for(pdf_path: Path) -> tuple:
        native, probe_error = _pdf_has_native_text(pdf_path)
        if probe_error is not None:
            # The probe itself failed to even open the file -- `native` is
            # only pypdfium2's own fallback default here, not a real
            # "this PDF is scanned" verdict (see _pdf_has_native_text's own
            # docstring). Labeled distinctly from a genuine "OCR mode" so
            # the reason OCR was used for this specific file survives into
            # the metrics Notes column (see the call site below), not just
            # a console warning that scrolls past.
            mode_label = "OCR mode (native-text probe failed)"
        else:
            mode_label = "native-text mode" if native else "OCR mode"
        if native not in converters:
            converters[native] = _build_converter(native_text=native, images_scale=args.images_scale)
        return converters[native], mode_label, probe_error

    processed = 0
    output_folders_used = set()
    for pdf_path in pdf_files:
        output_folder = Path(args.output_folder) if args.output_folder else pdf_path.parent / "docling_output"
        output_folders_used.add(output_folder)
        metrics = _metrics_for(output_folder)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        start = time.perf_counter()
        try:
            converter, mode_label, probe_error = _converter_for(pdf_path)
            summary = process_pdf(converter, mode_label, pdf_path, output_folder, debug=args.debug)
            elapsed = time.perf_counter() - start

            notes = summary["notes"]
            if probe_error is not None:
                # Folded into the existing Notes column rather than a new
                # one -- this is the one place a reader can learn OCR was
                # used for this file because the native-text probe itself
                # failed (see _pdf_has_native_text), not because the PDF
                # was confirmed scanned.
                probe_note = f"OCR mode used because the native-text probe failed: {probe_error}"
                notes = f"{notes} -- {probe_note}" if notes else probe_note

            metrics.append(
                args.metrics_sheet_name, METRICS_COLUMNS,
                Timestamp=timestamp,
                File=pdf_path.name, Status=summary["status"], Notes=notes,
                **{
                    "Page Count": summary["page_count"], "Tables Found": summary["tables_found"],
                    "TableFormer Check": summary.get("table_check", ""),
                    "Layout Confidence Scores (by page) -- table_score proxy": summary.get(
                        "layout_confidence", ""
                    ),
                    "OCR/Text Confidence Scores (by page)": summary.get("ocr_confidence", ""),
                    "Pages Flagged (Uncaptured Text)": summary["pages_flagged"],
                    "Pages Flagged (Low Docling Confidence)": summary["low_confidence_pages"],
                    "Run Time (s)": round(elapsed, 2),
                },
            )
            if summary.get("low_confidence_skip"):
                metrics.append(
                    LOW_CONFIDENCE_SHEET_NAME, LOW_CONFIDENCE_COLUMNS,
                    Timestamp=timestamp,
                    File=pdf_path.name,
                    **{"Low Confidence Page(s)": summary.get("low_confidence_pages_detail", "")},
                    Notes="Confidence score was low -- document not parsed; a copy was placed in "
                          "'manual review/' for a human to look at (native file left in place).",
                )
            processed += 1
        except Exception as e:
            elapsed = time.perf_counter() - start
            print(f"  Error processing {pdf_path.name}: {e}", file=sys.stderr)
            # An exception here (e.g. Docling's own RuntimeError on a
            # conversion failure -- see process_pdf's ConversionStatus.FAILURE
            # check) means pdf_path was never routed anywhere by process_pdf
            # itself, so it's still sitting untouched at its original
            # location with nothing to signal that -- place a COPY in
            # "manual review/" (native file stays exactly where it was) so a
            # reviewer actually finds it instead of it silently sitting in
            # the input folder.
            if pdf_path.exists():
                manual_dir = output_folder / "manual review"
                manual_dir.mkdir(parents=True, exist_ok=True)
                dest_path = _copy_verified(pdf_path, manual_dir / pdf_path.name)
                print(f"  Copied to {dest_path} for manual review (original left in place).")
                notes = f"{e} -- MANUAL REVIEW REQUIRED (copy placed at {dest_path})"
            else:
                notes = f"{e} -- MANUAL REVIEW REQUIRED (original file no longer at {pdf_path})"
            metrics.append(
                args.metrics_sheet_name, METRICS_COLUMNS,
                Timestamp=timestamp,
                File=pdf_path.name, Status="Error", Notes=notes,
                **{"Run Time (s)": round(elapsed, 2)},
            )

    sheet_specs = [
        (args.metrics_sheet_name, METRICS_COLUMNS),
        (LOW_CONFIDENCE_SHEET_NAME, LOW_CONFIDENCE_COLUMNS),
    ]
    for metrics in metrics_logs.values():
        metrics.save(sheet_specs)

    folders_str = ", ".join(str(f) for f in sorted(output_folders_used))
    print(f"\nDone. {processed}/{len(pdf_files)} file(s) processed. Output in: {folders_str}")
    metrics_paths_str = ", ".join(str(p) for p in sorted(metrics_logs.keys()))
    print(f"Metrics log(s): {metrics_paths_str}")


if __name__ == "__main__":
    main()
