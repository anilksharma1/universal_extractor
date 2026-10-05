#!/usr/bin/env python3
"""triage_pipeline.py — combined PDF + Excel triage pipeline.

Runs copy_split_tables.py (Excel) and doc_reader_v2.py (PDF -- Azure Document
Intelligence, with an Azure OpenAI LLM call as a further per-page fallback),
both in this same folder, back-to-back over one mixed input folder, and sorts
every result into one unified "tabular" tree and one unified "not tabular"
tree.

Pipeline
    Stage 1 -- split source_folder's top-level files by type into
               source_folder/pdf/ and source_folder/excel/. Every file that
               lands in either subfolder STAYS there for the rest of the
               run, untouched -- these two subfolders only ever hold the
               original source documents, nothing else.
    Stage 2 -- run copy_split_tables.py's own table-splitting logic (imported
               directly -- no subprocess, no Azure dependency) over
               source_folder/excel/, writing its single_dir/modified_dir
               working copies into a throwaway temp folder rather than
               nesting them inside source_folder/excel/ itself. Every file
               it successfully classifies lands in <output>/tabular/ as a
               PAIR: its noted, one-table-per-sheet output copy, renamed
               "<stem>_intermediate.xlsx", plus a COPY of the native source
               file under its own original name. A file copy_split_tables
               couldn't open or process at all, that has at least one sheet
               with Excel's own "Protect Sheet" turned on (see
               copy_split_tables.py's _protected_sheet_names -- a protected
               sheet might be hiding a row/column copy_split_tables has no
               way to confirm isn't there), OR where no sheet has even one
               real table -- either because it was too small a fragment, too
               fragmented (3+ unmatched skipped fragments on one sheet, or
               more than 10 summed across the whole workbook -- the latter
               overriding even a real table found on some other, cleaner
               sheet), because one sheet alone produced more than 5 tables
               (treated as a likely detection mistake rather than a genuine
               multi-table sheet, also overriding every other sheet in the
               same file), or because the sheet contains a merged cell
               and/or an embedded image and copy_split_tables refuses to run
               table detection on it at all (see copy_split_tables.py's
               _is_real_table and its per-sheet/file-level gates in
               process_file -- e.g. a workbook that's nothing but
               narrative/merged-cell "banner" text or a scanned image, no
               actual grid data anywhere), gets a COPY placed in
               <output>/not tabular/ instead -- the original stays in
               source_folder/excel/ either way; the source file is never
               removed from there, only ever copied out of it.
    Stage 3 -- run doc_reader_v2.py as a subprocess directly against
               source_folder/pdf/ (see stage3_process_pdfs). Unlike a
               relocate-by-delete tool, doc_reader_v2.py never moves or
               deletes the PDF it's given (see its own docstring), so it's
               pointed straight at source_folder/pdf/ itself -- no throwaway
               copy of the input is needed the way Stage 2 needs one for its
               own working files. It needs its own Azure Document
               Intelligence / Azure OpenAI credentials from a .env next to
               it. doc_reader_v2.py is run with --output-folder pointed at a
               throwaway temp folder (never nested inside source_folder/pdf/
               itself) since it writes its own "tabular data"/"non-tabular
               data" subfolders there; this function then moves the
               contents of those two subfolders into <output>/tabular/ and
               <output>/not tabular/ respectively. Every PDF doc_reader_v2.py
               successfully parses lands in "tabular data/" as a PAIR, same
               convention as stage 2: its one-table-per-sheet .xlsx output,
               renamed "<stem>_intermediate.xlsx", plus a COPY of the native
               source PDF under its own original name -- moved wholesale
               into <output>/tabular/. A PDF identified as not tabular
               (no table found, a government/tax form, substantially
               handwritten, low OCR confidence, etc. -- see doc_reader_v2.py's
               own docstring for the full list) gets a COPY in "non-tabular
               data/", moved wholesale into <output>/not tabular/; which
               specific reason applies to a given file is recorded on
               doc_reader_v2.py's own metrics row for it, not by where its
               copy landed. A PDF left entirely unresolved (an unresolved
               DI/LLM-fallback error, not a confirmed "no tables" -- see
               doc_reader_v2.py's own docstring's "Partial" status), or never
               reached at all because the subprocess died first, is detected
               by the absence of a matching output workbook AND a matching
               non-tabular copy; this function then copies the untouched
               original straight from source_folder/pdf/ into
               <output>/not tabular/ itself, reported as "not tabular" rather
               than a distinct "needs reprocessing" outcome. doc_reader_v2.py
               is pointed, via --metrics-path and --metrics-sheet-name,
               straight at <output>/metrics.xlsx's own "PDF Parsing" sheet
               for its persistent run log -- not at a path inside the
               throwaway temp folder -- so its append-then-consolidate
               durability guarantee (see doc_reader_v2.py's MetricsWriter)
               actually holds under this pipeline, and a second run against
               the same --output-folder appends into the same log instead of
               producing a second, differently-named copy. Stage 2 logs its
               own run into the SAME metrics.xlsx, on its own "Excel Parsing"
               sheet (see stage2_process_excels / copy_split_tables.py's
               ExcelMetricsWriter). Either way, the original in
               source_folder/pdf/ is untouched.
    Stage 4 -- no extra step: stages 2-3 already write straight into
               <output>/tabular/ and <output>/not tabular/, so both land as
               siblings inside one output folder.

Usage
    python triage_pipeline.py [source_folder] [--output-folder PATH]

    source_folder   Folder containing a mix of PDFs and Excel files
                     (.xlsx/.xlsm/.xls/.xlsb/.csv), top level only. If
                     omitted, a folder-picker dialog opens (via tkinter),
                     same convention as copy_split_tables.py.
    --output-folder  Defaults to a sibling of source_folder named
                     "<source_folder name> output".

Output layout (inside --output-folder)
    tabular/
        <stem>_intermediate.xlsx   -- parsed output, one per source file
        <native source file>      -- copy of the file that produced it
    not tabular/
        flat, no further subfolders -- PDF and Excel rejects side by side;
        which reason applies to a given file is recorded on its own row in
        metrics.xlsx (see below), not by where its copy landed
    metrics.xlsx
        "PDF Parsing"   -- one row per processed PDF, from doc_reader_v2.py's
                           own MetricsWriter (models used, pages analyzed,
                           OCR confidence, tables found, token usage,
                           estimated cost, run time, status, error)
        "Excel Parsing" -- one row per processed Excel file: timestamp, file
                           name, tables found, run time, status (parsed --
                           single- or multi-table -- or sent for manual
                           review), error, and notes (see
                           copy_split_tables.py's ExcelMetricsWriter)

Handling note: this pipeline's working folders (source_folder/pdf,
source_folder/excel) and its entire output folder may contain real PII/PHI --
the same original documents and extracted data doc_reader_v2.py and
copy_split_tables.py themselves handle. Keep them in approved secure storage
only; never commit them to source control. copy_split_tables.py's own
Excel-COM formula-recalculation temp copies (see its _fill_blank_formula_
values / _recalc_tmp_dir_for_pid) are redirected into this output folder's
own storage controls too, same as every other working copy here -- see
stage2_process_excels' own temp_base_dir=tmp.
"""
from __future__ import annotations

import argparse
import importlib.util
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:
    tk = None

_HERE = Path(__file__).resolve().parent

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"}


def _pick_folder(initial_dir: str) -> "Path | None":
    """Open a folder-picker dialog, same convention as copy_split_tables.py's
    own _pick_folder / doc_reader_v2.py's select_pdf_files. Returns None if
    tkinter isn't available or the user cancels."""
    if tk is None:
        return None
    root = None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(
            initialdir=initial_dir, title="Select folder containing PDFs and Excel files to triage"
        )
    finally:
        if root is not None:
            root.destroy()
    return Path(folder) if folder else None


def _load_copy_split_tables():
    """Import copy_split_tables.py from this same folder by file path,
    regardless of the current working directory the pipeline happens to be
    invoked from -- rather than a plain top-level `import copy_split_tables`,
    which would move the import to this module's own load time, before
    main()'s Stage 2 try/except exists to catch a failure there (a missing
    dependency, say) and let Stage 1's results / Stage 3 still proceed. This
    keeps that failure lazy and scoped to Stage 2 alone, exactly where it's
    called from below.

    Registered into sys.modules under its own spec.name BEFORE exec_module
    runs (the same order Python's own import system uses, in case the
    module's own top-level code ever re-enters its own import) -- this is
    the one thing this function's split_tables.py predecessor didn't need
    and got away without: stage2_process_excels below calls
    copy_split_tables.py's own multiprocessing-based
    _process_file_with_timeout, and Windows' spawn start method pickles that
    call's target function by looking it up as
    sys.modules[func.__module__].<func.__qualname__> -- a module loaded via
    module_from_spec + exec_module alone is never entered into sys.modules
    at all, so that lookup fails and mp.Process.start() raises a
    PicklingError before a child process is ever created. A plain
    module_from_spec(spec) call three lines below does NOT run the module's
    top-level code by itself -- exec_module does that -- so registering the
    (still nearly-empty) module object first and running exec_module second
    is safe and matches CPython's own import order."""
    spec = importlib.util.spec_from_file_location("copy_split_tables", _HERE / "copy_split_tables.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_NUMBERED_DEST_SUFFIX_RE = re.compile(r"^(.*) \((\d+)\)$")


def _highest_existing_dest_suffix(dest_dir: Path, stem: str, suffix: str) -> int:
    """Scan dest_dir ONCE (a single directory listing) to find the highest
    already-used " (N)" numbered variant of stem+suffix -- lets
    unique_dest's reservation loop seed its counter just past whatever's
    already there, instead of walking up one number at a time via a failed
    os.open() syscall for every number already taken.

    Re-running this pipeline into an already-populated output folder meant
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


def unique_dest(dest_dir: Path, name: str) -> Path:
    """Atomically claim a destination path in dest_dir that doesn't collide
    with an existing file, even across two concurrently running pipeline
    instances writing into the same output folder -- auto-renaming with a
    " (2)", " (3)", ... suffix, same convention as sort_files_by_type.py /
    extract_files_by_extension.py in this repo.

    A plain "does candidate exist? if not, use it" check (the previous
    implementation of this function) is a textbook check-then-act race: two
    processes can both see the same free candidate and both proceed to write
    it, with the loser's write silently clobbering the winner's -- and since
    move_into's source file is already gone (moved, not copied) by that
    point, the loser's deliverable AND its source are both destroyed with no
    recovery. os.O_CREAT | os.O_EXCL asks the OS to create-or-fail
    atomically, so at most one process can ever win a given candidate name;
    every other racer gets FileExistsError and moves on to the next
    candidate instead of clobbering. Mirrors doc_reader_v2.py's own
    _reserve_unique_destination.

    Returns a path to an already-created (empty) file -- the caller must
    write into that same path (move_into / copy_into do), never re-resolve a
    fresh one, or the reservation is pointless."""
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


def move_into(src: Path, dest_dir: Path) -> Path:
    dest = unique_dest(dest_dir, src.name)
    shutil.move(str(src), str(dest))
    return dest


def copy_into(src: Path, dest_dir: Path) -> Path:
    """Like move_into, but leaves src in place -- for a file that must stay
    in the pipeline's input folder (pdf/ or excel/) even after a copy of it
    has been sent to the 'not tabular' output folder."""
    dest = unique_dest(dest_dir, src.name)
    shutil.copy2(str(src), str(dest))
    return dest


def move_into_renamed(src: Path, dest_dir: Path, dest_name: str) -> Path:
    """Like move_into, but under an explicit destination filename rather than
    src's own name -- used to land a parsed workbook next to its native
    source file under the "<stem>_intermediate.xlsx" naming convention (see
    tabular_dir in main)."""
    dest = unique_dest(dest_dir, dest_name)
    shutil.move(str(src), str(dest))
    return dest


def _cleanup_temp_dir(path: Path, label: str) -> None:
    """Best-effort delete of a temp folder that held PII/PHI copies (source
    PDFs/Excel files, or a fully-recalculated workbook) -- warns loudly on
    stderr instead of silently leaving a full copy behind if something (e.g.
    Excel still holding a file handle right after a COM SaveAs) prevents
    complete removal. Mirrors doc_reader_v2.py's own _cleanup_images_dir.
    Note this is still best-effort only: nothing runs this on a hard kill
    (SIGKILL/process crash), same limitation _cleanup_images_dir has."""
    if not path.exists():
        return
    shutil.rmtree(path, ignore_errors=True)
    if path.exists():
        print(
            f"WARNING: could not fully delete temp folder containing PII/PHI ({label}): {path}",
            file=sys.stderr,
        )


# "copy_split_tables_excel_recalc_pid" is excel_mod's own Excel-COM
# recalculation temp dir name (see stage2_process_excels' own docstring) --
# unlike stage2_process_excels' own "excel_*" working-copy folder below, it
# is NEVER created under output_folder (excel_mod has no temp_base_dir
# parameter to redirect it there), so only the OS-temp-dir call to this
# function below ever actually finds one; listed here anyway since both
# calls share this one prefix list.
_ORPHANED_TEMP_DIR_PREFIXES = ("excel_", "doc_reader_", "copy_split_tables_excel_recalc_pid")


def _sweep_orphaned_temp_dirs(base_dir: Path, label: str) -> None:
    """Best-effort delete of any leftover excel_*/doc_reader_*/
    copy_split_tables_excel_recalc_pid* folder found directly under base_dir
    -- each one is a working copy of source PDFs/Excel files (or a
    fully-recalculated workbook) created by stage2_process_excels/
    stage3_process_pdfs/excel_mod's own _fill_blank_formula_values (see
    their own mkdtemp calls, all pointed at output_folder precisely so
    cleanup stays inside approved storage -- except excel_mod's own
    recalculation temp dir, which is NOT redirectable and always lands in
    the OS temp dir instead; see stage2_process_excels' own docstring).
    Those functions already delete their own folder via _cleanup_temp_dir
    when they finish normally; this sweep only ever finds
    something when a PREVIOUS run was hard-killed (SIGKILL/crash/power loss)
    before its own `finally` block could run -- _cleanup_temp_dir has no
    reach in that case, so the folder (real PII/PHI) is left sitting on disk
    indefinitely until something like this runs. Called once at startup
    against both output_folder and the OS temp directory, so a folder
    orphaned before this sweep existed (when these dirs still landed in the
    OS temp dir unconditionally) gets caught too, not just future ones."""
    if not base_dir.exists():
        return
    for prefix in _ORPHANED_TEMP_DIR_PREFIXES:
        for orphan in base_dir.glob(f"{prefix}*"):
            if not orphan.is_dir():
                continue
            shutil.rmtree(orphan, ignore_errors=True)
            if orphan.exists():
                print(
                    f"WARNING: found orphaned temp folder from a previous run, possibly "
                    f"containing PII/PHI, but could not fully delete it ({label}): {orphan}",
                    file=sys.stderr,
                )
            else:
                print(f"Cleaned up orphaned temp folder from a previous run ({label}): {orphan}")


def move_tree_contents(src_dir: Path, dest_dir: Path) -> int:
    """Move every top-level entry out of src_dir into dest_dir. Returns the
    count moved; a no-op (returns 0) if src_dir doesn't exist."""
    if not src_dir.is_dir():
        return 0
    count = 0
    for entry in sorted(src_dir.iterdir()):
        move_into(entry, dest_dir)
        count += 1
    return count


# Mirrors doc_reader_v2.py's OWN _NUMBERED_DEST_SUFFIX_RE -- its
# _reserve_unique_destination appends "_N" (not this file's own " (N)"
# convention above -- that one's for triage_pipeline's own unique_dest, a
# completely separate reservation happening on the OUTPUT side of this
# pipeline, not doc_reader_v2.py's) when its own output folder already has a
# file under a given stem. Needed by _pdf_parsed_stems and _matched_stems
# below to recognize a collision-renamed file for what it is, since
# doc_reader_v2.py runs as a subprocess here (see stage3_process_pdfs) and
# can't hand back the out_path its own process_pdf already returns
# internally -- this has to be re-derived from the directory listing instead.
_DOC_READER_NUMBERED_STEM_RE = re.compile(r"^(.*)_(\d+)$")

# doc_reader_v2.py's own tabular-success path names its output workbook
# "<stem>_intermediate.xlsx" ITSELF (see process_pdf's own tabular_dir
# branch) -- unlike the old only_docling.py wrapper, whose plain "<stem>.xlsx"
# this pipeline used to rename to "_intermediate.xlsx" on the way into
# tabular_dir. A workbook's stem is therefore "<pdf_stem>_intermediate", or,
# on a doc_reader_v2.py-side collision (_reserve_unique_destination appends
# "_N" to the WHOLE stem it's given, i.e. after "_intermediate" is already
# part of it), "<pdf_stem>_intermediate_N" -- both stripped by this pattern
# to recover the original pdf_stem.
_INTERMEDIATE_STEM_RE = re.compile(r"^(.*)_intermediate(?:_\d+)?$")


def _pdf_parsed_stems(tabular_src: Path, pdf_stems: set) -> set:
    """Returns the subset of pdf_stems that doc_reader_v2.py successfully
    parsed into tabular_src (its own "tabular data/") -- recognized via its
    own "<stem>_intermediate.xlsx" naming convention (see
    _INTERMEDIATE_STEM_RE). stage3_process_pdfs moves tabular_src's contents
    into this pipeline's tabular_dir wholesale (every file there is already
    a correctly-named half of a pair doc_reader_v2.py wrote itself, native
    PDF copy included -- nothing to re-pair); this function only answers
    "which source PDFs does that cover", so the caller can tell those apart
    from a genuinely unresolved PDF in its own not-tabular sweep."""
    if not tabular_src.is_dir():
        return set()
    parsed: set = set()
    for xlsx_path in tabular_src.glob("*.xlsx"):
        m = _INTERMEDIATE_STEM_RE.match(xlsx_path.stem)
        if m and m.group(1) in pdf_stems:
            parsed.add(m.group(1))
    return parsed


def _matched_stems(dir_path: Path, suffix: str, pdf_stems: set) -> set:
    """Like _pdf_parsed_stems, but for a flat reject folder (e.g.
    doc_reader_v2.py's own "non-tabular data/") where every file is a plain
    COPY of a source PDF under its own name -- no "_intermediate" tag to
    strip, just "does a copy for this pdf_stem exist here, allowing for
    doc_reader_v2.py's own '_N' collision suffix." Returns the subset of
    pdf_stems already accounted for in dir_path, so stage3_process_pdfs can
    tell which source PDFs still need a copy sent to not_tabular_dir itself
    versus which are already covered by doc_reader_v2.py's own output."""
    if not dir_path.is_dir():
        return set()
    present: set = set()
    for path in dir_path.glob(f"*{suffix}"):
        if path.stem in pdf_stems:
            present.add(path.stem)
            continue
        m = _DOC_READER_NUMBERED_STEM_RE.match(path.stem)
        if m and m.group(1) in pdf_stems:
            present.add(m.group(1))
    return present


# --------------------------------------------------------------------------
# Stage 1 -- split source_folder's top-level files by type
# --------------------------------------------------------------------------

def stage1_split_by_type(source: Path) -> tuple[Path, Path, list[str]]:
    """Move source's top-level .pdf files into source/pdf/ and its top-level
    Excel files into source/excel/. Anything else at the top level is left in
    place and returned in `skipped` so nothing is silently dropped. Safe to
    re-run: only acts on files currently sitting at source's top level, never
    descends into pdf/ or excel/ themselves."""
    pdf_dir = source / "pdf"
    excel_dir = source / "excel"
    pdf_dir.mkdir(exist_ok=True)
    excel_dir.mkdir(exist_ok=True)

    skipped: list[str] = []
    for entry in sorted(source.iterdir()):
        if entry in (pdf_dir, excel_dir) or not entry.is_file():
            continue
        if entry.name.startswith("~$"):
            continue
        suffix = entry.suffix.lower()
        if suffix == ".pdf":
            move_into(entry, pdf_dir)
        elif suffix in EXCEL_SUFFIXES:
            move_into(entry, excel_dir)
        else:
            skipped.append(entry.name)

    return pdf_dir, excel_dir, skipped


# --------------------------------------------------------------------------
# Stage 2 -- Excel, via copy_split_tables.py (imported directly)
# --------------------------------------------------------------------------

# Reviewer-facing Status text for the "Excel Parsing" metrics sheet (see
# stage2_process_excels) -- collapses excel_mod.process_file's own internal
# status vocabulary ("single"/"multi"/"error"/"skipped"/"protected"/
# "no_tables") down to what a reviewer actually wants to know at a glance:
# whether the file was parsed (and how), or sent for manual review. The
# Error/Notes columns carry the specific reason in the manual-review case --
# for "no_tables" specifically, that no sheet had anything excel_mod's own
# table detector accepted as a real table -- too small a fragment, too
# fragmented (per-sheet or whole-workbook unmatched-fragment ceiling), one
# sheet alone producing an implausible number of tables (likely a detection
# mistake), or a whole sheet skipped outright for containing a merged cell
# and/or an embedded image (see its _is_real_table and process_file's own
# per-sheet/file-level gates) -- not that the file failed to open or process.
_EXCEL_STATUS_LABELS = {
    "single": "Parsed - Single-table",
    "multi": "Parsed - Multi-table",
    "error": "Sent for Manual Review",
    "skipped": "Sent for Manual Review",
    "protected": "Sent for Manual Review",
    "no_tables": "Sent for Manual Review",
}


def stage2_process_excels(
    excel_dir: Path, tabular_dir: Path, not_tabular_dir: Path, output_folder: Path, excel_mod, log
) -> dict:
    """excel_dir is the pipeline's input folder -- every file already there
    stays there, untouched, no matter what happens to it below. excel_mod
    still needs somewhere to write its single_dir/modified_dir working copies
    (see process_file); those go in a throwaway temp folder instead of nested
    inside excel_dir itself, so excel_dir never ends up holding anything but
    the original source files.

    Every source file excel_mod successfully parses lands in tabular_dir
    as a PAIR: the parsed workbook, renamed "<stem>_intermediate.xlsx", plus a
    COPY of the native source file under its own original name -- so a
    reviewer opening tabular_dir always has both side by side. Which of
    single_dir/modified_dir a given call actually wrote into is read straight
    off ProcessResult.dest -- process_file already knows exactly which path
    it saved to, so there's no need to re-derive or guess it here. (An
    earlier version of this function instead diffed each directory's listing
    before/after the call: besides needlessly re-deriving something
    process_file already had on hand, that had two real failure modes --
    a "single"/"multi" result whose save somehow didn't produce a new file
    left the diff empty, with the file then logged as "Parsed" while no
    workbook existed anywhere for it and nothing downstream noticed; and if
    more than one new file had ever appeared in that diff, an arbitrary
    set.pop() would have silently paired the wrong workbook to the wrong
    source instead of detecting the mismatch.)

    Every file's outcome -- timestamp, table count, run time, reviewer-facing
    status, error, and any note (see process_file's ProcessResult) -- is
    durably logged as one row on output_folder/metrics.xlsx's "Excel Parsing"
    sheet (see excel_mod.ExcelMetricsWriter), the same shared workbook
    doc_reader_v2.py's own MetricsWriter logs the PDF side into under its own
    "PDF Parsing" sheet (see stage3_process_pdfs) -- one workbook covers the
    whole pipeline's run history.

    Calls excel_mod's own _process_file_with_timeout, not its plain
    process_file -- runs each file in its own short-lived child process with
    a hard wall-clock kill limit (see PROCESS_FILE_TIMEOUT_SECONDS in
    copy_split_tables.py), so one stuck file (e.g. hung Excel COM automation
    -- see _fill_blank_formula_values) can't block the rest of this batch the
    way a bare process_file() call would. Only safe to call this way because
    _load_copy_split_tables() registers copy_split_tables.py into
    sys.modules before exec'ing it -- see that function's own docstring for
    why an unregistered module breaks multiprocessing's spawn start method
    on Windows. temp_base_dir=tmp below redirects the disposable
    Excel-recalculated copy _fill_blank_formula_values creates for a
    .xlsx/.xlsm file with an uncached formula value into this same
    output_folder-scoped working folder, instead of the OS temp dir."""
    files = excel_mod.collect_input_files(excel_dir)
    counts = {"single": 0, "multi": 0, "error": 0, "skipped": 0, "protected": 0, "no_tables": 0}
    unparsed_files: list[Path] = []
    n_parsed = 0

    metrics = excel_mod.ExcelMetricsWriter(output_folder / "metrics.xlsx")

    # dir=output_folder keeps this working copy of every source Excel file
    # (real PII/PHI) inside output_folder's own storage controls instead of
    # the OS temp dir -- see _sweep_orphaned_temp_dirs in main() for the
    # backstop if a hard kill still leaves it behind.
    tmp = Path(tempfile.mkdtemp(prefix="excel_", dir=output_folder))
    try:
        single_dir = tmp / excel_mod.SINGLE_DIRNAME
        modified_dir = tmp / excel_mod.MODIFIED_DIRNAME
        single_dir.mkdir()
        modified_dir.mkdir()

        try:
            for idx, path in enumerate(files, 1):
                log(f"\n[Excel {idx}/{len(files)}] {path.name}")
                start = time.perf_counter()
                result = excel_mod._process_file_with_timeout(
                    path, single_dir, modified_dir, log_fn=log, temp_base_dir=tmp
                )
                elapsed = time.perf_counter() - start

                status = result.status
                error = result.error
                if status in ("single", "multi") and result.dest is None:
                    # process_file's own success path always sets dest to the
                    # workbook it just saved -- getting here means it
                    # returned "single"/"multi" without one, a contract
                    # violation this function has no output to harvest for.
                    # Reported as an error and routed to manual review
                    # rather than trusted at face value, which would
                    # otherwise log this file as successfully parsed while
                    # producing nothing anywhere for it (see this function's
                    # own docstring).
                    log(
                        f"  ERROR: process_file reported '{status}' for {path.name} but returned "
                        "no output path -- treating as an error and sending to manual review "
                        "instead of losing it silently."
                    )
                    status = "error"
                    error = error or f"process_file reported '{result.status}' with no output path"

                if status in ("single", "multi"):
                    # Moved into tabular_dir HERE -- before this file's status
                    # is counted or its metrics row is written below -- not
                    # deferred to a separate loop after every file in the
                    # batch has already been parsed (this function's previous
                    # shape). That deferred shape let metrics.append() commit
                    # a "Parsed" row for every successfully parsed file while
                    # their workbooks were still sitting unmoved in `tmp`; if
                    # the move loop then raised partway through a LATER file,
                    # the outer `finally`'s _cleanup_temp_dir below deletes
                    # the whole temp tree -- destroying every not-yet-moved
                    # workbook even though metrics.xlsx had already durably
                    # recorded each one as parsed, with no way for a reviewer
                    # to tell from metrics.xlsx alone that tabular_dir doesn't
                    # actually hold them. Moving first means a move failure
                    # is caught and reflected in THIS file's own status
                    # before it's ever recorded, and confines the failure to
                    # this one file instead of risking the rest of the batch.
                    try:
                        move_into_renamed(
                            result.dest, tabular_dir, f"{path.stem}_intermediate.xlsx"
                        )
                        copy_into(path, tabular_dir)
                        n_parsed += 1
                    except Exception as move_exc:
                        log(
                            f"  ERROR: parsed '{path.name}' successfully but could not move its "
                            f"output into '{tabular_dir}' ({move_exc}) -- sending the source to "
                            "manual review instead of losing it silently."
                        )
                        status = "error"
                        error = f"Move to tabular_dir failed: {move_exc}"

                counts[status] = counts.get(status, 0) + 1
                if status in ("error", "skipped", "protected", "no_tables"):
                    unparsed_files.append(path)
                metrics.append(**{
                    "Timestamp (UTC)": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                    "File Name": path.name,
                    "File Size (KB)": round(path.stat().st_size / 1024, 2),
                    "Tables Found": result.tables_found,
                    "Run Time (s)": round(elapsed, 2),
                    "Status": _EXCEL_STATUS_LABELS.get(status, status),
                    "Error": error,
                    "Notes": result.note,
                })
        finally:
            # Quits the shared hidden Excel instance excel_mod may have
            # started to recalculate formula workbooks in this batch -- a
            # no-op if it was never needed. In `finally` so a mid-batch error
            # still doesn't leave an orphan EXCEL.EXE process running.
            excel_mod.close_excel_app()
    finally:
        # tmp held every parsed/unparsed Excel working copy (real PII/PHI) --
        # verify the delete actually took (see _cleanup_temp_dir) rather than
        # trusting ignore_errors=True silently, since a file excel_mod just
        # saved via openpyxl could still be momentarily locked.
        _cleanup_temp_dir(tmp, "copy_split_tables working copies")
        # Consolidates every row buffered by metrics.append() above into
        # metrics.xlsx in one single write -- called once per batch, in
        # `finally`, so a mid-batch crash still gets everything durably
        # written up to the last successfully processed file (see
        # ExcelMetricsWriter's own docstring on why append() alone already
        # makes each row crash-safe before save() ever runs). Wrapped in its
        # own try/except: if the try block above is ALREADY propagating an
        # exception (a real bug, not a metrics problem) and this save() then
        # also raises (e.g. a still-contended metrics.xlsx.lock, or
        # metrics.xlsx open in Excel), Python's own finally semantics would
        # let save()'s exception silently REPLACE the original one -- the
        # caller would see a confusing metrics/locking error instead of the
        # actual bug that was crashing this stage. When that's the case,
        # this failure is only logged, not raised, so the original exception
        # keeps propagating; when nothing else is already failing, save()'s
        # own exception is the only error there is, so it's raised as normal.
        #
        # primary_exception_in_flight is captured BEFORE calling save(), not
        # inside the except block below -- sys.exc_info() inside an except
        # clause reports the exception THAT clause is handling (save_exc
        # itself), which would make this always look "already failing" and
        # never re-raise a save()-only error (see Case 3 in this fix's own
        # verification). Capturing it beforehand reads the state as it
        # actually was when save() was called.
        primary_exception_in_flight = sys.exc_info()[0] is not None
        try:
            metrics.save()
        except Exception as save_exc:
            if primary_exception_in_flight:
                print(
                    f"WARNING: metrics.save() also failed while another error was already "
                    f"propagating -- {type(save_exc).__name__}: {save_exc}. Pending rows remain "
                    "in the sidecar for the next run's save() to pick up.",
                    file=sys.stderr,
                )
            else:
                raise

    # A copy goes to the output "not tabular" folder; the original stays
    # in excel_dir (the input folder) -- see copy_into. Guarded the same way
    # stage3_process_pdfs' own equivalent loop is: left unguarded, a failure
    # on one file (disk full, AV lock, network-share hiccup) would propagate
    # out of this loop entirely, leaving every later file in unparsed_files
    # silently never copied even though each one's metrics row above already
    # claims "Sent for Manual Review".
    for path in unparsed_files:
        if path.exists():
            try:
                copy_into(path, not_tabular_dir)
            except Exception as copy_exc:
                log(
                    f"  ERROR: could not copy '{path.name}' to '{not_tabular_dir}' "
                    f"({copy_exc}) -- it may be missing from this run's output."
                )

    counts["parsed_files"] = n_parsed
    return counts


# --------------------------------------------------------------------------
# Stage 3 -- PDF, via doc_reader_v2.py (subprocess; Azure Document
# Intelligence, with its own Azure OpenAI LLM call as a further per-page
# fallback)
# --------------------------------------------------------------------------

def stage3_process_pdfs(pdf_dir: Path, tabular_dir: Path, not_tabular_dir: Path, output_folder: Path, log) -> dict:
    """pdf_dir is the pipeline's input folder -- every file already there
    stays there, untouched, no matter which category doc_reader_v2.py sorts
    it into. Unlike a relocate-by-delete tool, doc_reader_v2.py never moves
    or deletes the PDF it's given (see its own docstring, "The source PDF is
    NEVER moved or deleted, no matter the outcome") -- every folder it
    writes to only ever receives a COPY -- so pdf_dir is passed to it
    DIRECTLY, with no throwaway input copy needed the way stage2 needs one
    for its own working files.

    doc_reader_v2.py's own --output-folder writes two subfolders:
    "tabular data/" (a PAIR per successfully parsed PDF -- the parsed
    workbook, renamed "<stem>_intermediate.xlsx", plus a copy of the native
    source PDF) and "non-tabular data/" (a plain copy of every PDF it didn't
    produce a workbook for -- no table found, a government/tax form,
    substantially handwritten, OCR confidence too low, etc. -- see
    doc_reader_v2.py's own docstring for the full list of reasons, each
    recorded on that file's own metrics row rather than by subfolder). Those
    two subfolder NAMES are doc_reader_v2.py's own convention, not this
    pipeline's -- --output-folder here points at a throwaway temp folder
    (never inside pdf_dir itself) precisely so this function can move each
    subfolder's contents into this pipeline's own tabular_dir/not_tabular_dir
    afterward, merging with whatever stage2 already put there.

    A PDF doc_reader_v2.py leaves genuinely unresolved -- an unresolved
    DI/LLM-fallback error rather than a confirmed "no tables" (its own
    "Partial" status; see its docstring) -- gets neither a workbook in
    "tabular data/" nor a copy in "non-tabular data/". The same is true of
    any PDF the subprocess never got to at all because it died partway
    through the batch. Both are detected here by the absence of a matching
    output in either subfolder, and handled identically: since pdf_dir was
    never touched, this function copies the untouched ORIGINAL straight from
    pdf_dir into not_tabular_dir itself, reported as "not tabular" rather
    than a distinct "needs reprocessing" outcome.

    --metrics-path is pointed straight at output_folder / "metrics.xlsx" --
    NOT at anything inside the temp folder above -- with --metrics-sheet-name
    "PDF Parsing" so its rows land on their own sheet inside that shared
    workbook, next to stage2_process_excels' own "Excel Parsing" sheet (see
    excel_mod.ExcelMetricsWriter) rather than colliding with it.
    doc_reader_v2.py's MetricsWriter durably appends each processed file's
    row to a small sidecar immediately and only consolidates it into the
    xlsx at the very end of its run (see its own docstring) -- specifically
    so a mid-batch crash loses nothing. That guarantee is void if the
    sidecar itself lives inside a temp tree this function deletes in
    `finally` regardless of whether doc_reader_v2.py's subprocess exited
    cleanly: a crash would take the sidecar down with it, destroying the
    exact audit trail it exists to protect. Writing straight into
    output_folder also means a second run against the same --output-folder
    appends into the SAME metrics.xlsx (MetricsWriter merges into whatever's
    already there) instead of racing triage_pipeline.py's own unique_dest
    into a second, differently-named copy.

    A non-zero subprocess exit does NOT skip the harvest below, and does
    NOT raise -- doc_reader_v2.py processes one PDF at a time and durably
    writes each one's result as it goes (same append-then-consolidate
    guarantee as its own MetricsWriter), so whatever it completed before
    dying is already sitting in the temp output folder exactly like a clean
    run's output; only the PDF(s) it never got to are simply absent from
    both "tabular data/" and "non-tabular data/" -- the same signal this
    function already uses for its own "Partial"/never-reached case above, so
    no separate handling is needed. The caller is told about the failure via
    the returned "subprocess_returncode" key instead, once harvesting has
    actually run."""
    pdf_files = sorted(pdf_dir.glob("*.pdf"))
    if not pdf_files:
        log(f"\nNo PDFs found in {pdf_dir} -- skipping doc_reader_v2.py.")
        return {"parsed": 0, "not_tabular": 0, "subprocess_returncode": 0}

    doc_reader_script = _HERE / "doc_reader_v2.py"

    # dir=output_folder keeps this temp output (real PII/PHI) inside
    # output_folder's own storage controls instead of the OS temp dir -- see
    # _sweep_orphaned_temp_dirs in main() for the backstop if a hard kill
    # still leaves it behind.
    tmp_path = Path(tempfile.mkdtemp(prefix="doc_reader_", dir=output_folder))
    try:
        doc_reader_output = tmp_path / "doc_reader_output"
        metrics_path = output_folder / "metrics.xlsx"

        log(f"\nRunning doc_reader_v2.py on {len(pdf_files)} PDF(s) in {pdf_dir} ...")
        result = subprocess.run(
            [
                sys.executable, str(doc_reader_script), str(pdf_dir),
                "--output-folder", str(doc_reader_output),
                "--metrics-path", str(metrics_path),
                "--metrics-sheet-name", "PDF Parsing",
            ]
        )
        if result.returncode != 0:
            log(
                f"WARNING: doc_reader_v2.py exited with code {result.returncode} -- see its "
                "output above for details. Harvesting whatever it completed before failing; "
                "any PDF it never got to is sent to 'not tabular' below, same as an unresolved "
                "per-file error."
            )

        tabular_src = doc_reader_output / "tabular data"
        non_tabular_src = doc_reader_output / "non-tabular data"

        # Captured BEFORE moving anything out of tabular_src/non_tabular_src
        # -- _pdf_parsed_stems/_matched_stems tell apart "doc_reader_v2.py
        # already accounted for this source PDF" from "genuinely unresolved"
        # for the sweep further down; they don't gate the moves below at
        # all, since every file already sitting in either subfolder is
        # already correctly named and paired by doc_reader_v2.py itself
        # (see stage3_process_pdfs' own docstring) and gets moved wholesale
        # either way.
        pdf_stems = {p.stem for p in pdf_files}
        parsed_stems = _pdf_parsed_stems(tabular_src, pdf_stems)
        non_tabular_stems = _matched_stems(non_tabular_src, ".pdf", pdf_stems)

        # Every file moved below (out of the temp folder) is already a copy,
        # once removed from the real pdf_dir -- pdf_dir itself was never
        # touched, so moving rather than copying here costs nothing.
        # not_tabular_dir itself has no further nested subfolders -- every
        # reason a PDF ends up here is distinguished only via
        # doc_reader_v2.py's own metrics row for that file, not by where its
        # copy landed.
        move_tree_contents(tabular_src, tabular_dir)
        n_not_tabular = move_tree_contents(non_tabular_src, not_tabular_dir)
        # n_parsed counts source PDFs, not files moved -- move_tree_contents'
        # own count above is 2 files (native copy + "_intermediate.xlsx")
        # per successfully parsed PDF, which would double the reported
        # "parsed" total in the run summary.
        n_parsed = len(parsed_stems)

        # Every source PDF whose stem isn't in parsed_stems and isn't
        # already covered by non_tabular_stems is either left alone by
        # doc_reader_v2.py for an unresolved per-file error (its own
        # "Partial" case), or was never processed at all because the
        # subprocess died first (see the returncode check above) -- both are
        # indistinguishable from here, and neither is a confirmed "needs
        # reprocessing" outcome, so both simply get a COPY of the untouched
        # original sent to not_tabular_dir (flat, same as every other reject
        # above) and are reported as "not tabular".
        n_unresolved = 0
        for pdf_path in pdf_files:
            if pdf_path.stem in parsed_stems or pdf_path.stem in non_tabular_stems:
                continue
            try:
                copy_into(pdf_path, not_tabular_dir)
            except Exception as copy_exc:
                log(
                    f"  ERROR: could not copy '{pdf_path.name}' to '{not_tabular_dir}' "
                    f"({copy_exc}) -- it may be missing from this run's output."
                )
            n_unresolved += 1
        n_not_tabular += n_unresolved

        # Nothing to move for metrics.xlsx -- doc_reader_v2.py already wrote
        # (and appended) straight into metrics_path (see above), which lives
        # in output_folder itself, not in the temp tree being cleaned up
        # below.
    finally:
        # tmp_path held doc_reader_v2.py's full working output for this run
        # (real PII/PHI) -- verify the delete actually took (see
        # _cleanup_temp_dir) rather than trusting ignore_errors=True
        # silently.
        _cleanup_temp_dir(tmp_path, "doc_reader_v2.py PDF output")

    return {
        "parsed": n_parsed,
        "not_tabular": n_not_tabular,
        "subprocess_returncode": result.returncode,
    }


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Combined PDF + Excel triage pipeline: splits a mixed folder by file "
                    "type, runs copy_split_tables.py over the Excel files and doc_reader_v2.py "
                    "(Azure Document Intelligence, with its own Azure OpenAI LLM call as a "
                    "further per-page fallback) over the PDFs, and sorts every result into a "
                    "unified 'tabular' / 'not tabular' output tree."
    )
    parser.add_argument(
        "source_folder",
        nargs="?",
        default=None,
        help="Folder containing a mix of PDFs and Excel files (top level only). "
             "If omitted, a folder-picker dialog opens."
    )
    parser.add_argument(
        "--output-folder",
        default=None,
        help="Where tabular / not-tabular output goes. Defaults to a sibling of "
             "source_folder named '<source_folder name> output'."
    )
    args = parser.parse_args(argv)

    if args.source_folder:
        source = Path(args.source_folder)
    else:
        picked = _pick_folder(str(_HERE))
        if picked is None:
            print(
                "ERROR: no source folder given and the folder-picker is unavailable "
                "(tkinter missing) or was cancelled. Pass a folder path as an argument.",
                file=sys.stderr,
            )
            return 2
        source = picked

    source = source.resolve()
    if not source.is_dir():
        print(f"ERROR: not a folder: {source}", file=sys.stderr)
        return 2

    output_folder = (
        Path(args.output_folder).resolve() if args.output_folder
        else source.parent / f"{source.name} output"
    )
    # Created up front, before stage2/stage3's own tempfile.mkdtemp(dir=output_folder)
    # calls below need it to already exist -- mkdtemp does not create its
    # parent -- rather than relying on ExcelMetricsWriter's incidental
    # mkdir(parents=True) as a side effect of Stage 2 happening to run first.
    output_folder.mkdir(parents=True, exist_ok=True)

    tabular_dir = output_folder / "tabular"
    not_tabular_dir = output_folder / "not tabular"

    print(
        "NOTE: this pipeline's working folders and output may contain real PII/PHI "
        "(source documents and extracted data). Keep them in approved secure storage "
        f"only -- do not commit '{source}/pdf', '{source}/excel', or '{output_folder}' "
        "to source control.\n"
    )

    # Working copies of source PDFs/Excel files (see stage2_process_excels/
    # stage3_process_pdfs -- NOT excel_mod's own _fill_blank_formula_values,
    # whose recalculation temp dir always lands in the OS temp dir instead;
    # see stage2_process_excels' own docstring) are created via
    # tempfile.mkdtemp(dir=output_folder) below, precisely so a hard kill
    # leaves them inside output_folder's own storage controls rather than
    # the OS temp dir. That only protects folders created AFTER
    # this fix, though -- sweep for anything a previous crashed run (before
    # or after this fix) already left behind, in both places such a folder
    # could exist.
    _sweep_orphaned_temp_dirs(output_folder, "output folder")
    _sweep_orphaned_temp_dirs(Path(tempfile.gettempdir()), "OS temp dir")

    # Stages 1, 2, and 3 are each wrapped in their own try/except -- a
    # failure in any one of them (e.g. Stage 1 hitting a locked file, a
    # permission error, or a network-share hiccup while moving files --
    # this is the stage that actually touches the filesystem the most, and
    # the most failure-prone of the three -- or Stage 2/3's own scaffolding
    # failing, an openpyxl version mismatch, a TimeoutError from a stale
    # metrics lock) would otherwise abort main() entirely before either
    # remaining stage ever runs: every PDF (or every Excel file) would then
    # be silently never processed, with no summary and nothing but a
    # traceback to say why. process_file/process_pdf already get this right
    # per FILE inside each stage; this is the same guarantee one level up,
    # per STAGE. The summary is printed from `finally` so it always
    # appears, even when one or more stages failed outright.
    excel_counts: dict = {}
    pdf_counts: dict = {}
    stage1_error: Exception | None = None
    stage2_error: Exception | None = None
    stage3_error: Exception | None = None
    doc_reader_returncode = 0

    print(f"=== Stage 1: splitting {source} by file type ===")
    try:
        pdf_dir, excel_dir, skipped = stage1_split_by_type(source)
        if skipped:
            print(
                f"WARNING: {len(skipped)} file(s) matched neither .pdf nor a supported Excel "
                f"extension and were left at the source root: {', '.join(skipped)}"
            )
    except Exception as e:
        stage1_error = e
        # stage1_split_by_type creates pdf_dir/excel_dir (both simple,
        # deterministic paths derived from `source`) and only moves files
        # into them AFTER that -- so even a failure partway through the
        # move loop (the likeliest failure point: a locked file, a
        # permission error, a network-share hiccup on one specific file)
        # leaves both directories in place, already holding whatever got
        # moved before it failed. Recomputing them here rather than giving
        # up on Stage 2/3 entirely lets both stages still run against
        # whatever Stage 1 DID manage to sort before hitting the error --
        # the same "don't let one failure silently skip everything else"
        # guarantee this whole three-stage try/except structure exists for.
        pdf_dir = source / "pdf"
        excel_dir = source / "excel"
        print(
            f"ERROR: Stage 1 (splitting {source} by file type) failed partway through -- "
            f"{type(e).__name__}: {e}. Continuing with whatever was already sorted into "
            f"'{pdf_dir}'/'{excel_dir}' before the failure; anything not yet moved is still "
            f"sitting at '{source}'.",
            file=sys.stderr,
        )

    try:
        print(f"\n=== Stage 2: processing Excel files in {excel_dir} ===")
        try:
            excel_mod = _load_copy_split_tables()
            excel_counts = stage2_process_excels(
                excel_dir, tabular_dir, not_tabular_dir, output_folder, excel_mod, log=print
            )
        except Exception as e:
            stage2_error = e
            print(
                f"ERROR: Stage 2 (Excel) failed and was aborted -- {type(e).__name__}: {e}. "
                f"No Excel file in '{excel_dir}' was processed this run -- Stage 3 (PDF) still "
                "runs below.",
                file=sys.stderr,
            )

        print(f"\n=== Stage 3: processing PDFs in {pdf_dir} ===")
        try:
            pdf_counts = stage3_process_pdfs(pdf_dir, tabular_dir, not_tabular_dir, output_folder, log=print)
            # Popped rather than left in pdf_counts -- stage3_process_pdfs
            # already harvested everything doc_reader_v2.py completed before
            # this returncode was set (see its own docstring), so a non-zero
            # value here is reported below, not treated as reason to abort
            # before printing the summary -- the whole point is that the
            # counts printed below remain trustworthy even after a mid-batch
            # subprocess failure.
            doc_reader_returncode = pdf_counts.pop("subprocess_returncode", 0)
        except Exception as e:
            stage3_error = e
            print(
                f"ERROR: Stage 3 (PDF) failed and was aborted -- {type(e).__name__}: {e}. "
                f"No PDF in '{pdf_dir}' was processed this run.",
                file=sys.stderr,
            )
    finally:
        print("\n=== Summary ===")
        if stage1_error is not None:
            print(f"Split: STAGE FAILED ({type(stage1_error).__name__}: {stage1_error}) -- "
                  "see the ERROR above; Stage 2/3 below only ever saw whatever was already "
                  f"sorted into '{pdf_dir}'/'{excel_dir}' before the failure.")
        if stage2_error is not None:
            print(f"Excel: STAGE FAILED ({type(stage2_error).__name__}: {stage2_error}) -- "
                  "see the ERROR above; nothing in this folder was processed this run.")
        else:
            excel_unparsed = (
                excel_counts.get("error", 0) + excel_counts.get("skipped", 0)
                + excel_counts.get("protected", 0) + excel_counts.get("no_tables", 0)
            )
            print(
                f"Excel: {excel_counts.get('single', 0)} single-table, "
                f"{excel_counts.get('multi', 0)} multi-table (-> '{tabular_dir}'), "
                f"{excel_unparsed} could not be processed (-> '{not_tabular_dir}')"
            )
        if stage3_error is not None:
            print(f"PDF: STAGE FAILED ({type(stage3_error).__name__}: {stage3_error}) -- "
                  "see the ERROR above; nothing in this folder was processed this run.")
        else:
            print(
                f"PDF: {pdf_counts.get('parsed', 0)} parsed (-> '{tabular_dir}'), "
                f"{pdf_counts.get('not_tabular', 0)} not tabular (-> '{not_tabular_dir}')"
            )
        print(f"\nDone. Output folder: {output_folder}")

    if doc_reader_returncode != 0:
        print(
            f"\nWARNING: doc_reader_v2.py exited with code {doc_reader_returncode} -- the PDF "
            "counts above only cover what it completed before failing (see its output above "
            "for details). Any PDF it never got to was sent to 'not tabular', in "
            f"'{not_tabular_dir}'.",
            file=sys.stderr,
        )
    if (
        stage1_error is not None or stage2_error is not None or stage3_error is not None
        or doc_reader_returncode != 0
    ):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
