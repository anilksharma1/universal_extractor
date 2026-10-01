# text_from_pdf.py — Extract text from PDF files into a CSV.
# See text_from_pdf.md for usage, OCR modes, and dependencies.

import pdfplumber
import pandas as pd
from tqdm import tqdm
import os
import argparse
import sys
import contextlib
import multiprocessing
import threading
from concurrent.futures import ProcessPoolExecutor, as_completed


def resolve_pdf_files(path: str) -> list[str]:
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


def check_ocr_deps() -> None:
    """Verify pdf2image and pytesseract are importable; exit with guidance if not."""
    missing = []
    try:
        import pdf2image  # noqa: F401
    except ImportError:
        missing.append("pdf2image")
    try:
        import pytesseract  # noqa: F401
    except ImportError:
        missing.append("pytesseract")

    if missing:
        print(
            f"\nError: OCR requires these packages: {', '.join(missing)}\n"
            "Install with:\n"
            "  pip install pdf2image pytesseract\n"
            "\nAlso ensure Tesseract OCR is installed on your system:\n"
            "  Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "  Then add Tesseract to your PATH (or set pytesseract.pytesseract.tesseract_cmd)\n"
        )
        sys.exit(1)


def is_garbled(text: str, threshold: float = 0.3, single_letter_ratio: float = 0.4) -> bool:
    """
    Heuristic: detect text pdfplumber decoded incorrectly, via two failure modes.

    1. Font with no ToUnicode map — decodes to Private Use Area codepoints
       (or other unmapped control chars), rendering as tofu boxes even
       though the string is non-empty.
    2. Interleaved multi-column layout — narrow side-by-side text blocks
       (e.g. a W-2's box 13 checkbox labels: "Statutory employee",
       "Retirement plan", "Third-party sick pay") get character-interleaved
       when pdfplumber reconstructs reading order by position. The letters
       themselves decode correctly, but nearly every whitespace-split token
       ends up a single stray letter — a pattern real extracted text never
       produces. OCR reads the rendered page image instead, so it isn't
       fooled by this layout artifact.
    """
    if not text.strip():
        return False
    bad = sum(
        1 for ch in text
        if 0xE000 <= ord(ch) <= 0xF8FF  # Private Use Area
        or (ord(ch) < 0x20 and ch not in "\t\n\r")
    )
    if (bad / len(text)) > threshold:
        return True

    for line in text.split("\n"):
        tokens = [t for t in line.split() if t.isalpha()]
        if len(tokens) < 6:
            continue
        singles = sum(1 for t in tokens if len(t) == 1 and t.lower() not in ("a", "i"))
        if singles / len(tokens) > single_letter_ratio:
            return True
    return False


# Set these to match your machine so OCR runs with no path flags at all.
# Leave as None (or pass --poppler-path / --tesseract-cmd) if Poppler's bin
# folder and tesseract.exe are already on PATH.
DEFAULT_POPPLER_PATH = r"C:\Poppler\Release-26.02.0-0\poppler-26.02.0\Library\bin"
DEFAULT_TESSERACT_CMD = r"C:\Tessaract\tesseract.exe"


def ocr_page_image(
    pdf_path: str, page_num: int, lang: str, dpi: int, poppler_path: str = None, tesseract_cmd: str = None
) -> str:
    """Render a single PDF page to an image and return OCR text.

    tesseract_cmd is (re)applied here, not just once in the main process,
    because --workers > 1 runs this in separate worker processes (Windows
    multiprocessing uses 'spawn') that don't inherit main-process globals.
    """
    from pdf2image import convert_from_path
    import pytesseract

    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    images = convert_from_path(
        pdf_path, dpi=dpi, first_page=page_num, last_page=page_num, poppler_path=poppler_path
    )
    if not images:
        return ""
    return pytesseract.image_to_string(images[0], lang=lang)


def open_pdfium(pdf_path: str):
    """
    Open a PDF with PDFium (pypdfium2) — the same engine Edge/Chrome use to
    display PDFs, so if text can be selected and copied in the browser,
    PDFium extracts it too. Used as a fallback when pdfplumber (pdfminer)
    can't open the file, returns no text, or returns garbled text — e.g.
    unusual encryption, a malformed xref table, or font encodings pdfminer
    can't map but PDFium can. pypdfium2 is already a pdfplumber dependency.
    Returns None if unavailable or the file won't open.
    """
    try:
        import pypdfium2
        return pypdfium2.PdfDocument(pdf_path)
    except Exception:
        return None


def pdfium_page_text(doc, page_index: int) -> str:
    """Return one page's text layer via PDFium, or '' on any failure."""
    try:
        page = doc[page_index]
        try:
            textpage = page.get_textpage()
            try:
                text = textpage.get_text_range() or ""
            finally:
                textpage.close()
        finally:
            page.close()
    except Exception:
        return ""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def count_pages(pdf_path: str) -> int:
    """Page count via pdfplumber, falling back to PDFium; 0 if neither opens it."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            return len(pdf.pages)
    except Exception:
        pass
    doc = open_pdfium(pdf_path)
    if doc is None:
        return 0
    try:
        return len(doc)
    finally:
        doc.close()


MIN_PAGES_PER_CHUNK = 25


def plan_chunks(pdf_files: list[str], num_workers: int) -> list[dict]:
    """
    Split each PDF's pages into up to `num_workers` contiguous chunks, so a
    single huge file can be spread across the whole worker pool instead of
    only ever handing one worker a whole file. Small files (fewer than
    MIN_PAGES_PER_CHUNK pages) stay as one chunk — splitting them just adds
    reopen overhead for no parallelism benefit.
    """
    chunks = []
    scan_bar = tqdm(
        pdf_files,
        desc="Scanning files (counting pages)",
        unit="file",
        bar_format="{desc}: {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}] {bar}",
    )
    for file_index, pdf_path in enumerate(scan_bar, start=1):
        scan_bar.set_postfix_str(os.path.basename(pdf_path)[:40])
        total_pages = count_pages(pdf_path)

        n_splits = min(num_workers, max(1, total_pages // MIN_PAGES_PER_CHUNK)) if total_pages else 1
        base, rem = divmod(total_pages, n_splits)
        start = 0
        for i in range(n_splits):
            size = base + (1 if i < rem else 0)
            end = start + size
            chunks.append(
                {
                    "file_index": file_index,
                    "pdf_path": pdf_path,
                    "page_start": start,
                    "page_end": end,
                    "n_chunks": n_splits,
                    "chunk_no": i + 1,
                }
            )
            start = end
    scan_bar.close()
    return chunks


def extract_pdf_range(
    pdf_path: str,
    file_index: int,
    total_files: int,
    ocr_mode: str,
    lang: str,
    dpi: int,
    poppler_path: str = None,
    tesseract_cmd: str = None,
    page_start: int = 0,
    page_end: int = None,
    chunk_label: str = "",
    show_progress: bool = True,
    progress_queue=None,
    preserve_layout: bool = False,
) -> tuple[list[dict], dict]:
    """
    Extract text lines from a page range [page_start, page_end) of one PDF
    (page_end=None means "to the end"). A whole file is just the range
    covering all of its pages — chunking one huge file into several ranges
    is what lets it be split across multiple workers.

    ocr_mode:
      'never'  — pdfplumber, then PDFium for pages it can't read; no OCR
      'auto'   — pdfplumber, then PDFium; OCR fallback for pages with no text
      'always' — OCR every page (skips pdfplumber and PDFium)

    Pages read by PDFium get Method='pdfium' in the records but are counted
    as 'native' in stats (it's still the PDF's own text layer, not OCR).

    progress_queue, when given (worker-pool mode), replaces the local tqdm
    bar with ("start"/"tick"/"close", ...) messages consumed by a listener
    in the main process — workers never touch the terminal directly, since
    unsynchronized writes from multiple processes garble/freeze the display.

    preserve_layout, when True, extracts native text with pdfplumber's
    layout=True mode: instead of collapsing gaps into single spaces, it
    pads with extra spaces to keep the original column alignment (e.g. a
    table's columns stay lined up). Only affects native extraction — OCR
    output already reflects the image's own spacing.

    Returns (records, stats) where stats counts native/ocr/empty pages.
    """
    records = []
    stats = {"native": 0, "ocr": 0, "empty": 0, "error": 0}
    pdf_file = os.path.basename(pdf_path)
    ident = None

    def log(msg: str) -> None:
        if progress_queue is not None:
            progress_queue.put(("log", msg))
        else:
            tqdm.write(msg)

    try:
        with contextlib.ExitStack() as stack:
            try:
                pdf = stack.enter_context(pdfplumber.open(pdf_path))
                doc_pages = len(pdf.pages)
            except Exception as e:
                pdf = None
                doc_pages = 0
                plumber_error = e

            # PDFium is opened lazily — only when pdfplumber fails on a page
            # (or on the whole file) — so normal PDFs pay nothing for it.
            pdfium_state = {"doc": None, "tried": False}

            def get_pdfium():
                if not pdfium_state["tried"]:
                    pdfium_state["tried"] = True
                    pdfium_state["doc"] = open_pdfium(pdf_path)
                    if pdfium_state["doc"] is not None:
                        stack.callback(pdfium_state["doc"].close)
                return pdfium_state["doc"]

            if pdf is None:
                if get_pdfium() is None:
                    raise plumber_error
                doc_pages = len(pdfium_state["doc"])
                log(f"  Note: pdfplumber couldn't open '{pdf_file}' ({plumber_error}) — using PDFium.")

            pages = list(range(doc_pages))[page_start:page_end]
            total_pages = len(pages)
            label = f"{pdf_file[:32]}{chunk_label}"

            if progress_queue is not None:
                proc_ident = multiprocessing.current_process()._identity
                ident = proc_ident[0] if proc_ident else 1
                progress_queue.put(("start", ident, file_index, total_files, label, total_pages))
                page_iter = pages
            elif show_progress:
                page_iter = tqdm(
                    pages,
                    desc=f"  {label}",
                    unit="pg",
                    total=total_pages,
                    bar_format=(
                        f"  [{file_index}/{total_files}] {{desc}} | "
                        "{n_fmt}/{total_fmt} pages [{elapsed}<{remaining}, {rate_fmt}]"
                    ),
                )
            else:
                page_iter = pages

            for page_index in page_iter:
                page_num = page_index + 1
                try:
                    text = ""
                    method = "native"

                    if ocr_mode != "always":
                        if pdf is not None:
                            try:
                                text = pdf.pages[page_index].extract_text(layout=preserve_layout) or ""
                            except Exception:
                                text = ""
                        if not text.strip() or is_garbled(text):
                            # pdfplumber found nothing usable — try the
                            # browser's engine before resorting to OCR.
                            doc = get_pdfium()
                            alt = pdfium_page_text(doc, page_index) if doc is not None else ""
                            if alt.strip() and not is_garbled(alt):
                                text = alt
                                method = "pdfium"

                    needs_ocr = not text.strip() or (
                        ocr_mode == "auto" and is_garbled(text)
                    )
                    if needs_ocr and ocr_mode != "never":
                        # OCR fallback (or forced) — also catches pages
                        # whose font has no ToUnicode map, where
                        # extract_text() returns non-empty garbage
                        ocr_text = ocr_page_image(pdf_path, page_num, lang, dpi, poppler_path, tesseract_cmd)
                        if ocr_text.strip():
                            text = ocr_text
                            method = "ocr"

                    lines_added = 0
                    for line in text.split("\n"):
                        stripped = line.strip()
                        if stripped:
                            records.append(
                                {
                                    "File Name": pdf_file,
                                    "Page Number": page_num,
                                    "Extracted Text": stripped,
                                    "Method": method,
                                }
                            )
                            lines_added += 1

                    if lines_added == 0:
                        stats["empty"] += 1
                    elif method == "ocr":
                        stats["ocr"] += 1
                    else:
                        stats["native"] += 1

                except Exception as e:
                    stats["error"] += 1
                    msg = f"    Warning: page {page_num} skipped — {e}"
                    if progress_queue is not None:
                        progress_queue.put(("log", msg))
                    else:
                        tqdm.write(msg)
                finally:
                    if progress_queue is not None:
                        progress_queue.put(("tick", ident))

    except Exception as e:
        msg = f"  Failed to open '{pdf_file}': {e}"
        if progress_queue is not None:
            progress_queue.put(("log", msg))
        else:
            tqdm.write(msg)

    if ident is not None:
        progress_queue.put(("close", ident))

    return records, stats


def _progress_listener(progress_queue, stop_event) -> None:
    """
    Owns every per-worker page bar in the main process. Workers only ever
    send messages here — this is the sole writer to the terminal for
    per-page progress, so bars never race across process boundaries.
    """
    bars = {}
    while not stop_event.is_set() or not progress_queue.empty():
        try:
            msg = progress_queue.get(timeout=0.2)
        except Exception:
            continue

        kind = msg[0]
        if kind == "start":
            _, ident, file_index, total_files, pdf_file, total_pages = msg
            if ident in bars:
                bars[ident].close()
            bars[ident] = tqdm(
                total=total_pages,
                desc=f"  {pdf_file[:40]}",
                unit="pg",
                position=ident,
                leave=False,
                bar_format=(
                    f"  [{file_index}/{total_files}] {{desc}} | "
                    "{n_fmt}/{total_fmt} pages [{elapsed}<{remaining}, {rate_fmt}]"
                ),
            )
        elif kind == "tick":
            _, ident = msg
            if ident in bars:
                bars[ident].update(1)
        elif kind == "close":
            _, ident = msg
            if ident in bars:
                bars[ident].close()
                del bars[ident]
        elif kind == "log":
            _, text = msg
            tqdm.write(text)

    for bar in bars.values():
        bar.close()


def main():
    parser = argparse.ArgumentParser(
        description="Extract text from PDF(s) into a CSV, with optional Tesseract OCR.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python text_from_pdf.py report.pdf\n"
            "  python text_from_pdf.py ./docs/\n"
            "  python text_from_pdf.py ./docs/ --ocr auto\n"
            "  python text_from_pdf.py scan.pdf --ocr always --lang eng+ara\n"
            "  python text_from_pdf.py ./docs/ --output results.csv --dpi 400\n"
            "  python text_from_pdf.py scan.pdf --poppler-path \"C:\\Poppler\\Release-26.02.0-0\\poppler-26.02.0\\Library\\bin\"\n"
        ),
    )
    parser.add_argument(
        "path",
        nargs="+",
        help="Path to a single PDF file or a folder containing PDFs. Spaces in path are handled automatically.",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output CSV file name (default: derived from input path).",
    )
    parser.add_argument(
        "--ocr",
        choices=["never", "auto", "always"],
        default="auto",
        help=(
            "OCR mode: 'never' = pdfplumber only, no OCR; "
            "'auto' = OCR fallback for pages with no selectable text or "
            "garbled/unmapped text (default); "
            "'always' = force OCR on every page (for scanned PDFs)."
        ),
    )
    parser.add_argument(
        "--lang",
        default="eng",
        help="Tesseract language code(s), e.g. 'eng', 'eng+ara' (default: eng).",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI used when rendering pages to images for OCR (default: 300).",
    )
    parser.add_argument(
        "--poppler-path",
        default=DEFAULT_POPPLER_PATH,
        help=(
            "Folder containing pdftoppm.exe / pdftocairo.exe, if Poppler isn't "
            "on PATH, e.g. 'C:\\Poppler\\Release-26.02.0-0\\poppler-26.02.0\\Library\\bin'."
        ),
    )
    parser.add_argument(
        "--tesseract-cmd",
        default=DEFAULT_TESSERACT_CMD,
        help="Full path to tesseract.exe, if not on PATH.",
    )
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=4,
        help=(
            "Number of parallel worker processes (default: 4). Files are "
            "spread across workers; a single large PDF is also split into "
            "page-chunks so it can use every worker too."
        ),
    )
    parser.add_argument(
        "--per-file",
        action="store_true",
        help=(
            "Also write one CSV per input PDF, saved as soon as that file "
            "finishes (in addition to the combined output CSV at the end). "
            "Safer for large folders: a crash partway through only loses "
            "the file still in progress, not everything already extracted."
        ),
    )
    parser.add_argument(
        "--per-file-dir",
        default=None,
        help=(
            "Folder to write per-file CSVs into when --per-file is set "
            "(default: '<output_csv_stem>_per_file' next to the output CSV)."
        ),
    )
    args = parser.parse_args()
    input_path = " ".join(args.path)

    if args.ocr != "never":
        check_ocr_deps()
        if args.tesseract_cmd and not os.path.isfile(args.tesseract_cmd):
            print(f"Warning: --tesseract-cmd '{args.tesseract_cmd}' not found — falling back to PATH lookup.")
            args.tesseract_cmd = None

    pdf_files = resolve_pdf_files(input_path)

    if args.output:
        output_csv = args.output
    elif os.path.isfile(input_path):
        output_csv = os.path.splitext(os.path.basename(input_path))[0] + ".csv"
    else:
        folder_name = os.path.basename(os.path.normpath(input_path)) or "output"
        output_csv = folder_name + ".csv"

    total_files = len(pdf_files)
    ocr_label = {"never": "off", "auto": "fallback", "always": "forced"}[args.ocr]

    per_file_dir = None
    if args.per_file:
        per_file_dir = args.per_file_dir or (os.path.splitext(output_csv)[0] + "_per_file")
        os.makedirs(per_file_dir, exist_ok=True)

    print(
        f"\nFound {total_files} PDF file(s)  |  OCR: {ocr_label}"
        + (f"  lang={args.lang}  dpi={args.dpi}" if args.ocr != "never" else "")
        + f"  →  output: {output_csv}\n"
        + (f"  →  per-file CSVs: {per_file_dir}\n" if per_file_dir else "")
    )

    all_records = []
    total_stats = {"native": 0, "ocr": 0, "empty": 0, "error": 0}
    requested_workers = max(1, args.workers)

    def log_result(pdf_path: str, records: list, stats: dict) -> None:
        parts = [f"{len(records)} lines"]
        if stats["ocr"]:
            parts.append(f"{stats['ocr']} pg OCR")
        if stats["empty"]:
            parts.append(f"{stats['empty']} pg empty")

        if per_file_dir and records:
            per_file_csv = os.path.join(
                per_file_dir, os.path.splitext(os.path.basename(pdf_path))[0] + ".csv"
            )
            pd.DataFrame(records).to_csv(per_file_csv, index=False, encoding="utf-8-sig")
            parts.append(f"saved → {os.path.basename(per_file_csv)}")

        tqdm.write(f"  -> {', '.join(parts)}  —  {os.path.basename(pdf_path)}")

    chunks = plan_chunks(pdf_files, requested_workers) if requested_workers > 1 else []
    num_workers = min(requested_workers, len(chunks))

    if num_workers <= 1:
        file_bar = tqdm(
            pdf_files,
            desc="Overall",
            unit="file",
            bar_format="{desc}: {n_fmt}/{total_fmt} files [{elapsed}<{remaining}, {rate_fmt}] {bar}",
            position=0,
        )
        for idx, pdf_path in enumerate(file_bar, start=1):
            file_bar.set_description(f"Overall [{idx}/{total_files}]")
            records, stats = extract_pdf_range(
                pdf_path, idx, total_files, args.ocr, args.lang, args.dpi, args.poppler_path, args.tesseract_cmd
            )
            all_records.extend(records)
            for k in total_stats:
                total_stats[k] += stats[k]
            log_result(pdf_path, records, stats)
        file_bar.close()
    else:
        multi_chunk_files = len({c["file_index"] for c in chunks if c["n_chunks"] > 1})
        detail = f" ({multi_chunk_files} large file(s) split into page-chunks)" if multi_chunk_files else ""
        print(f"Processing with {num_workers} parallel workers across {len(chunks)} chunk(s){detail}...\n")

        results_by_idx = {}
        pending = {
            c["file_index"]: {"total": c["n_chunks"], "parts": {}} for c in chunks
        }
        file_bar = tqdm(
            total=total_files,
            desc="Overall",
            unit="file",
            bar_format="{desc}: {n_fmt}/{total_fmt} files [{elapsed}<{remaining}, {rate_fmt}] {bar}",
            position=0,
        )

        manager = multiprocessing.Manager()
        progress_queue = manager.Queue()
        stop_event = threading.Event()
        listener = threading.Thread(
            target=_progress_listener, args=(progress_queue, stop_event), daemon=True
        )
        listener.start()

        with ProcessPoolExecutor(max_workers=num_workers) as executor:
            futures = {}
            for chunk in chunks:
                label = (
                    f" [p{chunk['page_start'] + 1}-{chunk['page_end']}]"
                    if chunk["n_chunks"] > 1
                    else ""
                )
                fut = executor.submit(
                    extract_pdf_range,
                    chunk["pdf_path"],
                    chunk["file_index"],
                    total_files,
                    args.ocr,
                    args.lang,
                    args.dpi,
                    args.poppler_path,
                    args.tesseract_cmd,
                    chunk["page_start"],
                    chunk["page_end"],
                    label,
                    True,
                    progress_queue,
                )
                futures[fut] = chunk

            for future in as_completed(futures):
                chunk = futures[future]
                file_index = chunk["file_index"]
                pdf_path = chunk["pdf_path"]
                try:
                    records, stats = future.result()
                except Exception as e:
                    records, stats = [], {"native": 0, "ocr": 0, "empty": 0, "error": 1}
                    tqdm.write(f"  Failed chunk of '{os.path.basename(pdf_path)}': {e}")

                file_state = pending[file_index]
                file_state["parts"][chunk["chunk_no"]] = (records, stats)
                if len(file_state["parts"]) == file_state["total"]:
                    merged_records = []
                    merged_stats = {"native": 0, "ocr": 0, "empty": 0, "error": 0}
                    for chunk_no in sorted(file_state["parts"]):
                        r, s = file_state["parts"][chunk_no]
                        merged_records.extend(r)
                        for k in merged_stats:
                            merged_stats[k] += s[k]
                    results_by_idx[file_index] = (merged_records, merged_stats)
                    log_result(pdf_path, merged_records, merged_stats)
                    file_bar.update(1)

        stop_event.set()
        listener.join()
        manager.shutdown()
        file_bar.close()

        for idx in range(1, total_files + 1):
            records, stats = results_by_idx[idx]
            all_records.extend(records)
            for k in total_stats:
                total_stats[k] += stats[k]

    if not all_records:
        print("\nNo text extracted. CSV not written.")
        sys.exit(0)

    df = pd.DataFrame(all_records)
    df.to_csv(output_csv, index=False, encoding="utf-8-sig")

    print(f"\nDone. {len(all_records)} lines from {total_files} file(s) saved to: {output_csv}")
    if args.ocr != "never":
        print(
            f"  Pages — native: {total_stats['native']}  "
            f"ocr: {total_stats['ocr']}  "
            f"empty: {total_stats['empty']}  "
            f"errors: {total_stats['error']}"
        )


if __name__ == "__main__":
    main()
