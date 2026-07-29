"""
PDF Format Bucket Grouper

Scans a folder of PDFs, fingerprints the LAYOUT of each file's first pages
(no data extraction - only structural shape is analyzed), and groups files
that share a similar layout into the same "bucket". Output is a CSV with
File Name / Bucket Name.

Layout fingerprinting approach:
  - Text-based pages (a selectable text layer is present): for each of the
    first PAGES_TO_ANALYZE pages (fewer if the document is shorter), the
    text/image blocks are mapped onto a coarse grid based on their bounding
    boxes, producing an occupancy pattern that reflects the template shape
    (headers, columns, tables) independent of the actual words on the page.
    Each page's grid is concatenated into one combined fingerprint, so
    documents that look alike on page 1 but diverge on later pages (e.g.
    a W-2 vs. a similarly-laid-out 2222 form) still separate cleanly.
    Page orientation (portrait/landscape) is kept as a separate clustering
    key.
  - Alongside layout, the same pages' words are read into a "vocabulary"
    (alphabetic words, common stopwords removed - numbers/amounts/SSNs are
    dropped by construction, so this is a text-pattern signature, not the
    extracted data itself). Two documents whose vocabularies overlap
    strongly are recognized as the same template even when their layout
    grids don't line up exactly (e.g. due to scan jitter or reflow).
  - A file joins an existing bucket when EITHER its layout similarity meets
    the layout threshold OR its vocabulary overlap meets the text-pattern
    threshold (both configurable in the UI); otherwise it starts a new
    bucket.
  - Scanned/image-only pages (no usable text layer on page 1) are NOT
    bucketed - they are flagged with the note "scanned file" instead.
  - Rotated pages (PDF /Rotate other than 0, on any analyzed page) are NOT
    bucketed - they are flagged with the note "rotated file" instead.

Requires: pymupdf  (pip install pymupdf)
"""

import csv
import os
import queue
import threading
import traceback
from tkinter import (
    Tk, StringVar, DoubleVar, END, N, S, E, W,
    filedialog, messagebox,
)
from tkinter import ttk

import fitz  # PyMuPDF

GRID_COLS = 12
GRID_ROWS = 16
PAGES_TO_ANALYZE = 4           # look at up to this many leading pages per PDF
TEXT_CHAR_THRESHOLD = 15       # min chars on page 1 to treat the doc as "text" type
MIN_TOKEN_LEN = 4              # ignore very short words when building the vocabulary
ERROR_BUCKET_NAME = "Unreadable_Error"

STOPWORDS = {
    "the", "and", "for", "of", "to", "in", "on", "is", "are", "or", "by",
    "with", "this", "that", "as", "be", "an", "at", "from", "your", "you",
    "will", "not", "may", "any", "all", "such", "into", "than", "then",
    "have", "has", "was", "were", "been", "each", "date", "page",
}


def compute_fingerprint(pdf_path):
    """Return a dict describing the layout of a PDF's first pages."""
    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.needs_pass:
            if not doc.authenticate(""):
                return {"type": "error", "message": "Encrypted (password required)"}
        if doc.page_count == 0:
            return {"type": "error", "message": "No pages in file"}

        num_pages = min(PAGES_TO_ANALYZE, doc.page_count)
        pages = [doc[i] for i in range(num_pages)]

        first = pages[0]
        width, height = first.rect.width, first.rect.height
        if width <= 0 or height <= 0:
            return {"type": "error", "message": "Invalid page dimensions"}

        if any(p.rotation != 0 for p in pages):
            return {"type": "rotated"}

        orientation = "landscape" if width > height else "portrait"

        first_blocks = first.get_text("dict").get("blocks", [])
        if _count_chars(first_blocks) < TEXT_CHAR_THRESHOLD:
            return {"type": "scanned"}

        combined_grid = []
        tokens = set()
        for page in pages:
            blocks = page.get_text("dict").get("blocks", [])
            combined_grid.extend(_grid_from_blocks(blocks, page.rect.width, page.rect.height))
            tokens.update(_extract_tokens(page.get_text("text")))

        return {"type": "text", "orientation": orientation, "grid": combined_grid, "tokens": tokens}

    except Exception as exc:  # noqa: BLE001 - want to bucket ANY failure, not crash
        return {"type": "error", "message": f"{type(exc).__name__}: {exc}"}
    finally:
        if doc is not None:
            doc.close()


def _count_chars(blocks):
    total = 0
    for block in blocks:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                total += len(span.get("text", "").strip())
    return total


def _grid_from_blocks(blocks, width, height):
    grid = [0] * (GRID_COLS * GRID_ROWS)
    for block in blocks:
        bbox = block.get("bbox")
        if not bbox:
            continue
        x0, y0, x1, y1 = bbox
        col_start = _clamp(int((x0 / width) * GRID_COLS), 0, GRID_COLS - 1)
        col_end = _clamp(int((x1 / width) * GRID_COLS), 0, GRID_COLS - 1)
        row_start = _clamp(int((y0 / height) * GRID_ROWS), 0, GRID_ROWS - 1)
        row_end = _clamp(int((y1 / height) * GRID_ROWS), 0, GRID_ROWS - 1)
        for r in range(row_start, row_end + 1):
            for c in range(col_start, col_end + 1):
                grid[r * GRID_COLS + c] = 1
    return grid


def _clamp(value, lo, hi):
    return max(lo, min(hi, value))


def _similarity(grid_a, grid_b):
    matches = sum(1 for a, b in zip(grid_a, grid_b) if a == b)
    return matches / len(grid_a) if grid_a else 0.0


def _extract_tokens(text):
    tokens = set()
    for raw in text.split():
        word = "".join(ch for ch in raw.lower() if ch.isalpha())
        if len(word) >= MIN_TOKEN_LEN and word not in STOPWORDS:
            tokens.add(word)
    return tokens


def _jaccard(tokens_a, tokens_b):
    if not tokens_a or not tokens_b:
        return 0.0
    return len(tokens_a & tokens_b) / len(tokens_a | tokens_b)


class BucketAssigner:
    """Greedy clustering over text-type fingerprints only: each new
    fingerprint joins the first bucket it's similar enough to (within the
    same orientation) on EITHER layout or text-pattern grounds, else starts
    a new one."""

    def __init__(self, layout_threshold, text_threshold):
        self.layout_threshold = layout_threshold
        self.text_threshold = text_threshold
        self.buckets = []  # [{name, orientation, grid, tokens}]
        self._counter = 1

    def assign(self, fingerprint):
        for bucket in self.buckets:
            if bucket["orientation"] != fingerprint["orientation"]:
                continue
            layout_sim = _similarity(bucket["grid"], fingerprint["grid"])
            text_sim = _jaccard(bucket["tokens"], fingerprint["tokens"])
            if layout_sim >= self.layout_threshold or text_sim >= self.text_threshold:
                return bucket["name"]

        name = f"Bucket_{self._counter}"
        self._counter += 1
        self.buckets.append({
            "name": name,
            "orientation": fingerprint["orientation"],
            "grid": fingerprint["grid"],
            "tokens": fingerprint["tokens"],
        })
        return name


def process_folder(folder, layout_threshold, text_threshold, progress_callback):
    """Walk the top-level of `folder` for PDFs, bucket them, return row list.

    progress_callback(done, total, current_filename) is invoked after each file.
    Returns list of (filename, bucket_name, notes).
    """
    filenames = sorted(
        f for f in os.listdir(folder)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(folder, f))
    )

    assigner = BucketAssigner(layout_threshold, text_threshold)
    rows = []
    total = len(filenames)

    for i, filename in enumerate(filenames, start=1):
        full_path = os.path.join(folder, filename)
        fingerprint = compute_fingerprint(full_path)
        ftype = fingerprint["type"]

        if ftype == "error":
            bucket_name, notes = ERROR_BUCKET_NAME, fingerprint.get("message", "")
        elif ftype == "rotated":
            bucket_name, notes = "", "rotated file"
        elif ftype == "scanned":
            bucket_name, notes = "", "scanned file"
        else:
            bucket_name, notes = assigner.assign(fingerprint), ""

        rows.append((filename, bucket_name, notes))
        progress_callback(i, total, filename)

    return rows


def write_csv(rows, csv_path):
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["File Name", "Bucket Name", "Notes"])
        writer.writerows(rows)


class App:
    def __init__(self, root):
        self.root = root
        root.title("PDF Format Bucket Grouper")
        root.resizable(False, False)

        self.folder_var = StringVar()
        self.layout_threshold_var = DoubleVar(value=0.92)
        self.text_threshold_var = DoubleVar(value=0.60)
        self.status_var = StringVar(value="Choose a folder of PDFs to begin.")

        pad = {"padx": 8, "pady": 6}

        frame = ttk.Frame(root)
        frame.grid(row=0, column=0, sticky=(N, S, E, W))

        ttk.Label(frame, text="Input folder:").grid(row=0, column=0, sticky=W, **pad)
        ttk.Entry(frame, textvariable=self.folder_var, width=50, state="readonly").grid(
            row=0, column=1, **pad
        )
        ttk.Button(frame, text="Browse...", command=self.browse_folder).grid(
            row=0, column=2, **pad
        )

        ttk.Label(frame, text="Layout similarity threshold (0.50 - 1.00):").grid(
            row=1, column=0, sticky=W, **pad
        )
        ttk.Spinbox(
            frame, from_=0.50, to=1.00, increment=0.05, textvariable=self.layout_threshold_var, width=8
        ).grid(row=1, column=1, sticky=W, **pad)
        ttk.Label(
            frame, text="(higher = stricter, more buckets)", foreground="gray"
        ).grid(row=1, column=2, sticky=W, **pad)

        ttk.Label(frame, text="Text pattern threshold (0.10 - 1.00):").grid(
            row=2, column=0, sticky=W, **pad
        )
        ttk.Spinbox(
            frame, from_=0.10, to=1.00, increment=0.05, textvariable=self.text_threshold_var, width=8
        ).grid(row=2, column=1, sticky=W, **pad)
        ttk.Label(
            frame, text="(shared vocabulary can also merge a file into a bucket)", foreground="gray"
        ).grid(row=2, column=2, sticky=W, **pad)

        self.run_button = ttk.Button(frame, text="Run", command=self.run)
        self.run_button.grid(row=3, column=0, columnspan=3, pady=(4, 10))

        self.progress = ttk.Progressbar(frame, length=440, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, **pad)

        ttk.Label(frame, textvariable=self.status_var, wraplength=440, justify="left").grid(
            row=5, column=0, columnspan=3, sticky=W, **pad
        )

        self._queue = queue.Queue()

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if folder:
            self.folder_var.set(folder)

    def run(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("No folder", "Please choose a valid input folder first.")
            return

        try:
            layout_threshold = float(self.layout_threshold_var.get())
        except (ValueError, TypeError):
            layout_threshold = 0.92
        layout_threshold = _clamp(layout_threshold, 0.50, 1.00)

        try:
            text_threshold = float(self.text_threshold_var.get())
        except (ValueError, TypeError):
            text_threshold = 0.60
        text_threshold = _clamp(text_threshold, 0.10, 1.00)

        pdf_count = sum(
            1 for f in os.listdir(folder)
            if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(folder, f))
        )
        if pdf_count == 0:
            messagebox.showwarning("No PDFs found", "No .pdf files were found directly in that folder.")
            return

        default_name = os.path.join(folder, "pdf_bucket_results.csv")
        csv_path = filedialog.asksaveasfilename(
            title="Save bucket results CSV as",
            initialdir=folder,
            initialfile="pdf_bucket_results.csv",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not csv_path:
            return

        self.run_button.state(["disabled"])
        self.progress.configure(maximum=pdf_count, value=0)
        self.status_var.set(f"Processing 0/{pdf_count}...")

        worker = threading.Thread(
            target=self._worker,
            args=(folder, layout_threshold, text_threshold, csv_path),
            daemon=True,
        )
        worker.start()
        self.root.after(100, self._poll_queue)

    def _worker(self, folder, layout_threshold, text_threshold, csv_path):
        try:
            def progress_callback(done, total, filename):
                self._queue.put(("progress", done, total, filename))

            rows = process_folder(folder, layout_threshold, text_threshold, progress_callback)
            write_csv(rows, csv_path)
            bucket_names = sorted(set(r[1] for r in rows))
            self._queue.put(("done", csv_path, len(rows), len(bucket_names)))
        except Exception:
            self._queue.put(("error", traceback.format_exc()))

    def _poll_queue(self):
        try:
            while True:
                message = self._queue.get_nowait()
                kind = message[0]

                if kind == "progress":
                    _, done, total, filename = message
                    self.progress.configure(value=done)
                    self.status_var.set(f"Processing {done}/{total}: {filename}")

                elif kind == "done":
                    _, csv_path, file_count, bucket_count = message
                    self.status_var.set(
                        f"Done. {file_count} files grouped into {bucket_count} bucket(s).\n"
                        f"Saved to: {csv_path}"
                    )
                    self.run_button.state(["!disabled"])
                    messagebox.showinfo(
                        "Finished",
                        f"{file_count} files grouped into {bucket_count} bucket(s).\n\n{csv_path}",
                    )
                    return

                elif kind == "error":
                    _, tb = message
                    self.status_var.set("An error occurred. See popup for details.")
                    self.run_button.state(["!disabled"])
                    messagebox.showerror("Error", tb)
                    return

        except queue.Empty:
            pass

        self.root.after(100, self._poll_queue)


def main():
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
