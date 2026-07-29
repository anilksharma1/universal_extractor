"""
PDF Format Bucket Grouper

Scans a folder of PDFs, fingerprints the LAYOUT of each file's first page
(no data extraction - only structural shape is analyzed), and groups files
that share a similar layout into the same "bucket". Output is a CSV with
File Name / Bucket Name.

Layout fingerprinting approach:
  - Text-based pages (a selectable text layer is present): the first page's
    text/image blocks are mapped onto a coarse grid based on their bounding
    boxes, producing an occupancy pattern that reflects the template shape
    (headers, columns, tables) independent of the actual words on the page.
  - Scanned/image-only pages (no usable text layer): the page is rasterized
    and the same grid occupancy is derived from pixel darkness instead.
  - Page orientation (portrait/landscape) and page type (text/image) are
    kept as separate clustering keys, then bucket membership within a
    key is decided by Hamming similarity between occupancy grids.

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
TEXT_CHAR_THRESHOLD = 15       # min chars on page to treat it as "text" type
DARKNESS_OCCUPIED_CUTOFF = 25  # 0-255 scale; higher = stricter "has ink" test
SAMPLE_STEPS = 8               # pixel samples per cell edge when rasterizing
ERROR_BUCKET_NAME = "Unreadable_Error"


def compute_fingerprint(pdf_path):
    """Return a dict describing the layout of a PDF's first page."""
    doc = None
    try:
        doc = fitz.open(pdf_path)
        if doc.needs_pass:
            if not doc.authenticate(""):
                return {"type": "error", "message": "Encrypted (password required)"}
        if doc.page_count == 0:
            return {"type": "error", "message": "No pages in file"}

        page = doc[0]
        width, height = page.rect.width, page.rect.height
        if width <= 0 or height <= 0:
            return {"type": "error", "message": "Invalid page dimensions"}
        orientation = "landscape" if width > height else "portrait"

        text_dict = page.get_text("dict")
        blocks = text_dict.get("blocks", [])

        total_chars = 0
        for block in blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    total_chars += len(span.get("text", "").strip())

        if total_chars >= TEXT_CHAR_THRESHOLD:
            grid = _grid_from_blocks(blocks, width, height)
            page_type = "text"
        else:
            grid = _grid_from_pixels(page, width, height)
            page_type = "image"

        return {"type": page_type, "orientation": orientation, "grid": grid}

    except Exception as exc:  # noqa: BLE001 - want to bucket ANY failure, not crash
        return {"type": "error", "message": f"{type(exc).__name__}: {exc}"}
    finally:
        if doc is not None:
            doc.close()


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


def _grid_from_pixels(page, width, height):
    pix = page.get_pixmap(colorspace=fitz.csGRAY)
    samples = pix.samples
    pw, ph = pix.width, pix.height
    grid = [0] * (GRID_COLS * GRID_ROWS)

    cell_w = pw / GRID_COLS
    cell_h = ph / GRID_ROWS

    for r in range(GRID_ROWS):
        y0 = int(r * cell_h)
        y1 = max(y0 + 1, int((r + 1) * cell_h))
        step_y = max(1, (y1 - y0) // SAMPLE_STEPS)
        for c in range(GRID_COLS):
            x0 = int(c * cell_w)
            x1 = max(x0 + 1, int((c + 1) * cell_w))
            step_x = max(1, (x1 - x0) // SAMPLE_STEPS)

            total = 0
            count = 0
            for y in range(y0, y1, step_y):
                row_offset = y * pw
                for x in range(x0, x1, step_x):
                    total += samples[row_offset + x]
                    count += 1
            avg = (total / count) if count else 255
            grid[r * GRID_COLS + c] = 1 if (255 - avg) > DARKNESS_OCCUPIED_CUTOFF else 0
    return grid


def _clamp(value, lo, hi):
    return max(lo, min(hi, value))


def _similarity(grid_a, grid_b):
    matches = sum(1 for a, b in zip(grid_a, grid_b) if a == b)
    return matches / len(grid_a) if grid_a else 0.0


class BucketAssigner:
    """Greedy clustering: each new fingerprint joins the first bucket it's
    similar enough to (within the same type/orientation), else starts a new one."""

    def __init__(self, threshold):
        self.threshold = threshold
        self.buckets = []  # [{name, type, orientation, grid}]
        self._counter = 1

    def assign(self, fingerprint):
        if fingerprint["type"] == "error":
            return ERROR_BUCKET_NAME

        for bucket in self.buckets:
            if (
                bucket["type"] == fingerprint["type"]
                and bucket["orientation"] == fingerprint["orientation"]
                and _similarity(bucket["grid"], fingerprint["grid"]) >= self.threshold
            ):
                return bucket["name"]

        name = f"Bucket_{self._counter}"
        self._counter += 1
        self.buckets.append({
            "name": name,
            "type": fingerprint["type"],
            "orientation": fingerprint["orientation"],
            "grid": fingerprint["grid"],
        })
        return name


def process_folder(folder, threshold, progress_callback):
    """Walk the top-level of `folder` for PDFs, bucket them, return row list.

    progress_callback(done, total, current_filename) is invoked after each file.
    Returns list of (filename, bucket_name, notes).
    """
    filenames = sorted(
        f for f in os.listdir(folder)
        if f.lower().endswith(".pdf") and os.path.isfile(os.path.join(folder, f))
    )

    assigner = BucketAssigner(threshold)
    rows = []
    total = len(filenames)

    for i, filename in enumerate(filenames, start=1):
        full_path = os.path.join(folder, filename)
        fingerprint = compute_fingerprint(full_path)
        bucket_name = assigner.assign(fingerprint)
        notes = fingerprint.get("message", "") if fingerprint["type"] == "error" else ""
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
        self.threshold_var = DoubleVar(value=0.85)
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

        ttk.Label(frame, text="Similarity threshold (0.50 - 1.00):").grid(
            row=1, column=0, sticky=W, **pad
        )
        ttk.Spinbox(
            frame, from_=0.50, to=1.00, increment=0.05, textvariable=self.threshold_var, width=8
        ).grid(row=1, column=1, sticky=W, **pad)
        ttk.Label(
            frame, text="(higher = stricter, more buckets)", foreground="gray"
        ).grid(row=1, column=2, sticky=W, **pad)

        self.run_button = ttk.Button(frame, text="Run", command=self.run)
        self.run_button.grid(row=2, column=0, columnspan=3, pady=(4, 10))

        self.progress = ttk.Progressbar(frame, length=440, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=3, **pad)

        ttk.Label(frame, textvariable=self.status_var, wraplength=440, justify="left").grid(
            row=4, column=0, columnspan=3, sticky=W, **pad
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
            threshold = float(self.threshold_var.get())
        except (ValueError, TypeError):
            threshold = 0.85
        threshold = _clamp(threshold, 0.50, 1.00)

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
            target=self._worker, args=(folder, threshold, csv_path), daemon=True
        )
        worker.start()
        self.root.after(100, self._poll_queue)

    def _worker(self, folder, threshold, csv_path):
        try:
            def progress_callback(done, total, filename):
                self._queue.put(("progress", done, total, filename))

            rows = process_folder(folder, threshold, progress_callback)
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
