"""
=====================================================================
 PDF EXTRACTOR - GUI (Text / Tables / Both)
=====================================================================

 WHAT IT DOES
 ------------
 Extracts data from ALL PDF files in a folder into ONE Excel file:
   - Text mode  : each line of the PDF = one row, spacing preserved
   - Table mode : every table row = one row (Col_1, Col_2, ...)
   - Both       : two sheets -> "Text_Data" and "Table_Data"

 Output is saved INSIDE the selected folder as:
   <FolderName>_<MMDDYYYY>.xlsx     e.g.  New Folder_07042026.xlsx

 HOW TO INSTALL (one time)
 -------------------------
   pip install pdfplumber pandas openpyxl

   (tkinter comes with Python on Windows - no install needed)

 OCR FOR SCANNED PDFS (optional, only needed for image-only PDFs)
 -----------------------------------------------------------------
   1. pip install pymupdf pytesseract pillow
   2. Install the Tesseract OCR engine (separate program, not pip):
        Windows : https://github.com/UB-Mannheim/tesseract/wiki
                  (installs to C:\\Program Files\\Tesseract-OCR by default)
        macOS   : brew install tesseract
        Linux   : sudo apt install tesseract-ocr
   3. In the app, tick "Enable OCR for scanned pages" before starting.
      Pages that already contain real text are read normally (fast);
      only pages with no extractable text are sent through OCR (slower).

 HOW TO RUN
 ----------
   1. Open Command Prompt in the script folder (or use full path)
   2. Run:
        python pdf_extractor.py
   3. A window opens:
        - Click [Browse...] and select the folder with your PDFs
        - Choose extraction mode (Text / Tables / Both)
        - Choose date format for the output file name
        - Optionally tick "Enable OCR for scanned pages"
        - Click [START EXTRACTION]
   4. Watch progress in the log window.
      When done, the output path is shown and the folder can be opened.

 TIP: You can also double-click the .py file if Python is associated,
      or create a shortcut with:  pythonw pdf_extractor.py
      (pythonw = runs without the black console window)
=====================================================================
"""

import os
import re
import shutil
import threading
import subprocess
import sys
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pdfplumber
import pandas as pd

# ---- Optional OCR dependencies (only needed for scanned/image PDFs) ----
try:
    import fitz  # PyMuPDF - renders PDF pages to images without Poppler/ImageMagick
except ImportError:
    fitz = None

try:
    import pytesseract
    from PIL import Image
except ImportError:
    pytesseract = None
    Image = None

# Excel rejects some control characters - remove them safely
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def clean(value):
    if isinstance(value, str):
        return ILLEGAL.sub("", value)
    return value


# =====================================================================
#  OCR HELPERS (for scanned / image-only PDF pages)
# =====================================================================
def configure_tesseract():
    """Try to locate the Tesseract OCR binary. Returns True if usable."""
    if pytesseract is None:
        return False
    cmd = pytesseract.pytesseract.tesseract_cmd
    if cmd and os.path.isfile(cmd):
        return True
    for candidate in (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ):
        if os.path.isfile(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            return True
    return shutil.which("tesseract") is not None


def ocr_available():
    """True if all pieces needed for OCR fallback are installed."""
    return fitz is not None and pytesseract is not None and configure_tesseract()


def render_page_image(fitz_doc, page_num, dpi=300):
    """Render one page (1-indexed) of an open fitz document to a PIL image."""
    page = fitz_doc[page_num - 1]
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def ocr_page_text(fitz_doc, page_num, dpi=300):
    """OCR one page and return the recognized text (may be empty)."""
    img = render_page_image(fitz_doc, page_num, dpi)
    return pytesseract.image_to_string(img)


# =====================================================================
#  EXTRACTION LOGIC
# =====================================================================
def extract_text_rows(pdf_path, pdf_file, log, progress_page, use_ocr=False, ocr_dpi=300):
    """Each line = one row. layout=True keeps spacing as in the PDF.

    If a page has no extractable text (typical of a scanned/image PDF)
    and use_ocr is enabled, the page is rendered to an image and OCR'd.
    """
    rows = []
    fitz_doc = None
    if use_ocr:
        try:
            fitz_doc = fitz.open(pdf_path)
        except Exception as e:
            log(f"   [OCR] Could not open PDF for OCR rendering: {e}")

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, start=1):
            progress_page(page_num, total)
            try:
                text = page.extract_text(layout=True)  # preserve spaces
                lines = text.split("\n") if text else []
                if not any(line.strip() for line in lines) and fitz_doc is not None:
                    log(f"   [OCR] Page {page_num}: no embedded text - running OCR...")
                    try:
                        ocr_text = ocr_page_text(fitz_doc, page_num, ocr_dpi)
                        lines = ocr_text.split("\n") if ocr_text else []
                    except Exception as e:
                        log(f"   [OCR ERROR] page {page_num}: {e}")
                        lines = []
                for line in lines:
                    line = line.rstrip()  # keep LEADING spaces
                    if line:
                        rows.append({
                            "File Name": pdf_file,
                            "Page Number": page_num,
                            "Extracted Text": clean(line),
                        })
            except Exception as e:
                log(f"   [ERROR] Text - page {page_num}: {e}")

    if fitz_doc is not None:
        fitz_doc.close()
    return rows


def extract_table_rows(pdf_path, pdf_file, log, progress_page, use_ocr=False, ocr_dpi=300):
    """Every table row = one output row.

    If a page has no detectable table and use_ocr is enabled, the page is
    OCR'd and each recognized line is added as a single-column fallback row.
    """
    rows = []
    fitz_doc = None
    if use_ocr:
        try:
            fitz_doc = fitz.open(pdf_path)
        except Exception as e:
            log(f"   [OCR] Could not open PDF for OCR rendering: {e}")

    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for page_num, page in enumerate(pdf.pages, start=1):
            progress_page(page_num, total)
            try:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table:
                            rows.append([pdf_file, page_num] + [clean(c) for c in row])
                elif fitz_doc is not None:
                    log(f"   [OCR] Page {page_num}: no table detected - running OCR fallback...")
                    try:
                        ocr_text = ocr_page_text(fitz_doc, page_num, ocr_dpi)
                    except Exception as e:
                        log(f"   [OCR ERROR] page {page_num}: {e}")
                        ocr_text = ""
                    for line in (ocr_text.split("\n") if ocr_text else []):
                        line = line.rstrip()
                        if line:
                            rows.append([pdf_file, page_num, clean(line)])
            except Exception as e:
                log(f"   [ERROR] Tables - page {page_num}: {e}")

    if fitz_doc is not None:
        fitz_doc.close()
    return rows


# =====================================================================
#  GUI APPLICATION
# =====================================================================
class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("PDF Extractor - Text / Tables / Both")
        root.geometry("720x560")
        root.minsize(640, 500)

        self.folder_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="text")
        self.datefmt_var = tk.StringVar(value="%m%d%Y")
        self.ocr_var = tk.BooleanVar(value=False)
        self.running = False
        self.output_path = None

        pad = {"padx": 10, "pady": 5}

        # ---- 1) Folder selection -----------------------------------
        frm_folder = ttk.LabelFrame(root, text=" 1. Select PDF Folder ")
        frm_folder.pack(fill="x", **pad)
        ttk.Entry(frm_folder, textvariable=self.folder_var).pack(
            side="left", fill="x", expand=True, padx=(10, 5), pady=8)
        ttk.Button(frm_folder, text="Browse...",
                   command=self.browse_folder).pack(side="right", padx=10, pady=8)

        # ---- 2) Mode selection --------------------------------------
        frm_mode = ttk.LabelFrame(root, text=" 2. Extraction Mode ")
        frm_mode.pack(fill="x", **pad)
        ttk.Radiobutton(frm_mode, text="Text (line by line, keep spacing as per PDF)",
                        variable=self.mode_var, value="text").pack(anchor="w", padx=15, pady=2)
        ttk.Radiobutton(frm_mode, text="Tables",
                        variable=self.mode_var, value="tables").pack(anchor="w", padx=15, pady=2)
        ttk.Radiobutton(frm_mode, text="Both (Text + Tables in separate sheets)",
                        variable=self.mode_var, value="both").pack(anchor="w", padx=15, pady=(2, 8))

        # ---- 3) Output file date format -----------------------------
        frm_date = ttk.LabelFrame(root, text=" 3. Output File Date Format ")
        frm_date.pack(fill="x", **pad)
        ttk.Radiobutton(frm_date, text="MMDDYYYY  (e.g. FolderName_07042026.xlsx)",
                        variable=self.datefmt_var, value="%m%d%Y").pack(anchor="w", padx=15, pady=2)
        ttk.Radiobutton(frm_date, text="DDMMYYYY  (e.g. FolderName_04072026.xlsx)",
                        variable=self.datefmt_var, value="%d%m%Y").pack(anchor="w", padx=15, pady=(2, 8))

        # ---- 3b) OCR option ------------------------------------------
        frm_ocr = ttk.LabelFrame(root, text=" 4. Scanned PDFs ")
        frm_ocr.pack(fill="x", **pad)
        ttk.Checkbutton(
            frm_ocr,
            text="Enable OCR for scanned pages (slower, requires Tesseract OCR installed)",
            variable=self.ocr_var,
        ).pack(anchor="w", padx=15, pady=(2, 8))

        # ---- 4) Run button + progress -------------------------------
        frm_run = ttk.Frame(root)
        frm_run.pack(fill="x", **pad)
        self.btn_start = ttk.Button(frm_run, text="START  EXTRACTION",
                                    command=self.start_clicked)
        self.btn_start.pack(side="left", padx=(10, 5))
        self.btn_open = ttk.Button(frm_run, text="Open Output Folder",
                                   command=self.open_output_folder, state="disabled")
        self.btn_open.pack(side="left", padx=5)

        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 5))
        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(root, textvariable=self.status_var).pack(anchor="w", padx=12)

        # ---- 5) Log window ------------------------------------------
        frm_log = ttk.LabelFrame(root, text=" Log ")
        frm_log.pack(fill="both", expand=True, **pad)
        self.txt_log = tk.Text(frm_log, height=10, state="disabled",
                               font=("Consolas", 9))
        self.txt_log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scroll = ttk.Scrollbar(frm_log, command=self.txt_log.yview)
        scroll.pack(side="right", fill="y", pady=8)
        self.txt_log.configure(yscrollcommand=scroll.set)

    # ------------------------------------------------------------------
    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDF files")
        if folder:
            self.folder_var.set(folder)
            pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
            self.log(f"Folder selected: {folder}")
            self.log(f"PDF files found: {len(pdfs)}")

    def log(self, msg):
        def _append():
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", msg + "\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.root.after(0, _append)

    def set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def set_progress(self, value, maximum=None):
        def _set():
            if maximum is not None:
                self.progress.configure(maximum=maximum)
            self.progress.configure(value=value)
        self.root.after(0, _set)

    # ------------------------------------------------------------------
    def start_clicked(self):
        if self.running:
            return
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("No folder", "Please select a valid PDF folder first.")
            return
        pdf_files = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
        if not pdf_files:
            messagebox.showwarning("No PDFs", "No PDF files found in the selected folder.")
            return

        self.running = True
        self.btn_start.configure(state="disabled")
        self.btn_open.configure(state="disabled")
        self.set_progress(0, maximum=len(pdf_files))
        # Run in a background thread so the window stays responsive
        threading.Thread(target=self.run_extraction,
                         args=(folder, pdf_files), daemon=True).start()

    def run_extraction(self, folder, pdf_files):
        mode = self.mode_var.get()
        text_records, table_records = [], []
        self.log("=" * 60)
        self.log(f"Starting extraction | Mode: {mode.upper()} | Files: {len(pdf_files)}")

        use_ocr = self.ocr_var.get()
        if use_ocr:
            if ocr_available():
                self.log("OCR enabled - scanned pages will be recognized via Tesseract.")
            else:
                use_ocr = False
                self.log("[WARNING] OCR was requested but is not available.")
                self.log("          Install with: pip install pymupdf pytesseract pillow")
                self.log("          and install the Tesseract OCR engine, then restart.")
                self.log("          Continuing WITHOUT OCR - scanned pages will be skipped.")

        for idx, pdf_file in enumerate(pdf_files, start=1):
            pdf_path = os.path.join(folder, pdf_file)
            self.log(f"\n[{idx}/{len(pdf_files)}] Processing: {pdf_file}")
            self.set_status(f"Processing {idx}/{len(pdf_files)}: {pdf_file}")

            def page_progress(p, total, name=pdf_file, i=idx):
                self.set_status(f"[{i}/{len(pdf_files)}] {name} - page {p}/{total}")

            try:
                if mode in ("text", "both"):
                    rows = extract_text_rows(pdf_path, pdf_file, self.log, page_progress,
                                              use_ocr=use_ocr)
                    text_records.extend(rows)
                    self.log(f"   Text lines extracted : {len(rows)}")
                if mode in ("tables", "both"):
                    rows = extract_table_rows(pdf_path, pdf_file, self.log, page_progress,
                                               use_ocr=use_ocr)
                    table_records.extend(rows)
                    self.log(f"   Table rows extracted : {len(rows)}")
            except Exception as e:
                self.log(f"[FAILED] Could not open {pdf_file}: {e}")

            self.set_progress(idx)

        if not text_records and not table_records:
            self.log("\nNo data extracted.")
            self.finish(None)
            return

        # Output: inside selected folder -> FolderName_<date>.xlsx
        folder_name = os.path.basename(os.path.normpath(folder))
        date_str = datetime.now().strftime(self.datefmt_var.get())
        output_path = os.path.join(folder, f"{folder_name}_{date_str}.xlsx")

        self.log("\nSaving Excel file...")
        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                if text_records:
                    pd.DataFrame(text_records).to_excel(
                        writer, sheet_name="Text_Data", index=False)
                if table_records:
                    max_cols = max(len(r) for r in table_records)
                    columns = ["File Name", "Page Number"] + \
                              [f"Col_{i+1}" for i in range(max_cols - 2)]
                    padded = [r + [None] * (max_cols - len(r)) for r in table_records]
                    pd.DataFrame(padded, columns=columns).to_excel(
                        writer, sheet_name="Table_Data", index=False)
        except PermissionError:
            self.log("[ERROR] Output file is open in Excel. Close it and run again.")
            self.finish(None)
            return
        except Exception as e:
            self.log(f"[ERROR] Could not save Excel: {e}")
            self.finish(None)
            return

        self.log("=" * 60)
        self.log(f"DONE  ->  {output_path}")
        if text_records:
            self.log(f"  Text rows  : {len(text_records)}")
        if table_records:
            self.log(f"  Table rows : {len(table_records)}")
        self.finish(output_path)

    def finish(self, output_path):
        self.output_path = output_path
        def _done():
            self.running = False
            self.btn_start.configure(state="normal")
            if output_path:
                self.btn_open.configure(state="normal")
                self.set_status("Done. Output: " + output_path)
                messagebox.showinfo("Extraction complete",
                                    f"Saved to:\n{output_path}")
            else:
                self.set_status("Finished with no output.")
        self.root.after(0, _done)

    def open_output_folder(self):
        if self.output_path and os.path.exists(self.output_path):
            folder = os.path.dirname(self.output_path)
            if sys.platform.startswith("win"):
                subprocess.Popen(f'explorer /select,"{self.output_path}"')
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")  # nicer look on Windows; ignored elsewhere
    except Exception:
        pass
    PDFExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
