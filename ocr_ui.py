"""
ocr_ui.py
---------
Tkinter GUI for batch OCR extraction using EasyOCR.

Run:
    python ocr_ui.py
"""

import csv
import os
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}

LANG_OPTIONS = [
    ("English",                  "en"),
    ("Filipino / Tagalog",       "tl"),
    ("French",                   "fr"),
    ("Spanish",                  "es"),
    ("German",                   "de"),
    ("Italian",                  "it"),
    ("Portuguese",               "pt"),
    ("Chinese Simplified",       "ch_sim"),
    ("Chinese Traditional",      "ch_tra"),
    ("Japanese",                 "ja"),
    ("Korean",                   "ko"),
    ("Arabic",                   "ar"),
    ("Hindi",                    "hi"),
    ("Russian",                  "ru"),
    ("Thai",                     "th"),
]


# ─── Extraction logic (same as ocr_batch_extract.py) ────────────────────────

def collect_images(input_dir: str, recursive: bool) -> list:
    root = Path(input_dir)
    pattern = "**/*" if recursive else "*"
    return sorted(
        p for p in root.glob(pattern)
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def extract_text(reader, image_path: Path, paragraph: bool, min_confidence: float) -> list:
    results = reader.readtext(str(image_path), paragraph=paragraph, detail=1)
    rows = []
    for item in results:
        bbox, text, confidence = item
        if confidence >= min_confidence:
            rows.append({"text": text, "confidence": round(confidence, 4), "bbox": bbox})
    return rows


def write_text_file(rows: list, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        for row in rows:
            f.write(row["text"] + "\n")


def write_summary_csv(summary: list, output_dir: Path) -> None:
    csv_path = output_dir / "ocr_summary.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["file", "text", "confidence", "bbox"])
        writer.writeheader()
        writer.writerows(summary)
    return csv_path


# ─── GUI ────────────────────────────────────────────────────────────────────

class OCRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EasyOCR — Batch Extractor")
        self.resizable(False, False)
        self.configure(bg="#f0f2f5")
        self._build_ui()
        self._extraction_running = False

    # ── Layout ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        PAD = {"padx": 16, "pady": 8}

        # ── Header ──────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#1e3a5f", pady=14)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="EasyOCR  Batch Extractor",
            font=("Segoe UI", 16, "bold"),
            fg="#ffffff", bg="#1e3a5f"
        ).pack()
        tk.Label(
            hdr, text="Extract text from images in a folder using EasyOCR",
            font=("Segoe UI", 9), fg="#aac4e0", bg="#1e3a5f"
        ).pack()

        # ── Main body ───────────────────────────────────────────────────────
        body = tk.Frame(self, bg="#f0f2f5", padx=20, pady=14)
        body.pack(fill="both")

        # Input folder
        self._input_var = tk.StringVar()
        self._make_folder_row(
            body, "Input Image Folder", self._input_var,
            self._browse_input, row=0
        )

        # Output folder
        self._output_var = tk.StringVar()
        self._make_folder_row(
            body, "Output Folder", self._output_var,
            self._browse_output, row=1
        )

        # ── Options row ─────────────────────────────────────────────────────
        opt_frame = tk.LabelFrame(
            body, text=" Options ", font=("Segoe UI", 9),
            bg="#f0f2f5", fg="#1e3a5f", bd=1, relief="groove"
        )
        opt_frame.grid(row=2, column=0, columnspan=3, sticky="ew",
                       padx=0, pady=(8, 4))
        opt_frame.columnconfigure(0, weight=1)
        opt_frame.columnconfigure(1, weight=1)
        opt_frame.columnconfigure(2, weight=1)
        opt_frame.columnconfigure(3, weight=1)

        # Language selector
        tk.Label(opt_frame, text="Language:", bg="#f0f2f5",
                 font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        # _lang_display holds the combobox label; _lang_code holds the bare code
        self._lang_display = tk.StringVar()
        self._lang_code = "en"
        lang_combo = ttk.Combobox(
            opt_frame, textvariable=self._lang_display,
            values=[f"{lbl} [{code}]" for lbl, code in LANG_OPTIONS],
            state="readonly", width=22
        )
        lang_combo.current(0)
        self._lang_display.set(f"{LANG_OPTIONS[0][0]} [{LANG_OPTIONS[0][1]}]")
        lang_combo.grid(row=0, column=1, sticky="w", padx=4, pady=6)
        lang_combo.bind("<<ComboboxSelected>>", self._on_lang_select)

        # GPU toggle
        self._gpu_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opt_frame, text="Use GPU (if available)",
            variable=self._gpu_var
        ).grid(row=0, column=2, sticky="w", padx=12, pady=6)

        # Paragraph mode
        self._para_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame, text="Paragraph mode",
            variable=self._para_var
        ).grid(row=0, column=3, sticky="w", padx=12, pady=6)

        # Recursive + confidence on second row
        self._recursive_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_frame, text="Scan sub-folders recursively",
            variable=self._recursive_var
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 6))

        tk.Label(opt_frame, text="Min. confidence:", bg="#f0f2f5",
                 font=("Segoe UI", 9)).grid(row=1, column=2, sticky="e", padx=(12, 4), pady=(0, 6))
        self._conf_var = tk.DoubleVar(value=0.0)
        conf_spin = ttk.Spinbox(
            opt_frame, from_=0.0, to=1.0, increment=0.05,
            textvariable=self._conf_var, width=7, format="%.2f"
        )
        conf_spin.grid(row=1, column=3, sticky="w", padx=4, pady=(0, 6))

        # ── Begin Extraction button ──────────────────────────────────────────
        self._btn = tk.Button(
            body,
            text="▶  Begin Extraction",
            font=("Segoe UI", 11, "bold"),
            bg="#1e6f3e", fg="#ffffff",
            activebackground="#155c33", activeforeground="#ffffff",
            relief="flat", cursor="hand2",
            padx=24, pady=10,
            command=self._start_extraction
        )
        self._btn.grid(row=3, column=0, columnspan=3, pady=(14, 4))

        # ── Progress bar ────────────────────────────────────────────────────
        self._progress = ttk.Progressbar(body, orient="horizontal",
                                          length=560, mode="determinate")
        self._progress.grid(row=4, column=0, columnspan=3, pady=(4, 0), sticky="ew")
        self._progress_label = tk.Label(
            body, text="", bg="#f0f2f5",
            font=("Segoe UI", 8), fg="#57606a"
        )
        self._progress_label.grid(row=5, column=0, columnspan=3, sticky="w")

        # ── Log panel ───────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(
            body, text=" Log ", font=("Segoe UI", 9),
            bg="#f0f2f5", fg="#1e3a5f", bd=1, relief="groove"
        )
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew",
                       pady=(10, 0))
        body.rowconfigure(6, weight=1)
        body.columnconfigure(1, weight=1)

        self._log = tk.Text(
            log_frame, height=14, width=80,
            font=("Consolas", 9),
            bg="#1b1f23", fg="#c9d1d9",
            insertbackground="#c9d1d9",
            relief="flat", state="disabled",
            wrap="word"
        )
        scrollbar = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=scrollbar.set)
        self._log.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        scrollbar.pack(side="right", fill="y", pady=4)

        # log colour tags
        self._log.tag_config("info",  foreground="#79c0ff")
        self._log.tag_config("ok",    foreground="#56d364")
        self._log.tag_config("warn",  foreground="#e3b341")
        self._log.tag_config("error", foreground="#f85149")
        self._log.tag_config("dim",   foreground="#8b949e")

        # ── Footer ──────────────────────────────────────────────────────────
        tk.Label(
            self, text="Made with IBM Bob",
            font=("Segoe UI", 8), fg="#8b949e", bg="#f0f2f5"
        ).pack(pady=(6, 8))

    def _make_folder_row(self, parent, label, var, cmd, row):
        tk.Label(
            parent, text=label + ":", bg="#f0f2f5",
            font=("Segoe UI", 9, "bold"), fg="#1e3a5f", width=18, anchor="w"
        ).grid(row=row, column=0, sticky="w", pady=5)

        entry = tk.Entry(
            parent, textvariable=var,
            font=("Segoe UI", 9), width=46,
            relief="solid", bd=1
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(6, 6), pady=5)

        tk.Button(
            parent, text="Browse…",
            font=("Segoe UI", 9),
            bg="#3b82d4", fg="#ffffff",
            activebackground="#2563c0", activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=10,
            command=cmd
        ).grid(row=row, column=2, pady=5)

    # ── Event handlers ───────────────────────────────────────────────────────

    def _browse_input(self):
        folder = filedialog.askdirectory(title="Select Input Image Folder")
        if folder:
            self._input_var.set(folder)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self._output_var.set(folder)

    def _on_lang_select(self, event):
        selected = event.widget.get()
        # extract bare code from "Label [code]"
        self._lang_code = selected.split("[")[-1].rstrip("]")

    def _log_write(self, msg: str, tag: str = ""):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _set_progress(self, current: int, total: int, label: str = ""):
        pct = int((current / total) * 100) if total else 0
        self._progress["value"] = pct
        self._progress_label.config(
            text=label or f"{current} / {total}  ({pct}%)"
        )
        self.update_idletasks()

    # ── Extraction ───────────────────────────────────────────────────────────

    def _start_extraction(self):
        if self._extraction_running:
            return

        input_dir = self._input_var.get().strip()
        output_dir = self._output_var.get().strip()

        if not input_dir:
            messagebox.showwarning("Missing Input", "Please select an Input Image Folder.")
            return
        if not os.path.isdir(input_dir):
            messagebox.showerror("Invalid Folder", f"Input folder does not exist:\n{input_dir}")
            return
        if not output_dir:
            messagebox.showwarning("Missing Output", "Please select an Output Folder.")
            return

        # Clear log
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

        self._btn.config(state="disabled", text="⏳  Extracting…", bg="#555")
        self._extraction_running = True

        thread = threading.Thread(target=self._run_extraction, daemon=True)
        thread.start()

    def _run_extraction(self):
        input_dir  = self._input_var.get().strip()
        output_dir = self._output_var.get().strip()
        lang_code  = self._lang_code
        use_gpu    = self._gpu_var.get()
        paragraph  = self._para_var.get()
        recursive  = self._recursive_var.get()
        min_conf   = round(self._conf_var.get(), 2)

        try:
            import easyocr
        except ImportError:
            self._log_write("[ERROR] EasyOCR is not installed. Run: pip install easyocr", "error")
            self._finish_extraction(success=False)
            return

        images = collect_images(input_dir, recursive)
        if not images:
            self._log_write(f"[WARN] No supported images found in: {input_dir}", "warn")
            self._finish_extraction(success=False)
            return

        Path(output_dir).mkdir(parents=True, exist_ok=True)

        self._log_write(f"[INFO] Found {len(images)} image(s) to process.", "info")
        self._log_write(f"[INFO] Language   : {lang_code}", "dim")
        self._log_write(f"[INFO] GPU        : {use_gpu}", "dim")
        self._log_write(f"[INFO] Paragraph  : {paragraph}", "dim")
        self._log_write(f"[INFO] Min conf   : {min_conf}", "dim")
        self._log_write(f"[INFO] Recursive  : {recursive}", "dim")
        self._log_write("[INFO] Loading model (first run downloads weights)…", "info")

        self._set_progress(0, len(images), "Loading model…")

        try:
            reader = easyocr.Reader([lang_code], gpu=use_gpu)
        except Exception as exc:
            self._log_write(f"[ERROR] Failed to load model: {exc}", "error")
            self._finish_extraction(success=False)
            return

        self._log_write("[INFO] Model loaded. Starting extraction.\n", "ok")

        summary = []
        total_start = time.time()

        for idx, img_path in enumerate(images, start=1):
            self._log_write(f"[{idx}/{len(images)}]  {img_path.name}", "info")
            self._set_progress(idx - 1, len(images), f"Processing {idx} of {len(images)}: {img_path.name}")
            t0 = time.time()

            try:
                rows = extract_text(reader, img_path, paragraph, min_conf)
            except Exception as exc:
                self._log_write(f"        [SKIP] Error: {exc}", "error")
                continue

            elapsed = time.time() - t0
            relative = img_path.relative_to(Path(input_dir))
            txt_output = Path(output_dir) / relative.with_suffix(".txt")
            write_text_file(rows, txt_output)

            for row in rows:
                summary.append({
                    "file": str(relative),
                    "text": row["text"],
                    "confidence": row["confidence"],
                    "bbox": str(row["bbox"]),
                })

            self._log_write(
                f"        → {len(rows)} block(s)  |  "
                f"conf avg: {(sum(r['confidence'] for r in rows)/len(rows)):.2f}  |  "
                f"{elapsed:.1f}s  →  {txt_output.name}",
                "ok" if rows else "warn"
            )

        self._set_progress(len(images), len(images), "Writing summary CSV…")
        csv_path = None
        if summary:
            csv_path = write_summary_csv(summary, Path(output_dir))
            self._log_write(f"\n[INFO] Summary CSV → {csv_path}", "ok")
        else:
            self._log_write("\n[WARN] No text extracted. Check image quality or confidence threshold.", "warn")

        total_elapsed = time.time() - total_start
        self._log_write(
            f"[INFO] Done. {len(images)} image(s) in {total_elapsed:.1f}s.", "ok"
        )
        self._set_progress(len(images), len(images),
                           f"Complete — {len(images)} image(s) processed in {total_elapsed:.1f}s")

        self._finish_extraction(success=True, csv_path=csv_path)

    def _finish_extraction(self, success: bool, csv_path=None):
        self._extraction_running = False
        self._btn.config(state="normal", text="▶  Begin Extraction", bg="#1e6f3e")
        if success and csv_path:
            if messagebox.askyesno(
                "Extraction Complete",
                f"Extraction complete!\n\nSummary CSV saved to:\n{csv_path}\n\nOpen output folder?"
            ):
                os.startfile(str(Path(csv_path).parent))


# ─── Entry point ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = OCRApp()
    app.mainloop()
