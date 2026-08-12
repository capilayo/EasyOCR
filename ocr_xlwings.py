"""
ocr_xlwings.py
--------------
EasyOCR bridge for Microsoft Excel via xlwings RunPython.

SETUP (one-time, no Admin required):
  1. pip install xlwings easyocr
  2. In Excel: Tools > References > tick "xlwings"
     (or install the xlwings addin: xlwings addin install)
  3. Copy the VBA from ocr_xlwings_vba.bas into your workbook.

USAGE from Excel VBA:
  ' Extract text from a single image:
  RunPython "import ocr_xlwings; ocr_xlwings.read_image()"

  ' Or call with arguments via the helper functions in the .bas file.
"""

import os
import sys
import json
from pathlib import Path

# ── xlwings import (free tier — no license needed for RunPython) ─────────────
try:
    import xlwings as xw
except ImportError:
    raise ImportError("xlwings not installed. Run: pip install xlwings")

# ── Lazy EasyOCR reader cache ────────────────────────────────────────────────
_readers: dict = {}


def _get_reader(lang: str = "en", gpu: bool = True):
    """Return a cached easyocr.Reader for the given language + GPU setting."""
    try:
        import easyocr
    except ImportError:
        raise ImportError("easyocr not installed. Run: pip install easyocr")
    key = (lang, bool(gpu))
    if key not in _readers:
        _readers[key] = easyocr.Reader([lang], gpu=bool(gpu))
    return _readers[key]


SUPPORTED = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


# ── Public functions (called via RunPython) ──────────────────────────────────

def read_image():
    """
    Read the image path from cell B1 of the active sheet.
    Optional settings in C1:D1 (lang, min_confidence).
    Writes extracted text lines into column A starting at A3.

    Excel layout expected:
      B1 = full path to image file
      C1 = language code (default: en)
      D1 = min confidence 0.0-1.0 (default: 0.0)
    """
    wb  = xw.Book.caller()
    sht = wb.sheets.active

    img_path = str(sht.range("B1").value or "").strip()
    lang     = str(sht.range("C1").value or "en").strip()
    min_conf = float(sht.range("D1").value or 0.0)

    if not img_path:
        sht.range("A3").value = "[ERROR] Enter an image path in cell B1."
        return

    if not os.path.isfile(img_path):
        sht.range("A3").value = f"[ERROR] File not found: {img_path}"
        return

    sht.range("A3").value = "Processing… please wait"
    wb.app.screen_updating = True

    try:
        reader  = _get_reader(lang, gpu=True)
        results = reader.readtext(img_path, detail=1)
        rows = [
            (text, round(conf, 4))
            for (_bbox, text, conf) in results
            if conf >= min_conf
        ]
    except Exception as exc:
        sht.range("A3").value = f"[ERROR] {exc}"
        return

    # Clear old results
    last_row = sht.cells.last_cell.row
    sht.range(f"A3:B{last_row}").clear_contents()

    if not rows:
        sht.range("A3").value = "[No text found — try lowering Min Confidence]"
        return

    # Write results: col A = text, col B = confidence
    sht.range("A2").value = "Extracted Text"
    sht.range("B2").value = "Confidence"
    for i, (text, conf) in enumerate(rows):
        sht.range(f"A{i + 3}").value = text
        sht.range(f"B{i + 3}").value = conf


def read_image_from_cell(cell_address: str = "B1"):
    """
    Variant: reads image path from a named cell address (default B1).
    Writes all extracted text joined by newlines into the cell directly
    below the path cell.
    """
    wb  = xw.Book.caller()
    sht = wb.sheets.active

    img_path = str(sht.range(cell_address).value or "").strip()
    if not img_path or not os.path.isfile(img_path):
        return

    try:
        reader  = _get_reader("en", gpu=True)
        results = reader.readtext(img_path, detail=1)
        text    = "\n".join(t for (_, t, c) in results if c >= 0.0)
    except Exception as exc:
        text = f"[ERROR] {exc}"

    # Write result one row below the input cell
    col   = sht.range(cell_address).column
    row   = sht.range(cell_address).row + 1
    sht.cells(row, col).value = text


def batch_folder():
    """
    Read input/output folder paths from the active sheet:
      B1 = input folder
      B2 = output folder
      C1 = language (default: en)
      D1 = min confidence (default: 0.0)

    Processes all images, writes a summary table starting at A5,
    and saves ocr_summary.csv to the output folder.
    """
    import csv

    wb  = xw.Book.caller()
    sht = wb.sheets.active

    input_dir  = str(sht.range("B1").value or "").strip()
    output_dir = str(sht.range("B2").value or "").strip()
    lang       = str(sht.range("C1").value or "en").strip()
    min_conf   = float(sht.range("D1").value or 0.0)

    if not input_dir or not os.path.isdir(input_dir):
        sht.range("A5").value = f"[ERROR] Input folder not found: {input_dir}"
        return
    if not output_dir:
        sht.range("A5").value = "[ERROR] Enter an output folder path in B2."
        return

    images = sorted(
        p for p in Path(input_dir).glob("*")
        if p.is_file() and p.suffix.lower() in SUPPORTED
    )
    if not images:
        sht.range("A5").value = f"[ERROR] No supported images in: {input_dir}"
        return

    Path(output_dir).mkdir(parents=True, exist_ok=True)
    sht.range("A5").value = f"Processing {len(images)} image(s)…"
    wb.app.screen_updating = True

    try:
        reader = _get_reader(lang, gpu=True)
    except Exception as exc:
        sht.range("A5").value = f"[ERROR] {exc}"
        return

    summary = []
    for img_path in images:
        try:
            results = reader.readtext(str(img_path), detail=1)
            for (bbox, text, conf) in results:
                if conf >= min_conf:
                    summary.append({
                        "file": img_path.name,
                        "text": text,
                        "confidence": round(conf, 4),
                        "bbox": str(bbox),
                    })
            # Write per-image text file
            txt_path = Path(output_dir) / img_path.with_suffix(".txt").name
            with open(txt_path, "w", encoding="utf-8") as f:
                for row in summary:
                    if row["file"] == img_path.name:
                        f.write(row["text"] + "\n")
        except Exception:
            pass

    # Write CSV
    csv_path = Path(output_dir) / "ocr_summary.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f, fieldnames=["file", "text", "confidence", "bbox"]
        )
        writer.writeheader()
        writer.writerows(summary)

    # Write summary to sheet
    last_row = sht.cells.last_cell.row
    sht.range(f"A5:D{last_row}").clear_contents()
    sht.range("A4").value = [["File", "Text", "Confidence", "BBox"]]
    for i, row in enumerate(summary):
        sht.range(f"A{i + 5}").value = [
            row["file"], row["text"], row["confidence"], row["bbox"]
        ]

    sht.range("A3").value = f"Done — {len(summary)} text blocks from {len(images)} image(s). CSV: {csv_path}"


def get_version():
    """Write server/library version info into cell A1."""
    wb  = xw.Book.caller()
    sht = wb.sheets.active
    try:
        import easyocr
        ev = easyocr.__version__
    except Exception:
        ev = "not installed"
    try:
        import torch
        gpu = str(torch.cuda.is_available())
    except Exception:
        gpu = "unknown"

    sht.range("A1").value = (
        f"EasyOCR {ev} | xlwings {xw.__version__} | GPU: {gpu}"
    )
