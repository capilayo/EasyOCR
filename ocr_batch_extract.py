"""
ocr_batch_extract.py
--------------------
Batch OCR extraction using EasyOCR.

Scans a folder for image files, runs OCR on each one, and writes the
extracted text to individual .txt files alongside a combined summary CSV.

Supported input formats: .jpg, .jpeg, .png, .bmp, .tiff, .tif, .webp

Usage
-----
    python ocr_batch_extract.py --input_dir ./my_images --output_dir ./ocr_output

    # Specify languages (default: English only)
    python ocr_batch_extract.py --input_dir ./docs --output_dir ./out --lang en tl

    # Use CPU even if GPU is available
    python ocr_batch_extract.py --input_dir ./docs --output_dir ./out --gpu False

    # Combine all text boxes into readable paragraphs
    python ocr_batch_extract.py --input_dir ./docs --output_dir ./out --paragraph

    # Set a minimum confidence threshold (0.0–1.0)
    python ocr_batch_extract.py --input_dir ./docs --output_dir ./out --min_confidence 0.5
"""

import argparse
import csv
import os
import sys
import time
from pathlib import Path

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Batch OCR extraction from a folder of images using EasyOCR."
    )
    parser.add_argument(
        "--input_dir",
        required=True,
        type=str,
        help="Path to the folder containing images to process.",
    )
    parser.add_argument(
        "--output_dir",
        required=True,
        type=str,
        help="Path to the folder where extracted text files will be saved.",
    )
    parser.add_argument(
        "--lang",
        nargs="+",
        default=["en"],
        type=str,
        help="Language code(s) for OCR recognition (default: en). "
             "Examples: en, tl, fr, ch_sim. Multiple values allowed.",
    )
    parser.add_argument(
        "--gpu",
        type=lambda x: x.lower() != "false",
        default=True,
        help="Use GPU if available (default: True). Pass False to force CPU.",
    )
    parser.add_argument(
        "--paragraph",
        action="store_true",
        default=False,
        help="Combine detected text boxes into paragraphs (default: False).",
    )
    parser.add_argument(
        "--min_confidence",
        type=float,
        default=0.0,
        help="Minimum confidence score (0.0–1.0) to include a text result (default: 0.0).",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        default=False,
        help="Recursively scan sub-folders inside input_dir (default: False).",
    )
    return parser.parse_args()


def collect_images(input_dir: str, recursive: bool) -> list[Path]:
    """Return a sorted list of image paths found in input_dir."""
    root = Path(input_dir)
    if not root.is_dir():
        print(f"[ERROR] Input directory not found: {input_dir}")
        sys.exit(1)

    pattern = "**/*" if recursive else "*"
    images = sorted(
        p for p in root.glob(pattern)
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    return images


def extract_text(reader, image_path: Path, paragraph: bool, min_confidence: float) -> list[dict]:
    """
    Run OCR on a single image and return a list of result dicts.
    Each dict has: text, confidence, bbox.
    """
    results = reader.readtext(str(image_path), paragraph=paragraph, detail=1)
    rows = []
    for item in results:
        bbox, text, confidence = item
        if confidence >= min_confidence:
            rows.append({
                "text": text,
                "confidence": round(confidence, 4),
                "bbox": bbox,
            })
    return rows


def write_text_file(rows: list[dict], output_path: Path) -> None:
    """Write extracted text lines to a plain .txt file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        for row in rows:
            f.write(row["text"] + "\n")


def write_summary_csv(summary: list[dict], output_dir: Path) -> None:
    """Append all results to a single summary CSV in output_dir."""
    csv_path = output_dir / "ocr_summary.csv"
    fieldnames = ["file", "text", "confidence", "bbox"]
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(summary)
    print(f"\n[INFO] Summary CSV written to: {csv_path}")


def main():
    args = parse_args()

    # Lazy import so the script fails clearly if easyocr is not installed
    try:
        import easyocr
    except ImportError:
        print("[ERROR] EasyOCR is not installed. Run: pip install easyocr")
        sys.exit(1)

    images = collect_images(args.input_dir, args.recursive)
    if not images:
        print(f"[WARN] No supported image files found in: {args.input_dir}")
        sys.exit(0)

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"[INFO] Found {len(images)} image(s) to process.")
    print(f"[INFO] Languages : {args.lang}")
    print(f"[INFO] GPU       : {args.gpu}")
    print(f"[INFO] Paragraph : {args.paragraph}")
    print(f"[INFO] Min conf  : {args.min_confidence}")
    print("[INFO] Loading EasyOCR model (first run downloads model weights)…\n")

    reader = easyocr.Reader(args.lang, gpu=args.gpu)

    summary: list[dict] = []
    total_start = time.time()

    for idx, img_path in enumerate(images, start=1):
        print(f"[{idx}/{len(images)}] Processing: {img_path.name}")
        t0 = time.time()

        try:
            rows = extract_text(reader, img_path, args.paragraph, args.min_confidence)
        except Exception as exc:
            print(f"         [SKIP] Error reading {img_path.name}: {exc}")
            continue

        elapsed = time.time() - t0

        # Write individual text file (preserving sub-folder structure if recursive)
        relative = img_path.relative_to(Path(args.input_dir))
        txt_output = output_dir / relative.with_suffix(".txt")
        write_text_file(rows, txt_output)

        # Accumulate for summary CSV
        for row in rows:
            summary.append({
                "file": str(relative),
                "text": row["text"],
                "confidence": row["confidence"],
                "bbox": str(row["bbox"]),
            })

        word_count = len(rows)
        print(f"         → {word_count} text block(s) extracted in {elapsed:.1f}s → {txt_output}")

    total_elapsed = time.time() - total_start
    print(f"\n[INFO] All done. {len(images)} image(s) processed in {total_elapsed:.1f}s.")

    if summary:
        write_summary_csv(summary, output_dir)
    else:
        print("[WARN] No text was extracted. Check image quality or confidence threshold.")


if __name__ == "__main__":
    main()
