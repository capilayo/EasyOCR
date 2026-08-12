# EasyOCR — Excel COM Server (OCX/DLL-style)

Expose **EasyOCR** image text-extraction to **Microsoft Excel** (and any
Windows COM client) without a single line of C++.  The Python COM server
registers as a Windows automation object; Excel's VBA calls it exactly like
a `.dll` or `.ocx` reference.

---

## Architecture

```
Excel VBA  ──CreateObject("EasyOCR.Server")──▶  Windows COM Registry
                                                        │
                                                        ▼
                                               ocr_com_server.py
                                               (pywin32 COM server)
                                                        │
                                                        ▼
                                                EasyOCR (Python)
                                                        │
                                                        ▼
                                             PyTorch / model weights
```

| File | Purpose |
|---|---|
| `ocr_com_server.py` | COM server — the core component |
| `register_com_server.bat` | One-click register / unregister |
| `build_com_server.bat` | Package into a standalone `.exe` (no Python needed on end-user machines) |
| `EasyOCR_COMServer.spec` | PyInstaller spec used by the build script |
| `excel_vba_sample.bas` | Ready-to-import VBA module with 6 examples |

---

## Quick Start (Developer Machine — Python installed)

### 1. Install dependencies

```powershell
pip install pywin32 easyocr torch torchvision
python -m pywin32_postinstall -install   # only needed once after fresh pip install
```

### 2. Register the COM Server (run once, as Administrator)

Double-click **`register_com_server.bat`** → choose **1 (Register)**.

Or from an elevated prompt:

```bat
python ocr_com_server.py --register
```

### 3. Use from Excel VBA

Open Excel → `Alt+F11` → Import `excel_vba_sample.bas` (or paste any example):

```vb
Sub ExtractInvoiceText()
    Dim ocr As Object
    Set ocr = CreateObject("EasyOCR.Server")

    Dim text As String
    text = ocr.ReadTextFromImage("C:\invoices\inv001.png", "en", True, 0.3, False)

    Sheet1.Range("A1").Value = text
    Set ocr = Nothing
End Sub
```

---

## Public API

### `ReadTextFromImage(imagePath, lang, useGPU, minConfidence, paragraph) → String`

Returns all extracted text as newline-separated lines.

| Parameter | Type | Default | Description |
|---|---|---|---|
| `imagePath` | String | — | Full path to image (.jpg .png .bmp .tiff .webp) |
| `lang` | String | `"en"` | EasyOCR language code |
| `useGPU` | Boolean | `True` | Use CUDA GPU if available |
| `minConfidence` | Double | `0.0` | Discard results below this score (0.0–1.0) |
| `paragraph` | Boolean | `False` | Merge text boxes into paragraphs |

---

### `ReadTextFromImageJSON(imagePath, lang, useGPU, minConfidence, paragraph) → String`

Same as above but returns a **JSON string** with full detail per text block:

```json
[
  { "text": "INVOICE", "confidence": 0.9981, "bbox": [[10,5],[120,5],[120,25],[10,25]] },
  ...
]
```

---

### `ReadTextFromFolder(inputFolder, outputFolder, lang, useGPU, minConfidence, paragraph, recursive) → String`

Batch-process an entire folder.  Returns the path to `ocr_summary.csv`.

---

### `GetVersion() → String`

Returns `"1.0"`.

### `IsGPUAvailable() → Boolean`

Returns `True` if a CUDA-capable GPU is detected.

### `GetSupportedExtensions() → String`

Returns `".bmp,.jpg,.jpeg,.png,.tiff,.tif,.webp"`.

---

## Distribute to Machines Without Python

Use the build script to create a self-contained `.exe`:

```bat
build_com_server.bat
```

This runs PyInstaller and produces `dist\EasyOCR_COMServer.exe`.  On the
target machine (Administrator prompt):

```bat
EasyOCR_COMServer.exe --register
```

No Python, pip, or any other dependency is needed on the end-user machine.

---

## Language Codes

| Language | Code | Language | Code |
|---|---|---|---|
| English | `en` | Filipino/Tagalog | `tl` |
| French | `fr` | Spanish | `es` |
| German | `de` | Italian | `it` |
| Portuguese | `pt` | Arabic | `ar` |
| Hindi | `hi` | Russian | `ru` |
| Chinese (Simplified) | `ch_sim` | Chinese (Traditional) | `ch_tra` |
| Japanese | `ja` | Korean | `ko` |
| Thai | `th` | | |

Full list: https://www.jaided.ai/easyocr/

---

## Unregister

```bat
register_com_server.bat   → choose 2 (Unregister)
```

Or:

```bat
python ocr_com_server.py --unregister
```

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ActiveX component can't create object` | Run `register_com_server.bat` as Administrator |
| `[ERROR] EasyOCR is not installed` | `pip install easyocr` |
| `[ERROR] torch not found` | `pip install torch torchvision` |
| GPU not used | Install CUDA-enabled PyTorch from https://pytorch.org |
| Slow first run | EasyOCR downloads model weights (~100 MB) on first use — subsequent calls are fast |
