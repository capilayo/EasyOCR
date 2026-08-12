"""
ocr_com_server.py
-----------------
Windows COM Server that exposes EasyOCR as an out-of-process automation object.

Excel (or any COM client) can call it like a DLL:

    Dim ocr As Object
    Set ocr = CreateObject("EasyOCR.Server")
    MsgBox ocr.ReadTextFromImage("C:\\images\\invoice.png", "en", True)

Registration (run once as Administrator):
    python ocr_com_server.py --register

Unregistration:
    python ocr_com_server.py --unregister

After registration, in Excel VBA:
    Tools > References > "EasyOCR OCR Server"
    — or use late binding: CreateObject("EasyOCR.Server")

Implementation note
-------------------
Registration writes the full absolute path of the current Python interpreter
and this script file directly into the Windows registry (LocalServer32).
This avoids the 0x80004005 / E_FAIL error that occurs when Windows resolves
a bare "pythonw.exe" to a different interpreter than the one pywin32 was
installed into.
"""

import sys
import os
import json
import csv
import time
import traceback
from pathlib import Path

# coinit_flags must only be set when running as the COM server entry point,
# NOT at module-import time (MakePyFactory re-imports this module inside the
# already-running server process, and resetting coinit_flags mid-flight crashes COM).
# It is set inside __main__ before any pythoncom import happens.

# ── COM boilerplate ──────────────────────────────────────────────────────────

# {CLSID} — regenerated once; keep stable so existing workbooks don't break
CLSID       = "{9F2A4C1B-3E7D-4A8F-B6C2-0D5E1F3A9B7C}"
PROG_ID     = "EasyOCR.Server"
DESCRIPTION = "EasyOCR OCR Server — extract text from images via COM"
VERSION     = "1.0"

SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}


# ── Core OCR helpers (shared with ocr_batch_extract / ocr_ui) ────────────────

def _collect_images(folder: str, recursive: bool) -> list:
    root = Path(folder)
    pattern = "**/*" if recursive else "*"
    return sorted(
        p for p in root.glob(pattern)
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def _get_reader(lang: str, gpu: bool):
    """Return a cached easyocr.Reader, reusing it across calls for speed."""
    import easyocr
    key = (lang, gpu)
    if not hasattr(_get_reader, "_cache"):
        _get_reader._cache = {}
    if key not in _get_reader._cache:
        _get_reader._cache[key] = easyocr.Reader([lang], gpu=gpu)
    return _get_reader._cache[key]


# ── COM server class ─────────────────────────────────────────────────────────

class EasyOCRServer:
    """
    COM-visible class.  All public methods must return COM-safe types
    (str, int, float, bool — NOT Python-specific objects).
    """

    # ── pywin32 COM registration metadata ────────────────────────────────────
    # _reg_clsid_ / _reg_progid_ are used by UseCommandLine (kept for compat).
    _reg_clsid_        = CLSID
    _reg_progid_       = PROG_ID
    _reg_desc_         = DESCRIPTION

    # ── pywin32 DesignatedWrapPolicy requirements ────────────────────────────
    # _public_methods_ tells pywin32 which methods to expose over COM.
    # Without this, DesignatedWrapPolicy._wrap_() raises ValueError and
    # Excel gets HRESULT 0x80004005 (E_FAIL).
    _public_methods_ = [
        "ReadTextFromImage",
        "ReadTextFromImageJSON",
        "ReadTextFromFolder",
        "GetVersion",
        "IsGPUAvailable",
        "GetSupportedExtensions",
    ]

    # ── Public API ───────────────────────────────────────────────────────────

    def ReadTextFromImage(
        self,
        image_path: str,
        lang: str = "en",
        use_gpu: bool = True,
        min_confidence: float = 0.0,
        paragraph: bool = False,
    ) -> str:
        """
        Run OCR on a single image file and return all extracted text as a
        newline-separated string.

        Parameters
        ----------
        image_path      : Full path to the image file.
        lang            : EasyOCR language code (default "en").
        use_gpu         : Use GPU if available (default True).
        min_confidence  : Minimum confidence threshold 0.0–1.0 (default 0.0).
        paragraph       : Merge text boxes into paragraphs (default False).

        Returns
        -------
        str  — extracted text lines joined by newlines, or an error string
               prefixed with "[ERROR]".
        """
        try:
            reader = _get_reader(lang, bool(use_gpu))
            results = reader.readtext(
                str(image_path),
                paragraph=bool(paragraph),
                detail=1,
            )
            lines = [
                text
                for (_bbox, text, conf) in results
                if conf >= float(min_confidence)
            ]
            return "\n".join(lines)
        except Exception as exc:
            return f"[ERROR] {exc}"

    def ReadTextFromImageJSON(
        self,
        image_path: str,
        lang: str = "en",
        use_gpu: bool = True,
        min_confidence: float = 0.0,
        paragraph: bool = False,
    ) -> str:
        """
        Same as ReadTextFromImage but returns a JSON string with full detail:
        [{"text": "...", "confidence": 0.98, "bbox": [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]}, ...]

        Useful for reading individual cell coordinates.
        """
        try:
            reader = _get_reader(lang, bool(use_gpu))
            results = reader.readtext(
                str(image_path),
                paragraph=bool(paragraph),
                detail=1,
            )
            items = []
            for (bbox, text, conf) in results:
                if conf >= float(min_confidence):
                    items.append({
                        "text": text,
                        "confidence": round(conf, 4),
                        "bbox": [list(pt) for pt in bbox],
                    })
            return json.dumps(items, ensure_ascii=False)
        except Exception as exc:
            return json.dumps({"error": str(exc)})

    def ReadTextFromFolder(
        self,
        input_folder: str,
        output_folder: str,
        lang: str = "en",
        use_gpu: bool = True,
        min_confidence: float = 0.0,
        paragraph: bool = False,
        recursive: bool = False,
    ) -> str:
        """
        Batch-process all images in input_folder, write individual .txt files
        to output_folder, and return the path to a summary CSV.

        Parameters
        ----------
        input_folder    : Folder containing image files.
        output_folder   : Folder where .txt results and summary CSV are written.
        lang            : EasyOCR language code (default "en").
        use_gpu         : Use GPU if available (default True).
        min_confidence  : Minimum confidence threshold (default 0.0).
        paragraph       : Merge text boxes into paragraphs (default False).
        recursive       : Scan sub-folders (default False).

        Returns
        -------
        str  — path to the generated ocr_summary.csv, or an "[ERROR]…" string.
        """
        try:
            reader = _get_reader(lang, bool(use_gpu))
            images = _collect_images(input_folder, bool(recursive))
            if not images:
                return f"[ERROR] No supported images found in: {input_folder}"

            out_dir = Path(output_folder)
            out_dir.mkdir(parents=True, exist_ok=True)

            summary = []
            for img_path in images:
                try:
                    results = reader.readtext(
                        str(img_path),
                        paragraph=bool(paragraph),
                        detail=1,
                    )
                    rows = [
                        {"text": t, "confidence": round(c, 4), "bbox": str(b)}
                        for (b, t, c) in results
                        if c >= float(min_confidence)
                    ]
                except Exception as exc:
                    rows = []

                relative = img_path.relative_to(Path(input_folder))
                txt_path = out_dir / relative.with_suffix(".txt")
                txt_path.parent.mkdir(parents=True, exist_ok=True)
                with open(txt_path, "w", encoding="utf-8") as f:
                    for row in rows:
                        f.write(row["text"] + "\n")

                for row in rows:
                    summary.append({
                        "file": str(relative),
                        "text": row["text"],
                        "confidence": row["confidence"],
                        "bbox": row["bbox"],
                    })

            csv_path = out_dir / "ocr_summary.csv"
            with open(csv_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(
                    f, fieldnames=["file", "text", "confidence", "bbox"]
                )
                writer.writeheader()
                writer.writerows(summary)

            return str(csv_path)

        except Exception as exc:
            return f"[ERROR] {exc}"

    def GetVersion(self) -> str:
        """Return the server version string."""
        return VERSION

    def IsGPUAvailable(self) -> bool:
        """Return True if a CUDA-capable GPU is detected."""
        try:
            import torch
            return bool(torch.cuda.is_available())
        except Exception:
            return False

    def GetSupportedExtensions(self) -> str:
        """Return supported image extensions as a comma-separated string."""
        return ",".join(sorted(SUPPORTED_EXTENSIONS))


# ── Registry helpers ─────────────────────────────────────────────────────────

def _reg_delete_tree(key_path: str, hive=None) -> None:
    """Delete a registry key and all its subkeys, silently ignoring missing keys."""
    import winreg
    if hive is None:
        hive = winreg.HKEY_LOCAL_MACHINE

    def _delete_recursive(hive, path):
        try:
            with winreg.OpenKey(hive, path, 0, winreg.KEY_ALL_ACCESS) as k:
                while True:
                    try:
                        subkey = winreg.EnumKey(k, 0)
                        _delete_recursive(hive, f"{path}\\{subkey}")
                    except OSError:
                        break
            winreg.DeleteKey(hive, path)
        except FileNotFoundError:
            pass
        except OSError:
            pass

    _delete_recursive(hive, key_path)


# ── Entry point (registration / unregistration) ──────────────────────────────

def _do_register() -> None:
    """
    Register using pywin32's own RegisterServer with CLSCTX_LOCAL_SERVER only.

    This writes exactly the keys pywin32 needs (including AppID, PythonCOM,
    PythonCOMPath, LocalServer32) without ever writing InprocServer32.
    """
    try:
        import pythoncom
        import pywintypes
        import win32com.server.register as reg
    except ImportError:
        print("[ERROR] pywin32 is not installed.  Run:  pip install pywin32")
        sys.exit(1)

    script_path  = os.path.abspath(__file__)
    python_exe   = sys.executable
    pythonw_exe  = os.path.join(os.path.dirname(python_exe), "pythonw.exe")
    if not os.path.isfile(pythonw_exe):
        pythonw_exe = python_exe

    module_name  = os.path.splitext(os.path.basename(script_path))[0]
    python_class = f"{module_name}.{EasyOCRServer.__name__}"
    script_dir   = os.path.dirname(script_path)
    local_server = f'"{pythonw_exe}" "{script_path}"'

    clsid_obj = pywintypes.IID(CLSID)

    reg.RegisterServer(
        clsid          = clsid_obj,
        pythonInstString = python_class,
        desc           = DESCRIPTION,
        progID         = PROG_ID,
        verProgID      = PROG_ID,
        threadingModel = "Apartment",
        clsctx         = pythoncom.CLSCTX_LOCAL_SERVER,   # NO InprocServer32
        addnPath       = script_dir,
    )

    # Overwrite LocalServer32 with our full absolute pythonw.exe path.
    # RegisterServer writes a bare "pythonw.exe" — we need the full path.
    import winreg
    key_path = f"CLSID\\{CLSID}\\LocalServer32"
    with winreg.CreateKeyEx(
        winreg.HKEY_CLASSES_ROOT, key_path, 0, winreg.KEY_SET_VALUE
    ) as k:
        winreg.SetValueEx(k, "", 0, winreg.REG_SZ, local_server)

    print(f"Registered: {PROG_ID}")
    print(f"  CLSID   : {CLSID}")
    print(f"  Launcher: {local_server}")
    print(f"  Class   : {python_class}")
    print(f"  Path    : {script_dir}")


def _do_unregister() -> None:
    """Remove ALL registry entries for this CLSID — including stale pywin32 keys."""
    try:
        import pywintypes
        import win32com.server.register as reg
    except ImportError:
        print("[ERROR] pywin32 is not installed.  Run:  pip install pywin32")
        sys.exit(1)

    clsid_obj = pywintypes.IID(CLSID)
    reg.UnregisterServer(clsid_obj, PROG_ID)

    # Also nuke any leftover keys from manual or old registrations
    import winreg
    clsid_hkcr = f"CLSID\\{CLSID}"
    progid_hkcr = PROG_ID

    for sub in (
        "LocalServer32", "InprocServer32", "InprocHandler32",
        "ProgID", "VersionIndependentProgID", "Programmable",
        "PythonCOM", "PythonCOMPath", "Debugging", "AppID",
        r"Implemented Categories\{B3EF80D0-68E2-11D0-A689-00C04FD658FF}",
        "Implemented Categories",
    ):
        _reg_delete_tree(f"CLSID\\{CLSID}\\{sub}", winreg.HKEY_CLASSES_ROOT)
    _reg_delete_tree(clsid_hkcr, winreg.HKEY_CLASSES_ROOT)
    _reg_delete_tree(progid_hkcr, winreg.HKEY_CLASSES_ROOT)

    # Also clean HKLM mirror
    hklm_base = f"SOFTWARE\\Classes\\CLSID\\{CLSID}"
    for sub in (
        "LocalServer32", "InprocServer32", "InprocHandler32",
        "ProgID", "VersionIndependentProgID", "Programmable",
        "PythonCOM", "PythonCOMPath", "Debugging",
        r"Implemented Categories\{B3EF80D0-68E2-11D0-A689-00C04FD658FF}",
        "Implemented Categories",
    ):
        _reg_delete_tree(f"{hklm_base}\\{sub}")
    _reg_delete_tree(hklm_base)
    _reg_delete_tree(f"SOFTWARE\\Classes\\{PROG_ID}\\CLSID")
    _reg_delete_tree(f"SOFTWARE\\Classes\\{PROG_ID}")

    print(f"Unregistered: {PROG_ID}")


# ── Log helper (pythonw.exe has no console — write errors to a file) ─────────

_LOG_PATH = Path(os.environ.get("TEMP", os.path.expanduser("~"))) / "easyocr_com_server.log"


def _log(msg: str) -> None:
    """Append a timestamped line to the log file."""
    import datetime
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")
    except Exception:
        pass  # never crash the server trying to log


def _serve_localserver() -> None:
    """
    LocalServer32 COM entry point — called when Windows COM launches us.

    Uses pywin32's MakePyFactory which reads PythonCOM / PythonCOMPath from
    the registry (written by RegisterServer) to find and instantiate our class.
    """
    _log(f"_serve_localserver() called  argv={sys.argv}")
    try:
        import pythoncom
        import pywintypes
        from win32com.server import factory as _factory
        import win32api
    except ImportError as exc:
        _log(f"ImportError: {exc}  — run: pip install pywin32")
        sys.exit(1)

    try:
        # RegisterClassFactories expects plain strings (it does clsid[0] internally)
        infos = _factory.RegisterClassFactories([CLSID])
        _log("RegisterClassFactories OK")

        pythoncom.EnableQuitMessage(win32api.GetCurrentThreadId())
        pythoncom.CoResumeClassObjects()
        _log("Entering PumpMessages loop")

        pythoncom.PumpMessages()

        _factory.RevokeClassFactories(infos)
        pythoncom.CoUninitialize()
        _log("Server exited cleanly")

    except Exception:
        _log("EXCEPTION in _serve_localserver:\n" + traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    # Set STA apartment model BEFORE any pythoncom/win32com import.
    # This must be here (not at module level) so MakePyFactory's secondary
    # import of this module does NOT reset coinit_flags mid-flight.
    sys.coinit_flags = 2

    args = sys.argv[1:]  # everything after the script name

    if "--register" in args:
        try:
            _do_register()
            sys.exit(0)
        except Exception as exc:
            print(f"[ERROR] Registration failed: {exc}")
            sys.exit(1)

    elif "--unregister" in args:
        try:
            _do_unregister()
            sys.exit(0)
        except Exception as exc:
            print(f"[ERROR] Unregistration failed: {exc}")
            sys.exit(1)

    elif "--debug" in args:
        # Run the server in a visible console for troubleshooting.
        # python ocr_com_server.py --debug
        print(f"[DEBUG] Starting COM server in debug mode")
        print(f"[DEBUG] Log file: {_LOG_PATH}")
        print(f"[DEBUG] CLSID  : {CLSID}")
        print(f"[DEBUG] ProgID : {PROG_ID}")
        _serve_localserver()

    elif not args or args == ["-Embedding"] or (len(args) == 1 and args[0].startswith("-")):
        # Windows COM launches us with "-Embedding" (or no extra args).
        # Start the local server factory loop silently.
        _log(f"Launched by COM  argv={sys.argv}")
        _serve_localserver()

    else:
        print(
            f"EasyOCR COM Server  v{VERSION}\n"
            f"ProgID : {PROG_ID}\n"
            f"CLSID  : {CLSID}\n\n"
            "Usage:\n"
            "  Register   : python ocr_com_server.py --register\n"
            "  Unregister : python ocr_com_server.py --unregister\n"
            "  Debug mode : python ocr_com_server.py --debug\n"
            f"\nLog file   : {_LOG_PATH}\n"
        )
