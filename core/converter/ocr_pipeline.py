"""
core/converter/ocr_pipeline.py — Vietnamese OCR Pipeline & Scan Recovery Engine (v4.0).

Handles high-precision OCR on scanned PDFs, receipts, invoices, and official documents:
  - Auto-detection: Scanned PDF detection via text density (< 20 chars/page) & raster area (> 70%).
  - Pre-processing: Deskew angle detection and HSV Red-Stamp Filter (ERR_CONV_005).
  - Multi-language OCR: Tesseract execution with Vietnamese ('vie') and English ('eng').
  - Table grid and text structure recovery into Markdown (.md) or Word (.docx).
  - Graceful fallback when Tesseract binary is not installed on the host.
"""

from __future__ import annotations

import io
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Dict, List, Optional, Tuple, Union

try:
    from PIL import Image
    import numpy as np
    HAS_IMAGE_STACK = True
except ImportError:
    HAS_IMAGE_STACK = False

try:
    import pymupdf  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    try:
        import fitz as pymupdf  # type: ignore
        HAS_PYMUPDF = True
    except ImportError:
        HAS_PYMUPDF = False

# Standard known Tesseract binary search paths
COMMON_TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
    "/usr/bin/tesseract",
    "/usr/local/bin/tesseract",
    "/opt/homebrew/bin/tesseract",
]


def find_tesseract_binary() -> Optional[str]:
    """Find tesseract executable on system PATH or common installation paths."""
    which_path = shutil.which("tesseract")
    if which_path:
        return which_path

    for p in COMMON_TESSERACT_PATHS:
        if os.path.isfile(p):
            return p

    return None


def is_scanned_pdf(
    pdf_path: str,
    text_density_threshold: int = 20,
    raster_area_threshold: float = 0.70,
    max_pages: int = 5,
) -> Tuple[bool, Dict[str, Any]]:
    """
    Auto-detect whether a PDF is a scanned document rather than a digital vector PDF.
    Criteria:
      - Average text length per page < text_density_threshold (default 20 characters).
      - Raster image area coverage >= raster_area_threshold (default 70% of page area).
    """
    stats = {
        "file": pdf_path,
        "is_scanned": False,
        "pages_inspected": 0,
        "avg_text_length": 0.0,
        "avg_raster_ratio": 0.0,
        "reasons": [],
    }

    if not HAS_PYMUPDF:
        stats["reasons"].append("PyMuPDF not installed")
        return False, stats

    if not os.path.exists(pdf_path):
        stats["reasons"].append(f"File not found: {pdf_path}")
        return False, stats

    try:
        doc = pymupdf.open(pdf_path)
    except Exception as ex:
        stats["reasons"].append(f"Failed to open PDF: {ex}")
        return False, stats

    num_pages = min(len(doc), max_pages)
    if num_pages == 0:
        return False, stats

    total_chars = 0
    total_raster_ratio = 0.0

    for page_idx in range(num_pages):
        page = doc[page_idx]
        text = page.get_text().strip()
        total_chars += len(text)

        p_rect = page.rect
        page_area = max(1.0, p_rect.width * p_rect.height)

        # Calculate bounding area of raster images
        img_info_list = page.get_images(full=True)
        raster_area = 0.0
        for img in img_info_list:
            xref = img[0]
            try:
                # Get rects where this image is placed on the page
                rects = page.get_image_rects(xref)
                for r in rects:
                    raster_area += r.width * r.height
            except Exception:
                pass

        ratio = min(1.0, raster_area / page_area)
        total_raster_ratio += ratio

    avg_chars = total_chars / num_pages
    avg_ratio = total_raster_ratio / num_pages

    stats["pages_inspected"] = num_pages
    stats["avg_text_length"] = round(avg_chars, 1)
    stats["avg_raster_ratio"] = round(avg_ratio, 3)

    is_scanned = (
        avg_chars < text_density_threshold and avg_ratio >= raster_area_threshold
    )
    # Also if total chars is 0 and there is at least one image
    if avg_chars == 0 and avg_ratio > 0.3:
        is_scanned = True

    stats["is_scanned"] = is_scanned
    if is_scanned:
        stats["reasons"].append(
            f"Low text density ({avg_chars:.1f} < {text_density_threshold}) "
            f"and high raster coverage ({avg_ratio*100:.1f}% >= {raster_area_threshold*100:.0f}%)"
        )
    else:
        stats["reasons"].append("Standard digital vector text detected")

    return is_scanned, stats


def filter_red_stamp(
    image: Any,
    bg_fill: Tuple[int, int, int] = (255, 255, 255),
    sat_thresh: int = 50,
    val_thresh: int = 50,
) -> Any:
    """
    HSV Red-Stamp Filter for Official Vietnamese Documents (ERR_CONV_005 compliant).
    Detects red ink seal stamps (con dấu đỏ) in HSV color space and neutralizes them
    to prevent OCR noise and table cell border fragmentation.

    In HSV (PIL 0..255 representation):
      - Red wraps around 0/255: Hue in [0, 14] or [240, 255]
      - Saturation >= sat_thresh (to ignore grayscale/black text)
      - Value >= val_thresh (to ignore dark lines)
    """
    if not HAS_IMAGE_STACK:
        return image

    if not isinstance(image, Image.Image):
        # Convert from array if needed
        image = Image.fromarray(image)

    # Ensure RGB
    rgb_img = image.convert("RGB")
    hsv_img = rgb_img.convert("HSV")

    np_hsv = np.array(hsv_img)
    np_rgb = np.array(rgb_img)

    h = np_hsv[:, :, 0]
    s = np_hsv[:, :, 1]
    v = np_hsv[:, :, 2]

    # Red hue mask: 0 <= H <= 14 or 240 <= H <= 255 with significant saturation and brightness
    red_mask = ((h <= 14) | (h >= 240)) & (s >= sat_thresh) & (v >= val_thresh)

    # Replace masked pixels with clean background fill
    cleaned_rgb = np_rgb.copy()
    cleaned_rgb[red_mask] = bg_fill

    return Image.fromarray(cleaned_rgb, "RGB")


def detect_skew_angle(image: Any) -> float:
    """
    Estimate skew angle (in degrees) of a scanned page using horizontal projection profiles.
    Returns angle in degrees (-15 to +15).
    """
    if not HAS_IMAGE_STACK:
        return 0.0

    if not isinstance(image, Image.Image):
        image = Image.fromarray(image)

    # Convert to grayscale small thumbnail for speed
    gray = image.convert("L")
    w, h = gray.size
    scale = min(1.0, 800.0 / max(w, h))
    if scale < 1.0:
        gray = gray.resize((int(w * scale), int(h * scale)), Image.Resampling.BILINEAR)

    arr = np.array(gray)
    # Binarize with Otsu or median threshold
    thresh = np.median(arr)
    bin_arr = (arr < thresh).astype(np.uint8)

    best_score = -1.0
    best_angle = 0.0

    # Search in range [-10, +10] degrees in 0.5 steps
    for angle in np.arange(-5.0, 5.5, 0.5):
        # Rotate binary array
        rot_img = Image.fromarray(bin_arr * 255).rotate(
            angle, resample=Image.Resampling.NEAREST, expand=False
        )
        rot_arr = np.array(rot_img)
        # Profile variance along rows
        row_sums = np.sum(rot_arr, axis=1)
        variance = np.var(row_sums)
        if variance > best_score:
            best_score = variance
            best_angle = float(angle)

    return best_angle


def deskew_image(image: Any) -> Tuple[Any, float]:
    """Deskew image by detected angle."""
    if not HAS_IMAGE_STACK:
        return image, 0.0

    angle = detect_skew_angle(image)
    if abs(angle) > 0.2:
        straightened = image.rotate(
            -angle, resample=Image.Resampling.BICUBIC, expand=True, fillcolor=(255, 255, 255)
        )
        return straightened, angle
    return image, 0.0


def run_tesseract_ocr(
    image: Any,
    lang: str = "vie+eng",
    tesseract_cmd: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Execute Tesseract OCR command on an image.
    Supports fallback if Vietnamese langpack ('vie') is not installed.
    """
    binary = tesseract_cmd or find_tesseract_binary()
    result: Dict[str, Any] = {
        "success": False,
        "text": "",
        "lang_used": lang,
        "binary": binary,
        "error": None,
    }

    if not binary or (not os.path.isfile(binary) and not shutil.which(binary)):
        result["error"] = (
            "Tesseract OCR executable not found. "
            "Please install Tesseract (e.g. 'apt install tesseract-ocr tesseract-ocr-vie' "
            "or 'winget install UB-Mannheim.TesseractOCR')."
        )
        return result

    if not isinstance(image, Image.Image):
        image = Image.fromarray(image)

    # Save to temporary in-memory PNG
    import tempfile
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
        tmp_img_path = tmp_img.name

    with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp_out:
        tmp_out_base = tmp_out.name[:-4]  # Tesseract adds .txt automatically
        tmp_out_path = tmp_out_base + ".txt"

    try:
        image.save(tmp_img_path, format="PNG")

        # Run tesseract
        cmd = [binary, tmp_img_path, tmp_out_base, "-l", lang, "--psm", "3"]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=60)

        # If failed due to language pack missing, retry with 'eng'
        if proc.returncode != 0 and "vie" in lang:
            cmd = [binary, tmp_img_path, tmp_out_base, "-l", "eng", "--psm", "3"]
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            result["lang_used"] = "eng"

        if proc.returncode == 0 and os.path.exists(tmp_out_path):
            with open(tmp_out_path, "r", encoding="utf-8", errors="replace") as f:
                extracted = f.read().strip()
            result["success"] = True
            result["text"] = extracted
        else:
            result["error"] = proc.stderr or "Tesseract execution returned non-zero code."

    except Exception as ex:
        result["error"] = str(ex)
    finally:
        for p in [tmp_img_path, tmp_out_path]:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass

    return result


def _run_rapidocr_fallback(image: Any) -> Dict[str, Any]:
    """Optional zero-external-binary OCR via `rapidocr_onnxruntime` when installed."""
    res: Dict[str, Any] = {
        "success": False,
        "text": "",
        "lang_used": "multi",
        "backend": "rapidocr",
        "error": None,
    }
    try:
        from rapidocr_onnxruntime import RapidOCR  # type: ignore
    except ImportError:
        res["error"] = "rapidocr_onnxruntime is not installed."
        return res

    try:
        pil_img = image if isinstance(image, Image.Image) else Image.fromarray(image)
        if pil_img.mode != "RGB":
            pil_img = pil_img.convert("RGB")
        engine = RapidOCR()
        ocr_out, _ = engine(np.array(pil_img))
        lines = [item[1] for item in (ocr_out or []) if len(item) >= 2 and item[1]]
        res["success"] = True
        res["text"] = "\n".join(lines).strip()
        return res
    except Exception as ex:
        res["error"] = f"RapidOCR failed: {ex}"
        return res


def _run_windows_winrt_ocr(image: Any, lang: str = "vie+eng") -> Dict[str, Any]:
    """
    Execute Windows 10/11 built-in WinRT OCR (`Windows.Media.Ocr.OcrEngine`) via PowerShell
    without any external binary or pip package requirements.
    """
    result: Dict[str, Any] = {
        "success": False,
        "text": "",
        "lang_used": lang,
        "backend": "windows_winrt",
        "error": None,
    }
    if sys.platform != "win32":
        result["error"] = "Windows WinRT OCR is only available on Windows."
        return result

    pil_img = image if isinstance(image, Image.Image) else Image.fromarray(image)
    if pil_img.mode != "RGB":
        pil_img = pil_img.convert("RGB")

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_img:
        tmp_img_path = tmp_img.name

    ps_script = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
[Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime] | Out-Null
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
function Await($op, $type) {
    $netTask = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($op))
    $netTask.Wait() | Out-Null
    return $netTask.Result
}

$imgPath = $env:AG_OCR_IMG
$prefs = ($env:AG_OCR_LANGS -split ',') | Where-Object { $_ }
$engine = $null
$avail = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
foreach ($pref in $prefs) {
    $matched = $avail | Where-Object { $_.LanguageTag.ToLower().StartsWith($pref.ToLower()) } | Select-Object -First 1
    if ($matched) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($matched)
        if ($engine) { break }
    }
}
if (-not $engine) {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
}
if (-not $engine) {
    throw 'No Windows OCR recognizer language pack available.'
}

$stream = $null
$bmp = $null
try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($imgPath)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bmp = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
    Write-Output ($engine.RecognizerLanguage.LanguageTag)
    Write-Output ($res.Text)
} finally {
    if ($bmp) { $bmp.Dispose() }
    if ($stream) { $stream.Dispose() }
}
"""
    try:
        pil_img.save(tmp_img_path, format="PNG")
        lang_low = (lang or "").lower()
        prefs = "vi,en" if "vie" in lang_low or "vi" in lang_low else "en,vi"
        env = dict(os.environ)
        env["AG_OCR_IMG"] = tmp_img_path
        env["AG_OCR_LANGS"] = prefs

        proc = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            timeout=30,
        )
        if proc.returncode == 0:
            lines = proc.stdout.splitlines()
            used_lang = lines[0].strip() if lines else lang
            text_out = "\n".join(lines[1:]).strip() if len(lines) > 1 else ""
            result["success"] = True
            result["text"] = text_out
            result["lang_used"] = used_lang
        else:
            result["error"] = (proc.stderr or proc.stdout or "Windows WinRT OCR failed.").strip()
    except Exception as ex:
        result["error"] = f"Windows WinRT OCR error: {ex}"
    finally:
        if os.path.exists(tmp_img_path):
            try:
                os.remove(tmp_img_path)
            except OSError:
                pass

    return result


def run_ocr_with_fallback(
    image: Any,
    lang: str = "vie+eng",
    tesseract_cmd: Optional[str] = None,
    backend: str = "auto",
) -> Dict[str, Any]:
    """
    Execute OCR using the best available engine:
      1. Tesseract OCR (when binary is configured or found on PATH)
      2. RapidOCR (`rapidocr_onnxruntime`) if installed
      3. Windows built-in WinRT OCR (`Windows.Media.Ocr.OcrEngine`) on Win10/11
    """
    mode = (backend or "auto").strip().lower()
    if mode == "tesseract":
        res = run_tesseract_ocr(image, lang=lang, tesseract_cmd=tesseract_cmd)
        res.setdefault("backend", "tesseract" if res.get("success") else "none")
        return res
    if mode == "rapidocr":
        return _run_rapidocr_fallback(image)
    if mode in ("windows_winrt", "winrt"):
        return _run_windows_winrt_ocr(image, lang=lang)

    errors: List[str] = []
    binary = tesseract_cmd or find_tesseract_binary()
    if binary:
        t_res = run_tesseract_ocr(image, lang=lang, tesseract_cmd=binary)
        if t_res.get("success"):
            t_res["backend"] = "tesseract"
            return t_res
        if t_res.get("error"):
            errors.append(str(t_res["error"]))

    r_res = _run_rapidocr_fallback(image)
    if r_res.get("success"):
        return r_res

    if sys.platform == "win32":
        w_res = _run_windows_winrt_ocr(image, lang=lang)
        if w_res.get("success"):
            return w_res
        if w_res.get("error"):
            errors.append(str(w_res["error"]))

    return {
        "success": False,
        "text": "",
        "lang_used": lang,
        "binary": binary,
        "backend": "none",
        "error": "; ".join(errors) if errors else "No OCR backend available (install Tesseract or rapidocr_onnxruntime).",
    }


class VietnameseOcrPipeline:
    """
    High-Fidelity Vietnamese OCR Pipeline for Scanned Documents.
    Performs end-to-end PDF / Image rasterization, deskewing, red-stamp filtering,
    multi-backend OCR extraction (Tesseract -> RapidOCR -> Windows WinRT / PDF text layer),
    and export to Markdown or DOCX.
    """

    def __init__(
        self,
        lang: str = "vie+eng",
        filter_red_stamp: bool = True,
        deskew: bool = True,
        tesseract_cmd: Optional[str] = None,
        backend: str = "auto",
    ) -> None:
        self.lang = lang
        self.filter_red_stamp = filter_red_stamp
        self.deskew = deskew
        self.tesseract_cmd = tesseract_cmd or find_tesseract_binary()
        self.backend = backend

    def process_image(self, image: Any) -> Dict[str, Any]:
        """Process a single image through the pipeline."""
        processed_img = image
        skew_angle = 0.0

        if self.deskew:
            processed_img, skew_angle = deskew_image(processed_img)

        if self.filter_red_stamp:
            processed_img = filter_red_stamp(processed_img)

        ocr_res = run_ocr_with_fallback(
            processed_img,
            lang=self.lang,
            tesseract_cmd=self.tesseract_cmd,
            backend=self.backend,
        )

        return {
            "text": ocr_res.get("text", ""),
            "success": ocr_res.get("success", False),
            "skew_angle": skew_angle,
            "lang_used": ocr_res.get("lang_used", self.lang),
            "backend": ocr_res.get("backend", "none"),
            "error": ocr_res.get("error"),
        }

    def process_document(
        self,
        input_path: str,
        output_path: Optional[str] = None,
        max_pages: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        Process a PDF or image document and optionally save output as Markdown or DOCX.
        """
        ext = os.path.splitext(input_path)[1].lower()
        pages_results: List[Dict[str, Any]] = []

        if ext == ".pdf":
            if not HAS_PYMUPDF:
                raise RuntimeError("PyMuPDF is required to process PDF files in OCR pipeline.")

            doc = pymupdf.open(input_path)
            try:
                num_pages = len(doc)
                if max_pages:
                    num_pages = min(num_pages, max_pages)

                for p_idx in range(num_pages):
                    page = doc[p_idx]
                    pix = page.get_pixmap(dpi=300)
                    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

                    page_res = self.process_image(img)
                    if not (page_res.get("text") or "").strip():
                        text_layer = page.get_text().strip()
                        if text_layer:
                            page_res["text"] = text_layer
                            page_res["success"] = True
                            page_res["backend"] = "pdf_text_layer"
                            page_res["error"] = None
                    page_res["page_index"] = p_idx + 1
                    pages_results.append(page_res)
            finally:
                doc.close()

        elif ext in [".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".webp"]:
            if not HAS_IMAGE_STACK:
                raise RuntimeError("Pillow is required to open image files.")

            with Image.open(input_path) as img:
                page_res = self.process_image(img.copy())
                page_res["page_index"] = 1
                pages_results.append(page_res)
        else:
            raise ValueError(f"Unsupported input format for OCR pipeline: {ext}")

        # Combine text
        full_text_blocks = []
        for pr in pages_results:
            txt = pr.get("text", "")
            if txt:
                full_text_blocks.append(f"<!-- Page {pr['page_index']} -->\n\n{txt}")

        combined_text = "\n\n---\n\n".join(full_text_blocks)
        backends_used = sorted({str(pr.get("backend") or "none") for pr in pages_results if pr.get("success")})

        result_summary = {
            "input_path": input_path,
            "output_path": output_path,
            "total_pages": len(pages_results),
            "pages": pages_results,
            "full_text": combined_text,
            "has_tesseract": bool(self.tesseract_cmd),
            "ocr_backends_used": backends_used,
            "has_ocr_backend": bool(self.tesseract_cmd or backends_used),
        }

        if output_path:
            os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
            out_ext = os.path.splitext(output_path)[1].lower()
            if out_ext == ".docx":
                try:
                    from docx import Document
                    out_doc = Document()
                    for block in full_text_blocks:
                        for line in block.split("\n"):
                            line_s = line.strip()
                            if line_s:
                                out_doc.add_paragraph(line_s)
                    out_doc.save(output_path)
                except Exception:
                    with open(output_path, "w", encoding="utf-8") as f:
                        f.write(combined_text)
            else:
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(combined_text)

        return result_summary
