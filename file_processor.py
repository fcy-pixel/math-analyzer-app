"""
File processing utilities for AQP reports and exam papers.
Supports Excel, CSV, PDF, and image formats.
"""

import base64
import io
import os
from pathlib import Path
from typing import Dict, List

import pandas as pd
from PIL import Image

try:
    import fitz  # PyMuPDF

    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import pdfplumber

    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


# How many pages to bundle per vision-API call (Qwen VL handles up to ~10 images)
VISION_BATCH_SIZE = 6

# Qwen API limit: 10 MB per data-uri item.  Keep a generous safety margin.
_MAX_BASE64_BYTES = 7 * 1024 * 1024  # 7 MB (API max = 10 MB)


def _compress_image_b64(
    png_bytes: bytes,
    max_b64_bytes: int = _MAX_BASE64_BYTES,
) -> str:
    """Return a base64 string guaranteed to stay under the API data-uri limit.

    Strategy: try the original PNG first; if too large, re-encode as JPEG
    with progressively lower quality and smaller dimensions.
    """
    b64 = base64.b64encode(png_bytes).decode()
    if len(b64) <= max_b64_bytes:
        return b64

    # Open and get actual dimensions as starting point
    img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    w, h = img.size
    max_dim = max(w, h)

    for quality in (80, 65, 50, 35, 20):
        thumb = img.copy()
        thumb.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        thumb.save(buf, format="JPEG", quality=quality, optimize=True)
        b64 = base64.b64encode(buf.getvalue()).decode()
        if len(b64) <= max_b64_bytes:
            return b64
        # Shrink dimensions by 25% each iteration
        max_dim = int(max_dim * 0.75)
    # Last resort: return whatever we have (smallest attempt)
    return b64


class FileProcessor:
    """Process uploaded files into structured data for AI analysis."""

    IMAGE_DPI_SCALE = 1.5   # Render scale (~108 DPI) — good quality, manageable size
    MAX_IMAGE_SIZE = (1600, 1600)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def process_aqp(self, file_path: str) -> Dict:
        """Process an AQP report file and return structured data."""
        ext = Path(file_path).suffix.lower()
        if ext in (".xlsx", ".xls"):
            return self._process_excel(file_path)
        elif ext == ".csv":
            return self._process_csv(file_path)
        elif ext == ".pdf":
            return self._process_pdf_report(file_path)
        else:
            raise ValueError(f"不支援的檔案格式：{ext}")

    def process_exam(self, file_path: str) -> Dict:
        """Process an exam paper file and return image data for vision API."""
        ext = Path(file_path).suffix.lower()
        if ext == ".pdf":
            return self._process_pdf_exam(file_path)
        elif ext in (".jpg", ".jpeg", ".png"):
            return self._process_image(file_path)
        else:
            raise ValueError(f"不支援的檔案格式：{ext}")

    # ------------------------------------------------------------------
    # AQP processors
    # ------------------------------------------------------------------

    def _process_excel(self, file_path: str) -> Dict:
        xl = pd.ExcelFile(file_path)
        sheets: Dict[str, pd.DataFrame] = {}
        text_parts: List[str] = []

        for sheet_name in xl.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            # Drop entirely empty rows/columns
            df = df.dropna(how="all").dropna(axis=1, how="all")
            sheets[sheet_name] = df
            text_parts.append(f"【工作表：{sheet_name}】\n{df.to_string(index=False)}")

        return {
            "type": "excel",
            "sheets": {k: v.to_dict("records") for k, v in sheets.items()},
            "text_summary": "\n\n".join(text_parts),
            "images": [],
        }

    def _process_csv(self, file_path: str) -> Dict:
        # Try common encodings for Chinese text
        for enc in ("utf-8-sig", "utf-8", "big5", "gb18030"):
            try:
                df = pd.read_csv(file_path, encoding=enc)
                break
            except (UnicodeDecodeError, ValueError):
                continue
        else:
            df = pd.read_csv(file_path, encoding="utf-8", errors="replace")

        df = df.dropna(how="all").dropna(axis=1, how="all")
        return {
            "type": "csv",
            "sheets": {"data": df.to_dict("records")},
            "text_summary": df.to_string(index=False),
            "images": [],
        }

    def _process_pdf_report(self, file_path: str) -> Dict:
        """Extract ALL pages of an AQP report as both text and images."""
        result: Dict = {"type": "pdf", "text_summary": "", "images": [], "page_count": 0}

        if HAS_PYMUPDF:
            doc = fitz.open(file_path)
            total = len(doc)
            texts, images = [], []
            for page_num, page in enumerate(doc):
                texts.append(f"【第 {page_num + 1} / {total} 頁】\n{page.get_text()}")
                mat = fitz.Matrix(self.IMAGE_DPI_SCALE, self.IMAGE_DPI_SCALE)
                pix = page.get_pixmap(matrix=mat)
                images.append(_compress_image_b64(pix.tobytes("png")))
            doc.close()
            result["text_summary"] = "\n\n".join(texts)
            result["images"] = images
            result["page_count"] = total

        elif HAS_PDFPLUMBER:
            with pdfplumber.open(file_path) as pdf:
                total = len(pdf.pages)
                texts = []
                for i, page in enumerate(pdf.pages):
                    text = page.extract_text() or ""
                    page_text = f"【第 {i + 1} / {total} 頁】\n{text}"
                    for j, table in enumerate(page.extract_tables() or []):
                        df = pd.DataFrame(table)
                        page_text += f"\n表格{j + 1}：\n{df.to_string(index=False)}"
                    texts.append(page_text)
            result["text_summary"] = "\n\n".join(texts)
            result["page_count"] = total

        if not result["text_summary"] and not result["images"]:
            raise RuntimeError(
                "無法讀取 PDF，請安裝 PyMuPDF（pip install pymupdf）或 pdfplumber。"
            )

        return result

    # ------------------------------------------------------------------
    # Exam paper processors
    # ------------------------------------------------------------------

    def _process_pdf_exam(self, file_path: str) -> Dict:
        """Convert ALL pages of an exam paper PDF to base64 images for vision API."""
        result: Dict = {"type": "pdf_exam", "images": [], "text": "", "page_count": 0}

        if not HAS_PYMUPDF and not HAS_PDFPLUMBER:
            raise RuntimeError(
                "處理試卷 PDF 需要 PyMuPDF（pip install pymupdf）。"
            )

        if HAS_PYMUPDF:
            doc = fitz.open(file_path)
            total = len(doc)
            texts, images = [], []
            for page_num, page in enumerate(doc):
                mat = fitz.Matrix(self.IMAGE_DPI_SCALE, self.IMAGE_DPI_SCALE)
                pix = page.get_pixmap(matrix=mat)
                images.append(_compress_image_b64(pix.tobytes("png")))
                texts.append(f"【第 {page_num + 1} 頁】\n{page.get_text()}")
            doc.close()
            result["images"] = images
            result["text"] = "\n\n".join(texts)
            result["page_count"] = total

        elif HAS_PDFPLUMBER:
            # Fallback: extract text only (no images)
            with pdfplumber.open(file_path) as pdf:
                total = len(pdf.pages)
                texts = []
                for i, page in enumerate(pdf.pages):
                    texts.append(f"【第 {i + 1} 頁】\n{page.extract_text() or ''}")
            result["text"] = "\n\n".join(texts)
            result["page_count"] = total

        return result

    def _process_image(self, file_path: str) -> Dict:
        """Load a single image and encode to base64."""
        with open(file_path, "rb") as f:
            raw = f.read()

        img = Image.open(io.BytesIO(raw)).convert("RGB")
        img.thumbnail(self.MAX_IMAGE_SIZE, Image.Resampling.LANCZOS)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = _compress_image_b64(buf.getvalue())

        return {"type": "image", "images": [b64], "text": "", "page_count": 1}


def get_image_batches(images: List[str], batch_size: int = VISION_BATCH_SIZE) -> List[List[str]]:
    """Split a list of page images into batches for the vision API."""
    return [images[i: i + batch_size] for i in range(0, len(images), batch_size)]


def split_student_papers(pdf_bytes: bytes, pages_per_student: int) -> List[Dict]:
    """
    Split a single PDF (containing all students' scanned papers) into per-student chunks.

    Returns a list of dicts, each with:
        student_index  — 1-based
        page_range     — (first_page, last_page) 1-based inclusive
        images         — List[str]  base64-encoded PNG images (one per page)
    """
    if not HAS_PYMUPDF:
        raise RuntimeError("分割學生試卷需要 PyMuPDF（pip install pymupdf）。")

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    scale = FileProcessor.IMAGE_DPI_SCALE  # 1.8x → ~130 DPI

    students: List[Dict] = []
    for i, start in enumerate(range(0, total_pages, pages_per_student)):
        end = min(start + pages_per_student, total_pages)
        images: List[str] = []
        for page_num in range(start, end):
            page = doc[page_num]
            mat = fitz.Matrix(scale, scale)
            pix = page.get_pixmap(matrix=mat)
            images.append(_compress_image_b64(pix.tobytes("png")))
        students.append({
            "student_index": i + 1,
            "page_range": (start + 1, end),
            "images": images,
        })

    doc.close()
    return students


def get_student_count(pdf_bytes: bytes, pages_per_student: int) -> int:
    """Return the estimated number of students from the PDF size."""
    if not HAS_PYMUPDF:
        return 0
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = len(doc)
    doc.close()
    import math
    return math.ceil(total / pages_per_student)


def get_pdf_page_count(pdf_bytes: bytes) -> int:
    """Return total page count of a PDF bytes object."""
    if HAS_PYMUPDF:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        n = len(doc)
        doc.close()
        return n
    return 0


def check_pdf_support() -> str:
    """Return a string describing available PDF libraries."""
    libs = []
    if HAS_PYMUPDF:
        libs.append("PyMuPDF ✅")
    if HAS_PDFPLUMBER:
        libs.append("pdfplumber ✅")
    if not libs:
        return "⚠️ 未安裝 PDF 處理庫，請執行：pip install pymupdf"
    return "PDF 支援：" + "、".join(libs)
