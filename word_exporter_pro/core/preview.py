"""
Word Page Exporter Pro - Document Preview
Renders pages of Word documents and PDFs to PNG images for in-browser preview.
"""

import os
import threading
from typing import Tuple

import fitz  # PyMuPDF

from word_exporter_pro.core.com_engine import WordCOMContext
from word_exporter_pro.utils.logger import get_logger

logger = get_logger()

WD_EXPORT_PDF = 17  # wdExportFormatPDF

# Serializes Word COM preview generation so concurrent preview requests
# do not spawn an excessive number of Word instances.
_preview_lock = threading.Lock()


def render_page_preview(
    source_path: str, page: int = 1, max_width: int = 1000
) -> Tuple[bytes, int]:
    """
    Renders a single page of a Word/PDF document to PNG bytes.

    Args:
        source_path: Path to the source document (.pdf or Word file).
        page: 1-indexed page number to render.
        max_width: Approximate target width in pixels.

    Returns:
        (png_bytes, total_page_count)
    """
    src = _ensure_pdf(source_path)
    doc = fitz.open(src)
    try:
        total = doc.page_count
        idx = page - 1
        if idx < 0 or idx >= total:
            raise IndexError(
                f"Page {page} is out of range (document has {total} page(s))."
            )
        page_obj = doc.load_page(idx)
        zoom = max_width / page_obj.rect.width
        pix = page_obj.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return pix.tobytes("png"), total
    finally:
        doc.close()


def _preview_path(source_path: str) -> str:
    base = os.path.splitext(os.path.basename(source_path))[0]
    return os.path.join(os.path.dirname(source_path), f"{base}_preview.pdf")


def has_preview_pdf(source_path: str) -> bool:
    """True when the source is a PDF or its Word preview PDF is already generated."""
    if source_path.lower().endswith(".pdf"):
        return True
    preview_pdf = _preview_path(source_path)
    return os.path.exists(preview_pdf) and os.path.getmtime(preview_pdf) >= os.path.getmtime(source_path)


_preview_jobs: dict = {}
_preview_jobs_lock = threading.Lock()


def ensure_preview_async(source_path: str) -> bool:
    """
    Ensures a Word preview PDF is (or will be) generated in the background.

    Returns:
        True if the preview is already available, False if generation has started
        (client should poll the preview endpoint again).
    """
    if has_preview_pdf(source_path):
        return True

    key = os.path.abspath(source_path)
    with _preview_jobs_lock:
        if key in _preview_jobs:
            return False
        _preview_jobs[key] = True

    def worker():
        try:
            _ensure_pdf(source_path)
        except Exception as e:
            logger.error(f"Background preview generation failed for '{source_path}': {e}")
        finally:
            with _preview_jobs_lock:
                _preview_jobs.pop(key, None)

    threading.Thread(target=worker, daemon=True).start()
    return False


def _ensure_pdf(source_path: str) -> str:
    """Returns a PDF path for the source, converting Word files via a cached preview PDF."""
    if source_path.lower().endswith(".pdf"):
        return source_path

    preview_pdf = _preview_path(source_path)
    src_mtime = os.path.getmtime(source_path)
    with _preview_lock:
        if os.path.exists(preview_pdf) and os.path.getmtime(preview_pdf) >= src_mtime:
            return preview_pdf

        logger.info(f"Generating preview PDF for '{os.path.basename(source_path)}'...")
        with WordCOMContext(visible=False) as word_app:
            doc = word_app.Documents.Open(
                FileName=os.path.abspath(source_path),
                ReadOnly=True,
                ConfirmConversions=False,
                AddToRecentFiles=False,
            )
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=os.path.abspath(preview_pdf),
                    ExportFormat=WD_EXPORT_PDF,
                )
            finally:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass

    return preview_pdf
