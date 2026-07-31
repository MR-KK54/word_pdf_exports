"""
Word Page Exporter Pro - PDF Engine
Provides page extraction and inspection for PDF documents via PyMuPDF.
"""

import os
from typing import Dict, Any

import fitz  # PyMuPDF

from word_exporter_pro.utils.logger import get_logger

logger = get_logger()


class PdfInspector:
    """Inspects PDF document structure and metadata."""

    @staticmethod
    def get_info(file_path: str) -> Dict[str, Any]:
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Document file not found: {abs_path}")

        info = {
            "path": abs_path,
            "filename": os.path.basename(abs_path),
            "size_bytes": os.path.getsize(abs_path),
            "page_count": 0,
            "section_count": 1,
            "title": "",
            "author": "",
            "format": "pdf",
        }

        doc = fitz.open(abs_path)
        try:
            info["page_count"] = doc.page_count
            info["title"] = str(doc.metadata.get("title") or "")
            info["author"] = str(doc.metadata.get("author") or "")
        except Exception as e:
            logger.error(f"Error inspecting PDF '{abs_path}': {e}")
            raise RuntimeError(f"Could not open or inspect PDF document: {e}")
        finally:
            doc.close()

        return info


class PdfPageExtractor:
    """Core engine for extracting page ranges from PDF documents."""

    @staticmethod
    def extract_range(
        source_file: str,
        output_file: str,
        start_page: int,
        end_page: int
    ) -> str:
        """
        Extracts a page range from a source PDF into a new PDF file.

        Args:
            source_file: Source PDF file path.
            output_file: Target path to save the extracted PDF.
            start_page: 1-indexed start page.
            end_page: 1-indexed end page.

        Returns:
            Absolute path of created output file.
        """
        abs_source = os.path.abspath(source_file)
        abs_output = os.path.abspath(output_file)
        os.makedirs(os.path.dirname(abs_output), exist_ok=True)

        logger.info(
            f"Starting PDF page export [{start_page}-{end_page}] from "
            f"'{os.path.basename(abs_source)}' -> '{os.path.basename(abs_output)}'"
        )

        src = fitz.open(abs_source)
        try:
            total_pages = src.page_count
            if start_page < 1 or end_page > total_pages or start_page > end_page:
                raise ValueError(
                    f"Invalid range [{start_page}-{end_page}] for document with {total_pages} pages."
                )

            out = fitz.open()
            try:
                out.insert_pdf(src, from_page=start_page - 1, to_page=end_page - 1)
                out.save(abs_output, garbage=3, deflate=True)
            finally:
                out.close()
            logger.success(f"Exported pages [{start_page}-{end_page}] to '{abs_output}'")
        finally:
            src.close()

        return abs_output
