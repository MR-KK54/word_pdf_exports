"""
Word Page Exporter Pro - PDF Engine
Provides high-performance page extraction and inspection for PDF documents via PyMuPDF.
Preserves page sizes, rotations, fonts, images, links, annotations, and metadata.
"""

import os
from typing import Dict, Any, List, Tuple

import fitz  # PyMuPDF

from word_exporter_pro.utils.logger import get_logger

logger = get_logger()


class PdfInspector:
    """Inspects PDF document structure, metadata, and security status."""

    @staticmethod
    def get_info(file_path: str) -> Dict[str, Any]:
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Document file not found: {abs_path}")

        ext = os.path.splitext(abs_path)[1].lower()
        if ext != ".pdf":
            raise ValueError(f"Unsupported file format '{ext}'. Expected PDF file.")

        info = {
            "path": abs_path,
            "filename": os.path.basename(abs_path),
            "size_bytes": os.path.getsize(abs_path),
            "page_count": 0,
            "section_count": 1,
            "title": "",
            "author": "",
            "format": "pdf",
            "is_encrypted": False,
        }

        try:
            doc = fitz.open(abs_path)
        except (fitz.FileDataError, fitz.EmptyFileError) as e:
            logger.error(f"Corrupted PDF file '{abs_path}': {e}")
            raise RuntimeError(f"Corrupted or unreadable PDF document: '{os.path.basename(abs_path)}'")
        except Exception as e:
            logger.error(f"Error opening PDF '{abs_path}': {e}")
            raise RuntimeError(f"Could not open PDF document: {e}")

        try:
            if doc.is_encrypted and doc.needs_pass:
                info["is_encrypted"] = True
                raise RuntimeError(f"Document '{os.path.basename(abs_path)}' is password-protected and cannot be opened without a password.")

            info["page_count"] = doc.page_count
            if info["page_count"] == 0:
                raise RuntimeError(f"PDF document '{os.path.basename(abs_path)}' has 0 pages.")

            info["title"] = str(doc.metadata.get("title") or "")
            info["author"] = str(doc.metadata.get("author") or "")
        finally:
            doc.close()

        return info


class PdfPageExtractor:
    """Core engine for extracting page ranges from PDF documents while preserving full fidelity."""

    @staticmethod
    def extract_range(
        source_file: str,
        output_file: str,
        start_page: int,
        end_page: int
    ) -> str:
        """
        Extracts a page range from a source PDF into a new PDF file.
        Preserves page sizes, rotations, fonts, images, links, annotations, and metadata.
        """
        abs_source = os.path.abspath(source_file)
        abs_output = os.path.abspath(output_file)

        out_dir = os.path.dirname(abs_output)
        os.makedirs(out_dir, exist_ok=True)

        if not os.access(out_dir, os.W_OK):
            raise PermissionError(f"Insufficient write permissions for destination folder '{out_dir}'.")

        logger.info(
            f"Starting PDF page export [{start_page}-{end_page}] from "
            f"'{os.path.basename(abs_source)}' -> '{os.path.basename(abs_output)}'"
        )

        try:
            src = fitz.open(abs_source)
        except (fitz.FileDataError, fitz.EmptyFileError) as e:
            raise RuntimeError(f"Corrupted or unreadable PDF document: '{os.path.basename(abs_source)}'")
        except Exception as e:
            raise RuntimeError(f"Could not open PDF source file: {e}")

        try:
            if src.is_encrypted and src.needs_pass:
                raise RuntimeError(f"Document '{os.path.basename(abs_source)}' is password-protected.")

            total_pages = src.page_count
            if start_page < 1 or end_page > total_pages or start_page > end_page:
                raise ValueError(
                    f"Invalid range [{start_page}-{end_page}] for PDF document with {total_pages} pages."
                )

            out = fitz.open()
            try:
                # Copy selected page objects preserving size, rotation, links, annots, and graphics
                out.insert_pdf(
                    src,
                    from_page=start_page - 1,
                    to_page=end_page - 1,
                    rotate=-1,
                    links=True,
                    annots=True
                )
                
                # Preserve document metadata
                if src.metadata:
                    try:
                        out.set_metadata(src.metadata)
                    except Exception:
                        pass

                out.save(abs_output, garbage=3, deflate=True)
            except PermissionError:
                raise PermissionError(f"Insufficient write permissions for output file '{abs_output}'.")
            except Exception as e:
                raise RuntimeError(f"Failed to write output PDF file '{abs_output}': {e}")
            finally:
                out.close()

            logger.success(f"Exported PDF pages [{start_page}-{end_page}] to '{abs_output}'")
        finally:
            src.close()

        return abs_output

    @staticmethod
    def extract_ranges_batch(
        source_file: str,
        export_tasks: List[Tuple[str, int, int]]  # [(output_file_path, start_page, end_page), ...]
    ) -> List[str]:
        """
        High-Performance Single-Pass PDF Extractor:
        Opens source PDF document ONCE, computes page map ONCE, resolves all pages,
        and writes all output PDF files in a single pass.
        """
        abs_source = os.path.abspath(source_file)
        created_files = []

        try:
            src = fitz.open(abs_source)
        except Exception as e:
            raise RuntimeError(f"Could not open PDF source file: {e}")

        try:
            if src.is_encrypted and src.needs_pass:
                raise RuntimeError(f"Document '{os.path.basename(abs_source)}' is password-protected.")

            total_pages = src.page_count

            for output_file, start_p, end_p in export_tasks:
                abs_output = os.path.abspath(output_file)
                out_dir = os.path.dirname(abs_output)
                os.makedirs(out_dir, exist_ok=True)

                if not os.access(out_dir, os.W_OK):
                    raise PermissionError(f"Insufficient write permissions for destination directory '{out_dir}'.")

                start_clamped = max(1, min(start_p, total_pages))
                end_clamped = max(start_clamped, min(end_p, total_pages))

                out = fitz.open()
                try:
                    out.insert_pdf(
                        src,
                        from_page=start_clamped - 1,
                        to_page=end_clamped - 1,
                        rotate=-1,
                        links=True,
                        annots=True
                    )
                    if src.metadata:
                        try:
                            out.set_metadata(src.metadata)
                        except Exception:
                            pass
                    out.save(abs_output, garbage=3, deflate=True)
                    created_files.append(abs_output)
                finally:
                    out.close()

        finally:
            src.close()

        return created_files
