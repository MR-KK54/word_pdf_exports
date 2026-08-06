"""
Word Page Exporter Pro - Batch Processor
Executes multi-document and multi-range export jobs in background worker threads with progress tracking.
"""

import os
import threading
from dataclasses import dataclass
from typing import List, Tuple, Callable, Optional, Dict, Any

from word_exporter_pro.core.range_parser import PageRangeParser, RangeParseError
from word_exporter_pro.core.naming_formatter import NamingFormatter
from word_exporter_pro.core.com_engine import DocumentInspector, PageExporterEngine
from word_exporter_pro.core.pdf_engine import PdfInspector, PdfPageExtractor
from word_exporter_pro.utils.logger import get_logger

logger = get_logger()


@dataclass
class ExportJobConfig:
    source_files: List[str]
    range_expression: str
    output_dir: str
    export_format: str = "docx"
    naming_pattern: str = "{original_name}_pages_{start_page}-{end_page}"
    overwrite: bool = False
    engine_mode: str = "trimming"
    visible: bool = False
    clear_storage_after_export: bool = False


class BatchProcessor:
    """Manages background batch export operations with progress reporting and cancellation."""

    def __init__(self, config: ExportJobConfig):
        self.config = config
        self.cancel_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None

    def start_async(
        self,
        on_progress: Optional[Callable[[int, int, str, str], None]] = None,
        on_finished: Optional[Callable[[int, int, List[str]], None]] = None,
        on_file_created: Optional[Callable[[str], None]] = None
    ):
        """Starts processing jobs in a background worker thread."""
        self.cancel_event.clear()
        self.worker_thread = threading.Thread(
            target=self._run_job,
            args=(on_progress, on_finished, on_file_created),
            daemon=True
        )
        self.worker_thread.start()

    def cancel(self):
        """Requests cancellation of ongoing job."""
        self.cancel_event.set()
        logger.warning("Cancellation requested by user.")

    def _run_job(
        self,
        on_progress: Optional[Callable[[int, int, str, str], None]],
        on_finished: Optional[Callable[[int, int, List[str]], None]],
        on_file_created: Optional[Callable[[str], None]] = None
    ):
        logger.info(f"Starting batch export job across {len(self.config.source_files)} source document(s)...")

        total_export_tasks = 0
        job_tasks: List[Tuple[str, int, Tuple[int, int], bool]] = [] # (file_path, total_pages, (start, end), is_pdf)
        errors: List[str] = []
        success_count = 0
        fail_count = 0

        # Step 1: Pre-process document statistics & parse page ranges
        for file_idx, file_path in enumerate(self.config.source_files, 1):
            if self.cancel_event.is_set():
                break

            try:
                is_pdf = os.path.splitext(file_path)[1].lower() == ".pdf"
                if is_pdf:
                    info = PdfInspector.get_info(file_path)
                else:
                    info = DocumentInspector.get_info(file_path, visible=self.config.visible)
                total_pages = info["page_count"]
                
                # Parse range expression for this document without clamping to preserve user's range naming
                parsed_ranges = PageRangeParser.parse(self.config.range_expression, total_pages, clamp_to_total=False)
                
                for pr in parsed_ranges:
                    job_tasks.append((file_path, total_pages, pr, is_pdf))
            except Exception as e:
                err_msg = f"Failed pre-processing '{os.path.basename(file_path)}': {e}"
                logger.error(err_msg)
                errors.append(err_msg)

        total_export_tasks = len(job_tasks)

        if total_export_tasks == 0:
            logger.warning("No valid export tasks generated.")
            if on_finished:
                on_finished(0, len(errors), errors)
            return

        # Step 2: Execute page exports
        for task_idx, (file_path, total_pages, pr, is_pdf) in enumerate(job_tasks, 1):
            if self.cancel_event.is_set():
                logger.warning("Batch job execution halted due to user cancellation.")
                break

            base_name = os.path.basename(file_path)
            start_p, end_p = pr
            status_desc = f"Exporting {base_name} [pages {start_p}-{end_p}]"

            if on_progress:
                try:
                    on_progress(task_idx - 1, total_export_tasks, base_name, status_desc)
                except Exception as e:
                    logger.error(f"Error in progress callback: {e}")

            # PDF sources are always exported as PDF; "same"/"source" matches original file extension
            src_ext = os.path.splitext(file_path)[1].lower().lstrip(".")
            if is_pdf or self.config.export_format.lower() in ("same", "source"):
                fmt = src_ext if src_ext else "docx"
            else:
                fmt = self.config.export_format


            # Generate output file path (using the original range name)
            filename = NamingFormatter.generate_filename(
                pattern=self.config.naming_pattern,
                original_filepath=file_path,
                page_range=pr,
                total_pages=total_pages,
                output_ext=fmt,
                batch_index=task_idx
            )

            output_file_path = NamingFormatter.resolve_output_path(
                output_dir=self.config.output_dir,
                filename=filename,
                overwrite=self.config.overwrite
            )

            if is_pdf and start_p > total_pages:
                fail_count += 1
                err_msg = f"Failed exporting {base_name} [pages {start_p}-{end_p}]: Requested start page {start_p} exceeds total document page count ({total_pages})."
                logger.error(err_msg)
                errors.append(err_msg)
                continue

            # Clamp range limits safely for extraction engines
            clamped_start = max(1, start_p)
            if is_pdf:
                clamped_start = min(clamped_start, total_pages)
                clamped_end = max(clamped_start, min(end_p, total_pages))
            else:
                clamped_end = max(clamped_start, end_p)

            # Perform export
            try:
                if is_pdf:
                    PdfPageExtractor.extract_range(
                        source_file=file_path,
                        output_file=output_file_path,
                        start_page=clamped_start,
                        end_page=clamped_end
                    )
                else:
                    PageExporterEngine.export_range(
                        source_file=file_path,
                        output_file=output_file_path,
                        start_page=clamped_start,
                        end_page=clamped_end,
                        export_format=fmt,
                        mode=self.config.engine_mode,
                        visible=self.config.visible,
                        total_pages=total_pages
                    )
                success_count += 1
                if on_file_created:
                    try:
                        on_file_created(output_file_path)
                    except Exception as e:
                        logger.error(f"Error in file-created callback: {e}")
            except Exception as e:
                fail_count += 1
                err_msg = f"Failed exporting {base_name} [pages {start_p}-{end_p}]: {e}"
                logger.error(err_msg)
                errors.append(err_msg)

            if on_progress:
                try:
                    on_progress(task_idx, total_export_tasks, base_name, status_desc)
                except Exception as e:
                    pass

        # Auto-clear source files from server storage after export if requested
        if self.config.clear_storage_after_export:
            for sf in self.config.source_files:
                try:
                    if os.path.exists(sf) and os.path.isfile(sf):
                        os.remove(sf)
                        logger.info(f"Auto-cleared source document from server: {os.path.basename(sf)}")
                except Exception as e:
                    logger.warning(f"Could not auto-clear source file '{sf}': {e}")

        # Final reporting
        logger.info(f"Batch processing completed. Success: {success_count}, Failures: {fail_count}")
        if on_finished:
            on_finished(success_count, fail_count, errors)
