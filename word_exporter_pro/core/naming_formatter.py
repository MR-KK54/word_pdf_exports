"""
Word Page Exporter Pro - Naming Formatter
Formats output filenames using dynamic template strings and handles filesystem path safety.
"""

import os
import re
from datetime import datetime
from typing import Tuple


class NamingFormatter:
    """Evaluates dynamic filename pattern strings against document & range metadata."""

    DEFAULT_PATTERN = "{original_name}_pages_{start_page}-{end_page}"
    SINGLE_PAGE_PATTERN = "{original_name}_page_{start_page}"

    INVALID_CHAR_REGEX = re.compile(r'[\\/*?:"<>|]')

    @classmethod
    def sanitize_filename(cls, name: str) -> str:
        """Removes illegal Windows filename characters."""
        cleaned = cls.INVALID_CHAR_REGEX.sub("_", name)
        # Remove trailing periods or spaces which cause Windows issues
        return cleaned.strip(" .")

    @classmethod
    def generate_filename(
        cls,
        pattern: str,
        original_filepath: str,
        page_range: Tuple[int, int],
        total_pages: int,
        output_ext: str,
        batch_index: int = 1
    ) -> str:
        """
        Generates a sanitized output filename.

        Args:
            pattern: E.g., "{original_name}_pages_{start_page}-{end_page}"
            original_filepath: Full path or basename of source Word doc.
            page_range: (start_page, end_page)
            total_pages: Total pages in original document.
            output_ext: Output format extension without leading dot (e.g. 'docx', 'pdf').
            batch_index: Sequential job index.

        Returns:
            Sanitized filename string with extension.
        """
        base_filename = os.path.basename(original_filepath)
        original_name, source_ext = os.path.splitext(base_filename)
        source_ext = source_ext.lstrip(".")

        start_page, end_page = page_range
        page_count = end_page - start_page + 1

        if start_page == end_page:
            range_str = f"page_{start_page}"
        else:
            range_str = f"pages_{start_page}-{end_page}"

        now = datetime.now()
        timestamp_str = now.strftime("%Y%m%d_%H%M%S")
        date_str = now.strftime("%Y%m%d")
        time_str = now.strftime("%H%M%S")

        # Substitution mapping
        replacements = {
            "{original_name}": original_name,
            "{ext}": source_ext,
            "{range}": range_str,
            "{start_page}": str(start_page),
            "{end_page}": str(end_page),
            "{page_count}": str(page_count),
            "{total_pages}": str(total_pages),
            "{index}": f"{batch_index:03d}",
            "{timestamp}": timestamp_str,
            "{date}": date_str,
            "{time}": time_str,
        }

        # Formatted string
        result = pattern
        for placeholder, value in replacements.items():
            result = result.replace(placeholder, value)

        # Sanitize result
        clean_name = cls.sanitize_filename(result)
        
        # Ensure correct extension
        clean_ext = output_ext.lstrip(".")
        if not clean_name.lower().endswith(f".{clean_ext.lower()}"):
            clean_name = f"{clean_name}.{clean_ext}"

        return clean_name

    @classmethod
    def resolve_output_path(
        cls,
        output_dir: str,
        filename: str,
        overwrite: bool = False
    ) -> str:
        """
        Resolves destination path, generating unique names if file exists and overwrite=False.
        """
        os.makedirs(output_dir, exist_ok=True)
        target_path = os.path.join(output_dir, filename)

        if not os.path.exists(target_path) or overwrite:
            return target_path

        # Handle collision
        base_name, ext = os.path.splitext(filename)
        counter = 1
        while True:
            new_filename = f"{base_name} ({counter}){ext}"
            new_target_path = os.path.join(output_dir, new_filename)
            if not os.path.exists(new_target_path):
                return new_target_path
            counter += 1
