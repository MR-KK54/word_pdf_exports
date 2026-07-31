"""
Word Page Exporter Pro - Page Range Parser
Parses and validates page range specification strings into tuples of (start_page, end_page).
"""

import re
from typing import List, Tuple, Union


class RangeParseError(ValueError):
    """Exception raised for invalid page range expressions."""
    pass


class PageRangeParser:
    """Parses user input range strings into structured (start_page, end_page) list."""

    @staticmethod
    def parse(range_str: str, total_pages: int) -> List[Tuple[int, int]]:
        """
        Parses a page range specification string.

        Args:
            range_str: E.g., "1-5", "1, 3, 5-8", "3-end", "even", "odd", "all", "all-individual"
            total_pages: Total number of pages in the document.

        Returns:
            List of (start_page, end_page) 1-indexed tuples.

        Raises:
            RangeParseError: If syntax is invalid or out of bounds.
        """
        if total_pages <= 0:
            raise RangeParseError(f"Invalid total page count: {total_pages}. Document must have at least 1 page.")

        cleaned = range_str.strip().lower()
        if not cleaned:
            raise RangeParseError("Page range input cannot be empty.")

        # Preset keywords
        if cleaned == "all" or cleaned == "1-end":
            return [(1, total_pages)]
        
        if cleaned == "all-individual":
            return [(p, p) for p in range(1, total_pages + 1)]

        if cleaned == "even":
            evens = [p for p in range(1, total_pages + 1) if p % 2 == 0]
            if not evens:
                raise RangeParseError(f"Document has {total_pages} page(s), no even pages found.")
            return [(p, p) for p in evens]

        if cleaned == "odd":
            odds = [p for p in range(1, total_pages + 1) if p % 2 != 0]
            return [(p, p) for p in odds]

        # Comma-separated parts
        parts = [p.strip() for p in cleaned.split(",") if p.strip()]
        result_ranges: List[Tuple[int, int]] = []

        for part in parts:
            if "-" in part:
                tokens = part.split("-")
                if len(tokens) != 2:
                    raise RangeParseError(f"Invalid range segment '{part}'. Expected format 'X-Y'.")
                
                raw_start, raw_end = tokens[0].strip(), tokens[1].strip()

                # Start page
                try:
                    start_page = int(raw_start)
                except ValueError:
                    raise RangeParseError(f"Invalid start page '{raw_start}' in segment '{part}'.")

                # End page
                if raw_end == "end":
                    end_page = total_pages
                else:
                    try:
                        end_page = int(raw_end)
                    except ValueError:
                        raise RangeParseError(f"Invalid end page '{raw_end}' in segment '{part}'.")

                # Validation
                if start_page < 1:
                    raise RangeParseError(f"Start page {start_page} cannot be less than 1.")
                if end_page > total_pages:
                    raise RangeParseError(f"End page {end_page} exceeds document total pages ({total_pages}).")
                if start_page > end_page:
                    raise RangeParseError(f"Start page {start_page} is greater than end page {end_page} in range '{part}'.")

                result_ranges.append((start_page, end_page))

            else:
                # Single page number
                if part == "end":
                    page_num = total_pages
                else:
                    try:
                        page_num = int(part)
                    except ValueError:
                        raise RangeParseError(f"Invalid page number or identifier '{part}'.")

                if page_num < 1 or page_num > total_pages:
                    raise RangeParseError(f"Page number {page_num} is out of valid bounds (1-{total_pages}).")

                result_ranges.append((page_num, page_num))

        return result_ranges

    @staticmethod
    def format_range_summary(ranges: List[Tuple[int, int]]) -> str:
        """Formats a list of ranges into a compact display string like '1-3, 5, 7-9'."""
        formatted_parts = []
        for start, end in ranges:
            if start == end:
                formatted_parts.append(str(start))
            else:
                formatted_parts.append(f"{start}-{end}")
        return ", ".join(formatted_parts)
