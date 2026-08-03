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
    def parse(range_str: str, total_pages: int, clamp_to_total: bool = True) -> List[Tuple[int, int]]:
        """
        Parses a page range specification string.

        Args:
            range_str: E.g., "1-5", "1, 3, 5-8", "3-end", "even", "odd", "all", "all-individual"
            total_pages: Total number of pages in the document.
            clamp_to_total: Whether to clamp range values to total_pages.

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

        # Keyword Aliases for document splitting
        individual_keywords = {
            "all-individual", "all_individual", "all individual",
            "individual", "split", "each", "every", "all-pages", "all_pages"
        }
        if cleaned in individual_keywords:
            return [(p, p) for p in range(1, total_pages + 1)]

        full_doc_keywords = {"all", "1-end", "full", "1 to end"}
        if cleaned in full_doc_keywords:
            return [(1, total_pages)]

        if cleaned in {"even", "evens"}:
            evens = [p for p in range(1, total_pages + 1) if p % 2 == 0]
            if not evens:
                raise RangeParseError(f"Document has {total_pages} page(s), no even pages found.")
            return [(p, p) for p in evens]

        if cleaned in {"odd", "odds"}:
            odds = [p for p in range(1, total_pages + 1) if p % 2 != 0]
            return [(p, p) for p in odds]

        # Normalize separators: replace semicolons and spaces with commas, and range connectors with hyphens
        normalized = cleaned.replace(";", ",")
        # Replace range connectors " to ", " through ", ":" with "-"
        normalized = re.sub(r"\s+(?:to|through)\s+", "-", normalized)
        normalized = normalized.replace(":", "-")

        parts = [p.strip() for p in normalized.split(",") if p.strip()]
        result_ranges: List[Tuple[int, int]] = []

        for part in parts:
            if "-" in part:
                tokens = [t.strip() for t in part.split("-") if t.strip()]
                if len(tokens) != 2:
                    raise RangeParseError(f"Invalid range segment '{part}'. Expected format 'X-Y'.")
                
                raw_start, raw_end = tokens[0], tokens[1]

                # Start page
                try:
                    start_page = int(raw_start)
                except ValueError:
                    raise RangeParseError(f"Invalid start page '{raw_start}' in segment '{part}'.")

                # End page
                if raw_end in ("end", "last"):
                    end_page = total_pages
                else:
                    try:
                        end_page = int(raw_end)
                    except ValueError:
                        raise RangeParseError(f"Invalid end page '{raw_end}' in segment '{part}'.")

                if clamp_to_total:
                    # Clamp pages safely to [1, total_pages]
                    start_page = max(1, min(start_page, total_pages))
                    end_page = max(start_page, min(end_page, total_pages))
                else:
                    start_page = max(1, start_page)
                    end_page = max(start_page, end_page)

                result_ranges.append((start_page, end_page))

            else:
                # Single page number or keyword
                if part in ("end", "last"):
                    page_num = total_pages
                else:
                    try:
                        page_num = int(part)
                    except ValueError:
                        raise RangeParseError(f"Invalid page number or identifier '{part}'.")

                if clamp_to_total:
                    page_num = max(1, min(page_num, total_pages))
                else:
                    page_num = max(1, page_num)
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
