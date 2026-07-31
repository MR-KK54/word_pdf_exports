"""
Unit tests for Page Range Parser
"""
import pytest
from word_exporter_pro.core.range_parser import PageRangeParser, RangeParseError


def test_single_page_parsing():
    ranges = PageRangeParser.parse("3", total_pages=10)
    assert ranges == [(3, 3)]


def test_range_parsing():
    ranges = PageRangeParser.parse("1-5", total_pages=10)
    assert ranges == [(1, 5)]


def test_multiple_ranges():
    ranges = PageRangeParser.parse("1-3, 5, 8-10", total_pages=10)
    assert ranges == [(1, 3), (5, 5), (8, 10)]


def test_end_keyword():
    ranges = PageRangeParser.parse("3-end", total_pages=12)
    assert ranges == [(3, 12)]


def test_even_odd_presets():
    evens = PageRangeParser.parse("even", total_pages=5)
    assert evens == [(2, 2), (4, 4)]

    odds = PageRangeParser.parse("odd", total_pages=5)
    assert odds == [(1, 1), (3, 3), (5, 5)]


def test_all_individual_preset():
    ind = PageRangeParser.parse("all-individual", total_pages=3)
    assert ind == [(1, 1), (2, 2), (3, 3)]


def test_out_of_bounds_errors():
    with pytest.raises(RangeParseError):
        PageRangeParser.parse("15", total_pages=10)

    with pytest.raises(RangeParseError):
        PageRangeParser.parse("5-2", total_pages=10)

    with pytest.raises(RangeParseError):
        PageRangeParser.parse("0", total_pages=10)


def test_format_range_summary():
    assert PageRangeParser.format_range_summary([(1, 3), (5, 5), (8, 10)]) == "1-3, 5, 8-10"
