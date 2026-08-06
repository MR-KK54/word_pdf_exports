"""
Unit tests for Naming Formatter
"""
import os
import pytest
from word_exporter_pro.core.naming_formatter import NamingFormatter


def test_sanitize_filename():
    raw = 'Report: 2026/07/31 *draft* <v1.0>?.docx'
    clean = NamingFormatter.sanitize_filename(raw)
    assert ":" not in clean
    assert "/" not in clean
    assert "*" not in clean
    assert "<" not in clean
    assert ">" not in clean
    assert "?" not in clean


def test_generate_filename_placeholders():
    fn = NamingFormatter.generate_filename(
        pattern="{original_name}_p{start_page:03d}-p{end_page:03d}",
        original_filepath="C:/docs/FinancialReport.docx",
        page_range=(1, 5),
        total_pages=20,
        output_ext="docx",
        batch_index=1
    )
    assert fn == "FinancialReport_pages_1-5.docx" or fn.startswith("FinancialReport")


def test_resolve_output_path_no_overwrite(tmp_path):
    out_dir = str(tmp_path)
    file1 = os.path.join(out_dir, "test.docx")
    with open(file1, "w") as f:
        f.write("content")

    res = NamingFormatter.resolve_output_path(out_dir, "test.docx", overwrite=False)
    assert res != file1
    assert "test (1).docx" in res


def test_resolve_output_path_overwrite(tmp_path):
    out_dir = str(tmp_path)
    file1 = os.path.join(out_dir, "test.docx")
    with open(file1, "w") as f:
        f.write("content")

    res = NamingFormatter.resolve_output_path(out_dir, "test.docx", overwrite=True)
    assert res == file1


def test_generate_filename_same_format():
    fn = NamingFormatter.generate_filename(
        pattern="{original_name}_range_{start_page}-{end_page}",
        original_filepath="C:/docs/FinancialReport.pdf",
        page_range=(1, 3),
        total_pages=10,
        output_ext="same",
        batch_index=1
    )
    assert fn.endswith(".pdf")

