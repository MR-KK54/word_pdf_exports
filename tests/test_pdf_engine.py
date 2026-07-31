"""
Unit tests for PDF Engine (PyMuPDF)
"""

import os
import tempfile

import pytest
import fitz

from word_exporter_pro.core.pdf_engine import PdfInspector, PdfPageExtractor


def _make_pdf(num_pages: int, path: str) -> str:
    doc = fitz.open()
    try:
        for _ in range(num_pages):
            page = doc.new_page()
            page.insert_text((72, 72), "Word & PDF Exporter Pro")
        doc.set_metadata({"title": "Test PDF", "author": "Tester"})
        doc.save(path)
    finally:
        doc.close()
    return path


@pytest.fixture
def pdf_path():
    fd, path = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    yield _make_pdf(5, path)
    if os.path.exists(path):
        os.remove(path)


def test_pdf_inspector(pdf_path):
    info = PdfInspector.get_info(pdf_path)
    assert info["page_count"] == 5
    assert info["format"] == "pdf"
    assert info["title"] == "Test PDF"
    assert info["author"] == "Tester"
    assert info["size_bytes"] > 0


def test_pdf_inspector_missing_file():
    with pytest.raises(FileNotFoundError):
        PdfInspector.get_info("C:/does/not/exist.pdf")


def test_pdf_extract_range(pdf_path):
    output_dir = tempfile.mkdtemp()
    output = os.path.join(output_dir, "extracted.pdf")
    result = PdfPageExtractor.extract_range(pdf_path, output, 2, 4)

    assert result == output
    assert os.path.exists(output)

    doc = fitz.open(output)
    try:
        assert doc.page_count == 3
        text = doc[0].get_text()
        assert "Word & PDF Exporter Pro" in text
    finally:
        doc.close()
    os.remove(output)
    os.rmdir(output_dir)


def test_pdf_extract_invalid_range(pdf_path):
    output_dir = tempfile.mkdtemp()
    output = os.path.join(output_dir, "bad.pdf")
    with pytest.raises(ValueError):
        PdfPageExtractor.extract_range(pdf_path, output, 4, 2)
    os.rmdir(output_dir)


def test_pdf_extract_single_page(pdf_path):
    output_dir = tempfile.mkdtemp()
    output = os.path.join(output_dir, "single.pdf")
    PdfPageExtractor.extract_range(pdf_path, output, 1, 1)
    doc = fitz.open(output)
    try:
        assert doc.page_count == 1
    finally:
        doc.close()
    os.remove(output)
    os.rmdir(output_dir)
