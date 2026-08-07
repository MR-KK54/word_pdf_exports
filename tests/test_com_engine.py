"""
Integration tests for MS Word COM Engine
"""
import os
import pytest
import win32com.client
import pythoncom
from word_exporter_pro.core.com_engine import WordCOMContext, DocumentInspector, PageExporterEngine


@pytest.fixture(scope="module")
def sample_word_doc(tmp_path_factory):
    """Creates a 3-page sample Word document using Word COM Automation."""
    tmp_dir = tmp_path_factory.mktemp("word_test")
    doc_path = os.path.join(str(tmp_dir), "SampleMultiPageDoc.docx")

    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Add()

        # Page 1
        doc.Content.Text = "Page 1 Content - Introduction\n"
        word.Selection.EndKey(6) # wdStory = 6
        word.Selection.InsertBreak(7) # wdPageBreak = 7

        # Page 2
        word.Selection.TypeText("Page 2 Content - Executive Summary\n")
        word.Selection.InsertBreak(7)

        # Page 3
        word.Selection.TypeText("Page 3 Content - Conclusion and Appendices\n")

        doc.SaveAs2(FileName=doc_path, FileFormat=16) # wdFormatXMLDocument
        doc.Close(SaveChanges=False)
    finally:
        word.Quit(SaveChanges=False)
        pythoncom.CoUninitialize()

    return doc_path


def test_document_inspection(sample_word_doc):
    info = DocumentInspector.get_info(sample_word_doc)
    assert info["filename"] == "SampleMultiPageDoc.docx"
    assert info["page_count"] == 3
    assert info["section_count"] >= 1
    assert info["format"] == "docx"


def test_export_single_page(sample_word_doc, tmp_path):
    out_file = os.path.join(str(tmp_path), "Exported_Page2.docx")
    res_path = PageExporterEngine.export_range(
        source_file=sample_word_doc,
        output_file=out_file,
        start_page=2,
        end_page=2,
        export_format="docx",
        mode="trimming"
    )
    assert os.path.exists(res_path)

    # Verify page count of exported file
    out_info = DocumentInspector.get_info(res_path)
    assert out_info["page_count"] == 1


def test_export_pdf_format(sample_word_doc, tmp_path):
    out_file = os.path.join(str(tmp_path), "Exported_Pages1-2.pdf")
    res_path = PageExporterEngine.export_range(
        source_file=sample_word_doc,
        output_file=out_file,
        start_page=1,
        end_page=2,
        export_format="pdf",
        mode="trimming"
    )
    assert os.path.exists(res_path)
    assert os.path.getsize(res_path) > 0


@pytest.fixture(scope="module")
def doc_with_formatted_footer(tmp_path_factory):
    """Creates a document with styled/formatted footers."""
    tmp_dir = tmp_path_factory.mktemp("formatted_footer_test")
    doc_path = os.path.join(str(tmp_dir), "FormattedFooterDoc.docx")

    with WordCOMContext(visible=False) as word:
        doc = word.Documents.Add()
        try:
            # Add content
            doc.Content.Text = "Page 1 Content\n"
            word.Selection.EndKey(6)
            word.Selection.InsertBreak(7)
            word.Selection.TypeText("Page 2 Content\n")

            # Format Footer
            footer = doc.Sections(1).Footers(1)
            f_range = footer.Range
            f_range.Text = "Confidential - Page "
            f_range.Font.Name = "Arial"
            f_range.Font.Size = 12
            f_range.Font.Bold = True
            f_range.Font.Color = 255 # Red
            f_range.ParagraphFormat.Alignment = 1 # Center
            
            f_range.Collapse(0)
            doc.Fields.Add(f_range, 33) # wdFieldPage

            doc.SaveAs2(FileName=doc_path, FileFormat=16)
        finally:
            doc.Close(SaveChanges=False)

    return doc_path


@pytest.mark.parametrize("engine_mode", ["trimming", "selection"])
def test_export_preserves_footer_formatting(doc_with_formatted_footer, tmp_path, engine_mode):
    out_file = os.path.join(str(tmp_path), f"Exported_Footer_{engine_mode}.docx")
    res_path = PageExporterEngine.export_range(
        source_file=doc_with_formatted_footer,
        output_file=out_file,
        start_page=2,
        end_page=2,
        export_format="docx",
        mode=engine_mode
    )
    assert os.path.exists(res_path)

    # Inspect footer formatting in exported file via WordCOMContext
    with WordCOMContext(visible=False) as word:
        exported_doc = word.Documents.Open(FileName=res_path, ReadOnly=True)
        try:
            footer = exported_doc.Sections(1).Footers(1)
            # Verify font formatting preserved
            assert footer.Range.Font.Bold is True or footer.Range.Font.Bold == -1
            assert footer.Range.Font.Color == 255
            assert "Confidential - Page" in footer.Range.Text
        finally:
            exported_doc.Close(SaveChanges=False)


def test_export_main_file_footer_matching(tmp_path):
    """Verifies that split files retain exact main file footers without '0 de' field corruptions."""
    doc_path = os.path.join(str(tmp_path), "UserMainDoc.docx")
    out_file = os.path.join(str(tmp_path), "UserSplitDoc.docx")

    with WordCOMContext(visible=False) as word:
        doc = word.Documents.Add()
        try:
            doc.Content.Text = "Page 1 Body Text\n"
            word.Selection.EndKey(6)
            word.Selection.InsertBreak(7)
            word.Selection.TypeText("Page 2 Body Text\n")

            footer = doc.Sections(1).Footers(1)
            f_range = footer.Range
            f_range.Text = "Data de Submissão do Projeto: 10/07/2026\nNome do Arquivo: PB_INFORMAÇÕES_BÁSICAS.pdf\nVersão do Projeto: 4\nPágina "
            
            f_range2 = footer.Range
            f_range2.Collapse(0)
            doc.Fields.Add(f_range2, 33) # wdFieldPage = 33

            f_range3 = footer.Range
            f_range3.Collapse(0)
            f_range3.Text = " de "
            f_range3.Collapse(0)
            doc.Fields.Add(f_range3, 26) # wdFieldNumPages = 26

            doc.SaveAs2(FileName=doc_path, FileFormat=16)
        finally:
            doc.Close(SaveChanges=False)

    # Perform page 2 export via trimming engine
    res_path = PageExporterEngine.export_range(
        source_file=doc_path,
        output_file=out_file,
        start_page=2,
        end_page=2,
        export_format="docx",
        mode="trimming"
    )
    assert os.path.exists(res_path)

    # Verify split file footer matches main file exactly and has no '0 de' bug
    with WordCOMContext(visible=False) as word:
        exported_doc = word.Documents.Open(FileName=res_path, ReadOnly=True)
        try:
            footer_text = exported_doc.Sections(1).Footers(1).Range.Text
            assert "Data de Submissão do Projeto: 10/07/2026" in footer_text
            assert "Nome do Arquivo: PB_INFORMAÇÕES_BÁSICAS.pdf" in footer_text
            assert "Versão do Projeto: 4" in footer_text
            assert "Página" in footer_text
            assert "0 de" not in footer_text
        finally:
            exported_doc.Close(SaveChanges=False)


def test_word_com_context_available():
    ctx = WordCOMContext()
    assert isinstance(ctx.available, bool)


def test_libreoffice_availability_check():
    from word_exporter_pro.core.com_engine import _is_libreoffice_available
    is_available = _is_libreoffice_available()
    assert isinstance(is_available, bool)


@pytest.fixture(scope="module")
def doc_with_trailing_section_break(tmp_path_factory):
    """Creates a document where the last KEPT page is in a landscape section whose
    margin differs from the trimmed-away trailing section:

        Section 1 (portrait, L=72): pages 1-2
        Section 2 (landscape, L=36): page 3   <- last kept page
        Section 3 (landscape, L=108): page 4  <- trimmed away

    A bug used to make the exported last page inherit Section 3's layout (L=108)
    after the trailing section break was removed, corrupting the last kept page.
    """
    tmp_dir = tmp_path_factory.mktemp("trailing_section_test")
    doc_path = os.path.join(str(tmp_dir), "TrailingSectionDoc.docx")

    with WordCOMContext(visible=False) as word:
        doc = word.Documents.Add()
        try:
            sel = word.Selection
            sel.TypeText("Page1 A\n")
            sel.TypeText("Page1 B\n")
            sel.EndKey(6)
            sel.InsertBreak(7)  # page break
            sel.TypeText("Page2 A\n")
            sel.TypeText("Page2 B\n")
            sel.EndKey(6)
            sel.InsertBreak(2)  # section break next page -> Section 2
            sec2 = doc.Sections(2)
            sec2.PageSetup.Orientation = 1  # landscape
            sec2.PageSetup.PageWidth = 792
            sec2.PageSetup.PageHeight = 612
            sec2.PageSetup.LeftMargin = 36
            sel.TypeText("Page3 A (landscape)\n")
            sel.TypeText("Page3 B (landscape)\n")
            sel.EndKey(6)
            sel.InsertBreak(2)  # section break next page -> Section 3
            sec3 = doc.Sections(3)
            sec3.PageSetup.Orientation = 1  # landscape
            sec3.PageSetup.PageWidth = 792
            sec3.PageSetup.PageHeight = 612
            sec3.PageSetup.LeftMargin = 108  # deliberately different from Section 2
            sel.TypeText("Page4 A (trailing)\n")
            sel.TypeText("Page4 B (trailing)\n")
            doc.SaveAs2(FileName=doc_path, FileFormat=16)
        finally:
            doc.Close(SaveChanges=False)
    return doc_path


def test_export_last_page_preserves_layout_with_trailing_section(doc_with_trailing_section_break, tmp_path):
    """Exporting pages 1-3 must keep page 3's landscape layout (Section 2, margin 36)
    after the trailing Section 3 break is removed, and must not leave a trailing page."""
    out_file = os.path.join(str(tmp_path), "TrailingSection_1-3.docx")
    res_path = PageExporterEngine.export_range(
        source_file=doc_with_trailing_section_break,
        output_file=out_file,
        start_page=1,
        end_page=3,
        export_format="docx",
        mode="trimming",
        total_pages=4,
    )
    assert os.path.exists(res_path)

    with WordCOMContext(visible=False) as word:
        exported_doc = word.Documents.Open(FileName=res_path, ReadOnly=True)
        try:
            final_pages = exported_doc.ComputeStatistics(2)  # wdStatisticPages
            assert final_pages == 3

            # The last kept page's section must keep the SOURCE Section 2 layout
            # (landscape, 36pt margin), not the trimmed-away Section 3's margin.
            last_sec = exported_doc.Sections(exported_doc.Sections.Count)
            assert last_sec.PageSetup.Orientation == 1  # wdOrientationLandscape
            assert abs(last_sec.PageSetup.LeftMargin - 36.0) < 0.01

            # No trailing empty/overflow page: the kept content must end cleanly.
            assert "Page3 B (landscape)" in exported_doc.Content.Text
            assert "Page4" not in exported_doc.Content.Text
        finally:
            exported_doc.Close(SaveChanges=False)


def test_export_intermediate_range_preserves_layout(doc_with_trailing_section_break, tmp_path):
    """Exporting pages 1-2 (which live in the portrait Section 1) must keep the
    portrait layout instead of inheriting the following landscape section's."""
    out_file = os.path.join(str(tmp_path), "TrailingSection_1-2.docx")
    res_path = PageExporterEngine.export_range(
        source_file=doc_with_trailing_section_break,
        output_file=out_file,
        start_page=1,
        end_page=2,
        export_format="docx",
        mode="trimming",
        total_pages=4,
    )
    assert os.path.exists(res_path)

    with WordCOMContext(visible=False) as word:
        exported_doc = word.Documents.Open(FileName=res_path, ReadOnly=True)
        try:
            final_pages = exported_doc.ComputeStatistics(2)
            assert final_pages == 2
            last_sec = exported_doc.Sections(exported_doc.Sections.Count)
            # Kept page 2 lives in the portrait Section 1 (LeftMargin 72).
            assert last_sec.PageSetup.Orientation == 0  # wdOrientationPortrait
            assert abs(last_sec.PageSetup.LeftMargin - 72.0) < 0.01
            assert "Page3" not in exported_doc.Content.Text
        finally:
            exported_doc.Close(SaveChanges=False)


def test_export_removes_trailing_manual_page_break(tmp_path):
    """A trailing manual page break at the end of the kept range must not create an
    extra blank page; the last kept content must end the document."""
    doc_path = os.path.join(str(tmp_path), "TrailingPageBreakDoc.docx")
    with WordCOMContext(visible=False) as word:
        doc = word.Documents.Add()
        try:
            sel = word.Selection
            sel.TypeText("Page1 A\n")
            sel.EndKey(6)
            sel.InsertBreak(7)
            sel.TypeText("Page2 A\n")
            sel.EndKey(6)
            sel.InsertBreak(7)  # page break that opens page 3 (trimmed away)
            sel.TypeText("Page3 A (trimmed)\n")
            doc.SaveAs2(FileName=doc_path, FileFormat=16)
        finally:
            doc.Close(SaveChanges=False)

    out_file = os.path.join(str(tmp_path), "TrailingPageBreak_1-2.docx")
    res_path = PageExporterEngine.export_range(
        source_file=doc_path,
        output_file=out_file,
        start_page=1,
        end_page=2,
        export_format="docx",
        mode="trimming",
        total_pages=3,
    )
    assert os.path.exists(res_path)

    with WordCOMContext(visible=False) as word:
        exported_doc = word.Documents.Open(FileName=res_path, ReadOnly=True)
        try:
            final_pages = exported_doc.ComputeStatistics(2)
            assert final_pages == 2
            text = exported_doc.Content.Text
            assert "Page2 A" in text
            assert "Page3 A (trimmed)" not in text
        finally:
            exported_doc.Close(SaveChanges=False)




