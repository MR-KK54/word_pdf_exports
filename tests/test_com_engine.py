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



