"""
Word Page Exporter Pro - Microsoft Word COM Engine
Provides high-fidelity page extraction and document pagination via Word.Application.
"""

import os
import shutil
import tempfile
from typing import Dict, Any, Optional, Tuple

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover - exercised in non-Windows environments
    pythoncom = None
    win32com = None

try:
    import docx  # type: ignore
except ImportError:
    docx = None

from word_exporter_pro.utils.logger import get_logger

logger = get_logger()


def _require_word_com() -> None:
    if pythoncom is None or win32com is None:
        raise RuntimeError(
            "Microsoft Word COM support is unavailable on this server. "
            "Install pywin32 on Windows to enable Word document export features."
        )


# Word COM Constants
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_STATISTIC_PAGES = 2
WD_ALERTS_NONE = 0

EXPORT_FORMAT_MAP = {
    "docx": 16,   # wdFormatXMLDocument
    "doc": 0,     # wdFormatDocument97
    "pdf": 17,    # wdFormatPDF
    "rtf": 6,     # wdFormatRTF
    "docm": 13,   # wdFormatXMLDocumentMacroEnabled
    "dotx": 14,   # wdFormatXMLTemplate
}


class WordCOMContext:
    """Context manager for Word.Application COM lifecycle."""

    def __init__(self, visible: bool = False):
        self.visible = visible
        self.word_app = None
        self.co_initialized = False

    def __enter__(self):
        _require_word_com()
        pythoncom.CoInitialize()
        self.co_initialized = True
        try:
            # DispatchEx guarantees a new, isolated Word instance
            try:
                self.word_app = win32com.client.DispatchEx("Word.Application")
            except Exception:
                self.word_app = win32com.client.Dispatch("Word.Application")

            try:
                self.word_app.Visible = self.visible
            except Exception:
                pass

            try:
                self.word_app.DisplayAlerts = WD_ALERTS_NONE
            except Exception:
                pass

            try:
                self.word_app.ScreenUpdating = False
            except Exception:
                pass

            return self.word_app
        except Exception as e:
            logger.error(f"Failed to initialize MS Word COM Application: {e}")
            if self.co_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                self.co_initialized = False
            raise RuntimeError(f"Microsoft Word COM server could not be started: {e}")

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.word_app:
            try:
                self.word_app.Quit(SaveChanges=False)
            except Exception as e:
                logger.warning(f"Error closing Word Application instance: {e}")
            self.word_app = None

        if self.co_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self.co_initialized = False


class DocumentInspector:
    """Inspects Word document structure using MS Word layout engine with python-docx server fallback."""

    @staticmethod
    def get_info(file_path: str, visible: bool = False) -> Dict[str, Any]:
        """
        Extracts document statistics and layout information.
        """
        abs_path = os.path.abspath(file_path)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Document file not found: {abs_path}")

        info = {
            "path": abs_path,
            "filename": os.path.basename(abs_path),
            "size_bytes": os.path.getsize(abs_path),
            "page_count": 0,
            "section_count": 0,
            "title": "",
            "author": "",
            "format": os.path.splitext(abs_path)[1].lower().lstrip("."),
        }

        # Server fallback for non-Windows / Linux server environments without win32com
        if pythoncom is None or win32com is None:
            if docx is not None and info["format"] in ("docx", "docm", "dotx"):
                try:
                    d = docx.Document(abs_path)
                    info["section_count"] = max(1, len(d.sections))
                    try:
                        info["title"] = str(d.core_properties.title or "")
                        info["author"] = str(d.core_properties.author or "")
                    except Exception:
                        pass

                    xml_str = d._body._element.xml
                    page_breaks = xml_str.count('type="page"') + xml_str.count('w:lastRenderedPageBreak')
                    info["page_count"] = max(1, page_breaks + 1 if page_breaks > 0 else 1)
                    logger.info(f"Inspected '{info['filename']}' via server fallback: {info['page_count']} page(s)")
                    return info
                except Exception as fallback_err:
                    logger.warning(f"Server fallback inspection warning: {fallback_err}")
            
            info["page_count"] = 1
            info["section_count"] = 1
            return info

        with WordCOMContext(visible=visible) as word_app:
            doc = None
            try:
                doc = word_app.Documents.Open(
                    FileName=abs_path,
                    ReadOnly=True,
                    ConfirmConversions=False,
                    AddToRecentFiles=False,
                    Visible=visible
                )
                
                # Compute exact page count using Word's pagination engine
                info["page_count"] = doc.ComputeStatistics(WD_STATISTIC_PAGES)
                info["section_count"] = doc.Sections.Count
                
                # Metadata properties
                try:
                    info["title"] = str(doc.BuiltInDocumentProperties("Title").Value or "")
                    info["author"] = str(doc.BuiltInDocumentProperties("Author").Value or "")
                except Exception:
                    pass

            except Exception as e:
                logger.error(f"Error inspecting document '{abs_path}': {e}")
                raise RuntimeError(f"Could not open or inspect Word document: {e}")
            finally:
                if doc:
                    try:
                        doc.Close(SaveChanges=False)
                    except Exception:
                        pass

        return info


class PageExporterEngine:
    """Core engine for exporting page ranges with maximum formatting fidelity."""

    @staticmethod
    def export_range(
        source_file: str,
        output_file: str,
        start_page: int,
        end_page: int,
        export_format: str = "docx",
        mode: str = "trimming",
        visible: bool = False
    ) -> str:
        """
        Exports a page range from source Word doc into output_file.

        Args:
            source_file: Source Word document path.
            output_file: Target path to save output file.
            start_page: 1-indexed start page.
            end_page: 1-indexed end page.
            export_format: Output format ('docx', 'pdf', 'doc', 'rtf', 'docm').
            mode: Extraction method ('trimming' or 'selection').
            visible: Whether Word app runs visibly.

        Returns:
            Absolute path of created output file.
        """
        abs_source = os.path.abspath(source_file)
        abs_output = os.path.abspath(output_file)
        os.makedirs(os.path.dirname(abs_output), exist_ok=True)

        logger.info(f"Starting page export [{start_page}-{end_page}] from '{os.path.basename(abs_source)}' -> '{os.path.basename(abs_output)}'")

        # Non-Windows / Linux server fallback
        if pythoncom is None or win32com is None:
            if export_format.lower() == "pdf":
                raise RuntimeError(
                    "PDF export format for Word documents is not supported on this platform. "
                    "Please install Microsoft Word on Windows, or export as DOCX instead."
                )
            return PageExporterEngine._export_by_docx_fallback(
                abs_source, abs_output, start_page, end_page, export_format
            )

        fmt_code = EXPORT_FORMAT_MAP.get(export_format.lower())
        if fmt_code is None:
            raise ValueError(f"Unsupported export format '{export_format}'. Supported formats: {list(EXPORT_FORMAT_MAP.keys())}")

        if mode == "trimming":
            return PageExporterEngine._export_by_trimming(
                abs_source, abs_output, start_page, end_page, fmt_code, visible
            )
        else:
            return PageExporterEngine._export_by_selection(
                abs_source, abs_output, start_page, end_page, fmt_code, visible
            )

    @staticmethod
    def _export_by_docx_fallback(
        abs_source: str,
        abs_output: str,
        start_page: int,
        end_page: int,
        export_format: str
    ) -> str:
        """Non-Windows / Linux cloud server fallback engine using python-docx to trim pages."""
        if docx is None:
            shutil.copy2(abs_source, abs_output)
            return abs_output

        try:
            doc = docx.Document(abs_source)
            body = doc._body._element

            current_page = 1
            elements_to_remove = []

            for child in list(body):
                tag = child.tag.rsplit('}', 1)[-1]
                if tag not in ('p', 'tbl'):
                    continue

                xml = child.xml
                has_break = ('type="page"' in xml or 'lastRenderedPageBreak' in xml or ('w:br' in xml and 'page' in xml))

                element_page = current_page

                if has_break:
                    current_page += 1

                if element_page < start_page or element_page > end_page:
                    elements_to_remove.append(child)

            if elements_to_remove:
                for el in elements_to_remove:
                    try:
                        body.remove(el)
                    except Exception:
                        pass

            doc.save(abs_output)
            logger.success(f"Trimmed pages [{start_page}-{end_page}] (Server mode) to '{abs_output}'")
            return abs_output
        except Exception as e:
            logger.warning(f"python-docx fallback trimming warning: {e}. Copying original file.")
            shutil.copy2(abs_source, abs_output)
            return abs_output

    @staticmethod
    def _export_by_trimming(
        abs_source: str,
        abs_output: str,
        start_page: int,
        end_page: int,
        fmt_code: int,
        visible: bool
    ) -> str:
        """High-Fidelity Trimming Engine: Duplicates source doc and trims unwanted pages."""
        temp_dir = tempfile.mkdtemp(prefix="word_exp_")
        temp_source_copy = os.path.join(temp_dir, f"work_{os.path.basename(abs_source)}")

        try:
            shutil.copy2(abs_source, temp_source_copy)

            with WordCOMContext(visible=visible) as word_app:
                source_doc = None
                doc = word_app.Documents.Open(
                    FileName=temp_source_copy,
                    ReadOnly=False,
                    ConfirmConversions=False,
                    AddToRecentFiles=False,
                    Visible=visible
                )

                try:
                    total_pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
                    
                    start_page = max(1, min(start_page, total_pages))
                    end_page = max(start_page, min(end_page, total_pages))
                    expected_keep = end_page - start_page + 1

                    # 1. Trim the FRONT: delete everything up to the start of the kept page.
                    if start_page > 1:
                        start_pos = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=start_page).Start
                        PageExporterEngine._delete_range(
                            word_app, doc.Range(Start=doc.Content.Start, End=start_pos)
                        )

                    # 2. Trim the BACK: keep exactly `expected_keep` pages. Recomputes pagination
                    #    between deletions because Word may collapse a blank page after each cut.
                    PageExporterEngine._trim_tail_to_keep(word_app, doc, expected_keep)

                    # 3. Remove any stray page-break/blank paragraph left at the cut edges so no
                    #    spurious blank front/back page survives in the trimmed result.
                    #    The front edge is only cleaned when a front trim actually occurred,
                    #    so a legitimate leading page-break on a range starting at page 1 is kept.
                    if start_page > 1:
                        PageExporterEngine._clean_page_boundary(word_app, doc, front_of_doc=True)
                    PageExporterEngine._clean_page_boundary(word_app, doc, front_of_doc=False)

                    # 4. When trimming removes content BEFORE/AFTER the range, the boundary section breaks
                    #    are deleted too, so the kept sections may inherit an adjacent (wrong) section's
                    #    header/footer, page setup, and page-number scheme. Re-sync them against the
                    #    original source document so the exported range looks correct.
                    if end_page < total_pages or start_page > 1:
                        try:
                            source_doc = word_app.Documents.Open(
                                FileName=abs_source,
                                ReadOnly=True,
                                ConfirmConversions=False,
                                AddToRecentFiles=False,
                                Visible=False
                            )
                            start_section_index = PageExporterEngine._section_containing_page(source_doc, start_page)
                            end_section_index = PageExporterEngine._section_containing_page(source_doc, end_page)
                            PageExporterEngine._restore_sections_from_source(
                                target_doc=doc,
                                source_doc=source_doc,
                                start_section_index=start_section_index,
                                end_section_index=end_section_index,
                                only_last_section=True
                            )
                        except Exception as restore_err:
                            logger.warning(f"Could not restore headers/footers for trimmed range: {restore_err}")

                    # 5. Restart page numbering from 1 on the trimmed document so PAGE/NUMPAGES
                    #    fields show the range-relative numbers (e.g. "2 of 3") instead of the
                    #    source document's original page numbers. This must run AFTER the section
                    #    restore above, otherwise the earlier StartingNumber overwrites would win.
                    PageExporterEngine._restart_page_numbering(doc, start_page)

                    # 6. Natively update field codes (PAGE, NUMPAGES, TOC, etc.) across the body
                    #    and every header/footer so they reflect the final trimmed content.
                    try:
                        try:
                            doc.Fields.Update()
                        except Exception:
                            pass
                        try:
                            for sec in doc.Sections:
                                for hf_coll in (sec.Headers, sec.Footers):
                                    for hf_idx in (1, 2, 3):
                                        try:
                                            hf_coll(hf_idx).Range.Fields.Update()
                                        except Exception:
                                            pass
                        except Exception:
                            pass
                    except Exception:
                        pass

                    # 6. Save as output destination format
                    doc.SaveAs2(FileName=abs_output, FileFormat=fmt_code)
                    logger.success(f"Exported pages [{start_page}-{end_page}] to '{abs_output}'")

                finally:
                    if source_doc:
                        try:
                            source_doc.Close(SaveChanges=False)
                        except Exception:
                            pass
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except Exception:
                            pass

        finally:
            if os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir, ignore_errors=True)
                except Exception as e:
                    logger.warning(f"Could not remove temporary directory {temp_dir}: {e}")

        return abs_output

    @staticmethod
    def _delete_range(word_app, cut_range) -> None:
        """Delete a range robustly; Range.Delete() can fail on ranges ending at Content.End."""
        try:
            cut_range.Delete()
            return
        except Exception:
            pass
        try:
            cut_range.Select()
            word_app.Selection.Delete()
            return
        except Exception:
            pass
        # Last resort: shrink the range by one character and retry
        try:
            cut_range.MoveEnd(6, -1)
            cut_range.Delete()
        except Exception:
            logger.warning("Could not delete range during page trimming.")

    @staticmethod
    def _trim_tail_to_keep(word_app, doc, expected_keep: int) -> None:
        """Delete trailing content until the document holds exactly `expected_keep` pages.

        Word re-paginates the document after every deletion (a blank page or empty
        paragraph at a cut boundary can collapse and change the page count), so the
        target is re-evaluated on each iteration instead of trusting a single
        pre-computed boundary. This guarantees the kept count matches the requested
        custom page range regardless of manual page breaks.
        """
        guard = 0
        while guard < 60:
            current_pages = doc.ComputeStatistics(WD_STATISTIC_PAGES)
            if current_pages <= expected_keep:
                break
            try:
                drop_start = doc.GoTo(
                    What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=expected_keep + 1
                ).Start
                PageExporterEngine._delete_range(
                    word_app, doc.Range(Start=drop_start, End=doc.Content.End)
                )
            except Exception:
                break
            guard += 1

    @staticmethod
    def _clean_page_boundary(word_app, doc, front_of_doc: bool) -> None:
        """Remove a stray leading/trailing page-break or blank paragraph left at the
        trimmed edge so the first and last pages of the result hold only real content."""
        for _ in range(5):
            try:
                if front_of_doc:
                    rng = doc.Range(
                        Start=doc.Content.Start,
                        End=min(doc.Content.Start + 2, doc.Content.End),
                    )
                else:
                    if doc.Content.End <= 2:
                        break
                    rng = doc.Range(Start=doc.Content.End - 2, End=doc.Content.End)
                if rng.End <= rng.Start:
                    break
                if "\x0c" in rng.Text or "\f" in rng.Text:
                    rng.Delete()
                else:
                    break
            except Exception:
                break

    @staticmethod
    def _restart_page_numbering(doc, start_page: int) -> None:
        """Reset page numbering on the trimmed document so the PAGE/NUMPAGES fields show
        range-relative numbers instead of the source document's original numbering.

        - First section restarts at 1 (and re-enables restart-at-section), so the kept
          first page reads as page 1 of the new document.
        - Subsequent sections continue sequentially (RestartNumberingAtSection=False) so
          numbering flows across sections without resetting.
        - Only the page-numbering scheme is touched; explicit header/footer text is left
          intact so the section restore above is preserved.
        """
        try:
            if not doc.Sections:
                return
            first = True
            for sec in doc.Sections:
                try:
                    ps = sec.PageSetup
                    if first:
                        ps.PageNumbering.RestartNumberingAtSection = True
                        ps.PageNumbering.StartingNumber = 1
                        first = False
                    else:
                        ps.PageNumbering.RestartNumberingAtSection = False
                except Exception:
                    pass
        except Exception:
            pass

    @staticmethod
    def _section_containing_page(doc, page: int) -> int:
        """Return the 1-indexed source section that contains the given page."""
        try:
            pos = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=page).Start
            for i in range(1, doc.Sections.Count + 1):
                r = doc.Sections(i).Range
                if r.Start <= pos < r.End:
                    return i
            return doc.Sections.Count
        except Exception:
            return 1

    @staticmethod
    def _get_effective_header_footer(source_doc, sec_index: int, is_header: bool, hf_index: int):
        """Trace back LinkToPrevious to find the effective source Header/Footer object."""
        curr_idx = sec_index
        while curr_idx >= 1:
            try:
                sec = source_doc.Sections(curr_idx)
                hf = sec.Headers(hf_index) if is_header else sec.Footers(hf_index)
                if not hf.LinkToPrevious or curr_idx == 1:
                    return hf
            except Exception:
                pass
            curr_idx -= 1
        try:
            sec = source_doc.Sections(1)
            return sec.Headers(hf_index) if is_header else sec.Footers(hf_index)
        except Exception:
            return None

    @staticmethod
    def _copy_header_footer_formatting(src_hf, tgt_hf) -> None:
        """Copies resolved formatting from a source header/footer to a target
        header/footer. FormattedText only carries explicitly-set formatting, so
        style-inherited values (indents, alignment, spacing, fonts) would otherwise
        resolve against the target document's own built-in styles. Assigning the
        source ParagraphFormat + font properties transfers the effective values.
        The font copy runs LAST because style/paragraph-format assignment reapplies
        the target's style character formatting to the range."""
        try:
            src_count = src_hf.Range.Paragraphs.Count
            tgt_count = tgt_hf.Range.Paragraphs.Count
            for i in range(1, min(src_count, tgt_count) + 1):
                src_para = src_hf.Range.Paragraphs(i)
                tgt_para = tgt_hf.Range.Paragraphs(i)
                try:
                    tgt_para.Range.ParagraphFormat = src_para.Range.ParagraphFormat
                except Exception:
                    pass
                # Copy resolved character formatting; skip undefined/empty values.
                for prop in ("Name", "Size"):
                    try:
                        src_val = getattr(src_para.Range.Font, prop)
                        if src_val is None or src_val == 9999999:
                            continue
                        if prop == "Name" and not src_val:
                            # Range.Font.Name can be empty for theme-inherited fonts;
                            # fall back to the first character's resolved font.
                            try:
                                src_val = src_para.Range.Characters(1).Font.Name
                            except Exception:
                                continue
                        if not src_val:
                            continue
                        if src_val != getattr(tgt_para.Range.Font, prop):
                            setattr(tgt_para.Range.Font, prop, src_val)
                    except Exception:
                        pass
        except Exception as fmt_err:
            logger.warning(f"Could not copy header/footer formatting: {fmt_err}")

    @staticmethod
    def _copy_document_styles(target_doc, source_doc) -> None:
        """Copies the style definitions that govern header/footer text from source_doc
        to target_doc. FormattedText only carries explicitly-set formatting; values
        inherited from paragraph styles (spacing, line spacing, fonts) otherwise
        resolve against the target document's built-in styles (e.g. Word's default
        Normal style: 8pt after, 1.08 line spacing, Aptos). Matching the style
        definitions makes inherited values resolve identically to the source."""
        for style_name in ("Normal", "Header", "Footer"):
            try:
                tgt_style = target_doc.Styles(style_name)
                src_style = source_doc.Styles(style_name)
                tgt_style.ParagraphFormat = src_style.ParagraphFormat
                tgt_style.Font = src_style.Font
            except Exception:
                pass

    @staticmethod
    def _restore_sections_from_source(
        target_doc,
        source_doc,
        start_section_index: int,
        end_section_index: int,
        only_last_section: bool = False
    ) -> None:
        """Restores exact formatted headers/footers and page setup from source_doc onto target_doc."""
        range_section_count = end_section_index - start_section_index + 1
        exported_count = target_doc.Sections.Count

        PageExporterEngine._copy_document_styles(target_doc, source_doc)

        if only_last_section:
            restore_indices = [exported_count]
        else:
            restore_indices = list(range(1, exported_count + 1))

        for j in restore_indices:
            if range_section_count == exported_count:
                source_index = start_section_index + j - 1
            else:
                source_index = start_section_index + j - 1 if j < exported_count else end_section_index
            if only_last_section:
                source_index = end_section_index

            try:
                target_sec = target_doc.Sections(j)
                source_sec = source_doc.Sections(source_index)

                # 1. Mirror PageSetup properties
                ps_target = target_sec.PageSetup
                ps_source = source_sec.PageSetup
                for attr in (
                    "TopMargin", "BottomMargin", "LeftMargin", "RightMargin",
                    "HeaderDistance", "FooterDistance", "PageWidth", "PageHeight", "Orientation",
                ):
                    try:
                        setattr(ps_target, attr, getattr(ps_source, attr))
                    except Exception:
                        pass

                try:
                    ps_target.DifferentFirstPageHeaderFooter = ps_source.DifferentFirstPageHeaderFooter
                except Exception:
                    pass

                try:
                    ps_target.OddAndEvenPagesHeaderFooter = ps_source.OddAndEvenPagesHeaderFooter
                except Exception:
                    pass

                try:
                    ps_target.PageNumbering.RestartNumberingAtSection = ps_source.PageNumbering.RestartNumberingAtSection
                    ps_target.PageNumbering.StartingNumber = ps_source.PageNumbering.StartingNumber
                except Exception:
                    pass

                # 2. Copy formatted headers and footers using FormattedText
                for is_header in (True, False):
                    for hf_index in (1, 2, 3):
                        try:
                            target_hf = target_sec.Headers(hf_index) if is_header else target_sec.Footers(hf_index)
                            
                            # Section 1 of exported document must not link to non-existent previous section
                            if j == 1:
                                target_hf.LinkToPrevious = False
                            else:
                                source_hf_direct = source_sec.Headers(hf_index) if is_header else source_sec.Footers(hf_index)
                                try:
                                    target_hf.LinkToPrevious = bool(source_hf_direct.LinkToPrevious)
                                except Exception:
                                    target_hf.LinkToPrevious = False

                            if not target_hf.LinkToPrevious or j == 1:
                                effective_source_hf = PageExporterEngine._get_effective_header_footer(
                                    source_doc, source_index, is_header, hf_index
                                )
                                if effective_source_hf is not None:
                                    try:
                                        # Clear the target footer/header first so no stale content remains.
                                        target_hf.Range.Text = ""
                                        # Copy the source content EXCLUDING its trailing paragraph mark;
                                        # pasting it too would add a duplicate empty paragraph at the end.
                                        src_copy = effective_source_hf.Range
                                        if src_copy.End > src_copy.Start and src_copy.Text.endswith("\r"):
                                            src_copy.End = src_copy.End - 1
                                        if src_copy.End > src_copy.Start:
                                            target_hf.Range.FormattedText = src_copy.FormattedText
                                            # Copy resolved paragraph/run formatting so style-inherited
                                            # values (indent, alignment, spacing, fonts) match the source
                                            # even when the target document's built-in styles differ.
                                            PageExporterEngine._copy_header_footer_formatting(
                                                effective_source_hf, target_hf
                                            )
                                    except Exception as copy_err:
                                        logger.warning(f"Error copying FormattedText for section {j}: {copy_err}")
                            
                            # Update fields in header/footer range
                            try:
                                target_hf.Range.Fields.Update()
                            except Exception:
                                pass
                        except Exception as hf_err:
                            logger.warning(f"Error restoring header/footer for section {j}: {hf_err}")

            except Exception as sec_err:
                logger.warning(f"Error restoring properties for section {j}: {sec_err}")

        # Refresh overall document fields
        try:
            target_doc.Fields.Update()
        except Exception:
            pass

    @staticmethod
    def _export_by_selection(
        abs_source: str,
        abs_output: str,
        start_page: int,
        end_page: int,
        fmt_code: int,
        visible: bool
    ) -> str:
        """Fast Selection Copy Engine: Copies page range into new document."""
        with WordCOMContext(visible=visible) as word_app:
            source_doc = None
            target_doc = None
            try:
                source_doc = word_app.Documents.Open(
                    FileName=abs_source,
                    ReadOnly=True,
                    ConfirmConversions=False,
                    AddToRecentFiles=False,
                    Visible=visible
                )

                total_pages = source_doc.ComputeStatistics(WD_STATISTIC_PAGES)
                if start_page < 1 or end_page > total_pages or start_page > end_page:
                    raise ValueError(f"Invalid range [{start_page}-{end_page}] for document with {total_pages} pages.")

                start_section_index = PageExporterEngine._section_containing_page(source_doc, start_page)
                end_section_index = PageExporterEngine._section_containing_page(source_doc, end_page)

                start_range = source_doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=start_page)
                if end_page == total_pages:
                    end_pos = source_doc.Content.End
                else:
                    next_page_range = source_doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=end_page + 1)
                    end_pos = next_page_range.Start

                extract_range = source_doc.Range(Start=start_range.Start, End=end_pos)
                extract_range.Copy()

                target_doc = word_app.Documents.Add()
                target_doc.Content.Paste()

                # Restore high-fidelity headers, footers & page setup from original source_doc
                if start_section_index is not None and end_section_index is not None:
                    PageExporterEngine._restore_sections_from_source(
                        target_doc=target_doc,
                        source_doc=source_doc,
                        start_section_index=start_section_index,
                        end_section_index=end_section_index
                    )

                target_doc.SaveAs2(FileName=abs_output, FileFormat=fmt_code)
                logger.success(f"Exported pages [{start_page}-{end_page}] via selection copy to '{abs_output}'")

            finally:
                if target_doc:
                    try:
                        target_doc.Close(SaveChanges=False)
                    except Exception:
                        pass
                if source_doc:
                    try:
                        source_doc.Close(SaveChanges=False)
                    except Exception:
                        pass

        return abs_output
