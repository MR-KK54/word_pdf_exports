# Microsoft Word & PDF Page Exporter Pro

**Microsoft Word & PDF Page Exporter Pro** is a production-ready application designed for high-fidelity page detection and extraction from Microsoft Word documents and PDF files. Word documents are powered directly by the native Microsoft Word COM rendering engine (`Word.Application`), guaranteeing 100% layout, style, margin, header/footer, and section structure preservation. PDFs are processed natively via PyMuPDF.

---

## Key Features

- 📑 **Comprehensive Format Support**: Opens and processes `.doc`, `.docx`, `.docm`, `.dotx`, `.dotm`, `.rtf`, and `.pdf` files.
- 🎯 **Native Word Pagination Engine**: Uses MS Word's exact layout engine to calculate precise page boundaries.
- 📄 **PDF Page Extraction**: Splits and extracts page ranges from PDF files via PyMuPDF.
- ✂️ **Flexible Page Extraction**:
  - Export single pages (e.g. `5`).
  - Export page ranges (e.g. `1-3`, `5-8`).
  - Export non-contiguous ranges (e.g. `1-3, 5, 8-12, 15-end`).
  - Quick Presets: `All Pages (Single Doc)`, `Every Page as Separate Document`, `Even Pages`, `Odd Pages`.
- 🛡️ **Highest Fidelity Preservation (Trimming Engine)**: Duplicates source documents and trims content outside the target range natively, preserving all section breaks, page geometry, headers/footers, styles, floating objects, footnotes, and TOC links.
- 📁 **Multiple Output Formats**: Save extracted pages as `.docx`, `.pdf`, `.doc`, `.rtf`, or `.docm` (PDF inputs always export as PDF).
- 🔤 **Dynamic Naming Templates**: Custom filename patterns with tokens like `{original_name}`, `{range}`, `{start_page}`, `{end_page}`, `{timestamp}`, `{index:03d}`, and auto-collision resolution.
- 🖥️ **Dual Interface**:
  - **Modern Desktop GUI**: Built with CustomTkinter (Dark/Light mode, live logging, drag-and-drop file queue, page count inspector, progress bar).
  - **Scriptable CLI**: Full command-line interface for batch script integration and automated pipelines.
- ⚡ **Batch Processing & Progress Tracking**: Multi-threaded execution keeping the UI responsive with real-time status reporting.
- 🪵 **Structured Logging**: Timed console logs, UI logger console, and log files.

---

## Installation

### Prerequisites
- Windows OS
- Microsoft Word (Office 2016, 2019, 2021, or Microsoft 365)
- Python 3.9+

### Install Dependencies
```powershell
pip install -r requirements.txt
```

---

## Quick Start

### 1. Graphical User Interface (GUI)
Launch the Desktop Application:
```powershell
python run_gui.py
```

#### GUI Workflow:
1. Click **+ Add Files** or **+ Add Folder** to load target Word documents.
2. Click **🔍 Inspect Doc** to inspect Word COM page statistics.
3. Choose a **Preset** or enter a **Range Spec** (e.g. `1-3, 5, 8-end`).
4. Select the target **Export Format** (e.g. `docx`, `pdf`) and **Output Directory**.
5. Customize the **Naming Template** (e.g. `{original_name}_pages_{start_page}-{end_page}`).
6. Click **🚀 START EXPORT PRO**.

---

### 2. Command Line Interface (CLI)

#### View Help & Available Flags
```powershell
python run_cli.py --help
```

#### Inspect Document Page Statistics
```powershell
python run_cli.py -i "Report.docx" --inspect
```

#### Export Pages 1 to 5 to PDF
```powershell
python run_cli.py -i "Report.docx" -r "1-5" -o "./output" -f pdf
```

#### Export Non-Contiguous Ranges
```powershell
python run_cli.py -i "AnnualReport.docx" -r "1-3, 7, 12-end" -o "./exported"
```

#### Split Every Page into Separate DOCX Sub-Documents
```powershell
python run_cli.py -i "Contract.docx" -r "all-individual" -o "./split_pages"
```

#### Batch Export All Documents in a Folder
```powershell
python run_cli.py -b "./documents/*.docx" -r "1-2" -o "./summaries" -f pdf
```

---

### 3. Web Interface (Flask)

Launch the browser interface (local web server):
```powershell
pip install -r requirements.txt
python run_web.py
```

Open `http://127.0.0.1:8000` in your browser. The web UI mirrors the desktop GUI:

1. **Upload** Word documents or PDFs via drag-and-drop or the file picker.
2. Click **Inspect** to detect page counts (Word COM for `.doc*`/`.rtf`, PyMuPDF for `.pdf`).
3. Choose a **Preset** or enter a **Range Spec** (e.g. `1-3, 5, 8-end`).
4. Select the **Export Format** and **Naming Template** (with live filename preview). PDF inputs always export as PDF.
5. Click **START EXPORT** and watch the progress bar, live log console, and Save buttons for each produced file.

Notes:
- On Windows, Microsoft Word is used when available. On Render, Word documents are processed with Aspose.Words and the deployment image includes Calibri-compatible and Arial-compatible fonts. This is required for stable cloud page boundaries; exact parity with a desktop still requires the original document fonts to be installed or embedded. PDF files are processed with PyMuPDF.
- Render deploys this project using the included `Dockerfile`. Do not switch the service to Render's native Python runtime, because it does not install the font packages used for Word pagination.
- Uploaded files and exports are kept in a server-temporary directory and are downloaded with the Save buttons in the Results panel. The browser cannot write directly to a folder on the user's computer.
- Hosted services have ephemeral storage: download exports before the service restarts. Set `WORD_EXPORTER_WEB_DATA_DIR` only when you provide persistent server storage.

---

## Architecture & Code Structure

```
WordPageExporterPro/
├── run_gui.py                         # GUI Application Launcher
├── run_cli.py                         # CLI Application Launcher
├── run_web.py                         # Web Interface Launcher (Flask)
├── requirements.txt                   # Dependency Specification
├── README.md                          # Documentation
├── word_exporter_pro/
│   ├── core/
│   │   ├── com_engine.py              # MS Word COM Lifecycle, Document Inspector & Trimming Engine
│   │   ├── pdf_engine.py              # PDF Inspection & Page Extraction (PyMuPDF)
│   │   ├── range_parser.py            # Page Range Syntax Parser (1-3, 5, 8-end, even, odd)
│   │   ├── naming_formatter.py        # Dynamic Output Filename Token Formatter
│   │   └── batch_processor.py         # Multi-Threaded Batch Worker & Progress System
│   ├── utils/
│   │   └── logger.py                  # Structured Logger & Event Listener
│   ├── ui/
│   │   └── gui.py                     # CustomTkinter Desktop User Interface
│   └── web/
│       ├── app.py                     # Flask application & REST API
│       ├── job_manager.py             # Background job tracking & progress/log state
│       ├── templates/index.html       # Web UI page
│       └── static/                    # CSS & JavaScript for the web UI
└── tests/
    ├── test_range_parser.py           # Unit tests for range parser
    ├── test_naming_formatter.py       # Unit tests for filename formatter
    ├── test_pdf_engine.py             # Unit tests for PDF engine
    └── test_com_engine.py             # Integration tests for COM Word pagination & extraction
```

---

## Running Automated Tests

Run the test suite via pytest:
```powershell
python -m pytest tests/ -v
```

---

## License & Credits
Built by Antigravity Team. Powered by Microsoft Word COM Automation & CustomTkinter.
