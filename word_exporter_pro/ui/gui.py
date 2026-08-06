"""
Word Page Exporter Pro - Graphical User Interface (GUI)
Modern Desktop UI built with CustomTkinter.
"""

import os
import sys
import queue
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from typing import List, Optional, Dict, Any

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

from word_exporter_pro.core.com_engine import DocumentInspector
from word_exporter_pro.core.pdf_engine import PdfInspector
from word_exporter_pro.core.range_parser import PageRangeParser, RangeParseError
from word_exporter_pro.core.naming_formatter import NamingFormatter
from word_exporter_pro.core.batch_processor import BatchProcessor, ExportJobConfig
from word_exporter_pro.utils.logger import get_logger

logger = get_logger()

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class WordExporterApp(ctk.CTk):
    """Main Desktop Application Window."""

    def __init__(self):
        super().__init__()

        self.title("Microsoft Word & PDF Page Exporter Pro")
        self.geometry("1100x820")
        self.minsize(950, 700)

        # State Variables
        self.source_files: List[str] = []
        self.current_processor: Optional[BatchProcessor] = None
        self.ui_queue = queue.Queue()

        # Document Page Preview State
        self.preview_page_idx = 1
        self.preview_total_pages = 1
        self.current_preview_file = None
        self.current_ctk_image = None

        # Build UI layout
        self._init_layout()

        # Connect logger listener
        logger.add_listener(self._on_log_emitted)

        # Periodic queue check for thread-safe UI updates
        self.after(100, self._process_ui_queue)

    def _init_layout(self):
        # Main Grid Layout (2 columns: left controls, right logger/inspector)
        self.grid_columnconfigure(0, weight=1, minsize=520)
        self.grid_columnconfigure(1, weight=1, minsize=420)
        self.grid_rowconfigure(1, weight=1)

        # ------------------ HEADER ------------------
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray85", "gray14"))
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        header_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header_frame,
            text="📄 Microsoft Word Page Exporter Pro",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold")
        )
        title_label.grid(row=0, column=0, padx=20, pady=(12, 2), sticky="w")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="High-Fidelity Page Detection & Range Extraction Engine powered by MS Word Native COM",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color="gray60"
        )
        subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 12), sticky="w")

        theme_btn = ctk.CTkButton(
            header_frame,
            text="🌓 Theme",
            width=80,
            command=self._toggle_theme
        )
        theme_btn.grid(row=0, column=1, rowspan=2, padx=20, pady=10, sticky="e")

        # ------------------ LEFT PANEL (CONTROLS) ------------------
        left_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        left_scroll.grid(row=1, column=0, sticky="nsew", padx=15, pady=15)
        left_scroll.grid_columnconfigure(0, weight=1)

        # 1. File Selection Card
        files_card = ctk.CTkFrame(left_scroll)
        files_card.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        files_card.grid_columnconfigure(0, weight=1)

        card_title1 = ctk.CTkLabel(
            files_card, text="1. Select Input Word Documents",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        card_title1.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="w")

        btn_row = ctk.CTkFrame(files_card, fg_color="transparent")
        btn_row.grid(row=1, column=0, padx=15, pady=5, sticky="ew")

        add_files_btn = ctk.CTkButton(
            btn_row, text="+ Add Files", width=120, command=self._add_files
        )
        add_files_btn.pack(side="left", padx=(0, 10))

        add_folder_btn = ctk.CTkButton(
            btn_row, text="+ Add Folder", width=120, fg_color="gray30", hover_color="gray40",
            command=self._add_folder
        )
        add_folder_btn.pack(side="left", padx=(0, 10))

        clear_files_btn = ctk.CTkButton(
            btn_row, text="Clear All", width=90, fg_color="#C0392B", hover_color="#962D22",
            command=self._clear_files
        )
        clear_files_btn.pack(side="right")

        # Files scrollable list
        self.files_box = ctk.CTkTextbox(files_card, height=110, font=ctk.CTkFont(family="Consolas", size=11))
        self.files_box.grid(row=2, column=0, padx=15, pady=(5, 12), sticky="ew")
        self.files_box.insert("1.0", "No documents loaded. Click '+ Add Files' or '+ Add Folder' to begin.")
        self.files_box.configure(state="disabled")

        # 2. Page Range Configuration Card
        range_card = ctk.CTkFrame(left_scroll)
        range_card.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        range_card.grid_columnconfigure(1, weight=1)

        card_title2 = ctk.CTkLabel(
            range_card, text="2. Configure Page Export Range",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        card_title2.grid(row=0, column=0, columnspan=2, padx=15, pady=(12, 5), sticky="w")

        # Preset Dropdown
        preset_lbl = ctk.CTkLabel(range_card, text="Preset:")
        preset_lbl.grid(row=1, column=0, padx=15, pady=5, sticky="w")

        self.preset_var = ctk.StringVar(value="All Pages (Single Doc)")
        preset_dropdown = ctk.CTkOptionMenu(
            range_card,
            variable=self.preset_var,
            values=[
                "All Pages (Single Doc)",
                "Every Page as Separate Document",
                "Even Pages Only",
                "Odd Pages Only",
                "Custom Page Range"
            ],
            command=self._on_preset_change
        )
        preset_dropdown.grid(row=1, column=1, padx=15, pady=5, sticky="ew")

        # Range Expression Input
        range_lbl = ctk.CTkLabel(range_card, text="Range Spec:")
        range_lbl.grid(row=2, column=0, padx=15, pady=5, sticky="w")

        self.range_entry = ctk.CTkEntry(
            range_card, placeholder_text="e.g. 1-3, 5, 8-end"
        )
        self.range_entry.grid(row=2, column=1, padx=15, pady=5, sticky="ew")
        self.range_entry.insert(0, "1-end")
        self.range_entry.bind("<KeyRelease>", self._update_preview)

        hint_lbl = ctk.CTkLabel(
            range_card,
            text="Examples: '1-3', '1, 5, 8-10', '3-end', 'even', 'odd', 'all-individual'",
            font=ctk.CTkFont(size=10), text_color="gray60"
        )
        hint_lbl.grid(row=3, column=1, padx=15, pady=(0, 12), sticky="w")

        # 3. Target Output Options Card
        output_card = ctk.CTkFrame(left_scroll)
        output_card.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        output_card.grid_columnconfigure(1, weight=1)

        card_title3 = ctk.CTkLabel(
            output_card, text="3. Target Export Settings",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        card_title3.grid(row=0, column=0, columnspan=2, padx=15, pady=(12, 5), sticky="w")

        # Format Dropdown
        fmt_lbl = ctk.CTkLabel(output_card, text="Export Format:")
        fmt_lbl.grid(row=1, column=0, padx=15, pady=5, sticky="w")

        self.format_var = ctk.StringVar(value="same")
        fmt_dropdown = ctk.CTkOptionMenu(
            output_card,
            variable=self.format_var,
            values=["same", "docx", "pdf", "doc", "rtf", "docm"],
            command=lambda v: self._update_preview()
        )
        fmt_dropdown.grid(row=1, column=1, padx=15, pady=5, sticky="w")

        # Destination Directory
        dest_lbl = ctk.CTkLabel(output_card, text="Output Directory:")
        dest_lbl.grid(row=2, column=0, padx=15, pady=5, sticky="w")

        dest_frame = ctk.CTkFrame(output_card, fg_color="transparent")
        dest_frame.grid(row=2, column=1, padx=15, pady=5, sticky="ew")
        dest_frame.grid_columnconfigure(0, weight=1)

        self.dest_entry = ctk.CTkEntry(dest_frame)
        self.dest_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        default_out_dir = os.path.abspath("./exported_pages")
        self.dest_entry.insert(0, default_out_dir)

        browse_btn = ctk.CTkButton(dest_frame, text="Browse...", width=80, command=self._browse_output_dir)
        browse_btn.grid(row=0, column=1)

        clear_out_btn = ctk.CTkButton(dest_frame, text="Clear", width=60, fg_color="gray30", hover_color="gray40", command=self._clear_output_dir)
        clear_out_btn.grid(row=0, column=2, padx=(5, 0))

        # Naming Pattern
        name_lbl = ctk.CTkLabel(output_card, text="Naming Template:")
        name_lbl.grid(row=3, column=0, padx=15, pady=5, sticky="w")

        self.naming_entry = ctk.CTkEntry(output_card)
        self.naming_entry.grid(row=3, column=1, padx=15, pady=5, sticky="ew")
        self.naming_entry.insert(0, "{original_name}_pages_{start_page}-{end_page}")
        self.naming_entry.bind("<KeyRelease>", self._update_preview)

        # Token insert chips
        chip_frame = ctk.CTkFrame(output_card, fg_color="transparent")
        chip_frame.grid(row=4, column=1, padx=15, pady=(0, 5), sticky="w")

        chips = [("{original_name}", "+ Name"), ("{range}", "+ Range"), ("{start_page}", "+ Start"), ("{timestamp}", "+ Time")]
        for token, label in chips:
            btn = ctk.CTkButton(
                chip_frame, text=label, width=65, height=22, font=ctk.CTkFont(size=10),
                fg_color="gray25", hover_color="gray35",
                command=lambda t=token: self._insert_token(t)
            )
            btn.pack(side="left", padx=2)

        # Live Naming Preview
        preview_lbl = ctk.CTkLabel(output_card, text="Sample Filename:", text_color="gray60")
        preview_lbl.grid(row=5, column=0, padx=15, pady=5, sticky="w")

        self.preview_val_lbl = ctk.CTkLabel(
            output_card, text="...", font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
            text_color="#2ECC71"
        )
        self.preview_val_lbl.grid(row=5, column=1, padx=15, pady=5, sticky="w")

        # Overwrite option
        self.overwrite_var = ctk.BooleanVar(value=True)
        overwrite_chk = ctk.CTkCheckBox(output_card, text="Overwrite existing files", variable=self.overwrite_var)
        overwrite_chk.grid(row=6, column=1, padx=15, pady=(5, 2), sticky="w")

        self.clear_storage_var = ctk.BooleanVar(value=True)
        clear_storage_chk = ctk.CTkCheckBox(output_card, text="Auto-clear uploaded documents from server after export", variable=self.clear_storage_var)
        clear_storage_chk.grid(row=7, column=1, padx=15, pady=(2, 12), sticky="w")

        # 4. Engine & Advanced Settings
        adv_card = ctk.CTkFrame(left_scroll)
        adv_card.grid(row=3, column=0, sticky="ew", pady=(0, 15))
        adv_card.grid_columnconfigure(1, weight=1)

        card_title4 = ctk.CTkLabel(
            adv_card, text="4. Advanced Engine Options",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        card_title4.grid(row=0, column=0, columnspan=2, padx=15, pady=(12, 5), sticky="w")

        engine_lbl = ctk.CTkLabel(adv_card, text="Extraction Engine:")
        engine_lbl.grid(row=1, column=0, padx=15, pady=5, sticky="w")

        self.engine_var = ctk.StringVar(value="trimming")
        engine_opt = ctk.CTkOptionMenu(
            adv_card,
            variable=self.engine_var,
            values=["trimming", "aspose", "selection"]
        )
        engine_opt.grid(row=1, column=1, padx=15, pady=5, sticky="w")

        self.visible_var = ctk.BooleanVar(value=False)
        visible_chk = ctk.CTkCheckBox(adv_card, text="Show MS Word application window while running (Debug)", variable=self.visible_var)
        visible_chk.grid(row=2, column=1, padx=15, pady=(5, 12), sticky="w")

        # ------------------ RIGHT PANEL (STATUS & LOGS) ------------------
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=15, pady=15)
        right_panel.grid_rowconfigure(2, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)

        # Inspection & Action Box
        action_card = ctk.CTkFrame(right_panel)
        action_card.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        action_card.grid_columnconfigure(0, weight=1)

        action_title = ctk.CTkLabel(action_card, text="Export Control & Progress", font=ctk.CTkFont(size=14, weight="bold"))
        action_title.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="w")

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(action_card)
        self.progress_bar.grid(row=1, column=0, padx=15, pady=10, sticky="ew")
        self.progress_bar.set(0.0)

        self.status_lbl = ctk.CTkLabel(action_card, text="Ready to export.", text_color="gray70")
        self.status_lbl.grid(row=2, column=0, padx=15, pady=(0, 10), sticky="w")

        act_btn_row = ctk.CTkFrame(action_card, fg_color="transparent")
        act_btn_row.grid(row=3, column=0, padx=15, pady=(0, 15), sticky="ew")

        self.start_btn = ctk.CTkButton(
            act_btn_row, text="🚀 START EXPORT PRO", font=ctk.CTkFont(size=14, weight="bold"),
            height=40, fg_color="#27AE60", hover_color="#1E8449",
            command=self._start_export
        )
        self.start_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.open_output_btn = ctk.CTkButton(
            act_btn_row, text="📂 Open Output Folder", font=ctk.CTkFont(size=12, weight="bold"),
            height=40, fg_color="#2980B9", hover_color="#1F618D",
            command=self._open_output_dir
        )
        self.open_output_btn.pack(side="left", padx=(0, 10))

        self.cancel_btn = ctk.CTkButton(
            act_btn_row, text="⏹ Cancel", width=90, height=40,
            fg_color="#C0392B", hover_color="#962D22", state="disabled",
            command=self._cancel_export
        )
        self.cancel_btn.pack(side="right")

        # Quick Inspector Card with Live Page Image Preview
        inspect_card = ctk.CTkFrame(right_panel)
        inspect_card.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        inspect_card.grid_columnconfigure(0, weight=1)

        ins_header = ctk.CTkFrame(inspect_card, fg_color="transparent")
        ins_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 5))
        ins_header.grid_columnconfigure(0, weight=1)

        ins_title = ctk.CTkLabel(ins_header, text="Document Inspection & Live Page Preview", font=ctk.CTkFont(size=13, weight="bold"))
        ins_title.grid(row=0, column=0, sticky="w")

        inspect_btn = ctk.CTkButton(ins_header, text="🔍 Inspect Doc", width=100, height=24, command=self._inspect_selected)
        inspect_btn.grid(row=0, column=1, sticky="e")

        self.inspect_txt = ctk.CTkLabel(
            inspect_card, text="Select a document to inspect and view page previews.",
            font=ctk.CTkFont(size=11), text_color="gray60", justify="left", wraplength=380
        )
        self.inspect_txt.grid(row=1, column=0, padx=15, pady=(0, 5), sticky="w")

        # Live Image Preview Container Box
        self.preview_box = ctk.CTkFrame(inspect_card, fg_color=("gray90", "gray18"))
        self.preview_box.grid(row=2, column=0, padx=15, pady=(5, 12), sticky="ew")
        self.preview_box.grid_columnconfigure(0, weight=1)

        self.preview_img_lbl = ctk.CTkLabel(
            self.preview_box, text="📄 Load a document to display visual page preview", font=ctk.CTkFont(size=11), text_color="gray50"
        )
        self.preview_img_lbl.grid(row=0, column=0, padx=10, pady=10)

        # Page Navigation Row
        p_nav_row = ctk.CTkFrame(self.preview_box, fg_color="transparent")
        p_nav_row.grid(row=1, column=0, padx=10, pady=(0, 8))

        self.prev_p_btn = ctk.CTkButton(p_nav_row, text="◀ Prev Page", width=80, height=24, font=ctk.CTkFont(size=11), state="disabled", command=self._prev_preview_page)
        self.prev_p_btn.pack(side="left", padx=5)

        self.page_num_lbl = ctk.CTkLabel(p_nav_row, text="Page 1 / 1", font=ctk.CTkFont(size=11, weight="bold"))
        self.page_num_lbl.pack(side="left", padx=10)

        self.next_p_btn = ctk.CTkButton(p_nav_row, text="Next Page ▶", width=80, height=24, font=ctk.CTkFont(size=11), state="disabled", command=self._next_preview_page)
        self.next_p_btn.pack(side="left", padx=5)

        # Log Console Card
        log_card = ctk.CTkFrame(right_panel)
        log_card.grid(row=2, column=0, sticky="nsew")
        log_card.grid_rowconfigure(1, weight=1)
        log_card.grid_columnconfigure(0, weight=1)

        log_title_row = ctk.CTkFrame(log_card, fg_color="transparent")
        log_title_row.grid(row=0, column=0, padx=15, pady=(10, 5), sticky="ew")

        log_title = ctk.CTkLabel(log_title_row, text="Execution Logs", font=ctk.CTkFont(size=13, weight="bold"))
        log_title.pack(side="left")

        clear_log_btn = ctk.CTkButton(log_title_row, text="Clear", width=50, height=22, font=ctk.CTkFont(size=10), command=self._clear_logs)
        clear_log_btn.pack(side="right", padx=(5, 0))

        # Log Text Box
        self.log_textbox = ctk.CTkTextbox(log_card, font=ctk.CTkFont(family="Consolas", size=10), wrap="word")
        self.log_textbox.grid(row=1, column=0, padx=15, pady=(0, 12), sticky="nsew")
        self.log_textbox.configure(state="disabled")

        self._update_preview()

    def _toggle_theme(self):
        curr = ctk.get_appearance_mode()
        ctk.set_appearance_mode("Light" if curr == "Dark" else "Dark")

    def _add_files(self):
        files = filedialog.askopenfilenames(
            title="Select Word Documents",
            filetypes=[("Word Documents", "*.docx;*.doc;*.docm;*.dotx;*.dotm;*.rtf"), ("All Files", "*.*")]
        )
        if files:
            for f in files:
                abs_f = os.path.abspath(f)
                if abs_f not in self.source_files:
                    self.source_files.append(abs_f)
            self._refresh_files_box()

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Select Folder containing Word Documents")
        if folder:
            import glob
            matched = glob.glob(os.path.join(folder, "*.[dD][oO][cC]*"))
            for m in matched:
                if not os.path.basename(m).startswith("~$"):
                    abs_m = os.path.abspath(m)
                    if abs_m not in self.source_files:
                        self.source_files.append(abs_m)
            self._refresh_files_box()

    def _clear_files(self):
        self.source_files.clear()
        self._refresh_files_box()

    def _refresh_files_box(self):
        self.files_box.configure(state="normal")
        self.files_box.delete("1.0", "end")
        if not self.source_files:
            self.files_box.insert("1.0", "No documents loaded. Click '+ Add Files' or '+ Add Folder' to begin.")
            self.inspect_txt.configure(text="Select a document and click 'Inspect Doc' to detect page counts.")
        else:
            for idx, f in enumerate(self.source_files, 1):
                sz = os.path.getsize(f) / 1024
                self.files_box.insert("end", f"{idx:02d}. {os.path.basename(f)} ({sz:.1f} KB)\n")
            self._inspect_selected(quiet=True)
        self.files_box.configure(state="disabled")
        self._update_preview()

    def _on_preset_change(self, choice: str):
        if choice == "All Pages (Single Doc)":
            self.range_entry.delete(0, "end")
            self.range_entry.insert(0, "1-end")
            self.naming_entry.delete(0, "end")
            self.naming_entry.insert(0, "{original_name}_pages_{start_page}-{end_page}")
        elif choice == "Every Page as Separate Document":
            self.range_entry.delete(0, "end")
            self.range_entry.insert(0, "all-individual")
            self.naming_entry.delete(0, "end")
            self.naming_entry.insert(0, "{original_name}_page_{start_page:03d}")
        elif choice == "Even Pages Only":
            self.range_entry.delete(0, "end")
            self.range_entry.insert(0, "even")
        elif choice == "Odd Pages Only":
            self.range_entry.delete(0, "end")
            self.range_entry.insert(0, "odd")
        self._update_preview()

    def _insert_token(self, token: str):
        self.naming_entry.insert("end", token)
        self._update_preview()

    def _browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="Select Output Directory")
        if dir_path:
            self.dest_entry.delete(0, "end")
            self.dest_entry.insert(0, os.path.abspath(dir_path))

    def _clear_output_dir(self):
        self.dest_entry.delete(0, "end")

    def _update_preview(self, event=None):
        pattern = self.naming_entry.get().strip() or "{original_name}_pages_{start_page}-{end_page}"
        fmt = self.format_var.get().strip() or "docx"
        sample_path = self.source_files[0] if self.source_files else "SampleReport.docx"
        try:
            preview_filename = NamingFormatter.generate_filename(
                pattern=pattern,
                original_filepath=sample_path,
                page_range=(1, 3),
                total_pages=10,
                output_ext=fmt,
                batch_index=1
            )
            self.preview_val_lbl.configure(text=preview_filename, text_color="#2ECC71")
        except Exception as e:
            self.preview_val_lbl.configure(text=f"[Pattern Error: {e}]", text_color="#E74C3C")

    def _open_output_dir(self):
        out_dir = self.dest_entry.get().strip() or os.path.abspath("exported_pages")
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(out_dir)
        elif sys.platform == "darwin":
            subprocess.run(["open", out_dir])
        else:
            subprocess.run(["xdg-open", out_dir])

    def _prev_preview_page(self):
        if self.preview_page_idx > 1:
            self.preview_page_idx -= 1
            if self.current_preview_file:
                self._load_page_preview(self.current_preview_file, self.preview_page_idx)

    def _next_preview_page(self):
        if self.preview_page_idx < self.preview_total_pages:
            self.preview_page_idx += 1
            if self.current_preview_file:
                self._load_page_preview(self.current_preview_file, self.preview_page_idx)

    def _load_page_preview(self, file_path: str, page_num: int):
        if not file_path or not os.path.exists(file_path):
            return

        self.current_preview_file = file_path
        import threading

        def worker():
            try:
                ext = os.path.splitext(file_path)[1].lower()
                pil_img = None
                total_p = 1

                if ext == ".pdf":
                    doc = fitz.open(file_path)
                    total_p = len(doc)
                    page_idx = max(0, min(page_num - 1, total_p - 1))
                    page = doc[page_idx]
                    pix = page.get_pixmap(dpi=120)
                    pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    doc.close()
                else:
                    info = DocumentInspector.get_info(file_path, visible=False)
                    total_p = max(1, info.get("page_count", 1))
                    
                    # Render a clean document page graphic preview thumbnail
                    w, h = 320, 420
                    img = Image.new("RGB", (w, h), (245, 247, 250))
                    draw = ImageDraw.Draw(img)
                    
                    draw.rectangle([10, 10, w - 10, h - 10], fill=(255, 255, 255), outline=(200, 205, 215), width=2)
                    draw.rectangle([20, 25, w - 20, 65], fill=(79, 70, 229))
                    
                    fn = os.path.basename(file_path)
                    if len(fn) > 28:
                        fn = fn[:25] + "..."
                    draw.text((30, 36), fn, fill=(255, 255, 255))
                    
                    for i in range(12):
                        y = 90 + (i * 24)
                        draw.line([30, y, w - 30, y], fill=(210, 215, 225), width=4)
                    
                    draw.text((w // 2 - 40, h - 35), f"Page {page_num} of {total_p}", fill=(100, 110, 125))
                    pil_img = img

                if pil_img:
                    target_h = 210
                    ratio = target_h / float(pil_img.height)
                    target_w = max(140, int(pil_img.width * ratio))
                    pil_img = pil_img.resize((target_w, target_h), Image.Resampling.LANCZOS)
                    
                    ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(target_w, target_h))
                    self.ui_queue.put(("preview_image", (ctk_img, page_num, total_p)))

            except Exception as e:
                logger.error(f"Error loading page preview: {e}")

        threading.Thread(target=worker, daemon=True).start()

    def _inspect_selected(self, quiet: bool = False):
        if not self.source_files:
            if not quiet:
                messagebox.showwarning("No Selection", "Please add at least one Word document or PDF first.")
            return

        first_file = self.source_files[0]
        self.inspect_txt.configure(text=f"Inspecting '{os.path.basename(first_file)}'...")
        self.preview_page_idx = 1
        self._load_page_preview(first_file, 1)
        
        # Async inspect
        import threading
        def worker():
            try:
                ext = os.path.splitext(first_file)[1].lower()
                if ext == ".pdf":
                    info = PdfInspector.get_info(first_file)
                else:
                    info = DocumentInspector.get_info(first_file, visible=self.visible_var.get())
                text_res = (
                    f"📄 {info['filename']}\n"
                    f"Total Pages: {info['page_count']}\n"
                    f"Sections: {info.get('section_count', 1)} | Format: {info['format'].upper()}\n"
                    f"Size: {info['size_bytes']/1024:.1f} KB"
                )
                self.ui_queue.put(("inspect_res", text_res))
            except Exception as e:
                self.ui_queue.put(("inspect_res", f"Error inspecting document: {e}"))

        threading.Thread(target=worker, daemon=True).start()

    def _start_export(self):
        if not self.source_files:
            messagebox.showwarning("No Documents", "Please add one or more Word documents before starting export.")
            return

        range_str = self.range_entry.get().strip()
        if not range_str:
            messagebox.showwarning("Empty Range", "Please enter a valid page range specification.")
            return

        out_dir = self.dest_entry.get().strip()
        if not out_dir:
            messagebox.showwarning("No Output Directory", "Please specify an output folder.")
            return

        config = ExportJobConfig(
            source_files=list(self.source_files),
            range_expression=range_str,
            output_dir=os.path.abspath(out_dir),
            export_format=self.format_var.get(),
            naming_pattern=self.naming_entry.get().strip(),
            overwrite=self.overwrite_var.get(),
            engine_mode=self.engine_var.get(),
            visible=self.visible_var.get(),
            clear_storage_after_export=self.clear_storage_var.get()
        )

        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.progress_bar.set(0.0)
        self.status_lbl.configure(text="Initializing Word COM Engine...")

        self.current_processor = BatchProcessor(config)
        
        def on_progress(completed: int, total: int, filename: str, status: str):
            self.ui_queue.put(("progress", (completed, total, filename, status)))

        def on_finished(success: int, fail: int, errors: List[str]):
            self.ui_queue.put(("finished", (success, fail, errors)))

        self.current_processor.start_async(on_progress, on_finished)

    def _cancel_export(self):
        if self.current_processor:
            self.current_processor.cancel()
            self.status_lbl.configure(text="Cancelling job...")

    def _on_log_emitted(self, timestamp: str, level: str, message: str):
        self.ui_queue.put(("log", (timestamp, level, message)))

    def _clear_logs(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def _process_ui_queue(self):
        while not self.ui_queue.empty():
            try:
                event_type, payload = self.ui_queue.get_nowait()

                if event_type == "inspect_res":
                    self.inspect_txt.configure(text=payload)

                elif event_type == "preview_image":
                    ctk_img, page_num, total_p = payload
                    self.current_ctk_image = ctk_img
                    self.preview_total_pages = total_p
                    self.preview_page_idx = page_num
                    self.preview_img_lbl.configure(image=ctk_img, text="")
                    self.page_num_lbl.configure(text=f"Page {page_num} / {total_p}")
                    self.prev_p_btn.configure(state="normal" if page_num > 1 else "disabled")
                    self.next_p_btn.configure(state="normal" if page_num < total_p else "disabled")

                elif event_type == "progress":
                    completed, total, filename, status = payload
                    pct = (completed / total) if total > 0 else 0.0
                    self.progress_bar.set(pct)
                    self.status_lbl.configure(text=f"[{completed}/{total}] {status}")

                elif event_type == "finished":
                    success, fail, errors = payload
                    self.start_btn.configure(state="normal")
                    self.cancel_btn.configure(state="disabled")
                    self.progress_bar.set(1.0 if fail == 0 else self.progress_bar.get())
                    self.status_lbl.configure(text=f"Finished: {success} succeeded, {fail} failed.")
                    if fail == 0 and success > 0:
                        self.open_output_btn.configure(fg_color="#27AE60", text="📂 Open Output Folder (Saved)")
                        messagebox.showinfo("Export Complete", f"Successfully exported {success} page document(s) to:\n{self.dest_entry.get()}")

                elif event_type == "log":
                    timestamp, level, msg = payload
                    self.log_textbox.configure(state="normal")
                    prefix = f"[{timestamp}] [{level}] "
                    self.log_textbox.insert("end", prefix + msg + "\n")
                    self.log_textbox.see("end")
                    self.log_textbox.configure(state="disabled")

            except Exception as e:
                print(f"Error handling UI queue: {e}")

        self.after(100, self._process_ui_queue)


def main():
    app = WordExporterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
