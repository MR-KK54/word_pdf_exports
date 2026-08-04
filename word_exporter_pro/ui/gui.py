"""
Word Page Exporter Pro - Graphical User Interface (GUI)
Modern Desktop UI built with CustomTkinter.
"""

import os
import sys
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from typing import List, Optional, Dict, Any

from word_exporter_pro.core.com_engine import DocumentInspector
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

        self.format_var = ctk.StringVar(value="docx")
        fmt_dropdown = ctk.CTkOptionMenu(
            output_card,
            variable=self.format_var,
            values=["docx", "pdf", "doc", "rtf", "docm"],
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
        self.overwrite_var = ctk.BooleanVar(value=False)
        overwrite_chk = ctk.CTkCheckBox(output_card, text="Overwrite existing files", variable=self.overwrite_var)
        overwrite_chk.grid(row=6, column=1, padx=15, pady=(5, 12), sticky="w")

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

        self.cancel_btn = ctk.CTkButton(
            act_btn_row, text="⏹ Cancel", width=90, height=40,
            fg_color="#C0392B", hover_color="#962D22", state="disabled",
            command=self._cancel_export
        )
        self.cancel_btn.pack(side="right")

        # Quick Inspector Card
        inspect_card = ctk.CTkFrame(right_panel)
        inspect_card.grid(row=1, column=0, sticky="ew", pady=(0, 15))
        inspect_card.grid_columnconfigure(0, weight=1)

        ins_header = ctk.CTkFrame(inspect_card, fg_color="transparent")
        ins_header.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 5))
        ins_header.grid_columnconfigure(0, weight=1)

        ins_title = ctk.CTkLabel(ins_header, text="Document Inspection", font=ctk.CTkFont(size=13, weight="bold"))
        ins_title.grid(row=0, column=0, sticky="w")

        inspect_btn = ctk.CTkButton(ins_header, text="🔍 Inspect Doc", width=100, height=24, command=self._inspect_selected)
        inspect_btn.grid(row=0, column=1, sticky="e")

        self.inspect_txt = ctk.CTkLabel(
            inspect_card, text="Select a document and click 'Inspect Doc' to detect Word rendering engine page counts.",
            font=ctk.CTkFont(size=11), text_color="gray60", justify="left", wraplength=380
        )
        self.inspect_txt.grid(row=1, column=0, padx=15, pady=(0, 10), sticky="w")

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
        else:
            for idx, f in enumerate(self.source_files, 1):
                sz = os.path.getsize(f) / 1024
                self.files_box.insert("end", f"{idx:02d}. {os.path.basename(f)} ({sz:.1f} KB)\n")
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

    def _inspect_selected(self):
        if not self.source_files:
            messagebox.showwarning("No Selection", "Please add at least one Word document first.")
            return

        first_file = self.source_files[0]
        self.inspect_txt.configure(text=f"Inspecting '{os.path.basename(first_file)}' via Word COM engine...")
        
        # Async inspect
        import threading
        def worker():
            try:
                info = DocumentInspector.get_info(first_file, visible=self.visible_var.get())
                text_res = (
                    f"📄 {info['filename']}\n"
                    f"Total Pages (Word COM): {info['page_count']}\n"
                    f"Sections: {info['section_count']} | Format: {info['format'].upper()}\n"
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
            visible=self.visible_var.get()
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
