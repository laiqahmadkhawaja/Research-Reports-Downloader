"""
gui.py - Modern CustomTkinter desktop GUI for UK FCA NSM Annual Report Downloader.
Provides responsive controls, live KPI metrics, company-wise organization, threaded execution,
real-time log console, and audit Excel reporting.
"""
import os
import sys
import threading
import queue
import subprocess
import time
from datetime import datetime
from typing import Optional, List, Dict, Any
import customtkinter as ctk
from tkinter import filedialog, messagebox

from config import (
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_RESULTS_FILE,
    DEFAULT_STATE_FILE,
    DEFAULT_LOG_FILE,
    DEFAULT_REQUEST_DELAY,
    DEFAULT_MAX_RETRIES,
    DEFAULT_SKIP_EXISTING,
    DEFAULT_STRICT_PDF_ONLY,
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_ANNUAL_REPORT,
    STATUS_SKIPPED_NO_PDF,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_NETWORK_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_PDF,
    STATUS_REVIEW_REQUIRED
)
from logger import get_logger, set_gui_callback
from excel_reader import inspect_file_headers, detect_columns, load_input_records
from progress_manager import ProgressManager
from results_writer import ResultsWriter
from nsm_client import NsmApiClient
from filing_matcher import find_best_filing_for_year
from downloader import download_pdf_report

logger = get_logger()

# Set Theme & Appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class UkNsmDownloaderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("UK FCA NSM Annual Report Downloader & Audit System")
        self.geometry("1120, 880")
        self.minsize(980, 760)

        # State & Managers
        self.progress_mgr = ProgressManager(DEFAULT_STATE_FILE)
        self.results_writer = ResultsWriter(DEFAULT_RESULTS_FILE)
        self.worker_thread: Optional[threading.Thread] = None
        self.log_queue = queue.Queue()
        self.input_records: List[Dict[str, Any]] = []
        self.processed_records_list: List[Dict[str, Any]] = self.progress_mgr.get_all_saved_records()

        # Thread Control Flags
        self._is_running = False
        self._is_paused = False
        self._pause_event = threading.Event()
        self._pause_event.set()

        # Variables
        self.excel_path_var = ctk.StringVar()
        self.output_dir_var = ctk.StringVar(value=str(DEFAULT_DOWNLOAD_DIR))
        self.strict_pdf_var = ctk.BooleanVar(value=DEFAULT_STRICT_PDF_ONLY)
        self.skip_existing_var = ctk.BooleanVar(value=DEFAULT_SKIP_EXISTING)
        self.retry_failed_var = ctk.BooleanVar(value=True)
        self.max_retries_var = ctk.IntVar(value=DEFAULT_MAX_RETRIES)
        self.delay_var = ctk.DoubleVar(value=DEFAULT_REQUEST_DELAY)

        # Progress & Status Variables
        self.status_msg_var = ctk.StringVar(value="Ready. Please load input Excel/CSV file.")
        self.current_company_var = ctk.StringVar(value="-")
        self.current_year_var = ctk.StringVar(value="-")
        self.current_code_var = ctk.StringVar(value="-")
        self.total_records_var = ctk.StringVar(value="0")
        self.processed_records_var = ctk.StringVar(value="0")
        self.success_count_var = ctk.StringVar(value="0")
        self.skipped_count_var = ctk.StringVar(value="0")
        self.failed_count_var = ctk.StringVar(value="0")
        self.review_count_var = ctk.StringVar(value="0")
        self.eta_var = ctk.StringVar(value="ETA: --:--")

        # Connect logger callback to queue
        set_gui_callback(self._queue_log_message)

        # Build UI
        self._build_ui()

        # Start periodic log drain and status poller
        self.after(100, self._drain_log_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Check for sample file
        sample_path = os.path.join(os.path.dirname(__file__), "input", "uk_companies.xlsx")
        if not os.path.exists(sample_path):
            sample_path = os.path.join(os.path.dirname(__file__), "input", "uk_companies.csv")
        if os.path.exists(sample_path):
            self.excel_path_var.set(sample_path)
            self._load_file_preview(sample_path)

        # Restore existing progress metrics if any
        metrics = self.progress_mgr.get_metrics()
        if metrics["processed"] > 0:
            self._update_metrics_ui(metrics)

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)  # Log section expands

        # =========================================================================
        # 1. Header Banner
        # =========================================================================
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1E3A8A", "#0F172A"), height=65)
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 10))
        header_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header_frame,
            text="UK FCA National Storage Mechanism (NSM) - Annual Report Downloader",
            font=ctk.CTkFont(size=19, weight="bold"),
            text_color="#FFFFFF"
        )
        title_label.grid(row=0, column=0, padx=20, pady=(8, 2), sticky="w")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Official Source: data.fca.org.uk/#/nsm/nationalstoragemechanism  |  Company-wise Folders & UniCode Naming",
            font=ctk.CTkFont(size=11),
            text_color="#94A3B8"
        )
        subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 8), sticky="w")

        # =========================================================================
        # 2. Input & Output Selection Frame
        # =========================================================================
        input_frame = ctk.CTkFrame(self, corner_radius=8)
        input_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=4)
        input_frame.grid_columnconfigure(1, weight=1)

        # Excel File Row
        ctk.CTkLabel(input_frame, text="Input Excel / CSV:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=12, pady=6, sticky="w"
        )
        excel_entry = ctk.CTkEntry(input_frame, textvariable=self.excel_path_var, placeholder_text="Select input file...")
        excel_entry.grid(row=0, column=1, padx=8, pady=6, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse File", width=110, command=self._browse_excel).grid(
            row=0, column=2, padx=12, pady=6
        )

        # Output Folder Row
        ctk.CTkLabel(input_frame, text="Output Folder:", font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=12, pady=6, sticky="w"
        )
        out_entry = ctk.CTkEntry(input_frame, textvariable=self.output_dir_var)
        out_entry.grid(row=1, column=1, padx=8, pady=6, sticky="ew")

        out_btn_frame = ctk.CTkFrame(input_frame, fg_color="transparent")
        out_btn_frame.grid(row=1, column=2, padx=12, pady=6, sticky="e")
        ctk.CTkButton(out_btn_frame, text="Browse", width=65, command=self._browse_output_dir).pack(side="left", padx=2)
        ctk.CTkButton(out_btn_frame, text="Open", width=55, fg_color="#475569", hover_color="#334155", command=self._open_output_dir).pack(side="left", padx=2)

        # Preview Banner
        self.preview_label = ctk.CTkLabel(
            input_frame,
            text="No file loaded.",
            font=ctk.CTkFont(size=11),
            text_color="#64748B"
        )
        self.preview_label.grid(row=2, column=0, columnspan=3, padx=12, pady=(2, 6), sticky="w")

        # =========================================================================
        # 3. Settings & Options Frame
        # =========================================================================
        config_frame = ctk.CTkFrame(self, corner_radius=8)
        config_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=4)
        config_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Switches Row
        pdf_switch = ctk.CTkSwitch(
            config_frame,
            text="Strict PDF Only (Skip Non-PDF)",
            variable=self.strict_pdf_var,
            onvalue=True,
            offvalue=False
        )
        pdf_switch.grid(row=0, column=0, padx=15, pady=8, sticky="w")

        skip_switch = ctk.CTkSwitch(
            config_frame,
            text="Skip Already Downloaded",
            variable=self.skip_existing_var
        )
        skip_switch.grid(row=0, column=1, padx=15, pady=8, sticky="w")

        retry_switch = ctk.CTkSwitch(
            config_frame,
            text="Retry Failed Records",
            variable=self.retry_failed_var
        )
        retry_switch.grid(row=0, column=2, padx=15, pady=8, sticky="w")

        # Rate Limiting & Retries Sliders
        sliders_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        sliders_frame.grid(row=1, column=0, columnspan=4, padx=15, pady=4, sticky="ew")
        sliders_frame.grid_columnconfigure((0, 1), weight=1)

        # Delay Slider
        delay_subframe = ctk.CTkFrame(sliders_frame, fg_color="transparent")
        delay_subframe.grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(delay_subframe, text="Request Delay:").pack(side="left", padx=(0, 6))
        self.delay_slider = ctk.CTkSlider(
            delay_subframe, from_=0.1, to=3.0, number_of_steps=29, variable=self.delay_var, width=150
        )
        self.delay_slider.pack(side="left", padx=4)
        self.delay_label = ctk.CTkLabel(delay_subframe, text=f"{self.delay_var.get():.1f}s")
        self.delay_label.pack(side="left", padx=4)
        self.delay_slider.configure(command=lambda val: self.delay_label.configure(text=f"{float(val):.1f}s"))

        # Max Retries Slider
        retries_subframe = ctk.CTkFrame(sliders_frame, fg_color="transparent")
        retries_subframe.grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(retries_subframe, text="Max Retries:").pack(side="left", padx=(0, 6))
        self.retries_slider = ctk.CTkSlider(
            retries_subframe, from_=1, to=5, number_of_steps=4, variable=self.max_retries_var, width=120
        )
        self.retries_slider.pack(side="left", padx=4)
        self.retries_label = ctk.CTkLabel(retries_subframe, text=f"{self.max_retries_var.get()}")
        self.retries_label.pack(side="left", padx=4)
        self.retries_slider.configure(command=lambda val: self.retries_label.configure(text=f"{int(float(val))}"))

        # =========================================================================
        # 4. Progress, Controls, and Metrics Frame
        # =========================================================================
        progress_frame = ctk.CTkFrame(self, corner_radius=8)
        progress_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=4)
        progress_frame.grid_columnconfigure(0, weight=1)

        # Control Action Buttons Row
        ctrl_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        ctrl_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        ctrl_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_start = ctk.CTkButton(
            ctrl_frame, text="▶ Start Download", fg_color="#10B981", hover_color="#059669",
            font=ctk.CTkFont(weight="bold", size=13), height=36, command=self._start_processing
        )
        self.btn_start.grid(row=0, column=0, padx=5, pady=4, sticky="ew")

        self.btn_pause = ctk.CTkButton(
            ctrl_frame, text="⏸ Pause", fg_color="#F59E0B", hover_color="#D97706",
            font=ctk.CTkFont(weight="bold", size=13), height=36, state="disabled", command=self._pause_processing
        )
        self.btn_pause.grid(row=0, column=1, padx=5, pady=4, sticky="ew")

        self.btn_resume = ctk.CTkButton(
            ctrl_frame, text="⏯ Resume", fg_color="#3B82F6", hover_color="#2563EB",
            font=ctk.CTkFont(weight="bold", size=13), height=36, state="disabled", command=self._resume_processing
        )
        self.btn_resume.grid(row=0, column=2, padx=5, pady=4, sticky="ew")

        self.btn_stop = ctk.CTkButton(
            ctrl_frame, text="⏹ Stop", fg_color="#EF4444", hover_color="#DC2626",
            font=ctk.CTkFont(weight="bold", size=13), height=36, state="disabled", command=self._stop_processing
        )
        self.btn_stop.grid(row=0, column=3, padx=5, pady=4, sticky="ew")

        # Progress Bar & Percentage
        pbar_line = ctk.CTkFrame(progress_frame, fg_color="transparent")
        pbar_line.grid(row=1, column=0, sticky="ew", padx=12, pady=(4, 2))
        pbar_line.grid_columnconfigure(0, weight=1)

        self.progress_bar = ctk.CTkProgressBar(pbar_line, height=14)
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.progress_bar.set(0.0)

        self.lbl_pct = ctk.CTkLabel(pbar_line, text="0.0%", font=ctk.CTkFont(size=12, weight="bold"), width=50)
        self.lbl_pct.grid(row=0, column=1)

        # Live Status & Currently Processing Badge
        status_line = ctk.CTkFrame(progress_frame, fg_color="transparent")
        status_line.grid(row=2, column=0, sticky="ew", padx=12, pady=(2, 6))
        status_line.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(status_line, textvariable=self.status_msg_var, font=ctk.CTkFont(size=11))
        self.lbl_status.grid(row=0, column=0, sticky="w")

        curr_badge = ctk.CTkFrame(status_line, fg_color=("#1E293B", "#1E293B"), corner_radius=6)
        curr_badge.grid(row=0, column=1, sticky="e")
        ctk.CTkLabel(curr_badge, text=" Processing: ", font=ctk.CTkFont(size=10, weight="bold"), text_color="#94A3B8").pack(side="left")
        ctk.CTkLabel(curr_badge, textvariable=self.current_company_var, font=ctk.CTkFont(size=10, weight="bold"), text_color="#38BDF8").pack(side="left", padx=2)
        ctk.CTkLabel(curr_badge, text="| Year:", font=ctk.CTkFont(size=10), text_color="#94A3B8").pack(side="left", padx=2)
        ctk.CTkLabel(curr_badge, textvariable=self.current_year_var, font=ctk.CTkFont(size=10, weight="bold"), text_color="#FDE047").pack(side="left", padx=2)
        ctk.CTkLabel(curr_badge, text="| Code:", font=ctk.CTkFont(size=10), text_color="#94A3B8").pack(side="left", padx=2)
        ctk.CTkLabel(curr_badge, textvariable=self.current_code_var, font=ctk.CTkFont(size=10, weight="bold"), text_color="#4ADE80").pack(side="left", padx=(2, 6))

        # KPI Metric Cards Frame
        kpi_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        kpi_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(2, 8))
        kpi_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        self._create_metric_card(kpi_frame, 0, "Total Records", self.total_records_var, "#3B82F6")
        self._create_metric_card(kpi_frame, 1, "Processed", self.processed_records_var, "#8B5CF6")
        self._create_metric_card(kpi_frame, 2, "Success (PDF)", self.success_count_var, "#10B981")
        self._create_metric_card(kpi_frame, 3, "Skipped", self.skipped_count_var, "#0EA5E9")
        self._create_metric_card(kpi_frame, 4, "Failed", self.failed_count_var, "#EF4444")
        self._create_metric_card(kpi_frame, 5, "Review Needed", self.review_count_var, "#F59E0B")

        # =========================================================================
        # 5. Log Console Section
        # =========================================================================
        log_frame = ctk.CTkFrame(self, corner_radius=8)
        log_frame.grid(row=4, column=0, sticky="nsew", padx=15, pady=(4, 12))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        # Log Header with Quick Action Buttons
        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 2))
        log_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(log_header, text="Execution & Activity Log:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w")

        btn_box = ctk.CTkFrame(log_header, fg_color="transparent")
        btn_box.grid(row=0, column=1, sticky="e")

        self.auto_scroll_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(btn_box, text="Auto-scroll", variable=self.auto_scroll_var, width=80).pack(side="left", padx=6)
        ctk.CTkButton(btn_box, text="Open Results Excel", width=125, fg_color="#1E3A8A", hover_color="#1E40AF", command=self._open_results_file).pack(side="left", padx=4)
        ctk.CTkButton(btn_box, text="Open Log File", width=95, fg_color="#334155", hover_color="#1E293B", command=self._open_log_file).pack(side="left", padx=4)
        ctk.CTkButton(btn_box, text="Clear Log", width=75, fg_color="#475569", hover_color="#334155", command=self._clear_log).pack(side="left", padx=4)

        # Log Text Box
        self.log_textbox = ctk.CTkTextbox(log_frame, font=ctk.CTkFont(family="Consolas", size=11), wrap="word")
        self.log_textbox.grid(row=1, column=0, sticky="nsew", padx=10, pady=(2, 8))

        # Color Tags in Textbox
        self.log_textbox.tag_config("INFO", foreground="#CBD5E1")
        self.log_textbox.tag_config("SUCCESS", foreground="#4ADE80")
        self.log_textbox.tag_config("WARNING", foreground="#FDE047")
        self.log_textbox.tag_config("ERROR", foreground="#F87171")
        self.log_textbox.tag_config("HIGHLIGHT", foreground="#38BDF8")

    def _create_metric_card(self, parent, col, title, text_var, accent_color):
        card = ctk.CTkFrame(parent, corner_radius=6, border_width=1, border_color=accent_color)
        card.grid(row=0, column=col, padx=4, pady=2, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=10, weight="bold"), text_color="#94A3B8").grid(
            row=0, column=0, padx=4, pady=(4, 0)
        )
        ctk.CTkLabel(card, textvariable=text_var, font=ctk.CTkFont(size=16, weight="bold"), text_color=accent_color).grid(
            row=1, column=0, padx=4, pady=(0, 4)
        )

    # =========================================================================
    # File & Folder Browser Callbacks
    # =========================================================================
    def _browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Input Excel or CSV File",
            filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv"), ("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self._load_file_preview(file_path)

    def _load_file_preview(self, file_path: str):
        try:
            records, mapping = load_input_records(file_path)
            self.input_records = records
            total = len(records)
            self.total_records_var.set(str(total))

            comp_col = mapping.get("Company Name", "firm_name")
            year_col = mapping.get("Year", "year")
            code_col = mapping.get("UniCode", "uni_code")

            unique_firms = len(set(r["firm_name"] for r in records))
            self.preview_label.configure(
                text=f"Loaded {total} records across {unique_firms} companies. Mappings: Company='{comp_col}', Year='{year_col}', UniCode='{code_col}'",
                text_color="#10B981"
            )
            self.status_msg_var.set(f"Loaded {total} records. Ready to begin downloading.")
            logger.info(f"Loaded input file: {file_path} ({total} records, {unique_firms} companies)")
        except Exception as e:
            self.preview_label.configure(text=f"Error reading file: {str(e)}", text_color="#EF4444")
            logger.error(f"Error inspecting input file: {e}")

    def _browse_output_dir(self):
        folder = filedialog.askdirectory(title="Select Output Directory for Downloaded Reports")
        if folder:
            self.output_dir_var.set(folder)

    def _open_output_dir(self):
        out_dir = self.output_dir_var.get()
        if os.path.exists(out_dir):
            if sys.platform == "win32":
                os.startfile(out_dir)
            else:
                subprocess.Popen(["xdg-open", out_dir])
        else:
            messagebox.showinfo("Folder Notice", f"Output folder does not exist yet:\n{out_dir}")

    def _open_results_file(self):
        res_file = str(DEFAULT_RESULTS_FILE)
        if os.path.exists(res_file):
            if sys.platform == "win32":
                os.startfile(res_file)
            else:
                subprocess.Popen(["xdg-open", res_file])
        else:
            messagebox.showinfo("Results File", f"No results report generated yet at:\n{res_file}")

    def _open_log_file(self):
        log_file = str(DEFAULT_LOG_FILE)
        if os.path.exists(log_file):
            if sys.platform == "win32":
                os.startfile(log_file)
            else:
                subprocess.Popen(["xdg-open", log_file])
        else:
            messagebox.showinfo("Log File", f"No log file found at:\n{log_file}")

    def _clear_log(self):
        self.log_textbox.delete("1.0", "end")

    # =========================================================================
    # Logging Drain Queue
    # =========================================================================
    def _queue_log_message(self, message: str, level: str):
        self.log_queue.put((message, level))

    def _drain_log_queue(self):
        try:
            while not self.log_queue.empty():
                msg, level = self.log_queue.get_nowait()
                tag = "INFO"
                if level == "WARNING":
                    tag = "WARNING"
                elif level == "ERROR":
                    tag = "ERROR"
                elif "SUCCESS" in msg or "Downloaded" in msg:
                    tag = "SUCCESS"
                elif "Downloading" in msg:
                    tag = "HIGHLIGHT"

                self.log_textbox.insert("end", msg + "\n", tag)
                if self.auto_scroll_var.get():
                    self.log_textbox.see("end")
        except Exception:
            pass

        self.after(100, self._drain_log_queue)

    # =========================================================================
    # Worker Thread & Execution Control
    # =========================================================================
    def _start_processing(self):
        if not self.input_records:
            excel_path = self.excel_path_var.get()
            if not excel_path or not os.path.exists(excel_path):
                messagebox.showerror("Error", "Please select a valid input Excel or CSV file first.")
                return
            try:
                self._load_file_preview(excel_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")
                return

        output_dir = self.output_dir_var.get()
        os.makedirs(output_dir, exist_ok=True)

        self._is_running = True
        self._is_paused = False
        self._pause_event.set()

        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal")
        self.btn_resume.configure(state="disabled")
        self.btn_stop.configure(state="normal")

        self.worker_thread = threading.Thread(target=self._run_downloader_loop, daemon=True)
        self.worker_thread.start()

    def _pause_processing(self):
        if self._is_running and not self._is_paused:
            self._is_paused = True
            self._pause_event.clear()
            self.btn_pause.configure(state="disabled")
            self.btn_resume.configure(state="normal")
            self.status_msg_var.set("Paused. Click Resume to continue.")
            logger.info("Execution paused by user.")

    def _resume_processing(self):
        if self._is_running and self._is_paused:
            self._is_paused = False
            self._pause_event.set()
            self.btn_pause.configure(state="normal")
            self.btn_resume.configure(state="disabled")
            self.status_msg_var.set("Resuming download pipeline...")
            logger.info("Execution resumed by user.")

    def _stop_processing(self):
        if self._is_running:
            self._is_running = False
            self._pause_event.set()
            self.status_msg_var.set("Stopping after current record...")
            logger.info("Stopping execution requested by user.")

    def _on_window_close(self):
        if self._is_running:
            if messagebox.askyesno("Exit Confirmation", "Download is currently active. Do you really want to stop and exit?"):
                self._is_running = False
                self._pause_event.set()
                self.destroy()
        else:
            self.destroy()

    def _update_metrics_ui(self, metrics: Dict[str, Any]):
        total = metrics.get("total", len(self.input_records))
        processed = metrics.get("processed", 0)
        self.total_records_var.set(str(total))
        self.processed_records_var.set(str(processed))
        self.success_count_var.set(str(metrics.get("success", 0)))
        self.skipped_count_var.set(str(metrics.get("skipped", 0)))
        self.failed_count_var.set(str(metrics.get("failed", 0)))
        self.review_count_var.set(str(metrics.get("review", 0)))

        pct = (processed / total) if total > 0 else 0.0
        self.progress_bar.set(pct)
        self.lbl_pct.configure(text=f"{pct * 100:.1f}%")

    # =========================================================================
    # Downloader Execution Loop
    # =========================================================================
    def _run_downloader_loop(self):
        logger.info("=== Starting UK FCA NSM Download Pipeline ===")
        total_records = len(self.input_records)
        output_dir = self.output_dir_var.get()
        strict_pdf = self.strict_pdf_var.get()
        skip_existing = self.skip_existing_var.get()
        delay = self.delay_var.get()
        max_retries = self.max_retries_var.get()

        client = NsmApiClient(delay_sec=delay, max_retries=max_retries)
        start_time = time.time()

        for idx, rec in enumerate(self.input_records, 1):
            if not self._is_running:
                logger.info("Pipeline stopped by user.")
                break

            self._pause_event.wait()

            firm_name = rec["firm_name"]
            year = rec["year"]
            uni_code = rec["uni_code"]
            ticker = rec.get("ticker", "")
            country = rec.get("country", "")

            # Update Current Row Badge
            self.current_company_var.set(firm_name[:25])
            self.current_year_var.set(str(year))
            self.current_code_var.set(str(uni_code))
            self.status_msg_var.set(f"[{idx}/{total_records}] Processing: {firm_name} ({year})")

            # Check if already completed in previous run
            if skip_existing and self.progress_mgr.is_record_completed(uni_code):
                logger.info(f"Record {uni_code} ({firm_name} - {year}) already completed in previous session. Skipping.")
                metrics = self.progress_mgr.get_metrics()
                self._update_metrics_ui(metrics)
                continue

            # 1. Fetch filings from NSM API (cached per company)
            filings = client.get_company_filings(firm_name)

            # 2. Match strict Annual Report for the target year
            match_status, best_filing, notes = find_best_filing_for_year(
                filings=filings,
                target_year=year,
                strict_pdf_only=strict_pdf
            )

            result_entry = {
                "serial_num": rec.get("serial_num", idx),
                "uni_code": uni_code,
                "firm_name": firm_name,
                "year": year,
                "ticker": ticker,
                "country": country,
                "nsm_doc_type": "",
                "nsm_headline": "",
                "doc_date": "",
                "sub_date": "",
                "doc_format": "",
                "status": match_status,
                "saved_filename": "",
                "saved_filepath": "",
                "file_size_mb": 0.0,
                "download_url": "",
                "error_details": notes,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            # 3. If matched, download the report
            if match_status == STATUS_SUCCESS and best_filing:
                result_entry["nsm_doc_type"] = best_filing.get("type", "")
                result_entry["nsm_headline"] = best_filing.get("headline", "")
                result_entry["doc_date"] = best_filing.get("doc_date", "")
                result_entry["sub_date"] = best_filing.get("sub_date", "")
                result_entry["doc_format"] = best_filing.get("format", "")
                result_entry["download_url"] = best_filing.get("download_url", "")

                # Download PDF
                dl_status, saved_fn, saved_fp, size_mb, dl_err = download_pdf_report(
                    doc_url=best_filing["download_url"],
                    output_dir=output_dir,
                    firm_name=firm_name,
                    uni_code=uni_code,
                    skip_existing=skip_existing,
                    max_retries=max_retries
                )

                result_entry["status"] = dl_status
                result_entry["saved_filename"] = saved_fn or ""
                result_entry["saved_filepath"] = saved_fp or ""
                result_entry["file_size_mb"] = size_mb
                if dl_err:
                    result_entry["error_details"] = f"{notes} | Download Error: {dl_err}"

            # 4. Save state & update results
            self.progress_mgr.record_result(result_entry)
            self.processed_records_list.append(result_entry)

            # Update Metrics UI
            metrics = self.progress_mgr.get_metrics()
            self._update_metrics_ui(metrics)

            # Periodic Excel Results update every 10 records or on last record
            if idx % 10 == 0 or idx == total_records:
                self.results_writer.write_results(self.progress_mgr.get_all_saved_records(), total_records)

        # Final write
        self.results_writer.write_results(self.progress_mgr.get_all_saved_records(), total_records)

        self._is_running = False
        self._is_paused = False
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled")
        self.btn_resume.configure(state="disabled")
        self.btn_stop.configure(state="disabled")

        elapsed = time.time() - start_time
        final_metrics = self.progress_mgr.get_metrics()
        self.status_msg_var.set(
            f"Finished! Processed {final_metrics['processed']}/{total_records} in {elapsed:.1f}s. "
            f"Success: {final_metrics['success']}, Skipped: {final_metrics['skipped']}, Failed: {final_metrics['failed']}"
        )
        logger.info(f"=== Pipeline completed in {elapsed:.1f} seconds ===")
        messagebox.showinfo(
            "Download Finished",
            f"UK FCA NSM Annual Report Download complete!\n\n"
            f"Total Processed: {final_metrics['processed']}\n"
            f"Downloaded (New): {final_metrics['success']}\n"
            f"Skipped: {final_metrics['skipped']}\n"
            f"Failed: {final_metrics['failed']}\n\n"
            f"Reports saved in:\n{output_dir}\n\n"
            f"Audit report saved in:\n{DEFAULT_RESULTS_FILE}"
        )
