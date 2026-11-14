"""
gui.py - Modern CustomTkinter desktop GUI for SEC EDGAR 10-K Downloader.
Provides responsive controls, live progress metrics, threaded execution, and real-time logging.
"""
import os
import sys
import asyncio
import threading
import queue
import subprocess
from datetime import datetime
from typing import Optional, List, Dict, Any
import customtkinter as ctk
from tkinter import filedialog, messagebox

from config import (
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_RESULTS_FILE,
    DEFAULT_STATE_FILE,
    DEFAULT_BROWSER_DELAY,
    DEFAULT_MAX_RETRIES,
    DEFAULT_SKIP_EXISTING,
    DEFAULT_HEADLESS,
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_MATCHING_10K,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_BROWSER_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_HTML,
    STATUS_REVIEW_REQUIRED
)
from logger import get_logger, set_gui_callback
from excel_reader import inspect_file_headers, detect_columns, load_input_records
from progress_manager import ProgressManager
from results_writer import ResultsWriter
from browser_automation import SecBrowserAutomation

logger = get_logger()

# Set Theme & Appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class SecDownloaderApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("SEC EDGAR 10-K Report Automation & Tracking System")
        self.geometry("1100, 850")
        self.minsize(980, 750)

        # State & Managers
        self.progress_mgr = ProgressManager(DEFAULT_STATE_FILE)
        self.results_writer = ResultsWriter(DEFAULT_RESULTS_FILE)
        self.worker_thread: Optional[threading.Thread] = None
        self.log_queue = queue.Queue()
        self.input_records: List[Dict[str, Any]] = []

        # Variables
        self.excel_path_var = ctk.StringVar()
        self.output_dir_var = ctk.StringVar(value=str(DEFAULT_DOWNLOAD_DIR))
        self.headless_var = ctk.BooleanVar(value=DEFAULT_HEADLESS)
        self.skip_existing_var = ctk.BooleanVar(value=DEFAULT_SKIP_EXISTING)
        self.retry_failed_var = ctk.BooleanVar(value=True)
        self.max_retries_var = ctk.IntVar(value=DEFAULT_MAX_RETRIES)
        self.delay_var = ctk.DoubleVar(value=DEFAULT_BROWSER_DELAY)
        
        # Progress vars
        self.status_msg_var = ctk.StringVar(value="Ready to load Excel file.")
        self.current_company_var = ctk.StringVar(value="-")
        self.current_year_var = ctk.StringVar(value="-")
        self.total_records_var = ctk.StringVar(value="0")
        self.processed_records_var = ctk.StringVar(value="0")
        self.success_count_var = ctk.StringVar(value="0")
        self.skipped_count_var = ctk.StringVar(value="0")
        self.failed_count_var = ctk.StringVar(value="0")
        self.review_count_var = ctk.StringVar(value="0")

        # Connect logger callback to queue
        set_gui_callback(self._queue_log_message)

        # Build UI
        self._build_ui()

        # Start periodic log drain and status poller
        self.after(100, self._drain_log_queue)
        self.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Check for sample file
        sample_path = os.path.join(os.path.dirname(__file__), "input", "sample_companies.xlsx")
        if os.path.exists(sample_path):
            self.excel_path_var.set(sample_path)
            self._load_file_preview(sample_path)

    def _build_ui(self):
        # Main scrollable or grid container
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)  # Log section expands

        # 1. Header Banner
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1A365D", "#0F2042"), height=60)
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 10))
        header_frame.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            header_frame,
            text="SEC EDGAR 10-K Report Downloader",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#FFFFFF"
        )
        title_label.grid(row=0, column=0, padx=20, pady=(10, 2), sticky="w")

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Automated Browser Extraction, Filing Matching, and Excel Audit Tracking",
            font=ctk.CTkFont(size=12),
            text_color="#CBD5E1"
        )
        subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")

        # 2. Input Section Frame
        input_frame = ctk.CTkFrame(self, corner_radius=8)
        input_frame.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        input_frame.grid_columnconfigure(1, weight=1)

        # Excel File Row
        ctk.CTkLabel(input_frame, text="Input Excel / CSV:", font=ctk.CTkFont(weight="bold")).grid(
            row=0, column=0, padx=12, pady=8, sticky="w"
        )
        excel_entry = ctk.CTkEntry(input_frame, textvariable=self.excel_path_var, placeholder_text="Select input file...")
        excel_entry.grid(row=0, column=1, padx=8, pady=8, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse File", width=110, command=self._browse_excel).grid(
            row=0, column=2, padx=12, pady=8
        )

        # Output Folder Row
        ctk.CTkLabel(input_frame, text="Output Folder:", font=ctk.CTkFont(weight="bold")).grid(
            row=1, column=0, padx=12, pady=8, sticky="w"
        )
        out_entry = ctk.CTkEntry(input_frame, textvariable=self.output_dir_var)
        out_entry.grid(row=1, column=1, padx=8, pady=8, sticky="ew")
        ctk.CTkButton(input_frame, text="Browse Folder", width=110, command=self._browse_output_dir).grid(
            row=1, column=2, padx=12, pady=8
        )

        # 3. Configuration Section Frame
        config_frame = ctk.CTkFrame(self, corner_radius=8)
        config_frame.grid(row=2, column=0, sticky="ew", padx=15, pady=5)
        config_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Row 0: Switches
        headless_switch = ctk.CTkSwitch(
            config_frame,
            text="Headless Mode (OFF recommended)",
            variable=self.headless_var,
            onvalue=True,
            offvalue=False
        )
        headless_switch.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        skip_switch = ctk.CTkSwitch(
            config_frame,
            text="Skip Existing Reports",
            variable=self.skip_existing_var
        )
        skip_switch.grid(row=0, column=1, padx=15, pady=10, sticky="w")

        retry_switch = ctk.CTkSwitch(
            config_frame,
            text="Retry Failed Records",
            variable=self.retry_failed_var
        )
        retry_switch.grid(row=0, column=2, padx=15, pady=10, sticky="w")

        # Row 1: Delay slider & Max retries
        delay_subframe = ctk.CTkFrame(config_frame, fg_color="transparent")
        delay_subframe.grid(row=1, column=0, columnspan=2, padx=15, pady=5, sticky="ew")
        ctk.CTkLabel(delay_subframe, text="Browser Delay:").pack(side="left", padx=(0, 8))
        self.delay_slider = ctk.CTkSlider(
            delay_subframe, from_=0.5, to=5.0, number_of_steps=9, variable=self.delay_var, width=180
        )
        self.delay_slider.pack(side="left", padx=5)
        self.delay_label = ctk.CTkLabel(delay_subframe, text=f"{self.delay_var.get():.1f}s")
        self.delay_label.pack(side="left", padx=5)
        self.delay_slider.configure(command=lambda val: self.delay_label.configure(text=f"{float(val):.1f}s"))

        retries_subframe = ctk.CTkFrame(config_frame, fg_color="transparent")
        retries_subframe.grid(row=1, column=2, columnspan=2, padx=15, pady=5, sticky="ew")
        ctk.CTkLabel(retries_subframe, text="Max Retries:").pack(side="left", padx=(0, 8))
        self.retries_slider = ctk.CTkSlider(
            retries_subframe, from_=1, to=5, number_of_steps=4, variable=self.max_retries_var, width=140
        )
        self.retries_slider.pack(side="left", padx=5)
        self.retries_label = ctk.CTkLabel(retries_subframe, text=f"{self.max_retries_var.get()}")
        self.retries_label.pack(side="left", padx=5)
        self.retries_slider.configure(command=lambda val: self.retries_label.configure(text=f"{int(float(val))}"))

        # 4. Processing & Progress Section
        progress_frame = ctk.CTkFrame(self, corner_radius=8)
        progress_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=5)
        progress_frame.grid_columnconfigure(0, weight=1)

        # Control Buttons Row
        ctrl_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        ctrl_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        ctrl_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_start = ctk.CTkButton(
            ctrl_frame, text="▶ Start Download", fg_color="#10B981", hover_color="#059669",
            font=ctk.CTkFont(weight="bold", size=13), height=38, command=self._start_processing
        )
        self.btn_start.grid(row=0, column=0, padx=6, pady=5, sticky="ew")

        self.btn_pause = ctk.CTkButton(
            ctrl_frame, text="⏸ Pause", fg_color="#F59E0B", hover_color="#D97706",
            font=ctk.CTkFont(weight="bold", size=13), height=38, state="disabled", command=self._pause_processing
        )
        self.btn_pause.grid(row=0, column=1, padx=6, pady=5, sticky="ew")

        self.btn_resume = ctk.CTkButton(
            ctrl_frame, text="⏯ Resume", fg_color="#3B82F6", hover_color="#2563EB",
            font=ctk.CTkFont(weight="bold", size=13), height=38, state="disabled", command=self._resume_processing
        )
        self.btn_resume.grid(row=0, column=2, padx=6, pady=5, sticky="ew")

        self.btn_stop = ctk.CTkButton(
            ctrl_frame, text="⏹ Stop", fg_color="#EF4444", hover_color="#DC2626",
            font=ctk.CTkFont(weight="bold", size=13), height=38, state="disabled", command=self._stop_processing
        )
        self.btn_stop.grid(row=0, column=3, padx=6, pady=5, sticky="ew")

        # Progress Bar & Live Status
        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=14)
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=15, pady=(8, 4))
        self.progress_bar.set(0.0)

        status_line = ctk.CTkFrame(progress_frame, fg_color="transparent")
        status_line.grid(row=2, column=0, sticky="ew", padx=15, pady=(2, 8))
        status_line.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(status_line, textvariable=self.status_msg_var, font=ctk.CTkFont(size=12))
        self.lbl_status.grid(row=0, column=0, sticky="w")

        curr_info = ctk.CTkLabel(
            status_line,
            text="Current: ",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        curr_info.grid(row=0, column=1, sticky="e", padx=(10, 2))

        self.lbl_curr = ctk.CTkLabel(
            status_line,
            textvariable=self.current_company_var,
            text_color="#38BDF8",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.lbl_curr.grid(row=0, column=2, sticky="e")

        # Metric Cards Frame
        cards_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        cards_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(2, 10))
        for i in range(6):
            cards_frame.grid_columnconfigure(i, weight=1)

        self._create_metric_card(cards_frame, 0, "Total Records", self.total_records_var, "#94A3B8")
        self._create_metric_card(cards_frame, 1, "Processed", self.processed_records_var, "#60A5FA")
        self._create_metric_card(cards_frame, 2, "Success", self.success_count_var, "#34D399")
        self._create_metric_card(cards_frame, 3, "Skipped", self.skipped_count_var, "#38BDF8")
        self._create_metric_card(cards_frame, 4, "Failed", self.failed_count_var, "#F87171")
        self._create_metric_card(cards_frame, 5, "Review Needed", self.review_count_var, "#FBBF24")

        # 5. Live Log Section Frame
        log_frame = ctk.CTkFrame(self, corner_radius=8)
        log_frame.grid(row=4, column=0, sticky="nsew", padx=15, pady=(5, 12))
        log_frame.grid_rowconfigure(1, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        log_top = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_top.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))
        log_top.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(log_top, text="Execution Logs", font=ctk.CTkFont(weight="bold", size=13)).grid(
            row=0, column=0, sticky="w"
        )

        btn_box = ctk.CTkFrame(log_top, fg_color="transparent")
        btn_box.grid(row=0, column=1, sticky="e")

        ctk.CTkButton(
            btn_box, text="Open Output Folder", width=130, height=26,
            command=self._open_output_folder
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_box, text="Open Results Excel", width=130, height=26,
            command=self._open_results_excel
        ).pack(side="left", padx=4)

        ctk.CTkButton(
            btn_box, text="Clear Log", width=80, height=26,
            command=self._clear_log
        ).pack(side="left", padx=4)

        self.log_textbox = ctk.CTkTextbox(
            log_frame, font=ctk.CTkFont(family="Consolas", size=11), wrap="word"
        )
        self.log_textbox.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))

    def _create_metric_card(self, parent, col, title, var, accent_color):
        card = ctk.CTkFrame(parent, corner_radius=6, border_width=1, border_color="#334155")
        card.grid(row=0, column=col, padx=4, pady=2, sticky="ew")
        
        lbl_title = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=10, weight="bold"), text_color="#94A3B8")
        lbl_title.pack(pady=(4, 0))

        lbl_val = ctk.CTkLabel(card, textvariable=var, font=ctk.CTkFont(size=18, weight="bold"), text_color=accent_color)
        lbl_val.pack(pady=(0, 4))

    def _browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Input Excel or CSV File",
            filetypes=[("Excel & CSV Files", "*.xlsx *.xls *.csv"), ("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self._load_file_preview(file_path)

    def _browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="Select Output Folder")
        if dir_path:
            self.output_dir_var.set(dir_path)

    def _load_file_preview(self, file_path: str):
        try:
            records, mapping = load_input_records(file_path)
            self.input_records = records
            self.total_records_var.set(str(len(records)))
            self.status_msg_var.set(
                f"Loaded {len(records)} records from {os.path.basename(file_path)} "
                f"(Mapped: {mapping['company_col']}, {mapping['year_col']}, {mapping['code_col']})"
            )
            logger.info(f"Loaded {len(records)} records from {file_path}")
        except Exception as e:
            messagebox.showerror("Error Reading Input File", str(e))
            self.status_msg_var.set(f"Error loading file: {e}")

    def _queue_log_message(self, message: str, level: str):
        self.log_queue.put((message, level))

    def _drain_log_queue(self):
        while not self.log_queue.empty():
            try:
                msg, level = self.log_queue.get_nowait()
                self.log_textbox.insert("end", msg + "\n")
                self.log_textbox.see("end")
            except queue.Empty:
                break
        self.after(100, self._drain_log_queue)

    def _clear_log(self):
        self.log_textbox.delete("1.0", "end")

    def _open_output_folder(self):
        folder = self.output_dir_var.get()
        if os.path.exists(folder):
            if sys.platform == "win32":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        else:
            messagebox.showinfo("Folder Not Found", f"Directory does not exist yet:\n{folder}")

    def _open_results_excel(self):
        path = str(DEFAULT_RESULTS_FILE)
        if os.path.exists(path):
            if sys.platform == "win32":
                os.startfile(path)
            else:
                subprocess.Popen(["xdg-open", path])
        else:
            messagebox.showinfo("File Not Found", f"Results Excel file does not exist yet:\n{path}")

    def _start_processing(self):
        excel_file = self.excel_path_var.get()
        if not excel_file or not os.path.exists(excel_file):
            messagebox.showwarning("File Required", "Please select a valid input Excel or CSV file.")
            return

        if not self.input_records:
            try:
                self._load_file_preview(excel_file)
            except Exception:
                return

        output_dir = self.output_dir_var.get()
        os.makedirs(output_dir, exist_ok=True)

        # Update button states
        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal")
        self.btn_resume.configure(state="disabled")
        self.btn_stop.configure(state="normal")

        # Launch background worker
        self.worker_thread = threading.Thread(
            target=self._run_automation_loop,
            daemon=True
        )
        self.worker_thread.start()

    def _pause_processing(self):
        self.progress_mgr.pause()
        self.btn_pause.configure(state="disabled")
        self.btn_resume.configure(state="normal")
        self.status_msg_var.set("Automation paused by user.")

    def _resume_processing(self):
        self.progress_mgr.resume()
        self.btn_pause.configure(state="normal")
        self.btn_resume.configure(state="disabled")
        self.status_msg_var.set("Automation resumed.")

    def _stop_processing(self):
        if messagebox.askyesno("Stop Processing", "Are you sure you want to stop processing? Progress will be saved."):
            self.progress_mgr.stop()
            self.status_msg_var.set("Stopping automation...")
            self.btn_pause.configure(state="disabled")
            self.btn_resume.configure(state="disabled")
            self.btn_stop.configure(state="disabled")

    def _update_ui_progress(self, current_comp: str, current_year: str):
        stats = self.progress_mgr.get_summary_stats()
        total = stats["total"] or len(self.input_records)
        processed = stats["processed"]

        self.total_records_var.set(str(total))
        self.processed_records_var.set(str(processed))
        self.success_count_var.set(str(stats["success"]))
        self.skipped_count_var.set(str(stats["skipped"]))
        self.failed_count_var.set(str(stats["failed"]))
        self.review_count_var.set(str(stats["review"]))

        self.current_company_var.set(f"{current_comp} ({current_year})")
        pct = (processed / total) if total > 0 else 0.0
        self.progress_bar.set(pct)
        self.status_msg_var.set(f"Processing: {processed}/{total} records ({pct*100:.1f}%)")

    def _on_captcha_detected(self):
        """Called when CAPTCHA or bot verification is detected."""
        def prompt_user():
            messagebox.showwarning(
                "Bot Verification Detected",
                "SEC EDGAR has displayed a verification or rate-limit challenge.\n\n"
                "Please switch to the visible browser window, complete the challenge manually, "
                "then click Resume in this application."
            )
            self._pause_processing()

        self.after(0, prompt_user)

    def _run_automation_loop(self):
        """Background thread executing the Playwright async event loop."""
        asyncio.run(self._async_automation_worker())

    async def _async_automation_worker(self):
        logger.info("Initializing browser automation worker...")
        output_dir = self.output_dir_var.get()
        skip_existing = self.skip_existing_var.get()
        headless = self.headless_var.get()
        delay_sec = self.delay_var.get()

        browser_auto = SecBrowserAutomation(
            headless=headless,
            delay_sec=delay_sec,
            on_captcha_callback=self._on_captcha_detected
        )

        try:
            await browser_auto.start()

            total_records = len(self.input_records)
            self.progress_mgr.state["total_records"] = total_records

            # Process records
            for record in self.input_records:
                if self.progress_mgr.is_stopped():
                    logger.info("Worker detected stop request. Halting loop.")
                    break

                self.progress_mgr.wait_if_paused()
                if self.progress_mgr.is_stopped():
                    break

                unique_code = record["unique_code"]
                company_name = record["company_name"]
                reporting_year = record["reporting_year"]

                # UI update
                self.after(0, self._update_ui_progress, company_name, str(reporting_year))

                # Check if already completed and skip_existing is True
                if skip_existing and self.progress_mgr.is_already_completed(unique_code):
                    cached_record = self.progress_mgr.get_record(unique_code)
                    if cached_record:
                        logger.info(f"Skipping [{unique_code}] {company_name} ({reporting_year}) - already completed.")
                        continue

                # Execute record download with retries
                max_retries = self.max_retries_var.get()
                record_result = None

                for attempt in range(max_retries + 1):
                    self.progress_mgr.wait_if_paused()
                    if self.progress_mgr.is_stopped():
                        break

                    record_result = await browser_auto.process_record(
                        record, output_dir=output_dir, skip_existing=skip_existing
                    )

                    status = record_result.get("download_status")
                    if status in [STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED, STATUS_SKIPPED_NO_MATCHING_10K, STATUS_REVIEW_REQUIRED]:
                        break

                    if attempt < max_retries:
                        logger.warning(f"Retrying [{unique_code}] {company_name} (Attempt {attempt+2}/{max_retries+1})...")
                        await asyncio.sleep(2.0)

                if record_result:
                    self.progress_mgr.save_record_result(unique_code, record_result)

                    # Update Results Excel in safe batches / per record
                    all_results = list(self.progress_mgr.state["records"].values())
                    self.results_writer.write_results(all_results, total_records)

                # Update UI counts
                self.after(0, self._update_ui_progress, company_name, str(reporting_year))

            logger.info("Automation run finished.")

        except Exception as e:
            logger.error(f"Fatal error in automation loop: {e}", exc_info=True)
            self.after(0, lambda: messagebox.showerror("Automation Error", f"Fatal error: {str(e)}"))

        finally:
            await browser_auto.close()
            self.after(0, self._on_automation_finished)

    def _on_automation_finished(self):
        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled")
        self.btn_resume.configure(state="disabled")
        self.btn_stop.configure(state="disabled")
        self.current_company_var.set("Complete")
        
        stats = self.progress_mgr.get_summary_stats()
        self.status_msg_var.set(
            f"Finished! Success: {stats['success']}, Skipped: {stats['skipped']}, Failed: {stats['failed']}, Review: {stats['review']}"
        )
        logger.info(f"Summary: Success={stats['success']}, Skipped={stats['skipped']}, Failed={stats['failed']}, Review={stats['review']}")

    def _on_window_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if messagebox.askyesno("Exit Application", "An automation task is running. Do you want to stop it and exit?"):
                self.progress_mgr.stop()
                self.destroy()
        else:
            self.destroy()
