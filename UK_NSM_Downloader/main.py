"""
main.py - Entry point for UK FCA NSM Annual Report Downloader.
Supports both desktop CustomTkinter GUI and headless command-line execution.
"""
import sys
import os
import argparse
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import DEFAULT_DOWNLOAD_DIR, DEFAULT_RESULTS_FILE, DEFAULT_STATE_FILE
from logger import get_logger

logger = get_logger()

def run_gui():
    """Launches the CustomTkinter modern desktop GUI."""
    logger.info("Starting UK NSM Downloader GUI...")
    from gui import UkNsmDownloaderApp
    app = UkNsmDownloaderApp()
    app.mainloop()

def run_cli(args):
    """Executes the download pipeline directly in terminal/headless mode."""
    logger.info("=== Running UK NSM Downloader in CLI Mode ===")
    from excel_reader import load_input_records
    from progress_manager import ProgressManager
    from results_writer import ResultsWriter
    from nsm_client import NsmApiClient
    from filing_matcher import find_best_filing_for_year
    from downloader import download_pdf_report
    from config import STATUS_SUCCESS

    input_file = args.input
    output_dir = args.output or str(DEFAULT_DOWNLOAD_DIR)
    strict_pdf = not args.allow_non_pdf
    skip_existing = not args.force_redownload
    delay = args.delay
    retries = args.retries

    if not os.path.exists(input_file):
        logger.error(f"Input file not found: {input_file}")
        sys.exit(1)

    records, mapping = load_input_records(input_file)
    total_records = len(records)
    logger.info(f"Loaded {total_records} records from {input_file}")

    progress_mgr = ProgressManager(DEFAULT_STATE_FILE)
    results_writer = ResultsWriter(args.results or DEFAULT_RESULTS_FILE)
    client = NsmApiClient(delay_sec=delay, max_retries=retries)

    for idx, rec in enumerate(records, 1):
        firm_name = rec["firm_name"]
        year = rec["year"]
        uni_code = rec["uni_code"]

        if skip_existing and progress_mgr.is_record_completed(uni_code):
            logger.info(f"[{idx}/{total_records}] Code {uni_code} ({firm_name} {year}) already completed. Skipping.")
            continue

        logger.info(f"[{idx}/{total_records}] Processing: {firm_name} - {year} (UniCode: {uni_code})")
        filings = client.get_company_filings(firm_name)

        match_status, best_filing, notes = find_best_filing_for_year(
            filings=filings,
            target_year=year,
            strict_pdf_only=strict_pdf
        )

        result_entry = {
            "serial_num": idx,
            "uni_code": uni_code,
            "firm_name": firm_name,
            "year": year,
            "ticker": rec.get("ticker", ""),
            "country": rec.get("country", ""),
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
            "error_details": notes
        }

        if match_status == STATUS_SUCCESS and best_filing:
            result_entry["nsm_doc_type"] = best_filing.get("type", "")
            result_entry["nsm_headline"] = best_filing.get("headline", "")
            result_entry["doc_date"] = best_filing.get("doc_date", "")
            result_entry["sub_date"] = best_filing.get("sub_date", "")
            result_entry["doc_format"] = best_filing.get("format", "")
            result_entry["download_url"] = best_filing.get("download_url", "")

            dl_status, saved_fn, saved_fp, size_mb, dl_err = download_pdf_report(
                doc_url=best_filing["download_url"],
                output_dir=output_dir,
                firm_name=firm_name,
                uni_code=uni_code,
                skip_existing=skip_existing,
                max_retries=retries
            )

            result_entry["status"] = dl_status
            result_entry["saved_filename"] = saved_fn or ""
            result_entry["saved_filepath"] = saved_fp or ""
            result_entry["file_size_mb"] = size_mb
            if dl_err:
                result_entry["error_details"] = f"{notes} | {dl_err}"

        progress_mgr.record_result(result_entry)
        if idx % 10 == 0 or idx == total_records:
            results_writer.write_results(progress_mgr.get_all_saved_records(), total_records)

    results_writer.write_results(progress_mgr.get_all_saved_records(), total_records)
    metrics = progress_mgr.get_metrics()
    logger.info(f"Done! Processed: {metrics['processed']}, Downloaded: {metrics['success']}, Skipped: {metrics['skipped']}, Failed: {metrics['failed']}")

def main():
    parser = argparse.ArgumentParser(description="UK FCA NSM Annual Report Downloader & Tracking System")
    parser.add_argument("--cli", action="store_true", help="Run in command-line mode without GUI")
    parser.add_argument("--input", type=str, help="Path to input Excel or CSV file")
    parser.add_argument("--output", type=str, help="Destination directory for company PDF folders")
    parser.add_argument("--results", type=str, help="Path for output results Excel workbook")
    parser.add_argument("--allow-non-pdf", action="store_true", help="Allow non-PDF formats (ESEF/XHTML)")
    parser.add_argument("--force-redownload", action="store_true", help="Force re-download even if already downloaded")
    parser.add_argument("--delay", type=float, default=0.5, help="Polite delay between API calls in seconds (default: 0.5)")
    parser.add_argument("--retries", type=int, default=3, help="Max retries for network calls (default: 3)")

    args = parser.parse_args()

    if args.cli:
        if not args.input:
            parser.error("--input is required when running with --cli")
        run_cli(args)
    else:
        run_gui()

if __name__ == "__main__":
    main()
