"""
main.py - Entry point for SEC EDGAR 10-K Downloader Desktop Application.
Supports both GUI mode (default) and CLI mode for headless batch runs.
"""
import sys
import os
import argparse
import asyncio
from pathlib import Path

# Ensure package root is in sys.path
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from logger import setup_logger, get_logger
from config import (
    DEFAULT_INPUT_DIR,
    DEFAULT_DOWNLOAD_DIR,
    DEFAULT_RESULTS_FILE,
    DEFAULT_STATE_FILE,
    DEFAULT_HEADLESS,
    DEFAULT_BROWSER_DELAY,
    DEFAULT_SKIP_EXISTING,
    DEFAULT_MAX_RETRIES,
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_MATCHING_10K,
    STATUS_REVIEW_REQUIRED
)
from excel_reader import load_input_records
from progress_manager import ProgressManager
from results_writer import ResultsWriter
from browser_automation import SecBrowserAutomation

logger = setup_logger()

async def run_cli_mode(args):
    """Executes the extraction workflow in CLI / headless batch mode."""
    logger.info("=" * 60)
    logger.info("Starting SEC EDGAR 10-K Downloader in CLI Mode")
    logger.info("=" * 60)

    input_file = args.input or str(DEFAULT_INPUT_DIR / "sample_companies.xlsx")
    if not os.path.exists(input_file):
        # Fallback to CSV if xlsx not present
        csv_file = str(DEFAULT_INPUT_DIR / "sample_companies.csv")
        if os.path.exists(csv_file):
            input_file = csv_file
        else:
            logger.error(f"Input file not found: {input_file}")
            sys.exit(1)

    output_dir = args.output or str(DEFAULT_DOWNLOAD_DIR)
    os.makedirs(output_dir, exist_ok=True)

    records, mapping = load_input_records(input_file)
    if args.limit and args.limit > 0:
        records = records[:args.limit]
        logger.info(f"Limiting execution to first {len(records)} records.")

    total_records = len(records)
    progress_mgr = ProgressManager(args.state or DEFAULT_STATE_FILE)
    progress_mgr.state["total_records"] = total_records
    results_writer = ResultsWriter(args.results or DEFAULT_RESULTS_FILE)

    browser = SecBrowserAutomation(
        headless=args.headless,
        delay_sec=args.delay
    )

    try:
        await browser.start()

        for idx, record in enumerate(records, 1):
            unique_code = record["unique_code"]
            company_name = record["company_name"]
            reporting_year = record["reporting_year"]

            logger.info(f"[{idx}/{total_records}] Processing: {company_name} - {reporting_year} (Code: {unique_code})")

            # Skip if already completed
            if args.skip_existing and progress_mgr.is_already_completed(unique_code):
                logger.info(f"Skipping [{unique_code}] - already completed.")
                continue

            # Process record with retries
            record_result = None
            for attempt in range(args.max_retries + 1):
                record_result = await browser.process_record(
                    record, output_dir=output_dir, skip_existing=args.skip_existing
                )
                status = record_result.get("download_status")
                if status in [STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED, 
                              STATUS_SKIPPED_NO_MATCHING_10K, STATUS_REVIEW_REQUIRED]:
                    break
                if attempt < args.max_retries:
                    logger.warning(f"Retrying [{unique_code}] {company_name} (Attempt {attempt+2}/{args.max_retries+1})...")
                    await asyncio.sleep(2.0)

            if record_result:
                progress_mgr.save_record_result(unique_code, record_result)
                all_results = list(progress_mgr.state["records"].values())
                results_writer.write_results(all_results, total_records)

        stats = progress_mgr.get_summary_stats()
        logger.info("=" * 60)
        logger.info(f"CLI Run Finished: Total={stats['total']}, Success={stats['success']}, "
                    f"Skipped={stats['skipped']}, Failed={stats['failed']}, Review={stats['review']}")
        logger.info(f"Results workbook saved to: {results_writer.file_path}")
        logger.info("=" * 60)

    finally:
        await browser.close()

def main():
    parser = argparse.ArgumentParser(description="SEC EDGAR 10-K Report Downloader & Excel Tracking System")
    parser.add_argument("--cli", action="store_true", help="Run in command-line mode without GUI")
    parser.add_argument("--input", type=str, help="Path to input Excel or CSV file")
    parser.add_argument("--output", type=str, help="Directory to save downloaded HTML reports")
    parser.add_argument("--results", type=str, help="Path to output results Excel workbook")
    parser.add_argument("--state", type=str, help="Path to progress state JSON file")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of records to process (for testing)")
    parser.add_argument("--headless", action="store_true", default=DEFAULT_HEADLESS, help="Run browser in headless mode")
    parser.add_argument("--no-skip-existing", dest="skip_existing", action="store_false", default=True, help="Redownload existing files")
    parser.add_argument("--delay", type=float, default=DEFAULT_BROWSER_DELAY, help="Delay between browser requests in seconds")
    parser.add_argument("--max-retries", type=int, default=DEFAULT_MAX_RETRIES, help="Maximum retry attempts for failed records")

    args = parser.parse_args()

    if args.cli:
        asyncio.run(run_cli_mode(args))
    else:
        try:
            from gui import SecDownloaderApp
            app = SecDownloaderApp()
            app.mainloop()
        except Exception as e:
            logger.error(f"Failed to start GUI: {e}", exc_info=True)
            print(f"Error starting GUI: {e}")
            print("To run in command-line mode, use: python main.py --cli")
            sys.exit(1)

if __name__ == "__main__":
    main()
