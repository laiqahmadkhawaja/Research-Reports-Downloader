"""
config.py - Configuration constants and default settings for SEC 10-K Downloader.
"""
import os
from pathlib import Path

# Base Paths
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = BASE_DIR / "input"
DEFAULT_DOWNLOAD_DIR = BASE_DIR / "downloaded_reports"
DEFAULT_RESULTS_DIR = BASE_DIR / "results"
DEFAULT_LOGS_DIR = BASE_DIR / "logs"
DEFAULT_STATE_DIR = BASE_DIR / "state"

# Ensure directories exist
for d in [DEFAULT_INPUT_DIR, DEFAULT_DOWNLOAD_DIR, DEFAULT_RESULTS_DIR, DEFAULT_LOGS_DIR, DEFAULT_STATE_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# Default File Paths
DEFAULT_RESULTS_FILE = DEFAULT_RESULTS_DIR / "SEC_Download_Results.xlsx"
DEFAULT_STATE_FILE = DEFAULT_STATE_DIR / "progress.json"
DEFAULT_LOG_FILE = DEFAULT_LOGS_DIR / "sec_downloader.log"

# SEC EDGAR URL
SEC_SEARCH_URL = "https://www.sec.gov/edgar/search/"

# Default Browser Headers & User-Agent
# Standard modern browser user agent that complies with SEC Fair Access policies
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

DEFAULT_HTTP_HEADERS = {
    "Accept-Language": "en-US,en;q=0.9"
}

# Status Constants (exact match with requirements)
STATUS_SUCCESS = "SUCCESS"
STATUS_SKIPPED_ALREADY_DOWNLOADED = "SKIPPED - ALREADY DOWNLOADED"
STATUS_SKIPPED_NO_MATCHING_10K = "SKIPPED - NO MATCHING 10-K"
STATUS_FAILED_COMPANY_NOT_FOUND = "FAILED - COMPANY NOT FOUND"
STATUS_FAILED_BROWSER_ERROR = "FAILED - BROWSER ERROR"
STATUS_FAILED_DOWNLOAD_ERROR = "FAILED - DOWNLOAD ERROR"
STATUS_FAILED_INVALID_HTML = "FAILED - INVALID HTML"
STATUS_REVIEW_REQUIRED = "REVIEW REQUIRED"

ALL_STATUSES = [
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_MATCHING_10K,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_BROWSER_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_HTML,
    STATUS_REVIEW_REQUIRED,
]

# Operational Defaults
DEFAULT_HEADLESS = False  # Headed is recommended for SEC EDGAR to avoid Akamai 403 blocks
DEFAULT_BROWSER_DELAY = 1.5  # Seconds between browser actions to respect fair access
DEFAULT_MAX_RETRIES = 2
DEFAULT_SKIP_EXISTING = True
DEFAULT_PAGE_TIMEOUT_MS = 45000
DEFAULT_ACTION_TIMEOUT_MS = 15000

# Results Excel Sheet Names
SHEET_DOWNLOAD_RESULTS = "Download Results"
SHEET_SUMMARY = "Summary"
SHEET_FAILED_AND_REVIEW = "Failed & Review Records"

# Results Excel Columns (Sheet 1)
RESULTS_COLUMNS = [
    "Serial Number",
    "Company Name",
    "Reporting Year",
    "Unique Report Code",
    "SEC Company Name Selected",
    "Filing Type",
    "Filing Date",
    "Reporting Period",
    "Accession Number",
    "Download Status",
    "Saved File Name",
    "Saved File Path",
    "SEC Filing URL",
    "Error Details",
    "Processing Date & Time"
]

# Failed & Review Columns (Sheet 3)
FAILED_REVIEW_COLUMNS = [
    "Serial Number",
    "Company Name",
    "Reporting Year",
    "Unique Report Code",
    "Failure or Review Reason",
    "Error Details",
    "Suggested Action"
]
