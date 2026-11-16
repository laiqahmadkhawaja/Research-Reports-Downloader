"""
config.py - Configuration constants and default settings for UK FCA NSM Annual Report Downloader.
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
DEFAULT_RESULTS_FILE = DEFAULT_RESULTS_DIR / "UK_NSM_Download_Results.xlsx"
DEFAULT_STATE_FILE = DEFAULT_STATE_DIR / "progress.json"
DEFAULT_LOG_FILE = DEFAULT_LOGS_DIR / "uk_nsm_downloader.log"

# FCA NSM API & Portal URLs
NSM_PORTAL_URL = "https://data.fca.org.uk/#/nsm/nationalstoragemechanism"
NSM_SEARCH_API_URL = "https://api.data.fca.org.uk/search?index=nsm-search"
NSM_ARTEFACT_BASE_URL = "https://data.fca.org.uk/artefacts/"

# HTTP Headers
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

DEFAULT_API_HEADERS = {
    "User-Agent": DEFAULT_USER_AGENT,
    "Accept": "application/json, text/plain, */*",
    "Content-Type": "application/json",
    "Origin": "https://data.fca.org.uk",
    "Referer": "https://data.fca.org.uk/"
}

# Status Constants
STATUS_SUCCESS = "SUCCESS"
STATUS_SKIPPED_ALREADY_DOWNLOADED = "SKIPPED - ALREADY DOWNLOADED"
STATUS_SKIPPED_NO_ANNUAL_REPORT = "SKIPPED - NO ANNUAL REPORT FOUND"
STATUS_SKIPPED_NO_PDF = "SKIPPED - NO PDF VERSION (ONLY HTML/XBRL)"
STATUS_FAILED_COMPANY_NOT_FOUND = "FAILED - COMPANY NOT FOUND"
STATUS_FAILED_NETWORK_ERROR = "FAILED - NETWORK ERROR"
STATUS_FAILED_DOWNLOAD_ERROR = "FAILED - DOWNLOAD ERROR"
STATUS_FAILED_INVALID_PDF = "FAILED - INVALID PDF"
STATUS_REVIEW_REQUIRED = "REVIEW REQUIRED"

ALL_STATUSES = [
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_ANNUAL_REPORT,
    STATUS_SKIPPED_NO_PDF,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_NETWORK_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_PDF,
    STATUS_REVIEW_REQUIRED,
]

# Operational Defaults
DEFAULT_MAX_RETRIES = 3
DEFAULT_REQUEST_DELAY = 0.5  # Seconds between requests (polite API usage)
DEFAULT_SKIP_EXISTING = True
DEFAULT_STRICT_PDF_ONLY = True
DEFAULT_TIMEOUT_SEC = 30
DEFAULT_CHUNK_SIZE = 65536  # 64KB for streaming downloads

# Excel Sheet Names
SHEET_DOWNLOAD_RESULTS = "Download Results"
SHEET_SUMMARY = "Summary"
SHEET_FAILED_AND_REVIEW = "Failed & Review Records"

# Results Excel Columns (Sheet 1)
RESULTS_COLUMNS = [
    "Serial Number",
    "UniCode",
    "Company Name",
    "Target Year",
    "Ticker",
    "Country",
    "NSM Document Type",
    "NSM Headline",
    "Document Date",
    "Submitted Date",
    "Document Format",
    "Download Status",
    "Saved File Name",
    "Saved File Path",
    "File Size (MB)",
    "FCA Download URL",
    "Error / Review Details",
    "Processing Date & Time"
]

# Failed & Review Columns (Sheet 3)
FAILED_REVIEW_COLUMNS = [
    "Serial Number",
    "UniCode",
    "Company Name",
    "Target Year",
    "Ticker",
    "Failure or Review Reason",
    "Error Details",
    "Suggested Action"
]
