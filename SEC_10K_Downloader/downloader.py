"""
downloader.py - Downloads and validates original HTML 10-K filings with unique-code naming.
Includes duplicate detection, atomic file writes, and integrity validation.
"""
import os
import re
from pathlib import Path
from typing import Tuple, Optional
from config import (
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_HTML
)
from logger import get_logger

logger = get_logger()

# Minimal threshold size for a genuine SEC 10-K HTML document (e.g. 2KB)
MIN_HTML_SIZE_BYTES = 2048

def is_valid_html_filing(file_path: str) -> Tuple[bool, str]:
    """
    Validates that a file exists, is non-empty, and contains legitimate SEC HTML content.
    Returns (is_valid, error_reason).
    """
    if not os.path.exists(file_path):
        return False, "File does not exist on disk."

    size = os.path.getsize(file_path)
    if size < MIN_HTML_SIZE_BYTES:
        return False, f"File size too small ({size} bytes). Minimum expected: {MIN_HTML_SIZE_BYTES} bytes."

    # Inspect first 4096 bytes and last 4096 bytes
    try:
        with open(file_path, "rb") as f:
            head = f.read(4096).decode("utf-8", errors="ignore").lower()
            f.seek(max(0, size - 4096))
            tail = f.read(4096).decode("utf-8", errors="ignore").lower()
    except Exception as e:
        return False, f"Could not read file for validation: {str(e)}"

    # Check for SEC Rate Limit / Error markers
    if "request rate threshold exceeded" in head or "rate threshold exceeded" in head:
        return False, "Downloaded content is SEC rate limit error page."
    if "your request originates from an undeclared automated tool" in head:
        return False, "Downloaded content is SEC automated tool block page."
    if "access denied" in head or "<title>403 forbidden" in head:
        return False, "Downloaded content is 403 Forbidden error page."

    # Check for genuine HTML / XML / XBRL tags
    html_markers = ["<html", "<!doctype", "<xml", "<document>", "<ix:", "<body", "<head"]
    has_marker = any(m in head for m in html_markers)
    if not has_marker:
        return False, "File lacks standard HTML / XBRL document markers."

    return True, ""

def check_existing_download(output_dir: str, unique_code: str) -> Tuple[bool, Optional[str]]:
    """
    Checks if a valid report for this unique code already exists in output_dir.
    Returns (exists_and_valid, existing_file_path).
    """
    sanitized_code = re.sub(r'[\\/*?:"<>|]', "_", str(unique_code).strip())
    expected_filename = f"{sanitized_code}.html"
    file_path = os.path.join(output_dir, expected_filename)

    if os.path.exists(file_path):
        is_valid, _ = is_valid_html_filing(file_path)
        if is_valid:
            return True, file_path

    return False, None

async def download_html_report(
    page,
    doc_url: str,
    output_dir: str,
    unique_code: str,
    skip_existing: bool = True
) -> Tuple[str, Optional[str], Optional[str], Optional[str]]:
    """
    Downloads the primary HTML document for a filing.
    
    Args:
        page: Playwright Page instance
        doc_url: URL to the primary 10-K HTML document
        output_dir: Destination directory
        unique_code: Unique report code from input Excel
        skip_existing: If True and valid file exists, skip download
        
    Returns:
        status: One of STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED, STATUS_FAILED_*
        saved_filename: e.g. "BM00125.html"
        saved_filepath: Absolute path to the saved file
        error_details: Error description if any
    """
    os.makedirs(output_dir, exist_ok=True)
    sanitized_code = re.sub(r'[\\/*?:"<>|]', "_", str(unique_code).strip())
    saved_filename = f"{sanitized_code}.html"
    saved_filepath = os.path.abspath(os.path.join(output_dir, saved_filename))
    temp_filepath = saved_filepath + ".tmp"

    # Duplicate prevention check
    if skip_existing:
        exists_and_valid, existing_path = check_existing_download(output_dir, unique_code)
        if exists_and_valid:
            logger.info(f"Skipping {saved_filename} - already downloaded and validated.")
            return STATUS_SKIPPED_ALREADY_DOWNLOADED, saved_filename, existing_path, ""

    logger.info(f"Downloading report for code [{unique_code}] from: {doc_url}")

    try:
        # Fetch document using the active browser session
        # Using page.request or navigating in page to preserve cookies and headers
        response = await page.request.get(doc_url, timeout=45000)
        
        if response.status != 200:
            err = f"HTTP {response.status} when fetching document."
            logger.error(f"Download failed for {unique_code}: {err}")
            return STATUS_FAILED_DOWNLOAD_ERROR, saved_filename, None, err

        content_bytes = await response.body()

        # Atomic write to temporary file
        with open(temp_filepath, "wb") as f:
            f.write(content_bytes)

        # Validate downloaded content
        is_valid, reason = is_valid_html_filing(temp_filepath)
        if not is_valid:
            if os.path.exists(temp_filepath):
                try:
                    os.remove(temp_filepath)
                except OSError:
                    pass
            logger.warning(f"Downloaded report failed validation for {unique_code}: {reason}")
            return STATUS_FAILED_INVALID_HTML, saved_filename, None, reason

        # Move temp file to final destination
        if os.path.exists(saved_filepath):
            try:
                os.remove(saved_filepath)
            except OSError:
                pass
        os.replace(temp_filepath, saved_filepath)

        size_kb = len(content_bytes) / 1024
        logger.info(f"SUCCESS: Report saved as {saved_filename} ({size_kb:.1f} KB) at {saved_filepath}")
        return STATUS_SUCCESS, saved_filename, saved_filepath, ""

    except Exception as e:
        if os.path.exists(temp_filepath):
            try:
                os.remove(temp_filepath)
            except OSError:
                pass
        err = f"Exception during report download: {str(e)}"
        logger.error(f"Download error for {unique_code}: {err}")
        return STATUS_FAILED_DOWNLOAD_ERROR, saved_filename, None, err
