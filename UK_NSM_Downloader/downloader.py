"""
downloader.py - Downloads and validates genuine PDF Annual Reports with company-wise folders
and UniCode naming. Supports atomic file writes, integrity verification, and resume skips.
"""
import os
import re
import time
import requests
from typing import Tuple, Optional
from config import (
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_PDF,
    DEFAULT_USER_AGENT,
    DEFAULT_CHUNK_SIZE,
    DEFAULT_TIMEOUT_SEC,
    DEFAULT_MAX_RETRIES
)
from logger import get_logger

logger = get_logger()

MIN_PDF_SIZE_BYTES = 5120  # 5 KB minimum threshold for an Annual Report PDF

def sanitize_name(name: str) -> str:
    """Sanitizes strings for safe Windows directory and file names."""
    sanitized = re.sub(r'[\\/*?:"<>|]', '_', str(name).strip())
    # Remove leading or trailing periods and whitespace
    sanitized = re.sub(r'^\.+|\.+$', '', sanitized).strip()
    return sanitized or "Unknown"

def is_valid_pdf_file(file_path: str) -> Tuple[bool, str]:
    """
    Validates that a file on disk is a genuine non-corrupted PDF.
    Returns (is_valid, reason).
    """
    if not os.path.exists(file_path):
        return False, "File does not exist."

    size = os.path.getsize(file_path)
    if size < MIN_PDF_SIZE_BYTES:
        return False, f"File size too small ({size} bytes). Expected >= {MIN_PDF_SIZE_BYTES} bytes."

    try:
        with open(file_path, "rb") as f:
            header = f.read(1024)
            # Standard PDF files start with %PDF- (or within first 1024 bytes)
            if b"%PDF-" not in header:
                return False, "File header lacks standard %PDF- magic bytes (likely HTML error page)."
    except Exception as e:
        return False, f"Cannot read file header: {e}"

    return True, ""

def check_existing_pdf(company_dir: str, uni_code: str) -> Tuple[bool, Optional[str], float]:
    """
    Checks if a valid PDF report for this UniCode already exists in the company directory.
    Returns (exists_and_valid, existing_file_path, size_mb).
    """
    safe_code = sanitize_name(uni_code)
    expected_filename = f"{safe_code}.pdf"
    file_path = os.path.join(company_dir, expected_filename)

    if os.path.exists(file_path):
        is_valid, _ = is_valid_pdf_file(file_path)
        if is_valid:
            size_mb = round(os.path.getsize(file_path) / (1024 * 1024), 2)
            return True, file_path, size_mb

    return False, None, 0.0

def download_pdf_report(
    doc_url: str,
    output_dir: str,
    firm_name: str,
    uni_code: str,
    skip_existing: bool = True,
    max_retries: int = DEFAULT_MAX_RETRIES,
    progress_callback = None
) -> Tuple[str, Optional[str], Optional[str], float, Optional[str]]:
    """
    Downloads the Annual Report PDF into a company-wise folder named after the company,
    renamed to <uni_code>.pdf.
    
    Returns:
        status: STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED, or STATUS_FAILED_*
        saved_filename: e.g. "28207.pdf"
        saved_filepath: Absolute path to saved file
        file_size_mb: Size of file in MB
        error_details: Reason if failed
    """
    safe_firm = sanitize_name(firm_name)
    safe_code = sanitize_name(uni_code)
    company_dir = os.path.join(output_dir, safe_firm)
    os.makedirs(company_dir, exist_ok=True)

    target_filename = f"{safe_code}.pdf"
    target_filepath = os.path.join(company_dir, target_filename)
    temp_filepath = os.path.join(company_dir, f"{safe_code}.pdf.tmp")

    # 1. Skip Existing check
    if skip_existing:
        exists, existing_path, size_mb = check_existing_pdf(company_dir, safe_code)
        if exists:
            logger.info(f"Skipping download: {safe_firm}/{target_filename} already exists ({size_mb} MB).")
            return STATUS_SKIPPED_ALREADY_DOWNLOADED, target_filename, existing_path, size_mb, ""

    headers = {
        "User-Agent": DEFAULT_USER_AGENT,
        "Accept": "*/*",
        "Referer": "https://data.fca.org.uk/"
    }

    last_error = ""

    for attempt in range(1, max_retries + 1):
        try:
            logger.info(f"Downloading PDF [{attempt}/{max_retries}] for {safe_firm} (Code: {safe_code}) from {doc_url}")
            
            with requests.get(doc_url, headers=headers, stream=True, timeout=DEFAULT_TIMEOUT_SEC) as resp:
                if resp.status_code != 200:
                    last_error = f"HTTP Error {resp.status_code} downloading from FCA NSM."
                    logger.warning(f"Download failed with HTTP {resp.status_code}, retrying...")
                    time.sleep(attempt * 1.5)
                    continue

                total_length = resp.headers.get('content-length')
                total_bytes = int(total_length) if total_length and total_length.isdigit() else 0
                downloaded_bytes = 0

                with open(temp_filepath, "wb") as f_out:
                    for chunk in resp.iter_content(chunk_size=DEFAULT_CHUNK_SIZE):
                        if chunk:
                            f_out.write(chunk)
                            downloaded_bytes += len(chunk)
                            if progress_callback and total_bytes > 0:
                                progress_callback(downloaded_bytes, total_bytes)

            # 2. Validate PDF file
            is_valid, validation_err = is_valid_pdf_file(temp_filepath)
            if not is_valid:
                if os.path.exists(temp_filepath):
                    try:
                        os.remove(temp_filepath)
                    except Exception:
                        pass
                last_error = f"PDF Integrity check failed: {validation_err}"
                logger.warning(f"{last_error} (attempt {attempt}/{max_retries})")
                time.sleep(attempt * 1.5)
                continue

            # 3. Atomic rename from .tmp to final .pdf
            if os.path.exists(target_filepath):
                try:
                    os.remove(target_filepath)
                except Exception:
                    pass

            os.rename(temp_filepath, target_filepath)
            final_size_mb = round(os.path.getsize(target_filepath) / (1024 * 1024), 2)
            logger.info(f"SUCCESS: Saved {target_filename} to {safe_firm}/ ({final_size_mb} MB)")
            return STATUS_SUCCESS, target_filename, target_filepath, final_size_mb, ""

        except Exception as e:
            last_error = f"Exception during download: {str(e)}"
            logger.warning(f"Error during attempt {attempt}/{max_retries}: {e}")
            time.sleep(attempt * 1.5)
        finally:
            if os.path.exists(temp_filepath):
                try:
                    os.remove(temp_filepath)
                except Exception:
                    pass

    return STATUS_FAILED_DOWNLOAD_ERROR, target_filename, None, 0.0, last_error
