"""
results_writer.py - Generates and updates the 3-sheet UK_NSM_Download_Results.xlsx workbook.
Uses openpyxl with professional styling, summary KPI metrics, and action recommendations.
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional
from datetime import datetime
from config import (
    DEFAULT_RESULTS_FILE,
    SHEET_DOWNLOAD_RESULTS,
    SHEET_SUMMARY,
    SHEET_FAILED_AND_REVIEW,
    RESULTS_COLUMNS,
    FAILED_REVIEW_COLUMNS,
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
from logger import get_logger

logger = get_logger()

# Palette
COLOR_NAVY_HEADER = "003366"
COLOR_WHITE = "FFFFFF"
COLOR_SUCCESS = "E2EFDA"  # Soft green
COLOR_SKIPPED = "DDEBF7"  # Soft light blue
COLOR_FAILED = "FCE4D6"   # Soft coral/red
COLOR_REVIEW = "FFF2CC"   # Soft yellow
COLOR_BORDER = "D9D9D9"

FONT_HEADER = Font(name="Segoe UI", size=11, bold=True, color=COLOR_WHITE)
FONT_BODY = Font(name="Segoe UI", size=10)
FONT_BODY_BOLD = Font(name="Segoe UI", size=10, bold=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

BORDER_THIN = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER)
)

def get_status_fill(status: str) -> Optional[PatternFill]:
    if status == STATUS_SUCCESS:
        return PatternFill(start_color=COLOR_SUCCESS, end_color=COLOR_SUCCESS, fill_type="solid")
    elif "SKIPPED" in status:
        return PatternFill(start_color=COLOR_SKIPPED, end_color=COLOR_SKIPPED, fill_type="solid")
    elif "FAILED" in status:
        return PatternFill(start_color=COLOR_FAILED, end_color=COLOR_FAILED, fill_type="solid")
    elif "REVIEW" in status:
        return PatternFill(start_color=COLOR_REVIEW, end_color=COLOR_REVIEW, fill_type="solid")
    return None

def suggest_action_for_record(status: str, error_details: str) -> str:
    """Provides actionable advice for failed or skipped records."""
    if status == STATUS_FAILED_COMPANY_NOT_FOUND:
        return "Verify official company name or LEI on FCA NSM portal (data.fca.org.uk)."
    elif status == STATUS_SKIPPED_NO_ANNUAL_REPORT:
        return "Check if company filed annual reports under a parent holding name or different fiscal year."
    elif status == STATUS_SKIPPED_NO_PDF:
        return "Filing exists only as ESEF/iXBRL/HTML package. Uncheck 'Strict PDF Only' if you wish to download XHTML."
    elif status == STATUS_REVIEW_REQUIRED:
        return "Multiple filings detected; inspect NSM portal directly."
    elif status == STATUS_FAILED_NETWORK_ERROR:
        return "Check internet connection and retry."
    elif status == STATUS_FAILED_DOWNLOAD_ERROR:
        return "Verify download URL and server response on FCA portal."
    elif status == STATUS_FAILED_INVALID_PDF:
        return "File was corrupted or truncated; re-download to verify."
    return "Check application log details."

class ResultsWriter:
    """Manages creation and incremental updating of UK_NSM_Download_Results.xlsx"""
    def __init__(self, file_path=DEFAULT_RESULTS_FILE):
        self.file_path = str(file_path)
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

    def write_results(self, records: List[Dict[str, Any]], total_input_count: int = 0):
        """Generates or overwrites the 3-sheet results workbook with current data."""
        wb = openpyxl.Workbook()

        # Sheet 1: Download Results
        ws_results = wb.active
        ws_results.title = SHEET_DOWNLOAD_RESULTS
        self._build_results_sheet(ws_results, records)

        # Sheet 2: Summary
        ws_summary = wb.create_sheet(title=SHEET_SUMMARY)
        self._build_summary_sheet(ws_summary, records, total_input_count)

        # Sheet 3: Failed & Review Records
        ws_failed = wb.create_sheet(title=SHEET_FAILED_AND_REVIEW)
        self._build_failed_sheet(ws_failed, records)

        # Atomic Save
        temp_path = f"{self.file_path}.tmp"
        try:
            wb.save(temp_path)
            if os.path.exists(self.file_path):
                os.replace(temp_path, self.file_path)
            else:
                os.rename(temp_path, self.file_path)
            logger.info(f"Results report successfully written to {self.file_path}")
        except Exception as e:
            logger.error(f"Failed to save results workbook: {e}")
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    def _build_results_sheet(self, ws, records: List[Dict[str, Any]]):
        ws.views.sheetView[0].showGridLines = True
        ws.row_dimensions[1].height = 28

        header_fill = PatternFill(start_color=COLOR_NAVY_HEADER, end_color=COLOR_NAVY_HEADER, fill_type="solid")
        for col_idx, col_name in enumerate(RESULTS_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = FONT_HEADER
            cell.fill = header_fill
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

        for row_idx, r in enumerate(records, 2):
            ws.row_dimensions[row_idx].height = 20
            status = r.get("status", "")
            status_fill = get_status_fill(status)

            row_values = [
                r.get("serial_num", row_idx - 1),
                r.get("uni_code", ""),
                r.get("firm_name", ""),
                r.get("year", ""),
                r.get("ticker", ""),
                r.get("country", ""),
                r.get("nsm_doc_type", ""),
                r.get("nsm_headline", ""),
                r.get("doc_date", ""),
                r.get("sub_date", ""),
                r.get("doc_format", ""),
                status,
                r.get("saved_filename", ""),
                r.get("saved_filepath", ""),
                r.get("file_size_mb", 0.0),
                r.get("download_url", ""),
                r.get("error_details", ""),
                r.get("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ]

            for col_idx, val in enumerate(row_values, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = FONT_BODY
                cell.border = BORDER_THIN

                # Status column gets special color fill
                if col_idx == 12 and status_fill:
                    cell.fill = status_fill
                    cell.font = FONT_BODY_BOLD

                # Alignments
                if col_idx in [1, 2, 4, 9, 10, 11, 12, 18]:
                    cell.alignment = ALIGN_CENTER
                elif col_idx == 15:
                    cell.alignment = ALIGN_RIGHT
                else:
                    cell.alignment = ALIGN_LEFT

        ws.auto_filter.ref = f"A1:{get_column_letter(len(RESULTS_COLUMNS))}{max(len(records) + 1, 1)}"
        self._auto_fit_columns(ws)

    def _build_summary_sheet(self, ws, records: List[Dict[str, Any]], total_input_count: int):
        ws.views.sheetView[0].showGridLines = True
        total = total_input_count or len(records)
        success = sum(1 for r in records if r.get("status") == STATUS_SUCCESS)
        skipped_dl = sum(1 for r in records if r.get("status") == STATUS_SKIPPED_ALREADY_DOWNLOADED)
        skipped_no_pdf = sum(1 for r in records if r.get("status") == STATUS_SKIPPED_NO_PDF)
        skipped_no_ar = sum(1 for r in records if r.get("status") == STATUS_SKIPPED_NO_ANNUAL_REPORT)
        failed = sum(1 for r in records if "FAILED" in r.get("status", ""))
        review = sum(1 for r in records if r.get("status") == STATUS_REVIEW_REQUIRED)
        total_effective = success + skipped_dl
        rate = (total_effective / total * 100.0) if total > 0 else 0.0

        title_cell = ws.cell(row=2, column=2, value="UK FCA NSM Annual Report Downloader - Summary Metrics")
        title_cell.font = Font(name="Segoe UI", size=15, bold=True, color="003366")

        metrics_table = [
            ("Metric Description", "Count / Value"),
            ("Total Records in Input", total),
            ("Reports Successfully Downloaded (New)", success),
            ("Reports Skipped (Already Downloaded)", skipped_dl),
            ("Total Available Reports (New + Existing)", total_effective),
            ("Skipped - No PDF Version (Only XHTML/ESEF)", skipped_no_pdf),
            ("Skipped - No Annual Report Found for Year", skipped_no_ar),
            ("Failed (Company Not Found / Download Errors)", failed),
            ("Review Required", review),
            ("Overall Success Rate (%)", f"{rate:.1f}%"),
            ("Report Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        ]

        header_fill = PatternFill(start_color=COLOR_NAVY_HEADER, end_color=COLOR_NAVY_HEADER, fill_type="solid")
        for i, (m_desc, m_val) in enumerate(metrics_table, 4):
            c1 = ws.cell(row=i, column=2, value=m_desc)
            c2 = ws.cell(row=i, column=3, value=m_val)

            if i == 4:
                c1.fill = header_fill
                c2.fill = header_fill
                c1.font = FONT_HEADER
                c2.font = FONT_HEADER
                c1.alignment = ALIGN_LEFT
                c2.alignment = ALIGN_CENTER
            else:
                c1.font = FONT_BODY_BOLD if "Total" in m_desc or "Rate" in m_desc else FONT_BODY
                c2.font = FONT_BODY_BOLD
                c1.alignment = ALIGN_LEFT
                c2.alignment = ALIGN_CENTER

            c1.border = BORDER_THIN
            c2.border = BORDER_THIN

        self._auto_fit_columns(ws)

    def _build_failed_sheet(self, ws, records: List[Dict[str, Any]]):
        ws.views.sheetView[0].showGridLines = True
        ws.row_dimensions[1].height = 28

        header_fill = PatternFill(start_color="800000", end_color="800000", fill_type="solid")
        for col_idx, col_name in enumerate(FAILED_REVIEW_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = FONT_HEADER
            cell.fill = header_fill
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

        row_idx = 2
        for r in records:
            status = r.get("status", "")
            if status in [STATUS_SUCCESS, STATUS_SKIPPED_ALREADY_DOWNLOADED]:
                continue

            ws.row_dimensions[row_idx].height = 20
            action = suggest_action_for_record(status, r.get("error_details", ""))

            row_values = [
                r.get("serial_num", row_idx - 1),
                r.get("uni_code", ""),
                r.get("firm_name", ""),
                r.get("year", ""),
                r.get("ticker", ""),
                status,
                r.get("error_details", ""),
                action
            ]

            for col_idx, val in enumerate(row_values, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = FONT_BODY
                cell.border = BORDER_THIN
                if col_idx == 6:
                    cell.font = FONT_BODY_BOLD
                    cell.fill = PatternFill(start_color=COLOR_FAILED, end_color=COLOR_FAILED, fill_type="solid")
                cell.alignment = ALIGN_CENTER if col_idx in [1, 2, 4, 5, 6] else ALIGN_LEFT

            row_idx += 1

        ws.auto_filter.ref = f"A1:{get_column_letter(len(FAILED_REVIEW_COLUMNS))}{max(row_idx - 1, 1)}"
        self._auto_fit_columns(ws)

    def _auto_fit_columns(self, ws):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)
