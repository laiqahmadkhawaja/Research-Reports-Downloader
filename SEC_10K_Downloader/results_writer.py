"""
results_writer.py - Generates and updates the 3-sheet SEC_Download_Results.xlsx workbook.
Uses openpyxl with professional styling, summary metrics, and status distribution chart.
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
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
    STATUS_SKIPPED_NO_MATCHING_10K,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_BROWSER_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_HTML,
    STATUS_REVIEW_REQUIRED
)
from logger import get_logger

logger = get_logger()

# Professional Palette
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
    """Provides actionable guidance for failed or review records."""
    if status == STATUS_FAILED_COMPANY_NOT_FOUND:
        return "Verify official company spelling, ticker, or CIK on sec.gov/edgar."
    elif status == STATUS_SKIPPED_NO_MATCHING_10K:
        return "Confirm company was public in this fiscal year; check if Form 20-F or other form was used."
    elif status == STATUS_REVIEW_REQUIRED:
        return "Inspect multiple filings for this year on EDGAR to confirm the correct annual report."
    elif status == STATUS_FAILED_BROWSER_ERROR:
        return "Check internet connection, solve any CAPTCHA prompt, and retry."
    elif status == STATUS_FAILED_DOWNLOAD_ERROR:
        return "Verify SEC access permissions and retry download."
    elif status == STATUS_FAILED_INVALID_HTML:
        return "Check file integrity; redownload if file was truncated or blocked."
    return "Check application log details."

class ResultsWriter:
    """
    Manages creation and incremental updating of SEC_Download_Results.xlsx
    """
    def __init__(self, file_path=DEFAULT_RESULTS_FILE):
        self.file_path = str(file_path)
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

    def write_results(self, records: List[Dict[str, Any]], total_input_count: int = 0):
        """
        Generates or overwrites the 3-sheet results workbook with current data.
        Safe atomic write using .tmp file.
        """
        wb = openpyxl.Workbook()
        
        # 1. Sheet 1: Download Results
        ws_results = wb.active
        ws_results.title = SHEET_DOWNLOAD_RESULTS
        self._build_results_sheet(ws_results, records)

        # 2. Sheet 2: Summary
        ws_summary = wb.create_sheet(title=SHEET_SUMMARY)
        self._build_summary_sheet(ws_summary, records, total_input_count)

        # 3. Sheet 3: Failed & Review Records
        ws_failed = wb.create_sheet(title=SHEET_FAILED_AND_REVIEW)
        self._build_failed_sheet(ws_failed, records)

        # Save workbook atomically
        temp_file = self.file_path + ".tmp"
        try:
            wb.save(temp_file)
            if os.path.exists(self.file_path):
                try:
                    os.remove(self.file_path)
                except OSError:
                    pass
            os.replace(temp_file, self.file_path)
            logger.info(f"Results Excel updated successfully: {self.file_path}")
        except Exception as e:
            logger.error(f"Failed to save results Excel: {e}")
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except OSError:
                    pass

    def _build_results_sheet(self, ws, records: List[Dict[str, Any]]):
        ws.views.sheetView[0].showGridLines = True
        
        # Header Row
        ws.row_dimensions[1].height = 28
        header_fill = PatternFill(start_color=COLOR_NAVY_HEADER, end_color=COLOR_NAVY_HEADER, fill_type="solid")
        
        for col_idx, col_name in enumerate(RESULTS_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = FONT_HEADER
            cell.fill = header_fill
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

        # Data Rows
        for row_idx, r in enumerate(records, 2):
            ws.row_dimensions[row_idx].height = 20
            status = r.get("download_status", "")
            status_fill = get_status_fill(status)

            row_data = [
                r.get("serial_num", row_idx - 1),
                r.get("company_name", ""),
                r.get("reporting_year", ""),
                r.get("unique_code", ""),
                r.get("sec_company_name", ""),
                r.get("filing_type", ""),
                r.get("filing_date", ""),
                r.get("reporting_period", ""),
                r.get("accession_number", ""),
                status,
                r.get("saved_filename", ""),
                r.get("saved_filepath", ""),
                r.get("filing_url", ""),
                r.get("error_details", ""),
                r.get("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            ]

            for col_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = FONT_BODY
                cell.border = BORDER_THIN
                if col_idx in [1, 3, 4, 6, 7, 8, 9, 10]:
                    cell.alignment = ALIGN_CENTER
                else:
                    cell.alignment = ALIGN_LEFT

                # Highlight Status column
                if col_idx == 10 and status_fill:
                    cell.fill = status_fill
                    cell.font = FONT_BODY_BOLD

        # Adjust column widths
        self._auto_fit_columns(ws, min_width=12, max_width=45)

    def _build_summary_sheet(self, ws, records: List[Dict[str, Any]], total_input_count: int):
        ws.views.sheetView[0].showGridLines = True
        
        total_records = max(len(records), total_input_count)
        success = sum(1 for r in records if r.get("download_status") == STATUS_SUCCESS)
        skipped_dl = sum(1 for r in records if r.get("download_status") == STATUS_SKIPPED_ALREADY_DOWNLOADED)
        skipped_no_10k = sum(1 for r in records if r.get("download_status") == STATUS_SKIPPED_NO_MATCHING_10K)
        failed = sum(1 for r in records if "FAILED" in r.get("download_status", ""))
        review = sum(1 for r in records if r.get("download_status") == STATUS_REVIEW_REQUIRED)
        total_saved_html = success + skipped_dl
        completion_pct = (len(records) / total_records * 100) if total_records > 0 else 0.0

        # Title
        ws.merge_cells("A1:D1")
        title_cell = ws["A1"]
        title_cell.value = "SEC EDGAR 10-K Download Automation Summary"
        title_cell.font = Font(name="Segoe UI", size=16, bold=True, color="003366")
        title_cell.alignment = ALIGN_LEFT
        ws.row_dimensions[1].height = 35

        # Subtitle Timestamp
        ws["A2"] = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ws["A2"].font = Font(name="Segoe UI", size=10, italic=True, color="595959")
        ws.row_dimensions[2].height = 20

        # Summary Table Headers
        headers = ["Metric", "Count", "Percentage"]
        header_fill = PatternFill(start_color=COLOR_NAVY_HEADER, end_color=COLOR_NAVY_HEADER, fill_type="solid")
        ws.row_dimensions[4].height = 24
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = FONT_HEADER
            cell.fill = header_fill
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

        stats_data = [
            ("Total Records", total_records, "100.0%"),
            ("Successfully Downloaded", success, f"{(success/total_records*100) if total_records else 0:.1f}%"),
            ("Skipped - Already Downloaded", skipped_dl, f"{(skipped_dl/total_records*100) if total_records else 0:.1f}%"),
            ("Skipped - No Matching 10-K", skipped_no_10k, f"{(skipped_no_10k/total_records*100) if total_records else 0:.1f}%"),
            ("Failed Downloads", failed, f"{(failed/total_records*100) if total_records else 0:.1f}%"),
            ("Review Required", review, f"{(review/total_records*100) if total_records else 0:.1f}%"),
            ("Total HTML Files Saved", total_saved_html, f"{(total_saved_html/total_records*100) if total_records else 0:.1f}%"),
            ("Overall Completion Percentage", f"{completion_pct:.1f}%", "-")
        ]

        for idx, (label, val, pct) in enumerate(stats_data, 5):
            ws.row_dimensions[idx].height = 20
            c1 = ws.cell(row=idx, column=1, value=label)
            c2 = ws.cell(row=idx, column=2, value=val)
            c3 = ws.cell(row=idx, column=3, value=pct)

            for c in (c1, c2, c3):
                c.font = FONT_BODY
                c.border = BORDER_THIN
            c1.alignment = ALIGN_LEFT
            c2.alignment = ALIGN_RIGHT
            c3.alignment = ALIGN_CENTER

            if label == "Successfully Downloaded":
                c1.fill = c2.fill = c3.fill = PatternFill(start_color=COLOR_SUCCESS, end_color=COLOR_SUCCESS, fill_type="solid")
            elif "Failed" in label:
                c1.fill = c2.fill = c3.fill = PatternFill(start_color=COLOR_FAILED, end_color=COLOR_FAILED, fill_type="solid")
            elif "Review" in label:
                c1.fill = c2.fill = c3.fill = PatternFill(start_color=COLOR_REVIEW, end_color=COLOR_REVIEW, fill_type="solid")
            elif label == "Overall Completion Percentage":
                c1.font = c2.font = c3.font = FONT_BODY_BOLD

        # Chart Table for Status Distribution
        chart_start_row = 15
        ws.cell(row=chart_start_row, column=1, value="Status Category").font = FONT_BODY_BOLD
        ws.cell(row=chart_start_row, column=2, value="Records").font = FONT_BODY_BOLD
        
        chart_data = [
            ("Success", success),
            ("Already Downloaded", skipped_dl),
            ("No Matching 10-K", skipped_no_10k),
            ("Failed", failed),
            ("Review Required", review),
        ]
        
        for i, (cat, cnt) in enumerate(chart_data, 1):
            ws.cell(row=chart_start_row + i, column=1, value=cat)
            ws.cell(row=chart_start_row + i, column=2, value=cnt)

        # Create Bar Chart
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Download Status Distribution"
        chart.y_axis.title = "Record Count"
        chart.x_axis.title = "Status"
        chart.height = 10
        chart.width = 16

        data_ref = Reference(ws, min_col=2, min_row=chart_start_row, max_row=chart_start_row + len(chart_data))
        cats_ref = Reference(ws, min_col=1, min_row=chart_start_row + 1, max_row=chart_start_row + len(chart_data))
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.legend = None

        ws.add_chart(chart, "E4")
        self._auto_fit_columns(ws, min_width=15, max_width=35)

    def _build_failed_sheet(self, ws, records: List[Dict[str, Any]]):
        ws.views.sheetView[0].showGridLines = True
        
        # Header Row
        ws.row_dimensions[1].height = 28
        header_fill = PatternFill(start_color=COLOR_NAVY_HEADER, end_color=COLOR_NAVY_HEADER, fill_type="solid")
        for col_idx, col_name in enumerate(FAILED_REVIEW_COLUMNS, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.font = FONT_HEADER
            cell.fill = header_fill
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_THIN

        # Filter records requiring attention (Failed, Review Required, and Skipped No 10-K)
        attention_records = [
            r for r in records if (
                "FAILED" in r.get("download_status", "") or 
                r.get("download_status") in [STATUS_REVIEW_REQUIRED, STATUS_SKIPPED_NO_MATCHING_10K]
            )
        ]

        for row_idx, r in enumerate(attention_records, 2):
            ws.row_dimensions[row_idx].height = 20
            status = r.get("download_status", "")
            err_det = r.get("error_details", "")
            action = suggest_action_for_record(status, err_det)
            status_fill = get_status_fill(status)

            row_data = [
                r.get("serial_num", row_idx - 1),
                r.get("company_name", ""),
                r.get("reporting_year", ""),
                r.get("unique_code", ""),
                status,
                err_det,
                action
            ]

            for col_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = FONT_BODY
                cell.border = BORDER_THIN
                if col_idx in [1, 3, 4, 5]:
                    cell.alignment = ALIGN_CENTER
                else:
                    cell.alignment = ALIGN_LEFT

                if col_idx == 5 and status_fill:
                    cell.fill = status_fill
                    cell.font = FONT_BODY_BOLD

        self._auto_fit_columns(ws, min_width=14, max_width=50)

    def _auto_fit_columns(self, ws, min_width=10, max_width=50):
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or "")
                if "\n" in val:
                    val = max(val.split("\n"), key=len)
                max_len = max(max_len, len(val))
            ws.column_dimensions[col_letter].width = max(min(max_len + 3, max_width), min_width)
