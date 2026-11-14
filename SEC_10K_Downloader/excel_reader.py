"""
excel_reader.py - Flexible Excel and CSV file reader for SEC 10-K input data.
Strictly read-only: never modifies the original file.
"""
import os
import pandas as pd
from typing import List, Dict, Tuple, Optional, Any
from logger import get_logger

logger = get_logger()

# Standard aliases for automatic column detection
COMPANY_ALIASES = [
    "company name", "company", "entity name", "entity", "company_name", 
    "firm name", "firm", "issuer", "organization"
]

YEAR_ALIASES = [
    "reporting year", "year", "reporting_year", "report year", "fiscal year", 
    "fiscal_year", "period year", "period_year", "fy"
]

CODE_ALIASES = [
    "unique code", "unique report code", "unique_code", "report code", 
    "report_code", "code", "report id", "id", "identifier"
]

def detect_columns(columns: List[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Attempts to identify Company Name, Reporting Year, and Unique Code columns from header list.
    """
    comp_col = None
    year_col = None
    code_col = None

    normalized = {col: str(col).strip().lower() for col in columns}

    for orig, norm in normalized.items():
        if comp_col is None and norm in COMPANY_ALIASES:
            comp_col = orig
        elif year_col is None and norm in YEAR_ALIASES:
            year_col = orig
        elif code_col is None and norm in CODE_ALIASES:
            code_col = orig

    # Secondary fuzzy check if exact match wasn't found
    if comp_col is None:
        for orig, norm in normalized.items():
            if "company" in norm or "entity" in norm:
                comp_col = orig
                break

    if year_col is None:
        for orig, norm in normalized.items():
            if "year" in norm:
                year_col = orig
                break

    if code_col is None:
        for orig, norm in normalized.items():
            if "code" in norm or "unique" in norm:
                code_col = orig
                break

    return comp_col, year_col, code_col

def inspect_file_headers(file_path: str) -> List[str]:
    """Returns the list of column headers from an Excel or CSV file without loading entire data."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls", ".xlsm"]:
        df = pd.read_excel(file_path, nrows=2)
    elif ext == ".csv":
        df = pd.read_csv(file_path, nrows=2)
    else:
        raise ValueError(f"Unsupported file format: {ext}. Please select an Excel or CSV file.")

    return [str(c).strip() for c in df.columns]

def load_input_records(
    file_path: str,
    company_col: Optional[str] = None,
    year_col: Optional[str] = None,
    code_col: Optional[str] = None
) -> Tuple[List[Dict[str, Any]], Dict[str, str]]:
    """
    Loads records from the input Excel or CSV file.
    Never modifies the original file.
    
    Returns:
        records: List of dicts with keys: serial_num, company_name, reporting_year, unique_code, original_row
        mapping_used: Dict indicating which columns were mapped
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file not found: {file_path}")

    logger.info(f"Loading input file: {file_path}")
    ext = os.path.splitext(file_path)[1].lower()

    if ext in [".xlsx", ".xls", ".xlsm"]:
        df = pd.read_excel(file_path)
    elif ext == ".csv":
        df = pd.read_csv(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}")

    available_cols = [str(c).strip() for c in df.columns]
    auto_comp, auto_year, auto_code = detect_columns(available_cols)

    final_comp = company_col or auto_comp
    final_year = year_col or auto_year
    final_code = code_col or auto_code

    missing = []
    if not final_comp or final_comp not in df.columns:
        missing.append("Company Name")
    if not final_year or final_year not in df.columns:
        missing.append("Reporting Year")
    if not final_code or final_code not in df.columns:
        missing.append("Unique Report Code")

    if missing:
        err_msg = f"Could not map required columns: {', '.join(missing)}. Available columns: {available_cols}"
        logger.error(err_msg)
        raise ValueError(err_msg)

    records = []
    skipped_rows = 0

    for idx, row in df.iterrows():
        comp_val = row[final_comp]
        year_val = row[final_year]
        code_val = row[final_code]

        # Check for NaN / Empty
        if pd.isna(comp_val) or pd.isna(year_val) or pd.isna(code_val):
            skipped_rows += 1
            continue

        company_str = str(comp_val).strip()
        if not company_str:
            skipped_rows += 1
            continue

        # Parse year as integer
        try:
            reporting_year = int(float(str(year_val).strip()))
        except (ValueError, TypeError):
            logger.warning(f"Row {idx+2}: Invalid year value '{year_val}', skipping.")
            skipped_rows += 1
            continue

        # Unique code as clean string (strip float decimal if integer)
        if isinstance(code_val, float) and code_val.is_integer():
            unique_code = str(int(code_val)).strip()
        else:
            unique_code = str(code_val).strip()

        if not unique_code:
            skipped_rows += 1
            continue

        records.append({
            "serial_num": len(records) + 1,
            "company_name": company_str,
            "reporting_year": reporting_year,
            "unique_code": unique_code,
            "original_row_idx": idx + 2  # 1-based + 1 for header
        })

    mapping_used = {
        "company_col": final_comp,
        "year_col": final_year,
        "code_col": final_code
    }

    logger.info(f"Loaded {len(records)} valid records (Skipped {skipped_rows} invalid/empty rows).")
    return records, mapping_used
