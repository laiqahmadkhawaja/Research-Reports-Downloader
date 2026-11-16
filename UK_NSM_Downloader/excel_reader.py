"""
excel_reader.py - Flexible Excel and CSV file reader for UK company input data.
Strictly read-only: never modifies the original file.
"""
import os
import re
import pandas as pd
from typing import List, Dict, Tuple, Optional, Any
from logger import get_logger

logger = get_logger()

# Standard aliases for automatic column detection
COMPANY_ALIASES = [
    "firm_name", "firm name", "company name", "company", "firm", 
    "entity name", "entity", "company_name", "issuer", "organization"
]

YEAR_ALIASES = [
    "year", "reporting year", "reporting_year", "report year", 
    "fiscal year", "fiscal_year", "period year", "period_year", "fy"
]

CODE_ALIASES = [
    "uni_code", "unicode", "unique code", "unique report code", "unique_code", 
    "report code", "report_code", "code", "report id", "id", "identifier"
]

TICKER_ALIASES = [
    "ticker", "symbol", "epic", "ric", "ticker_symbol", "stock_symbol"
]

COUNTRY_ALIASES = [
    "country", "nation", "jurisdiction", "market"
]

def detect_columns(columns: List[str]) -> Tuple[Optional[str], Optional[str], Optional[str], Optional[str], Optional[str]]:
    """
    Attempts to identify Company Name, Reporting Year, Unique Code, Ticker, and Country.
    Returns: (company_col, year_col, code_col, ticker_col, country_col)
    """
    comp_col = None
    year_col = None
    code_col = None
    ticker_col = None
    country_col = None

    normalized = {col: str(col).strip().lower() for col in columns}

    for orig, norm in normalized.items():
        if comp_col is None and norm in COMPANY_ALIASES:
            comp_col = orig
        elif year_col is None and norm in YEAR_ALIASES:
            year_col = orig
        elif code_col is None and norm in CODE_ALIASES:
            code_col = orig
        elif ticker_col is None and norm in TICKER_ALIASES:
            ticker_col = orig
        elif country_col is None and norm in COUNTRY_ALIASES:
            country_col = orig

    # Secondary fuzzy check if exact match wasn't found
    if comp_col is None:
        for orig, norm in normalized.items():
            if "firm" in norm or "company" in norm:
                comp_col = orig
                break

    if year_col is None:
        for orig, norm in normalized.items():
            if "year" in norm:
                year_col = orig
                break

    if code_col is None:
        for orig, norm in normalized.items():
            if "code" in norm or "uni" in norm:
                code_col = orig
                break

    return comp_col, year_col, code_col, ticker_col, country_col

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
    code_col: Optional[str] = None,
    ticker_col: Optional[str] = None,
    country_col: Optional[str] = None
) -> Tuple[List[Dict[str, Any]], Dict[str, str]]:
    """
    Loads records from the input Excel or CSV file.
    Never modifies the original file.
    
    Returns:
        records: List of dicts with keys: serial_num, firm_name, year, uni_code, ticker, country, original_row
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
        raise ValueError(f"Unsupported format: {ext}")

    all_cols = [str(c).strip() for c in df.columns]
    det_comp, det_year, det_code, det_ticker, det_country = detect_columns(all_cols)

    use_comp = company_col or det_comp
    use_year = year_col or det_year
    use_code = code_col or det_code
    use_ticker = ticker_col or det_ticker
    use_country = country_col or det_country

    if not use_comp:
        raise ValueError("Could not detect Company Name column. Please specify column manually.")
    if not use_year:
        raise ValueError("Could not detect Year column. Please specify column manually.")
    if not use_code:
        raise ValueError("Could not detect UniCode column. Please specify column manually.")

    mapping_used = {
        "Company Name": use_comp,
        "Year": use_year,
        "UniCode": use_code,
        "Ticker": use_ticker or "Not Mapped",
        "Country": use_country or "Not Mapped"
    }

    records = []
    serial = 1

    for idx, row in df.iterrows():
        comp_val = str(row[use_comp]).strip() if pd.notna(row[use_comp]) else ""
        year_raw = row[use_year]
        code_raw = row[use_code]

        if not comp_val or pd.isna(year_raw) or pd.isna(code_raw):
            continue

        # Parse year
        try:
            year_val = int(float(year_raw))
        except (ValueError, TypeError):
            continue

        # Parse unicode (handle float representations like 28207.0)
        try:
            if isinstance(code_raw, float) and code_raw.is_integer():
                code_val = str(int(code_raw))
            else:
                code_val = str(code_raw).strip()
                if code_val.endswith(".0"):
                    code_val = code_val[:-2]
        except Exception:
            code_val = str(code_raw).strip()

        ticker_val = str(row[use_ticker]).strip() if use_ticker and pd.notna(row[use_ticker]) else ""
        country_val = str(row[use_country]).strip() if use_country and pd.notna(row[use_country]) else ""

        records.append({
            "serial_num": serial,
            "firm_name": comp_val,
            "year": year_val,
            "uni_code": code_val,
            "ticker": ticker_val,
            "country": country_val,
            "original_row": idx + 2
        })
        serial += 1

    logger.info(f"Loaded {len(records)} valid records from {file_path}")
    return records, mapping_used
