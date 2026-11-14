"""
filing_matcher.py - Company name matching, filing type filtering, and reporting year matching logic.
"""
import re
from difflib import SequenceMatcher
from typing import List, Dict, Tuple, Optional, Any
from logger import get_logger

logger = get_logger()

# Common legal endings to normalize for matching
LEGAL_SUFFIXES = [
    r"\binc\b\.?", r"\bincorporated\b\.?", r"\bcorp\b\.?", r"\bcorporation\b\.?",
    r"\bco\b\.?", r"\bcompany\b\.?", r"\bllc\b\.?", r"\bltd\b\.?", r"\bplc\b\.?",
    r"\bholdings?\b\.?", r"\bgroup\b\.?", r"\benterprises?\b\.?", r"\btechnologies\b\.?",
    r"\bclass [a-z]\b\.?"
]

def clean_company_name(name: str) -> str:
    """Normalizes company name by stripping legal suffixes, punctuation, and extra spaces."""
    if not name:
        return ""
    cleaned = name.lower()
    for suffix in LEGAL_SUFFIXES:
        cleaned = re.sub(suffix, " ", cleaned)
    cleaned = re.sub(r"[^\w\s]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned

def parse_suggestion(suggestion_text: str) -> Dict[str, str]:
    """
    Parses a suggestion string from SEC EDGAR hints table.
    Example: "Apple Inc. (AAPL)\tCIK 0000320193" or "Owens Corning (OC)  CIK 0001370946"
    """
    text = suggestion_text.strip()
    cik_match = re.search(r"CIK\s*(\d{10}|\d+)", text, re.IGNORECASE)
    cik = cik_match.group(1).zfill(10) if cik_match else ""

    # Ticker in parentheses e.g. (AAPL)
    ticker_match = re.search(r"\(([A-Z]{1,5})\)", text)
    ticker = ticker_match.group(1) if ticker_match else ""

    # Remove CIK and ticker from name
    name_part = re.sub(r"CIK\s*\d+", "", text, flags=re.IGNORECASE)
    name_part = re.sub(r"\([A-Z]{1,5}\)", "", name_part)
    name_part = name_part.strip("\t -")

    return {
        "full_text": text,
        "company_name": name_part,
        "ticker": ticker,
        "cik": cik
    }

def match_company_suggestion(
    target_name: str,
    suggestions: List[str]
) -> Tuple[Optional[int], Optional[str], float, bool]:
    """
    Evaluates suggestion rows to find the best match for the target company name.
    
    Returns:
        best_index: Index in suggestions list (or None if no acceptable match)
        selected_name: Selected SEC company name
        confidence: Match score (0.0 to 1.0)
        is_uncertain: True if confidence is borderline or ambiguous, requiring review
    """
    if not suggestions:
        return None, None, 0.0, True

    parsed_list = [parse_suggestion(s) for s in suggestions]
    target_clean = clean_company_name(target_name)

    best_idx = None
    best_score = 0.0
    scores = []

    for i, p in enumerate(parsed_list):
        sug_clean = clean_company_name(p["company_name"])

        # 1. Exact normalized match
        if target_clean == sug_clean:
            return i, p["company_name"], 1.0, False

        # 2. Token set / sequence ratio
        ratio = SequenceMatcher(None, target_clean, sug_clean).ratio()

        # Bonus if target is exact prefix or substring of suggestion or vice-versa
        if target_clean in sug_clean or sug_clean in target_clean:
            len_ratio = min(len(target_clean), len(sug_clean)) / max(len(target_clean), len(sug_clean))
            ratio = max(ratio, 0.70 + 0.30 * len_ratio)

        scores.append((i, ratio, p["company_name"]))
        if ratio > best_score:
            best_score = ratio
            best_idx = i

    # Sort descending by score
    scores.sort(key=lambda x: x[1], reverse=True)

    if not scores or best_score < 0.60:
        # Confidence too low
        return None, None, best_score, True

    # Ambiguity check: if 2nd best is very close to top best
    is_uncertain = False
    if len(scores) > 1:
        if (scores[0][1] - scores[1][1]) < 0.05 and scores[0][1] < 0.95:
            is_uncertain = True

    if best_score < 0.85:
        is_uncertain = True

    return best_idx, parsed_list[best_idx]["company_name"], best_score, is_uncertain

def is_valid_10k_form(form_text: str) -> bool:
    """
    Strict filing type filter.
    Returns True ONLY if form is '10-K' (excludes '10-K/A', '10-KT', 'NT 10-K', etc.)
    """
    if not form_text:
        return False

    clean_form = form_text.strip().split()[0].upper()

    # Rule: Must be exactly "10-K"
    if clean_form == "10-K":
        # Extra guard: check if anywhere says 10-K/A
        if "/A" in form_text.upper():
            return False
        return True

    return False

def extract_accession_number(url: str) -> str:
    """
    Extracts and formats accession number from SEC filing URL.
    e.g. https://www.sec.gov/Archives/edgar/data/1370946/000137094615000061/oc-20141231x10k.htm
    Returns formatted '0001370946-15-000061'
    """
    if not url:
        return ""
    # Look for 18 consecutive digits in URL path
    match = re.search(r"/(\d{18})/", url)
    if match:
        raw = match.group(1)
        # Format: 10 digits - 2 digits - 6 digits
        return f"{raw[:10]}-{raw[10:12]}-{raw[12:]}"

    # Or already formatted
    match_hyphen = re.search(r"(\d{10}-\d{2}-\d{6})", url)
    if match_hyphen:
        return match_hyphen.group(1)

    return ""

def match_filing_for_reporting_year(
    filings: List[Dict[str, Any]],
    target_year: int
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Matches a filing record with the target reporting year.
    
    Requirements:
    - Exclude all 10-K/A amended filings.
    - Strictly compare filing_type == "10-K".
    - Reporting year must match the reporting period end date.
    - If multiple original 10-Ks match the same year, select the original filing (earliest filed date).
    - If reporting year cannot be confidently identified, flag as 'REVIEW REQUIRED'.
    
    Returns:
        matched_filing: Dict of matched filing or None
        review_reason: Reason string if review required or why skipped
    """
    candidates = []

    for f in filings:
        form = f.get("form", "")
        if not is_valid_10k_form(form):
            continue

        reporting_str = f.get("reporting", "") or ""
        filed_str = f.get("filed", "") or ""

        # Extract year from reporting period date (e.g. "2014-12-31" -> 2014)
        rep_year = None
        if reporting_str and len(reporting_str) >= 4:
            try:
                rep_year = int(reporting_str[:4])
            except ValueError:
                rep_year = None

        filed_year = None
        if filed_str and len(filed_str) >= 4:
            try:
                filed_year = int(filed_str[:4])
            except ValueError:
                filed_year = None

        # Handle January/February fiscal year ends
        # E.g., period 2015-01-31 is often designated as FY 2014 or FY 2015 depending on company
        # Check primary match: reporting period year == target_year
        if rep_year == target_year:
            candidates.append((f, "exact_period_year"))
        elif rep_year is None and filed_year is not None:
            # If reporting period is missing, can we infer?
            # Standard 10-K is filed in Q1 of following year (e.g. FY 2014 filed early 2015)
            if filed_year == target_year + 1 or filed_year == target_year:
                candidates.append((f, "inferred_filed_year"))

    if not candidates:
        return None, f"No original 10-K found with reporting year {target_year}."

    # Filter to exact_period_year first if available
    exact_candidates = [c[0] for c in candidates if c[1] == "exact_period_year"]
    if exact_candidates:
        if len(exact_candidates) == 1:
            return exact_candidates[0], None
        else:
            # Multiple original 10-Ks for the same reporting period!
            # Sort by filing date ascending (earliest original filing)
            exact_candidates.sort(key=lambda x: x.get("filed", ""))
            # Flag for review due to multiple filings for the same period
            return exact_candidates[0], "Multiple original 10-K filings found for reporting year; selected earliest original filing."

    # Only inferred candidates available
    inferred_candidates = [c[0] for c in candidates if c[1] == "inferred_filed_year"]
    if len(inferred_candidates) == 1:
        # Needs review because period date was missing
        return inferred_candidates[0], "Reporting period date missing in SEC table; matched based on filing date year."

    return None, f"Ambiguous 10-K match for reporting year {target_year}."
