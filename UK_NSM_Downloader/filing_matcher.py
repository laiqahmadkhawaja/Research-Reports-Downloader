"""
filing_matcher.py - Intelligent filing matching and filtering engine for UK FCA NSM.
Strictly isolates Annual Reports (excludes semi-annual, interim, preliminary, quarterly)
and matches the target fiscal year with format prioritization (PDF first).
"""
import re
from typing import List, Dict, Any, Optional, Tuple
from config import (
    STATUS_SUCCESS,
    STATUS_SKIPPED_NO_ANNUAL_REPORT,
    STATUS_SKIPPED_NO_PDF,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_REVIEW_REQUIRED,
    NSM_ARTEFACT_BASE_URL
)
from logger import get_logger

logger = get_logger()

# Disallowed keywords that indicate non-annual reports
EXCLUDED_PATTERNS = [
    r'\bhalf[\s-]?year\b',
    r'\bsemi[\s-]?annual\b',
    r'\binterim\b',
    r'\bquarter(ly)?\b',
    r'\bq[1-4]\b',
    r'\bpreliminary\b',
    r'\bprelim\b',
    r'\btrading\s*(statement|update)\b',
    r'\bdividend\b',
    r'\btotal\s*voting\b',
    r'\bdirector\b',
    r'\bpdmr\b',
    r'\bagm\b',
    r'\bnotice\s*of\b',
    r'\bcircular\b',
    r'\bprospectus\b',
    r'\bholding(s)?\b',
    r'\bpayments\s*to\s*governments\b'
]

COMPILED_EXCLUSIONS = [re.compile(p, re.IGNORECASE) for p in EXCLUDED_PATTERNS]

def is_strict_annual_report(doc_type: str, headline: str) -> bool:
    """
    Returns True if the filing represents a full Annual Report.
    Strictly rejects semi-annual, interim, quarterly, preliminary, and regulatory announcements.
    """
    t_low = (doc_type or "").strip().lower()
    h_low = (headline or "").strip().lower()
    combined = f"{t_low} {h_low}"

    # Must contain indication of Annual / Yearly report
    has_annual_keyword = (
        "annual" in combined or 
        "yearly" in combined or 
        "full year" in combined or 
        "full-year" in combined or
        "year ended" in combined
    )

    if not has_annual_keyword:
        return False

    # Check against excluded patterns
    for pat in COMPILED_EXCLUSIONS:
        if pat.search(combined):
            # Special exception: "Annual Report" might say "Annual Report and Accounts ... Notice of AGM"
            # But if it says "Notice of AGM" alone without "Annual Report", it was rejected above.
            # If it explicitly says "Half-year" or "Interim" or "Preliminary", reject!
            if any(term in pat.pattern for term in ["half", "semi", "interim", "quarter", "preliminary", "trading"]):
                return False

    return True

def extract_filing_year(headline: str, doc_date: str, sub_date: str) -> Optional[int]:
    """
    Extracts the fiscal reporting year from headline, document_date, or submitted_date.
    """
    # 1. Headline: look for "year ended ... 20XX" or "Annual Report 20XX"
    h_text = headline or ""
    
    # Pattern: "year ended [day month] (20[0-2][0-9])"
    m_ended = re.search(r'year\s+ended.*?(\b20[0-2][0-9]\b)', h_text, flags=re.IGNORECASE)
    if m_ended:
        return int(m_ended.group(1))

    # Pattern: "Annual Report [and Accounts] [for] (20[0-2][0-9])"
    m_ar = re.search(r'annual\s+report.*?\b(20[0-2][0-9])\b', h_text, flags=re.IGNORECASE)
    if m_ar:
        return int(m_ar.group(1))

    # Pattern: "(20[0-2][0-9]) Annual Report"
    m_ar_prefix = re.search(r'\b(20[0-2][0-9])\b.*?annual\s+report', h_text, flags=re.IGNORECASE)
    if m_ar_prefix:
        return int(m_ar_prefix.group(1))

    # Pattern: generic 4-digit year in headline
    m_gen = re.search(r'\b(20[0-2][0-9])\b', h_text)
    if m_gen:
        return int(m_gen.group(1))

    # 2. Document Date (e.g. 2020-12-31 or 2020-12-31T00:00:00Z)
    if doc_date and len(doc_date) >= 4:
        try:
            yr = int(doc_date[:4])
            if 2000 <= yr <= 2030:
                return yr
        except ValueError:
            pass

    # 3. Submitted Date (UK companies submit annual reports 1-4 months after fiscal year end)
    if sub_date and len(sub_date) >= 7:
        try:
            s_yr = int(sub_date[:4])
            s_mo = int(sub_date[5:7])
            # If submitted between Jan and June, it is almost always for the previous calendar year
            return (s_yr - 1) if s_mo <= 6 else s_yr
        except ValueError:
            pass

    return None

def score_candidate(candidate: Dict[str, Any], target_year: int) -> int:
    """
    Scores a candidate filing to find the best primary annual report.
    Higher score = better candidate.
    """
    score = 0
    h_low = (candidate.get("headline") or "").lower()
    t_low = (candidate.get("type") or "").lower()
    fmt = (candidate.get("format") or "").upper()
    link = (candidate.get("link") or "").lower()

    # Format Preference: PDF is strongly preferred
    if fmt == "PDF" or link.endswith(".pdf"):
        score += 1000

    # Exact annual report title matches
    if "annual report and accounts" in h_low or "annual report & accounts" in h_low:
        score += 300
    elif "annual report and financial statements" in h_low:
        score += 280
    elif "annual report" in h_low:
        score += 200
    elif "annual financial report" in t_low:
        score += 150

    # Year match precision
    if candidate.get("year") == target_year:
        score += 500

    # Penalize summary or secondary notices
    if "summary" in h_low:
        score -= 100
    if "notice" in h_low:
        score -= 150

    return score

def find_best_filing_for_year(
    filings: List[Dict[str, Any]],
    target_year: int,
    strict_pdf_only: bool = True
) -> Tuple[str, Optional[Dict[str, Any]], str]:
    """
    Finds the best matching annual report filing for the given target year.
    
    Returns:
        status: One of STATUS_SUCCESS, STATUS_SKIPPED_NO_ANNUAL_REPORT, STATUS_SKIPPED_NO_PDF, STATUS_FAILED_COMPANY_NOT_FOUND
        filing_info: Matched filing details dict, or None
        notes: Informative details or error reason
    """
    if not filings:
        return STATUS_FAILED_COMPANY_NOT_FOUND, None, "No financial results returned by FCA NSM for this company."

    annual_candidates = []

    for hit in filings:
        src = hit.get("_source", {})
        doc_type = src.get("type") or ""
        headline = src.get("headline") or ""
        fmt = (src.get("document_format") or "").upper()
        doc_date = src.get("document_date") or ""
        sub_date = src.get("submitted_date") or ""
        link = src.get("download_link") or ""

        if not link:
            continue

        if not is_strict_annual_report(doc_type, headline):
            continue

        extracted_yr = extract_filing_year(headline, doc_date, sub_date)
        if extracted_yr != target_year:
            continue

        is_pdf = (fmt == "PDF" or link.lower().endswith(".pdf"))

        candidate = {
            "id": hit.get("_id"),
            "type": doc_type,
            "headline": headline,
            "format": fmt,
            "doc_date": doc_date,
            "sub_date": sub_date,
            "link": link,
            "is_pdf": is_pdf,
            "year": extracted_yr,
            "download_url": f"{NSM_ARTEFACT_BASE_URL}{link.lstrip('/')}"
        }
        annual_candidates.append(candidate)

    if not annual_candidates:
        return STATUS_SKIPPED_NO_ANNUAL_REPORT, None, f"No Annual Report found on NSM for year {target_year}."

    # Filter for PDFs if strict_pdf_only is enabled
    pdf_candidates = [c for c in annual_candidates if c["is_pdf"]]

    if strict_pdf_only and not pdf_candidates:
        other_fmts = ", ".join(set(c["format"] or "HTML/ESEF" for c in annual_candidates))
        return STATUS_SKIPPED_NO_PDF, None, f"Annual report for {target_year} exists but only in non-PDF format ({other_fmts})."

    # Select best candidate
    pool = pdf_candidates if pdf_candidates else annual_candidates
    pool.sort(key=lambda c: score_candidate(c, target_year), reverse=True)
    best_candidate = pool[0]

    notes = f"Matched Annual Report: '{best_candidate['headline'][:60]}' (Format: {best_candidate['format']})"
    return STATUS_SUCCESS, best_candidate, notes
