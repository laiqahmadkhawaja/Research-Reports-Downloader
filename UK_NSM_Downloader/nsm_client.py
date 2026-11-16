"""
nsm_client.py - High-performance client for FCA National Storage Mechanism (NSM) API.
Queries the official FCA Elasticsearch backend directly with caching, retries, and rate limiting.
"""
import time
import re
import requests
from typing import List, Dict, Any, Optional
from config import (
    NSM_SEARCH_API_URL,
    DEFAULT_API_HEADERS,
    DEFAULT_REQUEST_DELAY,
    DEFAULT_MAX_RETRIES,
    DEFAULT_TIMEOUT_SEC
)
from logger import get_logger

logger = get_logger()

class NsmApiClient:
    """
    Communicates with the FCA NSM Elasticsearch API.
    Features:
      - Memory caching per company (resolves all years for a company with 1 search)
      - Corporate suffix stripping for broader matching
      - Exponential backoff retry logic
      - Polite rate limiting
    """
    def __init__(self, delay_sec: float = DEFAULT_REQUEST_DELAY, max_retries: int = DEFAULT_MAX_RETRIES):
        self.delay_sec = delay_sec
        self.max_retries = max_retries
        self.session = requests.Session()
        self.session.headers.update(DEFAULT_API_HEADERS)
        self._company_cache: Dict[str, List[Dict[str, Any]]] = {}
        self._last_request_time = 0.0

    def _wait_rate_limit(self):
        """Enforces delay between consecutive requests."""
        if self.delay_sec <= 0:
            return
        elapsed = time.time() - self._last_request_time
        if elapsed < self.delay_sec:
            time.sleep(self.delay_sec - elapsed)
        self._last_request_time = time.time()

    @staticmethod
    def clean_company_name(name: str) -> str:
        """Removes corporate entity suffixes to maximize search recall."""
        # e.g. "Anglo American PLC" -> "Anglo American"
        # "Barratt Developments P L C" -> "Barratt Developments"
        # "CNH Industrial NV" -> "CNH Industrial"
        cleaned = re.sub(
            r'\b(PLC|P\s*L\s*C|NV|N\.V\.|LTD|LIMITED|LLC|CORP|CORPORATION|INC|INCORPORATED|HOLDINGS?|GROUP)\b',
            '',
            name,
            flags=re.IGNORECASE
        ).strip()
        # Clean double spaces or trailing punctuation
        cleaned = re.sub(r'\s+', ' ', cleaned).strip(' .,-')
        return cleaned or name.strip()

    def search_filings_raw(self, search_term: str, size: int = 200) -> Optional[List[Dict[str, Any]]]:
        """
        Executes a search against the FCA NSM Elasticsearch API for 'Financial Results'.
        Returns list of hit dictionaries or None if request fails.
        """
        payload = {
            "from": 0,
            "size": size,
            "primarySearch": [{"name": "company", "value": search_term}],
            "latestFlag": "Y",
            "criteriaObj": {
                "criteria": [{"name": "category_group", "value": "Financial Results"}],
                "dateCriteria": []
            },
            "sort": "submitted_date",
            "sortorder": "desc"
        }

        for attempt in range(1, self.max_retries + 1):
            try:
                self._wait_rate_limit()
                response = self.session.post(
                    NSM_SEARCH_API_URL,
                    json=payload,
                    timeout=DEFAULT_TIMEOUT_SEC
                )
                if response.status_code == 200:
                    data = response.json()
                    hits = data.get("hits", {}).get("hits", [])
                    return hits
                elif response.status_code == 429:
                    wait_time = attempt * 3.0
                    logger.warning(f"Rate limited by FCA API (429). Retrying in {wait_time:.1f}s...")
                    time.sleep(wait_time)
                else:
                    logger.warning(f"FCA API returned status {response.status_code} for search '{search_term}' (attempt {attempt}/{self.max_retries})")
                    time.sleep(attempt * 1.0)
            except Exception as e:
                logger.warning(f"Network error querying FCA API for '{search_term}': {e} (attempt {attempt}/{self.max_retries})")
                time.sleep(attempt * 1.5)

        return None

    def get_company_filings(self, firm_name: str) -> List[Dict[str, Any]]:
        """
        Retrieves filings for a given company, checking cache first.
        Tries cleaned company name if exact name returns 0 hits.
        """
        firm_key = firm_name.strip().lower()
        if firm_key in self._company_cache:
            return self._company_cache[firm_key]

        cleaned_name = self.clean_company_name(firm_name)
        cleaned_key = cleaned_name.lower()
        if cleaned_key in self._company_cache:
            self._company_cache[firm_key] = self._company_cache[cleaned_key]
            return self._company_cache[firm_key]

        # 1. Search with cleaned name first (most effective on FCA NSM)
        logger.info(f"Querying FCA NSM for: '{cleaned_name}'")
        hits = self.search_filings_raw(cleaned_name, size=250)

        # 2. If 0 hits and cleaned name differed from original, try original
        if (hits is None or len(hits) == 0) and cleaned_name != firm_name:
            logger.info(f"Retrying FCA NSM with original name: '{firm_name}'")
            hits = self.search_filings_raw(firm_name, size=250)

        if hits is None:
            logger.error(f"Failed to query filings for company '{firm_name}' after {self.max_retries} attempts.")
            return []

        # Cache results for both names
        self._company_cache[firm_key] = hits
        self._company_cache[cleaned_key] = hits
        logger.info(f"Received {len(hits)} total financial results for '{firm_name}'")
        return hits

    def clear_cache(self):
        """Clears in-memory search cache."""
        self._company_cache.clear()
