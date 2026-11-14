"""
browser_automation.py - Playwright browser automation for SEC EDGAR 10-K searches and downloads.
Strictly uses browser automation (no SEC APIs or third-party filing APIs).
"""
import os
import asyncio
import time
from typing import List, Dict, Tuple, Optional, Any, Callable
from playwright.async_api import async_playwright, Browser, BrowserContext, Page, Error as PlaywrightError
from config import (
    SEC_SEARCH_URL,
    DEFAULT_USER_AGENT,
    DEFAULT_HTTP_HEADERS,
    STATUS_SUCCESS,
    STATUS_SKIPPED_ALREADY_DOWNLOADED,
    STATUS_SKIPPED_NO_MATCHING_10K,
    STATUS_FAILED_COMPANY_NOT_FOUND,
    STATUS_FAILED_BROWSER_ERROR,
    STATUS_FAILED_DOWNLOAD_ERROR,
    STATUS_FAILED_INVALID_HTML,
    STATUS_REVIEW_REQUIRED,
    DEFAULT_BROWSER_DELAY,
    DEFAULT_PAGE_TIMEOUT_MS,
    DEFAULT_ACTION_TIMEOUT_MS
)
from filing_matcher import (
    match_company_suggestion,
    match_filing_for_reporting_year,
    extract_accession_number,
    is_valid_10k_form
)
from downloader import download_html_report, check_existing_download
from logger import get_logger

logger = get_logger()

class SecBrowserAutomation:
    """
    Manages Playwright browser lifecycle and SEC EDGAR navigation workflows.
    """
    def __init__(
        self,
        headless: bool = False,
        delay_sec: float = DEFAULT_BROWSER_DELAY,
        on_captcha_callback: Optional[Callable[[], None]] = None
    ):
        self.headless = headless
        self.delay_sec = delay_sec
        self.on_captcha_callback = on_captcha_callback
        
        self._playwright = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None
        
        # In-memory cache of scraped filings by company name
        # key: cleaned company name -> {"selected_name": str, "filings": list, "timestamp": float}
        self._company_filings_cache: Dict[str, Dict[str, Any]] = {}
        self._search_form_configured = False

    async def start(self):
        """Launches Playwright browser with realistic configuration."""
        if self._browser is not None:
            return

        logger.info(f"Starting browser automation (Headless={self.headless})...")
        self._playwright = await async_playwright().start()

        launch_args = [
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-infobars"
        ]

        self._browser = await self._playwright.chromium.launch(
            headless=self.headless,
            args=launch_args
        )

        self._context = await self._browser.new_context(
            user_agent=DEFAULT_USER_AGENT,
            extra_http_headers=DEFAULT_HTTP_HEADERS,
            viewport={"width": 1400, "height": 900},
            locale="en-US"
        )

        self._page = await self._context.new_page()
        self._page.set_default_timeout(DEFAULT_ACTION_TIMEOUT_MS)

        # Stealth init script to mask automation
        await self._page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        """)

        logger.info("Browser initialized successfully.")

    async def close(self):
        """Closes browser session."""
        try:
            if self._context:
                await self._context.close()
            if self._browser:
                await self._browser.close()
            if self._playwright:
                await self._playwright.stop()
        except Exception as e:
            logger.warning(f"Error closing browser: {e}")
        finally:
            self._browser = None
            self._context = None
            self._page = None
            self._playwright = None
            self._search_form_configured = False

    async def _check_and_handle_bot_verification(self) -> bool:
        """
        Inspects page for rate-limiting, CAPTCHA, or anti-bot verification challenges.
        If detected, pauses and requests manual user resolution.
        """
        if not self._page:
            return False

        try:
            title = await self._page.title()
            content_sample = (await self._page.content())[:2000].lower()
            
            is_blocked = (
                "request rate threshold exceeded" in title.lower() or
                "request rate threshold exceeded" in content_sample or
                "undeclared automated tool" in content_sample or
                "403 forbidden" in title.lower() or
                "access denied" in content_sample
            )

            captcha_selectors = [
                "#challenge-running", ".g-recaptcha", "iframe[src*='recaptcha']",
                "iframe[src*='hcaptcha']", "iframe[src*='turnstile']"
            ]
            has_captcha = False
            for sel in captcha_selectors:
                if await self._page.locator(sel).count() > 0:
                    has_captcha = True
                    break

            if is_blocked or has_captcha:
                msg = "SEC rate-limit or bot challenge detected! Pausing automation for user manual completion."
                logger.error(msg)
                if self.on_captcha_callback:
                    self.on_captcha_callback()
                return True

        except Exception as e:
            logger.warning(f"Error checking bot challenge: {e}")

        return False

    async def _ensure_search_filters(self) -> bool:
        """
        Navigates to SEC EDGAR search (if needed) and verifies/ensures:
        1. Form options are expanded.
        2. Date Range is set to 'all' (since 2001).
        3. Category dropdown is set to 'custom' and #filing-types input is visible.
        4. #filing-types is strictly filled with '10-K' with DOM events dispatched.
        """
        if not self._page:
            await self.start()

        # Check current URL, navigate to SEC EDGAR search if not already there
        current_url = self._page.url or ""
        if not current_url.startswith(SEC_SEARCH_URL):
            logger.info("Navigating to SEC EDGAR Search page...")
            await self._page.goto(SEC_SEARCH_URL, wait_until="load", timeout=DEFAULT_PAGE_TIMEOUT_MS)
            try:
                await self._page.wait_for_function(
                    "() => typeof jQuery !== 'undefined' && $('#entity-full-form, #entity-short-form').length > 0",
                    timeout=15000
                )
            except Exception:
                pass
            await self._page.wait_for_timeout(2000)

        if await self._check_and_handle_bot_verification():
            return False

        # Expand search options if collapsed (e.g. short form visible)
        short_input = self._page.locator("#entity-short-form")
        if await short_input.is_visible():
            more_options = self._page.locator("#show-full-search-form")
            if await more_options.is_visible():
                await more_options.click()
                await self._page.wait_for_timeout(500)

        # 1. Ensure Date Range is "all" (since 2001)
        date_select = self._page.locator("#date-range-select")
        if await date_select.is_visible():
            curr_date = await date_select.input_value()
            if curr_date != "all":
                await date_select.select_option("all")
                logger.info("Configured date range: All (since 2001).")
                await self._page.wait_for_timeout(200)

        # 2. Ensure Category is 'custom' so #filing-types is visible
        filing_input = self._page.locator("#filing-types")
        if not await filing_input.is_visible():
            cat_btn = self._page.locator("#category-select")
            if await cat_btn.is_visible():
                await cat_btn.click()
                await self._page.wait_for_timeout(300)
                custom_li = self._page.locator("#category-type-grp li[value='custom']")
                if await custom_li.count() > 0:
                    await custom_li.click()
                    await self._page.wait_for_timeout(300)

        # 3. Ensure #filing-types is filled with '10-K' and events dispatched
        if await filing_input.is_visible():
            curr_val = (await filing_input.input_value()).strip()
            if curr_val != "10-K":
                logger.info(f"Setting filing filter to '10-K' (was '{curr_val}')...")
                await filing_input.click()
                await filing_input.fill("10-K")
                # Dispatch DOM events so jQuery/EDGAR filters register the value
                await self._page.evaluate("""() => {
                    const el = document.querySelector('#filing-types');
                    if (el) {
                        el.value = '10-K';
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', keyCode: 13, bubbles: true }));
                    }
                }""")
                await self._page.wait_for_timeout(300)
                logger.info(f"Verified filing filter: {await filing_input.input_value()}.")

        self._search_form_configured = True
        return True

    async def search_company_filings(self, company_name: str) -> Tuple[str, Optional[str], List[Dict[str, Any]], Optional[str]]:
        """
        Searches SEC EDGAR for the given company, selects from autocomplete suggestions,
        and retrieves all 10-K filings.
        
        Returns:
            status: SUCCESS or FAILED_*
            selected_sec_name: Company name selected from SEC suggestions
            filings: List of parsed filing dicts
            error_details: Error message if failed
        """
        if not self._page:
            await self.start()

        # Check in-memory cache first to avoid hammering SEC servers
        cache_key = company_name.strip().lower()
        if cache_key in self._company_filings_cache:
            cached = self._company_filings_cache[cache_key]
            logger.info(f"Using cached SEC filings for '{company_name}' ({len(cached['filings'])} filings).")
            return STATUS_SUCCESS, cached["selected_name"], cached["filings"], None

        try:
            # Ensure filters are actively configured (Date Range 'all', Form '10-K')
            ok = await self._ensure_search_filters()
            if not ok:
                return STATUS_FAILED_BROWSER_ERROR, None, [], "Failed to configure SEC search form."

            # Locate entity input (full form if visible, else short form)
            full_input = self._page.locator("#entity-full-form")
            short_input = self._page.locator("#entity-short-form")
            full_vis = await full_input.is_visible()
            short_vis = await short_input.is_visible()
            target_input = full_input if full_vis else short_input
            target_id = await target_input.get_attribute("id")
            logger.info(f"Target input id='{target_id}' (full_vis={full_vis}, short_vis={short_vis})")

            logger.info(f"Entering company name '{company_name}' into SEC EDGAR entity search...")
            await target_input.click()
            await self._page.wait_for_timeout(300)
            await self._page.keyboard.press("Control+A")
            await self._page.keyboard.press("Backspace")
            await self._page.wait_for_timeout(300)
            await target_input.press_sequentially(company_name, delay=80)

            # Wait for autocomplete hints table
            hints_selector = "table.entity-hints tr"
            try:
                await self._page.wait_for_selector(hints_selector, state="visible", timeout=8000)
            except PlaywrightError:
                await self._page.wait_for_timeout(2000)

            # Extract suggestions
            suggestion_rows = self._page.locator(hints_selector)
            sug_count = await suggestion_rows.count()

            if sug_count == 0:
                # Save debug screenshot and inspect DOM
                debug_img = os.path.join(os.path.dirname(__file__), "logs", "debug_no_suggestions.png")
                await self._page.screenshot(path=debug_img)
                val = await target_input.input_value()
                logger.warning(f"No SEC EDGAR suggestions found for '{company_name}'. Input value: '{val}'. Screenshot saved to {debug_img}.")
                return STATUS_FAILED_COMPANY_NOT_FOUND, None, [], f"No autocomplete suggestions found for '{company_name}'."

            suggestions_text = []
            for i in range(sug_count):
                txt = (await suggestion_rows.nth(i).inner_text()).strip()
                suggestions_text.append(txt)

            logger.info(f"Found {len(suggestions_text)} suggestions for '{company_name}': {suggestions_text[:3]}")

            # Match suggestion
            best_idx, selected_sec_name, score, is_uncertain = match_company_suggestion(company_name, suggestions_text)
            if best_idx is None:
                err = f"No confident match among SEC suggestions for '{company_name}'. Top suggestions: {suggestions_text[:3]}"
                logger.warning(err)
                return STATUS_FAILED_COMPANY_NOT_FOUND, None, [], err

            logger.info(f"Selected SEC suggestion [{best_idx}]: '{selected_sec_name}' (Confidence: {score:.2f})")
            await suggestion_rows.nth(best_idx).click()
            await self._page.wait_for_timeout(int(self.delay_sec * 500))

            # Re-ensure filing filter remains '10-K' before submitting search
            await self._ensure_search_filters()

            # Click SEARCH button
            search_btn = self._page.locator("#search")
            if await search_btn.is_visible():
                await search_btn.click()
                logger.info(f"Executed search for '{selected_sec_name}'. Waiting for filing results table...")
                await self._page.wait_for_timeout(int(self.delay_sec * 1500))

            # Wait for results rows
            results_row_selector = "table.table tbody tr"
            try:
                await self._page.wait_for_selector(results_row_selector, state="attached", timeout=12000)
            except PlaywrightError:
                pass

            # Scrape filing results table
            results = await self._page.eval_on_selector_all("table.table tbody tr", """rows => rows.map(r => {
                const tds = Array.from(r.querySelectorAll('td')).map(td => td.innerText.trim());
                const a = r.querySelector('td:first-child a');
                return {
                    form: tds[0] || '',
                    filed: tds[1] || '',
                    reporting: tds[2] || '',
                    entity: tds[3] || '',
                    link: a ? a.href : ''
                };
            })""")

            # Filter valid filing records
            valid_filings = [
                r for r in results if r.get("filed") and r.get("form") and is_valid_10k_form(r.get("form"))
            ]

            if results and not valid_filings:
                sample_forms = list({r.get('form', '') for r in results[:10]})
                logger.warning(f"No original 10-K forms in results for '{selected_sec_name}'. Sample forms in table: {sample_forms}")

            logger.info(f"Retrieved {len(valid_filings)} valid original 10-K filings for '{selected_sec_name}'.")

            # Cache results
            self._company_filings_cache[cache_key] = {
                "selected_name": selected_sec_name,
                "filings": valid_filings,
                "timestamp": time.time()
            }

            return STATUS_SUCCESS, selected_sec_name, valid_filings, None

        except Exception as e:
            err = f"Browser error while searching company '{company_name}': {str(e)}"
            logger.error(err)
            return STATUS_FAILED_BROWSER_ERROR, None, [], err

    async def process_record(
        self,
        record: Dict[str, Any],
        output_dir: str,
        skip_existing: bool = True
    ) -> Dict[str, Any]:
        """
        Executes complete search, match, and download workflow for a single input record.
        
        Returns:
            result_dict formatted for results_writer and progress_manager.
        """
        company_name = record["company_name"]
        reporting_year = record["reporting_year"]
        unique_code = record["unique_code"]

        result = {
            "serial_num": record.get("serial_num", 1),
            "company_name": company_name,
            "reporting_year": reporting_year,
            "unique_code": unique_code,
            "sec_company_name": "",
            "filing_type": "",
            "filing_date": "",
            "reporting_period": "",
            "accession_number": "",
            "download_status": "",
            "saved_filename": "",
            "saved_filepath": "",
            "filing_url": "",
            "error_details": "",
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        }

        # 1. Duplicate check before making any browser requests
        if skip_existing:
            already_done, existing_path = check_existing_download(output_dir, unique_code)
            if already_done:
                result["download_status"] = STATUS_SKIPPED_ALREADY_DOWNLOADED
                result["saved_filename"] = os.path.basename(existing_path)
                result["saved_filepath"] = existing_path
                logger.info(f"Record [{unique_code}] {company_name} ({reporting_year}) already downloaded.")
                return result

        # 2. Search company and get filings
        search_status, selected_name, filings, err = await self.search_company_filings(company_name)
        result["sec_company_name"] = selected_name or ""

        if search_status != STATUS_SUCCESS:
            result["download_status"] = search_status
            result["error_details"] = err or ""
            return result

        # 3. Match 10-K filing for reporting year
        matched_filing, review_note = match_filing_for_reporting_year(filings, reporting_year)

        if not matched_filing:
            result["download_status"] = STATUS_SKIPPED_NO_MATCHING_10K
            result["error_details"] = review_note or f"No 10-K filing found for reporting year {reporting_year}."
            logger.info(f"[{unique_code}] {company_name}: {result['download_status']} ({result['error_details']})")
            return result

        # Populate matched filing details
        doc_url = matched_filing.get("link", "")
        result["filing_type"] = "10-K"
        result["filing_date"] = matched_filing.get("filed", "")
        result["reporting_period"] = matched_filing.get("reporting", "")
        result["accession_number"] = extract_accession_number(doc_url)
        result["filing_url"] = doc_url

        if not doc_url:
            result["download_status"] = STATUS_FAILED_DOWNLOAD_ERROR
            result["error_details"] = "Primary HTML document URL is missing in filing results."
            return result

        # 4. Download HTML report
        dl_status, saved_file, saved_path, dl_err = await download_html_report(
            self._page,
            doc_url=doc_url,
            output_dir=output_dir,
            unique_code=unique_code,
            skip_existing=skip_existing
        )

        result["saved_filename"] = saved_file or ""
        result["saved_filepath"] = saved_path or ""

        if dl_status == STATUS_SUCCESS:
            if review_note:
                result["download_status"] = STATUS_REVIEW_REQUIRED
                result["error_details"] = review_note
            else:
                result["download_status"] = STATUS_SUCCESS
                result["error_details"] = ""
        else:
            result["download_status"] = dl_status
            result["error_details"] = dl_err or ""

        # Delay to respect fair access
        await self._page.wait_for_timeout(int(self.delay_sec * 1000))
        return result
