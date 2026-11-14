# SEC EDGAR 10-K Report Automation & Excel Tracking System

A robust, enterprise-grade Python desktop application designed to automatically search, match, and download annual **10-K financial reports** in original HTML format from the official [SEC EDGAR Search System](https://www.sec.gov/edgar/search/) using browser automation.

The application reads an input Excel or CSV file containing company names, reporting years, and unique report codes, tracks full extraction audit details in a dedicated Excel workbook (`SEC_Download_Results.xlsx`), and provides a clean desktop interface built with **CustomTkinter**.

---

## Key Features

1. **Strictly Browser Automation Based**:
   - Replicates genuine human navigation on `https://www.sec.gov/edgar/search/`.
   - **No SEC APIs or third-party filing endpoints are used.** All searching, suggestion selection, filtering, and report downloads are performed through automated browser interaction via Playwright Chromium.

2. **Intelligent Company Search & Autocomplete Selection**:
   - Types entity names into SEC EDGAR entity search.
   - Waits for autocomplete suggestions (`table.entity-hints`).
   - Normalizes corporate legal identifiers (Inc, Corp, LLC, Ltd, etc.) and performs string similarity matching.
   - Selects the exact or best matching company, flagging uncertain matches for manual review.

3. **Strict 10-K Filing Filtering**:
   - Filters filings by form category and input form `10-K`.
   - **Excludes all amended filings (`10-K/A`)**, quarterly reports (`10-Q`), or other forms.
   - Implements exact filing type checking (`filing_type == "10-K"`).

4. **Accurate Reporting Year Matching**:
   - Matches filings based on the **Period of Report / Reporting Period End Date** (e.g., period `2014-12-31` matches reporting year `2014`, regardless of filing submission date in 2015).
   - Identifies original filings and excludes duplicates/amendments.
   - Flags ambiguous dates or multiple original filings for review (`REVIEW REQUIRED`).

5. **Original HTML Document Downloads**:
   - Downloads the full primary HTML document directly.
   - Saves files using the **Unique Report Code** as the filename (e.g., `34147.html`, `21977.html`).
   - Validates file integrity, minimum size (>2 KB), and authentic HTML/XBRL document markup.

6. **Duplicate Prevention & Recovery**:
   - Checks existing files before making browser requests.
   - Skips already downloaded, validated files (`SKIPPED - ALREADY DOWNLOADED`).
   - Persists state to `state/progress.json` continuously to support pausing, resuming, and interruption recovery.

7. **Comprehensive 3-Sheet Results Excel Workbook (`SEC_Download_Results.xlsx`)**:
   - **Sheet 1 (`Download Results`)**: 15 audit columns recording every processed row.
   - **Sheet 2 (`Summary`)**: Aggregate statistics table and embedded status distribution chart.
   - **Sheet 3 (`Failed & Review Records`)**: Filtered records requiring attention with failure reasons and actionable suggestions.

8. **Modern Desktop Interface (CustomTkinter)**:
   - File & output directory pickers.
   - Headless Mode toggle (Headed mode recommended for SEC EDGAR).
   - Skip existing files & retry controls.
   - Live progress bar, metric counters, and real-time timestamped execution logs.
   - Start, Pause, Resume, and Stop controls.

---

## Technology Stack

- **Python**: 3.11+
- **Browser Automation**: [Playwright](https://playwright.dev/python/) (Chromium)
- **GUI Framework**: [CustomTkinter](https://customtkinter.tomschimansky.com/)
- **Spreadsheet Processing**: [openpyxl](https://openpyxl.readthedocs.io/) & [pandas](https://pandas.pydata.org/)
- **HTML Inspection**: [BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/)
- **Logging**: Python standard library `logging` with rotating file and GUI queue handlers

---

## Project Structure

```
SEC_10K_Downloader/
│
├── main.py                     # Application entry point (GUI and CLI modes)
├── gui.py                      # CustomTkinter desktop interface
├── browser_automation.py       # Playwright browser interaction engine
├── excel_reader.py             # Flexible Excel/CSV reader with column detection
├── filing_matcher.py           # Company suggestion & reporting year matching logic
├── downloader.py               # HTML report downloading, validation & naming
├── results_writer.py           # 3-sheet results Excel generation with chart
├── progress_manager.py         # Thread-safe persistent state & pause/resume
├── logger.py                   # Rotating file and GUI callback logging
├── config.py                   # Central configuration constants and paths
├── requirements.txt            # Python dependencies
│
├── input/
│   ├── sample_companies.csv    # Source CSV data (2,129 records)
│   └── sample_companies.xlsx   # Input Excel workbook
│
├── downloaded_reports/         # Output directory for saved {unique_code}.html files
├── results/
│   └── SEC_Download_Results.xlsx  # Multi-sheet audit results workbook
├── logs/
│   └── sec_downloader.log      # Application execution logs
└── state/
    └── progress.json           # Real-time state file for pause/resume
```

---

## Installation

### 1. Clone or Navigate to the Project Directory

```powershell
cd "c:\xampp\htdocs\report downloading"
```

### 2. Install Required Python Packages

```powershell
pip install -r SEC_10K_Downloader/requirements.txt
```

### 3. Install Playwright Chromium Browser

```powershell
python -m playwright install chromium
```

---

## How to Run

### Option A: Launch the Desktop GUI Application (Recommended)

Run from root:
```powershell
python run.py
```
Or from inside `SEC_10K_Downloader`:
```powershell
cd SEC_10K_Downloader
python main.py
```

#### GUI Workflow:
1. **Input Section**: Browse and select your input Excel or CSV file (the included `sample_companies.xlsx` is pre-selected by default).
2. **Output Section**: Choose destination folder for reports (defaults to `downloaded_reports/`).
3. **Configuration Section**:
   - **Headless Mode**: Leave **OFF** (recommended). The SEC EDGAR web portal employs Akamai bot management which blocks headless Chromium instances with HTTP 403.
   - **Skip Existing Reports**: Keep **ON** to avoid redownloading previously saved reports.
   - **Browser Delay**: Adjust request pacing (default: 1.5s) to adhere to SEC fair access guidelines.
4. **Processing**: Click **Start Download**. Monitor live progress metrics, current company/year, and timestamped log messages. Use **Pause**, **Resume**, or **Stop** at any time.

---

### Option B: Run in Command-Line / Batch Mode

The application also features a CLI interface for automated or scripted execution:

```powershell
# Run with a 10-record limit for quick verification
python run.py --cli --limit 10

# Run full batch processing
python run.py --cli --input "SEC_10K_Downloader/input/sample_companies.xlsx" --output "SEC_10K_Downloader/downloaded_reports"

# Additional CLI arguments
python run.py --cli --help
```

CLI options:
- `--cli`: Runs in terminal mode without loading the GUI.
- `--input PATH`: Custom input Excel or CSV file.
- `--output PATH`: Directory to save HTML files.
- `--limit N`: Process only the first `N` records.
- `--delay S`: Pacing delay between browser actions in seconds (default: 1.5).
- `--max-retries N`: Retry attempts for transient errors (default: 2).
- `--no-skip-existing`: Forces redownload of existing files.

---

## Results Workbook Documentation (`SEC_Download_Results.xlsx`)

The results file is updated after each record and contains three worksheets:

### Sheet 1: `Download Results`
Detailed log of every input record across 15 columns:
1. **Serial Number**: Sequence index.
2. **Company Name**: Entity name from input file.
3. **Reporting Year**: Fiscal reporting year requested.
4. **Unique Report Code**: Identifier used as file name (e.g., `34147`).
5. **SEC Company Name Selected**: Exact company name selected from SEC suggestions.
6. **Filing Type**: `10-K` (strictly verified).
7. **Filing Date**: Date filed with SEC (`YYYY-MM-DD`).
8. **Reporting Period**: Fiscal period end date (`YYYY-MM-DD`).
9. **Accession Number**: Formatted SEC accession number (e.g., `0001193125-13-066200`).
10. **Download Status**: Status classification (color-coded).
11. **Saved File Name**: e.g., `34147.html`.
12. **Saved File Path**: Full absolute path on disk.
13. **SEC Filing URL**: Primary filing document web link.
14. **Error Details**: Descriptive note if skipped, failed, or review needed.
15. **Processing Date & Time**: Timestamp of execution.

### Sheet 2: `Summary`
- High-level metric summary table (Total Records, Success, Skipped Already Downloaded, Skipped No Matching 10-K, Failed, Review Required, Total HTML Files Saved, Completion Percentage).
- An embedded **openpyxl column chart** illustrating the distribution of download statuses.

### Sheet 3: `Failed & Review Records`
Filtered list containing only records requiring operator attention:
- Company Name, Reporting Year, Unique Report Code
- Failure or Review Reason
- Error Details
- Suggested Action (e.g., verifying spelling, checking if company was public during requested fiscal year, etc.)

---

## Status Classifications

| Status Code | Meaning |
|---|---|
| `SUCCESS` | Primary original HTML 10-K report successfully downloaded, validated, and saved. |
| `SKIPPED - ALREADY DOWNLOADED` | Valid report already exists on disk with matching unique code. |
| `SKIPPED - NO MATCHING 10-K` | SEC suggestions found, but no original 10-K exists for that reporting year (e.g., pre-IPO years). |
| `FAILED - COMPANY NOT FOUND` | No matching company autocomplete suggestions on SEC EDGAR. |
| `FAILED - BROWSER ERROR` | Browser navigation failure or page timeout. |
| `FAILED - DOWNLOAD ERROR` | HTTP failure or network error fetching the primary HTML document. |
| `FAILED - INVALID HTML` | Downloaded file failed validation (truncated, corrupted, or error page). |
| `REVIEW REQUIRED` | Ambiguous match or multiple original filings for the same period. |

---

## Reliability & SEC Fair Access Compliance

- **Rate Limiting**: Configurable delay (default 1.5 seconds) between browser requests to comply with SEC's fair access threshold (maximum 10 requests per second).
- **In-Memory Filing Cache**: Searches an entity once on SEC EDGAR and caches its filing history in memory. Subsequent years for the same company are matched instantly without redundant network queries.
- **Bot Challenge Detection**: If SEC EDGAR displays a bot challenge or verification prompt, the automation pauses and notifies the user to solve it in the browser window before resuming.
