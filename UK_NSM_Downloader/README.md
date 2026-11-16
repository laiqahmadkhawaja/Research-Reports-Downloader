# UK FCA National Storage Mechanism (NSM) - Annual Report Downloader

A dedicated, enterprise-grade desktop automation tool to search, match, and download **Annual Report PDFs** for UK publicly listed companies from the official **UK Financial Conduct Authority (FCA) National Storage Mechanism (NSM)**:
[https://data.fca.org.uk/#/nsm/nationalstoragemechanism](https://data.fca.org.uk/#/nsm/nationalstoragemechanism)

---

## 🌟 Key Features

1. **Company-Wise Subfolders**:
   - Organizes downloaded reports strictly by company name:
     `downloaded_reports/<Firm Name>/<uni_code>.pdf`
   - Example:
     `downloaded_reports/Anglo American PLC/28207.pdf`
     `downloaded_reports/Anglo American PLC/28208.pdf`
     `downloaded_reports/Associated British Foods PLC/28231.pdf`

2. **UniCode File Naming**:
   - Automatically detects the `uni_code` column from the input Excel/CSV file and names the saved PDF strictly after that code (`<uni_code>.pdf`).

3. **Strict Annual Report Filtering (Excludes Semi-Annual / Interim)**:
   - Filters out non-annual filings: strictly rejects `Half-year Financial Report`, `Semi-annual`, `Interim`, `Quarterly`, `Preliminary Results`, `Trading Updates`, `Dividend Announcements`, and `Voting Rights`.
   - Downloads **ONLY** full Annual Financial Reports and Accounts.

4. **Direct Official FCA NSM API Integration**:
   - Communicates directly with the official FCA NSM Elasticsearch backend (`api.data.fca.org.uk/search?index=nsm-search`).
   - Blazing-fast performance: in-memory caching ensures that all fiscal years (e.g. 2012–2023) for a company are resolved from a single query.

5. **PDF Integrity Verification & Stream Download**:
   - Verifies the standard `%PDF-` file header signature before finalizing downloads to prevent corrupted or HTML error pages from being saved.
   - Atomic file writes using `.tmp` files.

6. **Modern Dark-Themed GUI (CustomTkinter)**:
   - Live KPI metric cards: *Total*, *Processed*, *Success (PDF)*, *Skipped*, *Failed*, *Review Needed*.
   - Live progress bar, percentage counter, and dynamic ETA calculation.
   - Interactive control buttons: *Start Download*, *Pause*, *Resume*, *Stop*.
   - Direct shortcuts: *Browse File*, *Open Output Folder*, *Open Results Excel*, *Open Log File*.
   - Real-time color-coded execution log console.

7. **Multi-Sheet Excel Audit Tracking**:
   - Automatically generates `results/UK_NSM_Download_Results.xlsx` with 3 formatted worksheets:
     - **Download Results**: Complete record audit with UniCode, Company Name, Year, NSM Headline, Document Date, Status, File Path, Size (MB), and URL.
     - **Summary**: Comprehensive KPI metrics, success rate percentage, and completion timestamps.
     - **Failed & Review Records**: Actionable troubleshooting steps for any skipped or unlocated filings.

8. **Pause, Resume & Crash Recovery**:
   - State is continuously tracked in `state/progress.json`. You can pause or stop at any time and resume later without re-downloading completed reports.

---

## 📁 Directory Structure

```text
UK_NSM_Downloader/
├── input/
│   ├── uk_companies.xlsx          # 741 records across 67 UK companies (Excel)
│   └── uk_companies.csv           # 741 records across 67 UK companies (CSV)
├── downloaded_reports/            # Output folder with company-wise subfolders
│   ├── Anglo American PLC/
│   │   ├── 28207.pdf              # 2012 Annual Report
│   │   ├── 28208.pdf              # 2013 Annual Report
│   │   └── ...
│   ├── Admiral Group PLC/
│   │   └── ...
├── results/
│   └── UK_NSM_Download_Results.xlsx # 3-Sheet formatted audit report
├── logs/
│   └── uk_nsm_downloader.log     # Rotating execution log file
├── state/
│   └── progress.json              # State tracking for resume capabilities
├── config.py                      # URLs, paths, status codes, defaults
├── logger.py                      # Rotating file logger and GUI callback handler
├── excel_reader.py                # Auto column detection (uni_code, firm_name, year)
├── nsm_client.py                  # Direct FCA NSM API client with caching & retries
├── filing_matcher.py              # Strict annual report isolation & year matching
├── downloader.py                  # Stream PDF downloader with atomic verification
├── progress_manager.py            # State persistence & KPI metric calculation
├── results_writer.py              # 3-Sheet Excel workbook generator
├── gui.py                         # CustomTkinter graphical desktop interface
├── main.py                        # Entry point (GUI and CLI modes)
└── requirements.txt               # Dependencies
```

---

## 🚀 Quickstart Guide

### 1. Launch the Desktop GUI

From the root project directory:
```bash
python run_uk_nsm.py
```

Or from inside `UK_NSM_Downloader`:
```bash
cd UK_NSM_Downloader
python main.py
```

### 2. How to Use the GUI:
1. **Input File**: The bundled `input/uk_companies.xlsx` is automatically pre-loaded. You can also click **"Browse File"** to select any other file.
2. **Output Folder**: Defaults to `UK_NSM_Downloader/downloaded_reports`. You can change this or click **"Open"** to view downloaded reports.
3. **Settings**:
   - `Strict PDF Only`: Enabled by default to ensure only genuine PDFs are saved.
   - `Skip Already Downloaded`: Enabled by default so existing valid PDFs are never re-downloaded.
   - `Request Delay`: Set to 0.5s by default for polite and reliable API calls.
4. **Start Download**: Click the green **"▶ Start Download"** button.
5. **Real-time Monitoring**: Watch the live progress bar, KPI metric cards, and log console.
6. **Open Results**: When complete, click **"Open Results Excel"** to inspect the full audit trail.

---

## 💻 CLI / Headless Mode (Optional)

You can also run the downloader directly in terminal / background mode:
```bash
python UK_NSM_Downloader/main.py --cli --input UK_NSM_Downloader/input/uk_companies.xlsx --output UK_NSM_Downloader/downloaded_reports
```

---

## 🇵🇰 اردو ہدایات (Urdu Guidance)

- **علیحدہ فولڈر**: یہ نیا ٹول مکمل طور پر ایک الگ فولڈر `UK_NSM_Downloader` میں بنایا گیا ہے۔
- **کمپنی وائز فولڈر**: ہر کمپنی کے لیے ایک الگ فولڈر بنے گا جیسے `downloaded_reports/Anglo American PLC/`۔
- **یونیکوڈ سے نام**: ہر ڈاؤنلوڈ ہونے والی سالانہ رپورٹ (Annual Report) کا نام ایکسل کے `uni_code` کالم کے مطابق رکھا جائے گا (جیسے `28207.pdf`)۔
- **صرف سالانہ رپورٹ (Strict Annual Report PDF)**: یہ سسٹم سیمی اینول، کوارٹرلی، یا دوسری خبروں کو مکمل طور پر فلٹر کر کے صرف اصلی اینول رپورٹ پی ڈی ایف ڈاؤنلوڈ کرتا ہے۔
- **چلانے کا طریقہ**: ٹرمینل میں صرف `python run_uk_nsm.py` رن کریں، خوبصورت ڈارک تھیم والا GUI کھل جائے گا، پھر **Start Download** دبائیں۔
