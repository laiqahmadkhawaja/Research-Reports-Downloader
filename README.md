# SEC EDGAR 10-K Report Automation & Excel Tracking System

Please refer to the complete documentation in [SEC_10K_Downloader/README.md](SEC_10K_Downloader/README.md).

## Quick Start

### 1. Install Requirements
```powershell
pip install -r requirements.txt
playwright install chromium
```

### 2. Launch the Desktop GUI
```powershell
python run.py
```

### 3. Run in CLI Batch Mode
```powershell
python run.py --cli --limit 5
```

### 4. Results & Download Folders
- Generated Excel Results: `SEC_10K_Downloader/results/SEC_Download_Results.xlsx`
- Downloaded Reports: `SEC_10K_Downloader/downloaded_reports/`
