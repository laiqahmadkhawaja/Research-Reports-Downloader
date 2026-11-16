# Research Reports Downloader Suite

A comprehensive financial report automation and Excel audit tracking suite for financial research, supporting both US (SEC EDGAR) and UK (FCA NSM) corporate disclosures.

---

## 📦 Included Tools

### 1. SEC EDGAR 10-K Report Downloader (`SEC_10K_Downloader`)
- Downloads US SEC EDGAR 10-K filings with unique-code naming.
- Full CustomTkinter desktop GUI, headless browser automation, and 3-sheet audit Excel reporting.
- Launcher: `python run.py`
- Documentation: [SEC_10K_Downloader/README.md](SEC_10K_Downloader/README.md)

### 2. UK FCA NSM Annual Report Downloader (`UK_NSM_Downloader`)
- Downloads UK Annual Reports in PDF format from the official **UK FCA National Storage Mechanism (NSM)** ([data.fca.org.uk/#/nsm/nationalstoragemechanism](https://data.fca.org.uk/#/nsm/nationalstoragemechanism)).
- Creates company-wise subfolders and renames downloaded files strictly using the Excel `uni_code` column (`<uni_code>.pdf`).
- Filters out non-annual filings (excludes semi-annual, interim, quarterly, preliminary, etc.).
- Direct official FCA Elasticsearch API integration with high-performance in-memory caching.
- CustomTkinter desktop GUI, live KPI metric cards, progress tracking, and multi-sheet audit Excel reporting.
- Launcher: `python run_uk_nsm.py`
- Documentation: [UK_NSM_Downloader/README.md](UK_NSM_Downloader/README.md)

---

## 🚀 Quick Start

### Installation
```powershell
pip install -r requirements.txt
playwright install chromium
```

### Launch Tools

#### To launch the UK FCA NSM Downloader:
```powershell
python run_uk_nsm.py
```

#### To launch the SEC EDGAR 10-K Downloader:
```powershell
python run.py
```
