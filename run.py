"""
Root launcher for SEC EDGAR 10-K Downloader.
"""
import sys
import os
from pathlib import Path

# Add SEC_10K_Downloader to sys.path
app_dir = Path(__file__).resolve().parent / "SEC_10K_Downloader"
if str(app_dir) not in sys.path:
    sys.path.insert(0, str(app_dir))

from SEC_10K_Downloader.main import main

if __name__ == "__main__":
    main()
