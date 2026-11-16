"""
Root launcher for UK FCA NSM Annual Report Downloader.
"""
import sys
import os
from pathlib import Path

# Add UK_NSM_Downloader to sys.path
app_dir = Path(__file__).resolve().parent / "UK_NSM_Downloader"
if str(app_dir) not in sys.path:
    sys.path.insert(0, str(app_dir))

from UK_NSM_Downloader.main import main

if __name__ == "__main__":
    main()
