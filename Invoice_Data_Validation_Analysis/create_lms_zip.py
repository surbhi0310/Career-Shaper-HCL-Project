"""
create_lms_zip.py - Packages the entire project into a submission-ready ZIP file
"""

import os
import zipfile
from pathlib import Path

BASE_DIR = Path(r"C:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")
ZIP_OUT = BASE_DIR / "PJ_IVD_2026_Invoice_Data_Validation_Complete_Project.zip"

exclude_dirs = {".pytest_cache", "__pycache__", ".git"}
exclude_exts = {".pyc"}

with zipfile.ZipFile(ZIP_OUT, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for root, dirs, files in os.walk(BASE_DIR):
        # Skip excluded dirs
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        for file in files:
            file_path = Path(root) / file
            if file_path.suffix in exclude_exts or file_path == ZIP_OUT:
                continue
            arcname = file_path.relative_to(BASE_DIR)
            zipf.write(file_path, arcname)

size_mb = os.path.getsize(ZIP_OUT) / (1024 * 1024)
print(f"Created ZIP: {ZIP_OUT} ({size_mb:.2f} MB)")
