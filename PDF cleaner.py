"""
delete_processed_pdfs.py

Deletes PDFs from a folder if their filenames already exist
in the ISO Data Collection Excel file.

SAFE:
- Excel is READ-ONLY
- Only deletes PDFs that are already processed
"""

import os
import pandas as pd

# =========================
# CONFIG
# =========================
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
EXCEL_PATH = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"

DRY_RUN = True  # <-- SET TO False TO ACTUALLY DELETE

# =========================
# LOAD EXCEL
# =========================
print("[INFO] Loading Excel...")
df = pd.read_excel(EXCEL_PATH)

if "File" not in df.columns:
    raise RuntimeError("Excel does not contain a 'File' column")

processed_files = set(
    df["File"]
    .dropna()
    .astype(str)
    .str.strip()
)

print(f"[INFO] Processed files in Excel: {len(processed_files)}")

# =========================
# SCAN PDF FOLDER
# =========================
pdfs = [
    f for f in os.listdir(PDF_FOLDER)
    if f.lower().endswith(".pdf")
]

print(f"[INFO] PDFs found in folder: {len(pdfs)}")

to_delete = [f for f in pdfs if f in processed_files]

print(f"[INFO] PDFs eligible for deletion: {len(to_delete)}\n")

# =========================
# DELETE
# =========================
deleted = 0

for f in to_delete:
    path = os.path.join(PDF_FOLDER, f)
    if DRY_RUN:
        print(f"[DRY-RUN] Would delete: {f}")
    else:
        os.remove(path)
        print(f"[DELETED] {f}")
        deleted += 1

# =========================
# SUMMARY
# =========================
print("\n===== SUMMARY =====")
print(f"Dry run mode   : {DRY_RUN}")
print(f"Matched PDFs  : {len(to_delete)}")
print(f"Deleted PDFs  : {deleted if not DRY_RUN else 0}")
print("===================\n")
