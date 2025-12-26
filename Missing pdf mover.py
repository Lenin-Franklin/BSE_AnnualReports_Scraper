import os
import shutil
import pandas as pd

# =========================
# CONFIG
# =========================
MISSING_CSV = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\missing_company_years.csv"
SOURCE_ROOT = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"
DEST_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\Missing_PDFs_FLAT2"

os.makedirs(DEST_FOLDER, exist_ok=True)

# =========================
# LOAD MISSING LIST
# =========================
df = pd.read_csv(MISSING_CSV)

copied = 0
not_found = []

# =========================
# PROCESS
# =========================
for _, row in df.iterrows():
    company = row["Company"].strip()
    years = [y.strip() for y in row["Missing_Years"].split(",")]

    company_folder = os.path.join(SOURCE_ROOT, company)

    if not os.path.isdir(company_folder):
        print(f"[WARN] Company folder not found: {company}")
        not_found.append((company, "ALL"))
        continue

    for year in years:
        pdf_name = f"{year}_{company}.pdf"
        src_pdf = os.path.join(company_folder, pdf_name)
        dst_pdf = os.path.join(DEST_FOLDER, pdf_name)

        if os.path.exists(src_pdf):
            shutil.copy2(src_pdf, dst_pdf)
            copied += 1
            print(f"[OK] Copied: {pdf_name}")
        else:
            print(f"[MISS] Not found: {pdf_name}")
            not_found.append((company, year))

# =========================
# SUMMARY
# =========================
print("\n========== SUMMARY ==========")
print(f"Total PDFs copied      : {copied}")
print(f"PDFs not found         : {len(not_found)}")

if not_found:
    print("\n[!] Missing PDFs:")
    for c, y in not_found[:20]:
        print(f"   - {c} ({y})")
    if len(not_found) > 20:
        print("   ...")

print("\n[✔] Flat missing-PDF folder ready")
print("================================")
