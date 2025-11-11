import os
import pandas as pd

# Paths
excel_path = r"C:\Users\lenin\OneDrive\Desktop\All company list.xlsx"
base_path = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

# Load Excel (col1 = code, col2 = name)
df = pd.read_excel(excel_path)
codes = df.iloc[:, 0].astype(str).str.strip()
names = df.iloc[:, 1].astype(str).str.strip()

# Map: code → company name
code_to_name = dict(zip(codes, names))

print("Starting PDF renaming...\n")

# Counters
total_files = 0
renamed_files = 0
skipped_no_underscore = 0
skipped_no_match = 0
skipped_other = 0

# Loop over all folders
for folder in os.listdir(base_path):
    folder_path = os.path.join(base_path, folder)

    if not os.path.isdir(folder_path):
        continue

    for file in os.listdir(folder_path):
        if not file.lower().endswith(".pdf"):
            continue

        total_files += 1
        old_file_path = os.path.join(folder_path, file)

        # Expect format: YEAR_CODE.pdf
        if "_" not in file:
            print(f"Skipping unexpected file (no underscore): {file}")
            skipped_no_underscore += 1
            continue

        try:
            year, rest = file.split("_", 1)
        except ValueError:
            print(f"Skipping malformed file: {file}")
            skipped_other += 1
            continue

        # Extract code and remove .pdf
        code = rest.replace(".pdf", "").strip()

        # Look up company name
        if code not in code_to_name:
            print(f"❌ No match found for code: {code} in file {file}")
            skipped_no_match += 1
            continue

        company_name = code_to_name[code]

        # Build new filename
        new_filename = f"{year}_{company_name}.pdf"
        new_file_path = os.path.join(folder_path, new_filename)

        if file == new_filename:
            # Already correctly named
            skipped_other += 1
            continue

        os.rename(old_file_path, new_file_path)
        renamed_files += 1
        print(f"✅ {file} → {new_filename}")

# ✅ Final Summary
print("\n============================")
print("📊 RENAME SUMMARY REPORT")
print("============================")
print(f"Total PDF files scanned:     {total_files}")
print(f"✅ Successfully renamed:      {renamed_files}")
print(f"⚠️ Skipped (no underscore):   {skipped_no_underscore}")
print(f"⚠️ Skipped (no code match):   {skipped_no_match}")
print(f"⚠️ Skipped (already ok/other): {skipped_other}")
print("============================")
print("✅ Finished!")
