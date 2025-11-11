import os
import pandas as pd

# Paths
excel_path = r"C:\Users\lenin\OneDrive\Desktop\All company list.xlsx"
folder_path = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

# Load Excel
df = pd.read_excel(excel_path)

# First column = code, second column = name
codes = df.iloc[:, 0].astype(str).str.strip()
names = df.iloc[:, 1].astype(str).str.strip()

# Create mapping dict
mapping = dict(zip(codes, names))

# Process folder names
for folder in os.listdir(folder_path):
    old_path = os.path.join(folder_path, folder)

    if not os.path.isdir(old_path):
        continue

    folder_clean = folder.strip()

    # If folder name matches a company code
    if folder_clean in mapping:
        new_name = mapping[folder_clean]
        new_path = os.path.join(folder_path, new_name)

        # Prevent overwriting existing folder
        if os.path.exists(new_path):
            print(f"Skipped (target exists): {folder} → {new_name}")
        else:
            os.rename(old_path, new_path)
            print(f"Renamed: {folder} → {new_name}")
    else:
        print(f"No match for: {folder}")
