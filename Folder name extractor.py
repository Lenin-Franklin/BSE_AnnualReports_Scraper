import os
from openpyxl import Workbook

# 📁 Target directory
directory_path = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

# 📂 Get only folder names
folder_names = [name for name in os.listdir(directory_path) if os.path.isdir(os.path.join(directory_path, name))]

# 📊 Create Excel workbook
wb = Workbook()
ws = wb.active
ws.title = "Folder Names"

# 📝 Write folder names
ws.append(["Folder Name"])
for folder in folder_names:
    ws.append([folder])

# 💾 Save Excel file
output_path = r"C:\Users\lenin\OneDrive\Desktop\folder_names.xlsx"
wb.save(output_path)

print(f"✅ Folder names saved to {output_path}")
