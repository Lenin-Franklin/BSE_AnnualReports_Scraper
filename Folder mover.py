import os
import shutil

# 📁 Source folder (NSE Scraper)
source_root = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

# 📁 Destination folder (Company_PDF)
destination_folder = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
os.makedirs(destination_folder, exist_ok=True)

# 🔍 Walk through all subdirectories in source
for root, dirs, files in os.walk(source_root):
    for file in files:
        if file.lower().endswith(".pdf"):
            source_path = os.path.join(root, file)
            destination_path = os.path.join(destination_folder, file)

            # 🧠 Handle duplicate filenames
            if os.path.exists(destination_path):
                base, ext = os.path.splitext(file)
                counter = 1
                while os.path.exists(destination_path):
                    new_name = f"{base}_{counter}{ext}"
                    destination_path = os.path.join(destination_folder, new_name)
                    counter += 1

            # 📦 Move the PDF
            shutil.move(source_path, destination_path)

print(f"✅ All PDFs moved to: {destination_folder}")
