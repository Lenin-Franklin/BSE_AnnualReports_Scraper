import os
import shutil

source = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"
destination = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"

# Create destination folder if it does not exist
os.makedirs(destination, exist_ok=True)

# Walk through all subfolders
for root, dirs, files in os.walk(source):
    for file in files:
        if file.lower().endswith(".pdf"):
            src_path = os.path.join(root, file)
            dst_path = os.path.join(destination, file)

            # If a file with the same name already exists, rename the new copy
            if os.path.exists(dst_path):
                base, ext = os.path.splitext(file)
                counter = 1
                while os.path.exists(dst_path):
                    dst_path = os.path.join(destination, f"{base}_{counter}{ext}")
                    counter += 1

            shutil.copy2(src_path, dst_path)

print("PDF copy completed.")
