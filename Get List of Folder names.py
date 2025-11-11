import os

path = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

if not os.path.exists(path):
    print("Path does not exist:", path)
else:
    print("Folders inside:", path)
    for item in os.listdir(path):
        full_path = os.path.join(path, item)
        if os.path.isdir(full_path):
            print(item)
