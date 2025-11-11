import os
from collections import Counter

# Path to your directory
path = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

# Get only folder names (ignore files)
folder_names = [
    name for name in os.listdir(path)
    if os.path.isdir(os.path.join(path, name))
]

# Count occurrences
counts = Counter(folder_names)

# Unique folders
unique_folders = [name for name, count in counts.items() if count == 1]

# Duplicate folders (count > 1)
duplicate_folders = {name: count for name, count in counts.items() if count > 1}

# Print results
print(f"Total folders found: {len(folder_names)}")
print(f"Unique folders: {len(unique_folders)}")
print(f"Repeated folders: {len(duplicate_folders)}\n")

if duplicate_folders:
    print("Duplicate folder names and counts:")
    for name, count in duplicate_folders.items():
        print(f"{name} → {count} times")
else:
    print("✅ No duplicate folders found!")
