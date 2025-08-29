# BSE_AnnualReports_Scraper
This code lets you download the annual reports of 500 companies between 2019 to 2025

Prerequisites:
Before running the script, make sure you have the following installed and configured:

1. Python
Install Python 3.8+ (recommended: latest stable release).
Verify installation:
  _python --version_

2. Required Python Libraries
Install dependencies using pip:
  _pip install pandas selenium requests_

pandas → for reading the CSV of company codes.
selenium → for automating the browser.
requests → for downloading the PDF files.

3. Google Chrome
Install the Google Chrome browser (latest version).
Verify installation:
  _Windows: launch Chrome and check chrome://settings/help_
  _Linux/Mac: google-chrome --version_

4. ChromeDriver
Selenium requires ChromeDriver that matches your Chrome version.
Download from: https://chromedriver.chromium.org/downloads
Add the chromedriver executable to your system PATH, or place it in the same folder as your script.
Verify installation:
  _chromedriver --version_

5. Input Data (CSV File)
Create a CSV file (Company_Names.csv) with one company code per line.
Example:
500325
532540
500180
   
Update the script with the correct path:
csv_path = r"<Give your CSV file with company codes>"

7. Output Folder
The script will automatically create an output directory (e.g., NSE Scraper) to save the annual report PDFs.
Each company will get its own subfolder.

✏️ Configuration (Important!)
Before running the script, edit these two variables in the code:
  # Location of your CSV file
    csv_path = r"C:\Users\YOUR_USERNAME\Desktop\Company_Names.csv"
  # Folder where PDFs will be stored
    base_dir = r"C:\Users\YOUR_USERNAME\Desktop\NSE Scraper"

csv_path → path to your company codes CSV.
base_dir → path where you want all downloaded reports saved.
A subfolder is created automatically for each company.

✅ Once everything is set up:
Run the script with:
 - _python bse_scraper.py_
At the end, you’ll get a summary report in the terminal with:
Total companies processed
Reports downloaded
Reports skipped (already exists)
Errors encountered


