import os
import requests
import time
import shutil
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

#-----------------------------------------------------------------
# 📘 BSE Annual Report Scraper (Optimized + Progress + Auto-stop)
# Downloads reports for 2016–2025 only.
# Deletes incomplete folders automatically after post-check.
# Stops automatically after 500 valid companies.
#-----------------------------------------------------------------


class CompanySearch:
    def __init__(self, driver, wait):
        self.driver = driver
        self.wait = wait

    def search_company(self, company_code: str):
        """Searches for a company using its scrip code"""
        try:
            search_box = self.wait.until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_SmartSearch_smartSearch"))
            )
            search_box.clear()
            search_box.send_keys(company_code)
            time.sleep(1)
            search_box.send_keys(Keys.ENTER)

            submit_btn = self.wait.until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnSubmit"))
            )
            submit_btn.click()

            self.wait.until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_gvData"))
            )
            print(f"✔️ Search completed for {company_code}")

        except Exception:
            print(f"⚠️ Timeout or element not found for {company_code}")


class AnnualReportDownloader:
    def __init__(self, driver, wait, base_dir, summary):
        self.driver = driver
        self.wait = wait
        self.base_dir = base_dir
        self.summary = summary
        os.makedirs(base_dir, exist_ok=True)

    def download_reports(self, company_code: str):
        """Extracts annual report table and downloads PDFs (2016–2025 only)"""
        try:
            report_table = self.wait.until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_grdAnnualReport"))
            )
            rows = report_table.find_elements(By.TAG_NAME, "tr")

            print(f"📑 Found {len(rows) - 1} reports for {company_code}")

            company_dir = os.path.join(self.base_dir, company_code.upper())
            os.makedirs(company_dir, exist_ok=True)

            seen_years = set()
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/116.0 Safari/537.36"
                )
            }
            cookies = {c['name']: c['value'] for c in self.driver.get_cookies()}

            # Filter valid report rows
            valid_rows = []
            for row in rows[1:]:
                cols = row.find_elements(By.TAG_NAME, "td")
                if not cols:
                    continue
                year = cols[0].text.strip()
                if year.isdigit() and 2016 <= int(year) <= 2025:
                    valid_rows.append(row)

            # Progress bar for downloads
            for row in tqdm(valid_rows, desc=f"Downloading {company_code}", unit="report", ncols=80):
                cols = row.find_elements(By.TAG_NAME, "td")
                year = cols[0].text.strip()
                if year in seen_years:
                    continue

                seen_years.add(year)
                pdf_link = cols[-1].find_element(By.TAG_NAME, "a").get_attribute("href")

                if pdf_link and pdf_link.endswith(".pdf"):
                    filepath = os.path.join(company_dir, f"{year}_{company_code}.pdf")

                    if not os.path.exists(filepath):
                        try:
                            response = requests.get(pdf_link, headers=headers, cookies=cookies, stream=True, timeout=15)
                            if response.headers.get("Content-Type", "").lower().startswith("application/pdf"):
                                with open(filepath, "wb") as f:
                                    for chunk in response.iter_content(1024):
                                        f.write(chunk)
                                self.summary["downloads"] += 1
                            else:
                                self.summary["errors"] += 1
                        except requests.RequestException:
                            self.summary["errors"] += 1
                    else:
                        self.summary["skipped"] += 1

        except Exception as e:
            print(f"⚠️ Error extracting reports for {company_code}: {e}")
            self.summary["errors"] += 1


# ---------------- MAIN PROGRAM ----------------
if __name__ == "__main__":
    csv_path = r"C:\Users\lenin\OneDrive\Desktop\Company_Names.csv"
    base_dir = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

    df = pd.read_csv(csv_path, header=None)
    company_codes = df[0].astype(str).tolist()

    summary = {"total_companies": 0, "downloads": 0, "skipped": 0, "errors": 0}
    valid_company_count = 0  # ✅ counter for valid companies

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")

    for company_code in company_codes:
        # Stop early if 500 valid companies already processed
        if valid_company_count >= 500:
            print("\n🛑 Reached 500 valid companies. Stopping further processing.")
            break

        print(f"\n🚀 Processing {company_code} ...")
        summary["total_companies"] += 1

        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 20)

        try:
            driver.get("https://www.bseindia.com/corporates/HistoricalAnnualreport.aspx")

            searcher = CompanySearch(driver, wait)
            searcher.search_company(company_code)

            downloader = AnnualReportDownloader(driver, wait, base_dir, summary)
            downloader.download_reports(company_code)

        except Exception as e:
            print(f"⚠️ Error with {company_code}: {e}")
            summary["errors"] += 1

        finally:
            driver.quit()
            time.sleep(1)

        # ---------------- POST-CHECK for this company only ----------------
        company_dir = os.path.join(base_dir, company_code.upper())
        if os.path.exists(company_dir):
            report_files = os.listdir(company_dir)
            years = []
            for file in report_files:
                for year in range(2016, 2026):
                    if str(year) in file:
                        years.append(year)

            if len(set(years)) == 10:
                valid_company_count += 1
                print(f"✅ {company_code} has all 10 reports ({valid_company_count}/500)")
            else:
                # Delete incomplete folders
                try:
                    shutil.rmtree(company_dir)
                    print(f"🗑️ Deleted folder for {company_code} (only {len(set(years))} reports found)")
                except Exception as e:
                    print(f"⚠️ Could not delete folder for {company_code}: {e}")

    # ---------------- FINAL SUMMARY ----------------
    print("\n📊 --- FINAL RUN SUMMARY ---")
    print(f"Total Companies Processed : {summary['total_companies']}")
    print(f"Total Reports Downloaded  : {summary['downloads']}")
    print(f"Total Reports Skipped     : {summary['skipped']}")
    print(f"Errors Encountered        : {summary['errors']}")
    print(f"✅ Valid Companies (10 reports): {valid_company_count}")
    print("🏁 Job Completed! Time for chai ☕")
