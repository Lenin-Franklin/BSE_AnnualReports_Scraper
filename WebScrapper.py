import os
import requests
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

#-----------------------------------------------------------------
# BSE Annual Report Scraper
# Selenium drives Chrome 🏎️, Requests grabs PDFs 📑,
# and Pandas keeps track of which companies to hunt down.
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
            time.sleep(2)  # give dropdown a chance to wake up ☕
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
            print(f"Hold on yerume 🐃")


class AnnualReportDownloader:
    def __init__(self, driver, wait, base_dir, summary):
        self.driver = driver
        self.wait = wait
        self.base_dir = base_dir
        self.summary = summary
        os.makedirs(base_dir, exist_ok=True)

    def download_reports(self, company_code: str, limit: int = 7):
        """Extracts annual report table and downloads only the latest N unique PDFs"""
        try:
            report_table = self.wait.until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_grdAnnualReport"))
            )
            rows = report_table.find_elements(By.TAG_NAME, "tr")

            print(f"📑 Found {len(rows)-1} reports for {company_code}")

            company_dir = os.path.join(self.base_dir, company_code.upper())
            os.makedirs(company_dir, exist_ok=True)

            latest_rows = []
            seen_years = set()

            for row in rows[1:]:
                cols = row.find_elements(By.TAG_NAME, "td")
                if not cols:
                    continue

                year = cols[0].text.strip()
                if not year or year in seen_years:
                    continue

                seen_years.add(year)
                latest_rows.append(row)

                if len(latest_rows) >= limit:
                    break

            # --- PDF download using Selenium cookies + headers ---
            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                              "AppleWebKit/537.36 (KHTML, like Gecko) "
                              "Chrome/116.0 Safari/537.36"
            }
            cookies = {c['name']: c['value'] for c in self.driver.get_cookies()}

            for row in latest_rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                year = cols[0].text.strip()
                pdf_link = cols[-1].find_element(By.TAG_NAME, "a").get_attribute("href")

                if pdf_link and pdf_link.endswith(".pdf"):
                    filepath = os.path.join(company_dir, f"{year}_{company_code}.pdf")

                    if not os.path.exists(filepath):
                        response = requests.get(pdf_link, headers=headers, cookies=cookies, stream=True)
                        # Make sure we actually got a PDF
                        if response.headers.get("Content-Type", "").lower().startswith("application/pdf"):
                            with open(filepath, "wb") as f:
                                for chunk in response.iter_content(1024):
                                    f.write(chunk)
                            print(f"✅ Downloaded {filepath}")
                            self.summary["downloads"] += 1
                        else:
                            print(f"⚠️ Failed to download {year} for {company_code} (not a valid PDF)")
                            self.summary["errors"] += 1
                    else:
                        print(f"⏩ Skipped {year} (already exists)")
                        self.summary["skipped"] += 1
                else:
                    print(f"⚠️ No valid PDF link for {year}")

        except Exception as e:
            print(f"⚠️ Error extracting reports for {company_code}: {e}")
            self.summary["errors"] += 1


# --- Main Program ---
if __name__ == "__main__":
    # ✏️ Edit these paths before running
    csv_path = r"C:\Users\lenin\OneDrive\Desktop\Company_Names.csv"
    base_dir = r"C:\Users\lenin\OneDrive\Desktop\NSE Scraper"

    df = pd.read_csv(csv_path, header=None)
    company_codes = df[0].astype(str).tolist()

    summary = {"total_companies": 0, "downloads": 0, "skipped": 0, "errors": 0}

    for company_code in company_codes:
        print(f"\n🚀 Processing {company_code} ...")
        summary["total_companies"] += 1

        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 15)

        try:
            driver.get("https://www.bseindia.com/corporates/HistoricalAnnualreport.aspx")

            searcher = CompanySearch(driver, wait)
            searcher.search_company(company_code)

            downloader = AnnualReportDownloader(driver, wait, base_dir, summary)
            downloader.download_reports(company_code, limit=7)

        except Exception as e:
            print(f"⚠️ Error with {company_code}: {e}")
            summary["errors"] += 1

        finally:
            time.sleep(3)
            driver.quit()

    print("\n📊 --- RUN SUMMARY ---")
    print(f"Total Companies Processed : {summary['total_companies']}")
    print(f"Total Reports Downloaded  : {summary['downloads']}")
    print(f"Total Reports Skipped     : {summary['skipped']}")
    print(f"Errors Encountered        : {summary['errors']}")
    print("✅ Job Completed! Time for chai ☕")