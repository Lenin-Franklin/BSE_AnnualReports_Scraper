import pandas as pd

# =========================
# CONFIG
# =========================
EXCEL_PATH = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
OUTPUT_CSV = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\missing_company_years.csv"

EXPECTED_YEARS = set(str(y) for y in range(2016, 2026))

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(EXCEL_PATH, dtype=str)

# Normalize
df["Company"] = df["Company"].str.strip()
df["Year"] = df["Year"].str.strip()

# =========================
# ANALYSIS
# =========================
results = []

companies = sorted(df["Company"].dropna().unique())

for company in companies:
    years_present = set(
        df.loc[df["Company"] == company, "Year"]
        .dropna()
        .unique()
    )

    missing_years = sorted(EXPECTED_YEARS - years_present)

    if missing_years:
        results.append({
            "Company": company,
            "Years_Present": ", ".join(sorted(years_present)),
            "Missing_Years": ", ".join(missing_years),
            "Missing_Count": len(missing_years)
        })

# =========================
# OUTPUT
# =========================
missing_df = pd.DataFrame(results).sort_values(
    by=["Missing_Count", "Company"],
    ascending=[False, True]
)

missing_df.to_csv(OUTPUT_CSV, index=False)

# =========================
# SUMMARY
# =========================
print("========== ISO DATA AUDIT ==========")
print(f"Total companies found        : {len(companies)}")
print(f"Companies with missing PDFs  : {len(missing_df)}")
print(f"Total missing PDFs detected : {missing_df['Missing_Count'].sum()}")
print("-----------------------------------")
print("Top 10 companies with most missing years:")
print(missing_df.head(10).to_string(index=False))
print("-----------------------------------")
print(f"Detailed report saved to:\n{OUTPUT_CSV}")
print("===================================")
