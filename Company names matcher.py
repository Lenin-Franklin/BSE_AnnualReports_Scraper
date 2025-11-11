import pandas as pd
from fuzzywuzzy import process

# -----------------------------
# File paths
# -----------------------------
file_main = r"C:\Users\lenin\OneDrive\Desktop\List of companies with 10 year data.xlsx"
file_reference = r"C:\Users\lenin\Downloads\Top2000Companies_as_on_31March2024_based_on_market_capitalisation_updated\Top 2000 Companies as on 31 March 2024 based on market capitalisation_Updated.xlsx"

# -----------------------------
# Load data
# -----------------------------
df_main = pd.read_excel(file_main)
df_ref = pd.read_excel(file_reference)

# -----------------------------
# Clean names in both datasets
# -----------------------------
df_main['Company Name Clean'] = df_main['Company Name'].str.strip().str.upper()
df_ref['SCRIP_LONG_NAME Clean'] = df_ref['SCRIP_LONG_NAME'].str.strip().str.upper()

# -----------------------------
# EXACT MATCH
# -----------------------------
df_main = df_main.merge(
    df_ref[['SCRIP_LONG_NAME Clean', 'SCRIP_CODE']],
    left_on='Company Name Clean',
    right_on='SCRIP_LONG_NAME Clean',
    how='left'
)

# Rename for clarity
df_main.rename(columns={'SCRIP_CODE': 'Exact_Match_Code'}, inplace=True)

# -----------------------------
# Identify rows with no exact match
# -----------------------------
unmatched = df_main[df_main['Exact_Match_Code'].isna()].copy()

print(f"Unmatched after exact match: {len(unmatched)}")

# -----------------------------
# FUZZY MATCH ONLY FOR UNMATCHED
# -----------------------------
ref_names = df_ref['SCRIP_LONG_NAME Clean'].tolist()
code_map = df_ref.set_index('SCRIP_LONG_NAME Clean')['SCRIP_CODE'].to_dict()

THRESHOLD = 90  # Adjust if required

def fuzzy_match(name):
    match, score = process.extractOne(name, ref_names)
    return match if score >= THRESHOLD else None

unmatched['Fuzzy_Matched_Name'] = unmatched['Company Name Clean'].apply(fuzzy_match)
unmatched['Fuzzy_Match_Code'] = unmatched['Fuzzy_Matched_Name'].map(code_map)

# -----------------------------
# Merge fuzzy results back into df_main
# -----------------------------
df_main = df_main.merge(
    unmatched[['Company Name Clean', 'Fuzzy_Match_Code']],
    on='Company Name Clean',
    how='left'
)

# -----------------------------
# Final company code: exact match first, else fuzzy
# -----------------------------
df_main['Final_Company_Code'] = df_main['Exact_Match_Code'].fillna(df_main['Fuzzy_Match_Code'])

# -----------------------------
# Optional: list of still-unmatched companies
# -----------------------------
still_unmatched = df_main[df_main['Final_Company_Code'].isna()][['Company Name']]

# -----------------------------
# Save outputs
# -----------------------------
output_file = r"C:\Users\lenin\OneDrive\Desktop\List of companies with 10 year data_FINAL_UPDATED.xlsx"
df_main.to_excel(output_file, index=False)

unmatched_report = r"C:\Users\lenin\OneDrive\Desktop\Unmatched_companies_report.xlsx"
still_unmatched.to_excel(unmatched_report, index=False)

print("Processing complete.")
print("Final updated file:", output_file)
print("Unmatched companies report:", unmatched_report)
