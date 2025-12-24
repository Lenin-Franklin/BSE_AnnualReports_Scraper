import os
import sys
import pandas as pd
from datetime import datetime
from colorama import init

# =========================
# NEON CONSOLE
# =========================
init(autoreset=True)

RESET   = "\033[0m"
GREEN   = "\033[92m"
CYAN    = "\033[96m"
RED     = "\033[91m"
YELLOW  = "\033[93m"
PURPLE  = "\033[95m"
BLUE    = "\033[94m"
DIM     = "\033[2m"
BOLD    = "\033[1m"

def cyber(msg, level="INFO", color=CYAN):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"{DIM}[{ts}]{RESET} {color}{level:<6}{RESET} :: {msg}", flush=True)

def phase(msg):
    print(f"\n{PURPLE}{BOLD}[⚡] >>> {msg.upper()} <<< {RESET}\n", flush=True)

# =========================
# CONFIG
# =========================
EXCEL_MAIN = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
QA_REPORT  = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO_Data_Integrity_Report.xlsx"

ISO_KEYS = [
    "A.5","A.6","A.7","A.8","A.9","A.10","A.11",
    "A.12","A.13","A.14","A.15","A.16","A.17","A.18"
]

VALID_STATUSES = {"OK", "PDF_READ_FAILED", "NO_TEXT"}

# =========================
# ENTRY POINT
# =========================
phase("initializing integrity scanner")

cyber("Booting audit engine", "BOOT", GREEN)

if not os.path.exists(EXCEL_MAIN):
    cyber("Master Excel not found — aborting scan", "ABORT", RED)
    sys.exit(1)

# =========================
# LOAD DATA
# =========================
phase("syncing primary datastore")

df = pd.read_excel(EXCEL_MAIN, engine="openpyxl")
cyber(f"Records loaded :: {len(df)}", "SYNC", GREEN)

issues = []

# =========================
# STRUCTURE CHECK
# =========================
phase("verifying schema integrity")

required_cols = {"Company","Year","File","Status","Total_Score","Processed_On"} | set(ISO_KEYS)
missing = required_cols - set(df.columns)

if missing:
    cyber(f"Missing columns detected :: {missing}", "FAULT", RED)
    issues.append({
        "Issue_Type": "MISSING_COLUMNS",
        "Details": ", ".join(sorted(missing))
    })
else:
    cyber("Schema integrity verified", "OK", GREEN)

# =========================
# DUPLICATE FILE CHECK
# =========================
phase("scanning duplicate payloads")

dupes = df[df.duplicated(subset=["File"], keep=False)]
for f in dupes["File"].unique():
    issues.append({
        "Issue_Type": "DUPLICATE_FILE",
        "File": f
    })

cyber(f"Duplicate files flagged :: {dupes['File'].nunique()}", "SCAN", YELLOW)

# =========================
# ROW-LEVEL VALIDATION
# =========================
phase("validating record payloads")

for _, row in df.iterrows():
    file = str(row.get("File", "")).strip()
    status = str(row.get("Status", "")).strip()
    total = row.get("Total_Score")

    domain_vals = []

    for k in ISO_KEYS:
        v = row.get(k)
        if pd.isna(v):
            issues.append({
                "Issue_Type": "MISSING_DOMAIN_SCORE",
                "File": file,
                "Domain": k
            })
        else:
            domain_vals.append(v)
            if v not in (0,1,2):
                issues.append({
                    "Issue_Type": "INVALID_DOMAIN_VALUE",
                    "File": file,
                    "Domain": k,
                    "Value": v
                })

    if domain_vals:
        expected = sum(domain_vals)
        if total != expected:
            issues.append({
                "Issue_Type": "TOTAL_SCORE_MISMATCH",
                "File": file,
                "Expected": expected,
                "Actual": total
            })

    if status == "OK" and total == 0:
        issues.append({
            "Issue_Type": "OK_WITH_ZERO_SCORE",
            "File": file
        })

    if status != "OK" and any(v > 0 for v in domain_vals):
        issues.append({
            "Issue_Type": "FAILED_WITH_SCORES",
            "File": file,
            "Status": status
        })

    if status not in VALID_STATUSES:
        issues.append({
            "Issue_Type": "UNKNOWN_STATUS",
            "File": file,
            "Status": status
        })

# =========================
# OUTLIER ANALYSIS
# =========================
phase("detecting temporal score drift")

grp = df[df["Status"] == "OK"].groupby(["Company","Year"])["Total_Score"]

for (company, year), scores in grp:
    if scores.max() - scores.min() >= 10:
        issues.append({
            "Issue_Type": "COMPANY_YEAR_SCORE_VARIANCE",
            "Company": company,
            "Year": year,
            "Min": int(scores.min()),
            "Max": int(scores.max())
        })

# =========================
# RESULTS
# =========================
phase("integrity verdict")

qa_df = pd.DataFrame(issues)

if qa_df.empty:
    cyber("No anomalies detected — dataset integrity intact", "CLEAN", GREEN)
else:
    counts = qa_df["Issue_Type"].value_counts()
    for k, v in counts.items():
        cyber(f"{k:<30} :: {v}", "FLAG", RED if "MISMATCH" in k or "INVALID" in k else YELLOW)

# =========================
# EXPORT REPORT
# =========================
phase("writing forensic artifact")

qa_df["Checked_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
qa_df.to_excel(QA_REPORT, index=False, engine="openpyxl")

cyber("Integrity report written successfully", "WRITE", GREEN)
cyber(QA_REPORT, "PATH", BLUE)

phase("scan complete — system stable")
