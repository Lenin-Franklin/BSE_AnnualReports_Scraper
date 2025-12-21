import os
import sys
import pandas as pd
from datetime import datetime
from colorama import init

# =========================
# ANSI / HACKER CONSOLE
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

def hacker(msg, level="INFO", color=CYAN):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"{DIM}[{ts}]{RESET} {color}{level:<6}{RESET} :: {msg}", flush=True)

def phase(msg):
    print(f"\n{PURPLE}{BOLD}[⚡] >>> {msg.upper()} <<< {RESET}\n", flush=True)

# =========================
# CONFIG
# =========================
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
EXCEL_MAIN = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
OUTPUT_REPORT = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO_Unprocessed_Forensics.xlsx"

FAILED_STATUSES = {"PDF_READ_FAILED", "NO_TEXT"}

# =========================
# HELPERS
# =========================
def normalize(name: str) -> str:
    return os.path.basename(str(name)).strip().lower()

# =========================
# ENTRY POINT
# =========================
phase("booting forensic subsystem")

hacker("Initializing read-only diagnostics", "BOOT", GREEN)

if not os.path.exists(EXCEL_MAIN):
    hacker("Master Excel not found — aborting mission", "ABORT", RED)
    sys.exit(1)

if not os.path.exists(PDF_FOLDER):
    hacker("PDF directory missing — aborting mission", "ABORT", RED)
    sys.exit(1)

# =========================
# LOAD EXCEL
# =========================
phase("syncing excel intelligence")

df = pd.read_excel(EXCEL_MAIN, engine="openpyxl")
hacker(f"Excel index loaded :: {len(df)} records", "SYNC", GREEN)

excel_files = df.get("File", pd.Series()).dropna().astype(str)
excel_norm = set(normalize(f) for f in excel_files)

# =========================
# LOAD PDFs
# =========================
phase("scanning filesystem")

pdfs = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
pdf_norm_map = {normalize(f): f for f in pdfs}

hacker(f"PDF payloads discovered :: {len(pdfs)}", "SCAN", GREEN)

# =========================
# FORENSIC ANALYSIS
# =========================
phase("correlating targets")

records = []

# ---- PDFs never processed ----
for norm, actual in pdf_norm_map.items():
    if norm not in excel_norm:
        records.append({
            "File": actual,
            "Status": "NOT_IN_EXCEL",
            "Reason": "PDF exists on disk but has never been processed"
        })

# ---- Excel rows with failures or missing PDFs ----
for _, row in df.iterrows():
    file = str(row.get("File", "")).strip()
    norm = normalize(file)
    status = str(row.get("Status", "OK")).strip()

    if not file:
        continue

    if norm not in pdf_norm_map:
        records.append({
            "File": file,
            "Status": "PDF_MISSING_ON_DISK",
            "Reason": "Excel entry exists but PDF file is missing"
        })
        continue

    if status in FAILED_STATUSES:
        records.append({
            "File": file,
            "Status": status,
            "Reason": "Previously attempted and failed — intentionally not retried"
        })

# =========================
# RESULTS
# =========================
phase("intel summary")

report_df = pd.DataFrame(records).drop_duplicates(subset=["File"])

if report_df.empty:
    hacker("No anomalies detected — dataset is clean", "OK", GREEN)
else:
    counts = report_df["Status"].value_counts()
    for k, v in counts.items():
        color = RED if "FAILED" in k or "MISSING" in k else YELLOW
        hacker(f"{k:<22} :: {v}", "FLAG", color)

# =========================
# EXPORT REPORT
# =========================
phase("writing forensic artifact")

report_df["Checked_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
report_df.to_excel(OUTPUT_REPORT, index=False, engine="openpyxl")

hacker("Forensic report written successfully", "WRITE", GREEN)
hacker(OUTPUT_REPORT, "PATH", BLUE)

phase("mission complete")
