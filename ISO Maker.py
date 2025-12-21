"""
ISO27001 Batch Scoring Script – OFFLINE STABLE FINAL

• Fully offline (no HuggingFace calls)
• Snapshot-safe model loading
• MuPDF errors suppressed
• Batch confirmation restored
• Infinite-loop proof
• Safe Excel backups
"""

import os
import re
import sys
import glob
import zipfile
import shutil
import logging
import warnings
import contextlib
import io
from datetime import datetime

import fitz  # PyMuPDF
import pandas as pd
import numpy as np
from tqdm import tqdm
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import nltk
from nltk.tokenize import sent_tokenize

# =========================
# CONFIG
# =========================
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
EXCEL_MAIN = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
WORK_DIR = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper"
LOG_FILE = os.path.join(WORK_DIR, "iso_processing_log.txt")

HF_MODEL_ROOT = (
    r"C:\Users\lenin\.cache\huggingface\hub"
    r"\models--sentence-transformers--all-mpnet-base-v2\snapshots"
)

BATCH_SIZE = 20
SIM_MENTION = 0.60
SIM_HIGH = 0.72
WINDOW = 1

warnings.filterwarnings("ignore")
logging.basicConfig(filename=LOG_FILE, level=logging.INFO)

# =========================
# LOGGING
# =========================
def log(msg):
    ts = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    print(f"{ts} {msg}")
    logging.info(msg)

# =========================
# NLTK
# =========================
def ensure_nltk():
    try:
        nltk.data.find("tokenizers/punkt")
    except LookupError:
        nltk.download("punkt")

ensure_nltk()

# =========================
# UTILITIES
# =========================
def normalize_filename(f):
    return os.path.basename(str(f)).strip().lower()

# =========================
# EXCEL INTEGRITY
# =========================
def newest_backup():
    files = glob.glob(os.path.join(WORK_DIR, "ISO Data Collection_backup_*.xlsx"))
    return max(files, key=os.path.getmtime) if files else None

def excel_is_valid(path):
    return os.path.exists(path) and os.path.getsize(path) > 1024 and zipfile.is_zipfile(path)

def load_master_excel():
    if excel_is_valid(EXCEL_MAIN):
        try:
            df = pd.read_excel(EXCEL_MAIN, engine="openpyxl")
            log(f"✅ Loaded MAIN Excel ({len(df)} rows)")
            return df
        except Exception as e:
            log(f"❌ Excel read failed: {e}")

    backup = newest_backup()
    if backup:
        shutil.copy2(backup, EXCEL_MAIN)
        log(f"🛠 Recovered Excel from backup: {os.path.basename(backup)}")
        return pd.read_excel(EXCEL_MAIN, engine="openpyxl")

    log("⚠️ No valid Excel found, starting fresh")
    return pd.DataFrame()

# =========================
# QUIET MuPDF
# =========================
@contextlib.contextmanager
def suppress_mupdf():
    stderr = sys.stderr
    try:
        sys.stderr = io.StringIO()
        yield
    finally:
        sys.stderr = stderr

def extract_text(pdf_path):
    try:
        with suppress_mupdf():
            doc = fitz.open(pdf_path)

        try:
            doc.set_option("widget.update-appearance", False)
        except Exception:
            pass

        text = []
        for page in doc:
            try:
                text.append(page.get_text("text") or "")
            except Exception:
                try:
                    text.append(page.get_text("raw") or "")
                except Exception:
                    pass
        doc.close()
        return "\n".join(text)

    except Exception as e:
        log(f"❌ PDF FAILED: {os.path.basename(pdf_path)} | {e}")
        return None

# =========================
# NLP HELPERS
# =========================
def split_sentences(txt):
    txt = re.sub(r"\n+", " ", txt)
    return [s for s in sent_tokenize(txt) if len(s.strip()) > 10]

def has_evidence(sents, idx, window):
    keywords = [
        "implemented", "established", "maintained", "audit", "certified",
        "monitored", "trained", "reviewed", "tested", "assessed"
    ]
    lo, hi = max(0, idx - window), min(len(sents) - 1, idx + window)
    return any(any(k in sents[i].lower() for k in keywords) for i in range(lo, hi + 1))

# =========================
# PROMPT
# =========================
def ask_continue(remaining):
    print("\n" + "-" * 60)
    print(f"📦 Batch complete. PDFs remaining: {remaining}")
    choice = input("➡️ Continue with next batch? [Y/n]: ").strip().lower()
    return choice in ("", "y", "yes")

# =========================
# MODEL (OFFLINE SNAPSHOT)
# =========================
snapshots = glob.glob(os.path.join(HF_MODEL_ROOT, "*"))
if not snapshots:
    log("❌ No offline model snapshot found.")
    sys.exit(1)

EMBEDDING_MODEL = max(snapshots, key=os.path.getmtime)

log("🧠 Loading embedding model (OFFLINE)")
model = SentenceTransformer(
    EMBEDDING_MODEL,
    local_files_only=True,
    device="cpu"
)

# =========================
# ISO DOMAINS
# =========================
ISO_DOMAINS = {
    "A.5": "Information security policies",
    "A.6": "Organization of information security",
    "A.7": "Human resource security",
    "A.8": "Asset management",
    "A.9": "Access control",
    "A.10": "Cryptography",
    "A.11": "Physical security",
    "A.12": "Operations security",
    "A.13": "Communications security",
    "A.14": "System development security",
    "A.15": "Supplier relationships",
    "A.16": "Incident management",
    "A.17": "Business continuity",
    "A.18": "Compliance",
}

keys = list(ISO_DOMAINS.keys())
iso_embeddings = model.encode(list(ISO_DOMAINS.values()), show_progress_bar=False)

# =========================
# LOAD STATE
# =========================
master_df = load_master_excel()
processed_files = {normalize_filename(f) for f in master_df.get("File", [])}

pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
remaining = [f for f in pdf_files if normalize_filename(f) not in processed_files]

log(f"📁 PDFs found: {len(pdf_files)}")
log(f"🆕 PDFs remaining: {len(remaining)}")

# =========================
# MAIN LOOP
# =========================
while remaining:
    batch = remaining[:BATCH_SIZE]
    rows = []

    for pdf in tqdm(batch, desc="Processing PDFs"):
        path = os.path.join(PDF_FOLDER, pdf)
        base = os.path.splitext(pdf)[0]
        company, year = (base.split("_", 1) + [""])[:2]

        text = extract_text(path)

        if text is None:
            rows.append({
                "Company": company,
                "Year": year,
                "File": pdf,
                "Processed_On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Total_Score": 0,
                "Status": "PDF_READ_FAILED",
            })
            continue

        sentences = split_sentences(text)
        if not sentences:
            rows.append({
                "Company": company,
                "Year": year,
                "File": pdf,
                "Processed_On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Total_Score": 0,
                "Status": "NO_TEXT",
            })
            continue

        sims = cosine_similarity(
            model.encode(sentences, show_progress_bar=False),
            iso_embeddings,
        )

        row = {
            "Company": company,
            "Year": year,
            "File": pdf,
            "Processed_On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Status": "OK",
        }

        for j, key in enumerate(keys):
            idx = int(np.argmax(sims[:, j]))
            score = 0
            if sims[idx, j] >= SIM_MENTION:
                score = 1
                if sims[idx, j] >= SIM_HIGH or has_evidence(sentences, idx, WINDOW):
                    score = 2
            row[key] = score

        row["Total_Score"] = sum(row[k] for k in keys)
        rows.append(row)

    df = pd.DataFrame(rows)

    if os.path.exists(EXCEL_MAIN):
        backup = EXCEL_MAIN.replace(".xlsx", f"_backup_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        shutil.copy2(EXCEL_MAIN, backup)

    master_df = pd.concat([master_df, df], ignore_index=True)
    master_df.drop_duplicates(subset=["File"], keep="last", inplace=True)
    master_df.to_excel(EXCEL_MAIN, index=False, engine="openpyxl")

    log(f"💾 Saved {len(df)} rows")

    remaining = remaining[BATCH_SIZE:]

    if remaining and not ask_continue(len(remaining)):
        log("⏹ User stopped safely")
        break

log("🎉 COMPLETE — SCRIPT FINISHED SAFELY")
