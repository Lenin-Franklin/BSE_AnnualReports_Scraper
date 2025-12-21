"""
ISO27001 Batch Scoring Script
CYBER OPS EDITION — ANSI NEON
Self-healing • Offline-safe • Stable • Hacker-style output
"""
from colorama import init
init(autoreset=True)

# =========================================================
# BOOT SEQUENCE
# =========================================================
import sys, os, time, glob, zipfile, shutil, logging, warnings, re, random
from datetime import datetime

# ---------------- ANSI COLORS ----------------
ANSI_ENABLED = sys.stdout.isatty()

def c(code):
    return code if ANSI_ENABLED else ""

RESET   = c("\033[0m")
GREEN   = c("\033[92m")
CYAN    = c("\033[96m")
RED     = c("\033[91m")
YELLOW  = c("\033[93m")
PURPLE  = c("\033[95m")
BLUE    = c("\033[94m")
DIM     = c("\033[2m")
BOLD    = c("\033[1m")

GLITCH = ["▌", "▐", "█", "▒", "░", "▓"]

LEVEL_COLOR = {
    "BOOT": PURPLE,
    "OK": GREEN,
    "INIT": CYAN,
    "SCAN": BLUE,
    "EXEC": PURPLE,
    "ARMED": GREEN,
    "WARN": YELLOW,
    "FAIL": RED,
    "REPAIR": YELLOW,
    "SYNC": CYAN,
    "SUCCESS": GREEN,
    "ABORT": RED,
}

def hacker(msg, level="INFO"):
    ts = datetime.now().strftime("%H:%M:%S")
    color = LEVEL_COLOR.get(level, CYAN)
    glitch = random.choice(GLITCH)
    print(
        f"{DIM}[{ts}]{RESET} "
        f"{color}{glitch} {level:<7}{RESET} :: "
        f"{BOLD}{msg}{RESET}",
        flush=True
    )
    logging.info(msg)

def phase(msg):
    print(
        f"\n{BOLD}{PURPLE}[⚡] >>> {msg.upper()} <<<{RESET}\n",
        flush=True
    )
    time.sleep(0.15)

# =========================================================
# CONFIG
# =========================================================
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
EXCEL_MAIN = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
WORK_DIR = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper"
LOG_FILE = os.path.join(WORK_DIR, "iso_processing_log.txt")

HF_MODEL_NAME = "sentence-transformers/all-mpnet-base-v2"
HF_CACHE_ROOT = os.path.expanduser(r"~\.cache\huggingface\hub")
HF_MODEL_DIR = os.path.join(
    HF_CACHE_ROOT, "models--sentence-transformers--all-mpnet-base-v2"
)

BATCH_SIZE = 20
MAX_SENTENCES = 200

SIM_MENTION = 0.60
SIM_HIGH = 0.72
WINDOW = 1

warnings.filterwarnings("ignore")
logging.basicConfig(filename=LOG_FILE, level=logging.INFO)

phase("boot sequence initiated")
hacker("Runtime parameters locked", "BOOT")

# =========================================================
# IMPORTS
# =========================================================
phase("injecting modules")

import fitz
fitz.TOOLS.mupdf_display_errors(False)
hacker("MuPDF bindings injected", "OK")

import pandas as pd
import numpy as np
from tqdm import tqdm
hacker("Data pipeline modules online", "OK")

from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
hacker("Neural inference libraries armed", "OK")

import nltk
from nltk.tokenize import sent_tokenize
hacker("Linguistic engine online", "OK")

# =========================================================
# NLTK CHECK
# =========================================================
phase("linguistic subsystem")
try:
    nltk.data.find("tokenizers/punkt")
    hacker("Tokenizer integrity verified", "OK")
except LookupError:
    hacker("Tokenizer missing — deploying payload", "REPAIR")
    nltk.download("punkt")

# =========================================================
# MODEL SELF-HEAL
# =========================================================
phase("neural core integrity check")

def model_is_valid():
    if not os.path.exists(HF_MODEL_DIR):
        return False
    snaps = glob.glob(os.path.join(HF_MODEL_DIR, "snapshots", "*"))
    for s in snaps:
        files = set(os.listdir(s))
        if {"config.json", "modules.json"}.issubset(files) and any(
            f.endswith((".bin", ".safetensors")) for f in files
        ):
            return True
    return False

def has_internet():
    try:
        import socket
        socket.create_connection(("huggingface.co", 443), timeout=3)
        return True
    except Exception:
        return False

if not model_is_valid():
    hacker("Neural cache corrupted", "WARN")
    if has_internet():
        hacker("Uplink detected — rebuilding core", "REPAIR")
        shutil.rmtree(HF_MODEL_DIR, ignore_errors=True)
        SentenceTransformer(HF_MODEL_NAME)
        hacker("Neural core reconstructed", "OK")
    else:
        hacker("Offline & corrupted — aborting mission", "FAIL")
        sys.exit(1)
else:
    hacker("Neural core integrity confirmed", "OK")

# =========================================================
# LOAD MODEL
# =========================================================
phase("activating neural core")
model = SentenceTransformer(HF_MODEL_NAME, local_files_only=True, device="cpu")
hacker("Embedding engine online", "ARMED")

# =========================================================
# ISO DOMAINS
# =========================================================
phase("loading control vectors")

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
iso_embeddings = model.encode(
    list(ISO_DOMAINS.values()),
    normalize_embeddings=True,
    show_progress_bar=False
)
hacker("ISO vectors injected into memory", "ARMED")

# =========================================================
# UTILITIES
# =========================================================
def normalize_filename(f):
    return os.path.basename(str(f)).strip().lower()

def split_sentences(txt):
    txt = re.sub(r"\n+", " ", txt)
    return [s for s in sent_tokenize(txt) if len(s.strip()) > 10][:MAX_SENTENCES]

def has_evidence(sents, idx, window):
    keywords = [
        "implemented", "established", "maintained", "audit", "certified",
        "monitored", "trained", "reviewed", "tested", "assessed"
    ]
    lo, hi = max(0, idx - window), min(len(sents) - 1, idx + window)
    return any(any(k in sents[i].lower() for k in keywords) for i in range(lo, hi + 1))

# =========================================================
# EXCEL LOAD
# =========================================================
phase("state synchronization")

def excel_is_valid(path):
    return os.path.exists(path) and os.path.getsize(path) > 1024 and zipfile.is_zipfile(path)

def load_master_excel():
    if excel_is_valid(EXCEL_MAIN):
        df = pd.read_excel(EXCEL_MAIN, engine="openpyxl")
        hacker(f"Excel index synchronized ({len(df)} records)", "SYNC")
        return df
    hacker("No valid Excel found — initializing empty state", "WARN")
    return pd.DataFrame()

master_df = load_master_excel()
processed = {normalize_filename(f) for f in master_df.get("File", [])}

pdfs = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
remaining = [f for f in pdfs if normalize_filename(f) not in processed]

hacker(f"Targets discovered: {len(pdfs)}", "SCAN")
hacker(f"Targets pending: {len(remaining)}", "SCAN")

# =========================================================
# PDF EXTRACTION
# =========================================================
def extract_text(path):
    try:
        doc = fitz.open(path)
        return "\n".join(p.get_text("text") or "" for p in doc)
    except Exception:
        return None

# =========================================================
# MAIN LOOP
# =========================================================
phase("execution phase")

while remaining:
    batch = remaining[:BATCH_SIZE]
    hacker(f"Deploying batch payload :: {len(batch)} targets", "EXEC")
    rows = []

    for pdf in tqdm(batch, desc="Neural Sweep"):
        text = extract_text(os.path.join(PDF_FOLDER, pdf))
        base = os.path.splitext(pdf)[0]
        company, year = (base.split("_", 1) + [""])[:2]

        if not text:
            hacker(f"Malformed payload detected — skipping {pdf}", "WARN")
            rows.append({
                "Company": company,
                "Year": year,
                "File": pdf,
                "Processed_On": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Total_Score": 0,
                "Status": "PDF_READ_FAILED",
            })
            continue

        sents = split_sentences(text)
        embeds = model.encode(sents, normalize_embeddings=True, show_progress_bar=False)
        sims = cosine_similarity(embeds, iso_embeddings)

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
                if sims[idx, j] >= SIM_HIGH or has_evidence(sents, idx, WINDOW):
                    score = 2
            row[key] = score

        row["Total_Score"] = sum(row[k] for k in keys)
        rows.append(row)

    df = pd.DataFrame(rows)
    master_df = pd.concat([master_df, df], ignore_index=True)
    master_df.drop_duplicates(subset=["File"], keep="last", inplace=True)
    master_df.to_excel(EXCEL_MAIN, index=False, engine="openpyxl")

    remaining = remaining[BATCH_SIZE:]
    hacker(f"Batch committed — {len(df)} records written", "OK")

    if remaining:
        if input(f"{PURPLE}>>> Continue breach? [Y/n]: {RESET}").lower().startswith("n"):
            hacker("Operator aborted execution", "ABORT")
            break

phase("mission complete")
hacker("All operations finalized successfully", "SUCCESS")
