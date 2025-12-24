# ============================================================
# OG ISO27001 SCRAPPER — FINAL SAFE VERSION
# (Logic preserved, safety hardened)
# ============================================================

import os
import re
import time
import shutil
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
from tqdm import tqdm
from datetime import datetime
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import nltk
from nltk.tokenize import sent_tokenize
import torch

# =========================
# CONFIG
# =========================
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
EXCEL_PATH = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
CSV_SHADOW = EXCEL_PATH.replace(".xlsx", ".csv")
LOG_FILE = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\iso_log.txt"
QUARANTINE = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\quarantine"

BATCH_SIZE = 1000
MAX_SENTENCES = 1500

SIM_MENTION = 0.60
SIM_HIGH = 0.72
WINDOW = 1

os.makedirs(QUARANTINE, exist_ok=True)

# =========================
# LOGGING
# =========================
def log(msg):
    ts = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    print(f"{ts} {msg}")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{ts} {msg}\n")

# =========================
# NLTK
# =========================
try:
    nltk.data.find("tokenizers/punkt")
except LookupError:
    nltk.download("punkt")

# =========================
# ISO DESCRIPTIONS (UNCHANGED)
# =========================
iso_descriptions = {
    "A.5 Information security policies":
        "Information security policies and governance; policy framework, approved and communicated.",
    "A.6 Organization of information security":
        "Organization, roles and responsibilities for information security; segregation of duties; contact with authorities and security committees.",
    "A.7 Human resource security":
        "Employee screening, onboarding, training, awareness, and termination procedures related to security.",
    "A.8 Asset management":
        "Inventory of assets, asset ownership, classification of information and acceptable use.",
    "A.9 Access control":
        "User access management, access rights, least privilege, MFA, password policy, privileged accounts.",
    "A.10 Cryptography":
        "Encryption, cryptographic controls, key management, digital signatures and related controls.",
    "A.11 Physical and environmental security":
        "Physical protection of facilities, secure areas, CCTV, environmental controls and access to premises.",
    "A.12 Operations security":
        "Operations procedures, change management, backups, logging, monitoring, malware protection.",
    "A.13 Communications security":
        "Network security, secure transmission, VPNs, firewalls, email security, secure protocols.",
    "A.14 System acquisition, development and maintenance":
        "Secure development lifecycle, security requirements, code review, and application security testing.",
    "A.15 Supplier relationships":
        "Third-party/vendor security, supplier agreements, monitoring and risk assessments.",
    "A.16 Information security incident management":
        "Incident response, reporting, management, investigations and root-cause analysis.",
    "A.17 Information security aspects of business continuity management":
        "Business continuity, disaster recovery, contingency plans, and continuity testing for information security.",
    "A.18 Compliance":
        "Legal, regulatory and contractual compliance, data protection laws, audits and certifications."
}

iso_keys = list(iso_descriptions.keys())

# =========================
# MODEL
# =========================
device = "cuda" if torch.cuda.is_available() else "cpu"
torch.set_grad_enabled(False)

log(f"Loading embedding model ({device.upper()})")
model = SentenceTransformer("all-mpnet-base-v2", device=device)
iso_embeddings = model.encode(list(iso_descriptions.values()), show_progress_bar=False)
log("Model ready")

# =========================
# SAFETY HELPERS
# =========================
def excel_safe(val):
    if not isinstance(val, str):
        return val
    return re.sub(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F\uD800-\uDFFF\uFFFE\uFFFF]",
        "",
        val
    )

def quarantine_pdf(pdf, reason):
    try:
        shutil.move(
            os.path.join(PDF_FOLDER, pdf),
            os.path.join(QUARANTINE, pdf)
        )
        log(f"[QUARANTINE] {pdf} :: {reason}")
    except Exception:
        pass

# =========================
# NLP HELPERS
# =========================
def extract_text(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = []
        for page in doc:
            text.append(page.get_text("text") or "")
        doc.close()
        return "\n".join(text)
    except Exception:
        return None

def split_sentences(txt):
    txt = re.sub(r"\n+", " ", txt)
    return [s for s in sent_tokenize(txt) if len(s.strip()) > 10][:MAX_SENTENCES]

def has_evidence(sents, idx):
    words = [
        "implemented", "established", "maintained",
        "audit", "certified", "monitored",
        "trained", "reviewed"
    ]
    lo, hi = max(0, idx - WINDOW), min(len(sents) - 1, idx + WINDOW)
    return any(any(w in sents[i].lower() for w in words) for i in range(lo, hi + 1))

# =========================
# LOAD EXISTING DATA (RESUME)
# =========================
if os.path.exists(EXCEL_PATH):
    df_existing = pd.read_excel(EXCEL_PATH)
    processed = set(df_existing["File"].astype(str))
else:
    df_existing = pd.DataFrame()
    processed = set()

pdfs = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
pending = [p for p in pdfs if p not in processed][:BATCH_SIZE]

log(f"Processing {len(pending)} PDFs")

# =========================
# MAIN LOOP
# =========================
rows = []

for pdf in tqdm(pending, desc="Scoring PDFs"):
    base = os.path.splitext(pdf)[0]
    try:
        year, company = base.split("_", 1)
    except ValueError:
        year, company = "", base

    path = os.path.join(PDF_FOLDER, pdf)
    text = extract_text(path)

    if not text:
        quarantine_pdf(pdf, "NO_TEXT")
        continue

    sentences = split_sentences(text)
    if not sentences:
        quarantine_pdf(pdf, "NO_SENTENCES")
        continue

    try:
        sent_emb = model.encode(sentences, batch_size=32, show_progress_bar=False)
    except Exception:
        quarantine_pdf(pdf, "EMBEDDING_FAIL")
        continue

    sims = cosine_similarity(sent_emb, iso_embeddings)

    row = {
        "Company": company,
        "Year": year,
        "File": pdf
    }

    total = 0

    for j, domain in enumerate(iso_keys):
        idx = int(np.argmax(sims[:, j]))
        sim = float(sims[idx, j])
        snippet = sentences[idx]

        score = 0
        reason = "no_match"

        if sim >= SIM_MENTION:
            score = 1
            reason = "scored"
            if sim >= SIM_HIGH or has_evidence(sentences, idx):
                score = 2

        # WRITE BOTH SCHEMA LAYERS
        row[domain] = score
        row[f"{domain}__score"] = score
        row[f"{domain}__sim"] = round(sim, 4)
        row[f"{domain}__snippet"] = snippet
        row[f"{domain}__reason"] = reason

        total += score

    row["Total_Score"] = total
    row["Processed_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row["Status"] = "OK"

    rows.append(row)

# =========================
# SAFE SAVE (ATOMIC)
# =========================
df_new = pd.DataFrame(rows)

for col in df_new.columns:
    df_new[col] = df_new[col].map(excel_safe)

df_final = pd.concat([df_existing, df_new], ignore_index=True)

tmp_path = EXCEL_PATH.replace(".xlsx", "_TMP.xlsx")
df_final.to_excel(tmp_path, index=False)
os.replace(tmp_path, EXCEL_PATH)

df_final.to_csv(CSV_SHADOW, index=False)

log(f"Saved {len(df_new)} rows safely")
log("SAFE TO CLOSE — RESUME READY")
