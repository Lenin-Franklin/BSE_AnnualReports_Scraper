# ============================================================
# ISO27001 OG SCRAPPER — KAGGLE SAFE VERSION
# ============================================================

import os
import re
import time
import fitz
import torch
import pandas as pd
import numpy as np
from tqdm import tqdm
from datetime import datetime
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import nltk
from nltk.tokenize import sent_tokenize
from openpyxl.utils.exceptions import IllegalCharacterError

# =========================
# PATHS (KAGGLE)
# =========================
PDF_FOLDER = "/kaggle/working/Company_PDF"
EXCEL_PATH = "/kaggle/working/ISO Data Collection.xlsx"
LOG_FILE = "/kaggle/working/iso_log.txt"

BATCH_SIZE = 200
SIM_MENTION = 0.60
SIM_HIGH = 0.72
WINDOW = 1
MAX_SENTENCES = 1500
EXCEL_WRITE_RETRIES = 3

# =========================
# DEVICE (AUTO)
# =========================
device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"[⚡] Using device: {device}")

# =========================
# NLTK
# =========================
nltk.download("punkt", quiet=True)
nltk.download("punkt_tab", quiet=True)

# =========================
# ISO DESCRIPTIONS (ORIGINAL)
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

iso_core_keywords = {
    "A.5 Information security policies": ["information security policy", "security policies", "policy framework"],
    "A.6 Organization of information security": ["roles and responsibilities", "segregation of duties", "security committee"],
    "A.7 Human resource security": ["security awareness", "background checks", "training and awareness"],
    "A.8 Asset management": ["asset inventory", "information classification", "asset register"],
    "A.9 Access control": ["access control", "least privilege", "multi-factor authentication", "mfa"],
    "A.10 Cryptography": ["encryption", "cryptographic", "key management"],
    "A.11 Physical and environmental security": ["physical security", "cctv", "secure areas"],
    "A.12 Operations security": ["change management", "logging and monitoring", "backup"],
    "A.13 Communications security": ["network security", "vpn", "firewall", "email security"],
    "A.14 System acquisition, development and maintenance": ["secure development", "secure coding", "code review"],
    "A.15 Supplier relationships": ["supplier", "third-party", "vendor management", "outsourcing"],
    "A.16 Information security incident management": ["incident management", "security incident", "breach response"],
    "A.17 Business continuity": ["business continuity", "disaster recovery", "continuity plan"],
    "A.18 Compliance": ["regulatory compliance", "data protection", "security audits", "certified"]
}

ISO_KEYS = list(iso_descriptions.keys())

# =========================
# HELPERS
# =========================
def clean_excel(val):
    if isinstance(val, str):
        return re.sub(r"[\x00-\x08\x0B-\x1F]", "", val)
    return val

def extract_text(pdf):
    try:
        with fitz.open(pdf) as doc:
            return "\n".join(p.get_text("text") or "" for p in doc)
    except Exception:
        return None

def split_sentences(text):
    text = re.sub(r"\n+", " ", text)
    return [s for s in sent_tokenize(text) if len(s.strip()) > 10][:MAX_SENTENCES]

def has_evidence(sents, idx):
    words = ["implemented","established","maintained","audit","certified",
             "monitored","trained","reviewed","tested","assessed"]
    lo, hi = max(0, idx-WINDOW), min(len(sents)-1, idx+WINDOW)
    return any(any(w in sents[i].lower() for w in words) for i in range(lo, hi+1))

# =========================
# MODEL
# =========================
print("[⚡] Loading embedding model")
model = SentenceTransformer("all-mpnet-base-v2", device=device)
iso_emb = model.encode(list(iso_descriptions.values()), show_progress_bar=False)
print("[⚡] Model ready")

# =========================
# LOAD EXCEL / RESUME
# =========================
if os.path.exists(EXCEL_PATH):
    df_existing = pd.read_excel(EXCEL_PATH)
    processed = set(df_existing["File"].astype(str))
else:
    df_existing = pd.DataFrame()
    processed = set()

all_pdfs = sorted(f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf"))
pending = [f for f in all_pdfs if f not in processed][:BATCH_SIZE]

print(f"[⚡] Processing {len(pending)} PDFs")

# =========================
# MAIN LOOP
# =========================
rows = []
start = time.time()

for i, pdf in enumerate(tqdm(pending, desc="Scoring PDFs")):
    text = extract_text(os.path.join(PDF_FOLDER, pdf))
    if not text:
        continue

    sents = split_sentences(text)
    if not sents:
        continue

    emb = model.encode(sents, show_progress_bar=False)
    sim = cosine_similarity(emb, iso_emb)

    base, year = os.path.splitext(pdf)[0].split("_",1) if "_" in pdf else (pdf,"")
    row = {"Company": base, "Year": year, "File": pdf}

    total = 0
    for j, k in enumerate(ISO_KEYS):
        idx = int(np.argmax(sim[:,j]))
        score = 0
        if sim[idx,j] >= SIM_MENTION:
            score = 1
            if sim[idx,j] >= SIM_HIGH or has_evidence(sents, idx):
                score = 2
        row[f"{k}__score"] = score
        row[f"{k}__sim"] = round(float(sim[idx,j]),4)
        row[f"{k}__snippet"] = sents[idx][:200]
        row[f"{k}__reason"] = "semantic"
        total += score

    row["Total_Score"] = total
    row["Processed_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row["Status"] = "OK"
    rows.append(row)

# =========================
# SAVE SAFELY
# =========================
df_new = pd.DataFrame(rows).applymap(clean_excel)
df_final = pd.concat([df_existing, df_new], ignore_index=True)

for _ in range(EXCEL_WRITE_RETRIES):
    try:
        df_final.to_excel(EXCEL_PATH, index=False)
        break
    except IllegalCharacterError:
        time.sleep(2)

print("[⚡] SAFE TO CLOSE — RESUME READY")
