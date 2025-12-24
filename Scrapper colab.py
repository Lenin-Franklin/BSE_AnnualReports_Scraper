# ============================================================
# ISO27001 OG SCRAPPER — FINAL CPU-SAFE COLAB VERSION
# ============================================================

import os, re, time, unicodedata
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
# CONFIG
# =========================
PDF_FOLDER = "/content/Company_PDF"
QUARANTINE_DIR = os.path.join(PDF_FOLDER, "_quarantine")
EXCEL_PATH = "/content/ISO Data Collection.xlsx"
LOG_FILE = "/content/iso_log.txt"

BATCH_SIZE = 100          # CPU-safe
MAX_SENTENCES = 1500
EMBED_BATCH = 32          # REQUIRED
SIM_MENTION = 0.60
SIM_HIGH = 0.72
WINDOW = 1
WRITE_RETRIES = 3

os.makedirs(QUARANTINE_DIR, exist_ok=True)
torch.set_grad_enabled(False)

# =========================
# NLTK
# =========================
nltk.download("punkt", quiet=True)

# =========================
# SILENCE MUPDF WIDGET WARNINGS
# =========================
try:
    fitz.TOOLS.set_annot_appearance(False)
except Exception:
    pass

# =========================
# LOGGING
# =========================
def log(msg):
    ts = datetime.now().strftime("[%H:%M:%S]")
    print(f"{ts} {msg}")
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(f"{ts} {msg}\n")

# =========================
# EXCEL SAFE STRING
# =========================
def excel_safe(val):
    if not isinstance(val, str):
        return val
    val = unicodedata.normalize("NFKD", val)
    return "".join(c for c in val if unicodedata.category(c)[0] != "C")

# =========================
# QUARANTINE
# =========================
def quarantine(pdf, reason):
    try:
        os.rename(
            os.path.join(PDF_FOLDER, pdf),
            os.path.join(QUARANTINE_DIR, pdf)
        )
    except Exception:
        pass
    log(f"[QUARANTINED] {pdf} :: {reason}")

# =========================
# HELPERS
# =========================
def extract_text(path):
    try:
        with fitz.open(path) as doc:
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

ISO_KEYS = list(iso_descriptions.keys())

# =========================
# MODEL
# =========================
log("Loading embedding model (CPU)")
model = SentenceTransformer("all-mpnet-base-v2", device="cpu")
iso_embeddings = model.encode(list(iso_descriptions.values()), show_progress_bar=False)
log("Model ready")

# =========================
# LOAD EXCEL / RESUME
# =========================
if os.path.exists(EXCEL_PATH):
    df_existing = pd.read_excel(EXCEL_PATH)
    processed = set(df_existing["File"].astype(str))
else:
    df_existing = pd.DataFrame()
    processed = set()

pdfs = sorted(f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf"))
pending = [p for p in pdfs if p not in processed][:BATCH_SIZE]

log(f"Processing {len(pending)} PDFs")

# =========================
# MAIN LOOP
# =========================
rows = []
start = time.time()

for i, pdf in enumerate(tqdm(pending, desc="Scoring PDFs")):
    eta = int(((time.time()-start)/(i+1))*(len(pending)-(i+1))) if i else 0
    log(f"{i+1}/{len(pending)} :: {pdf} | ETA {eta}s")

    try:
        text = extract_text(os.path.join(PDF_FOLDER, pdf))
        if not text:
            quarantine(pdf, "NO_TEXT")
            continue

        sentences = split_sentences(text)
        if not sentences:
            quarantine(pdf, "NO_SENTENCES")
            continue

        sent_emb = model.encode(sentences, batch_size=EMBED_BATCH, show_progress_bar=False)
        sim = cosine_similarity(sent_emb, iso_embeddings)

        base, year = os.path.splitext(pdf)[0].split("_",1) if "_" in pdf else (pdf,"")
        row = {"Company": base, "Year": year, "File": pdf}
        total = 0

        for j, key in enumerate(ISO_KEYS):
            idx = int(np.argmax(sim[:,j]))
            score = 0
            if sim[idx,j] >= SIM_MENTION:
                score = 1
                if sim[idx,j] >= SIM_HIGH or has_evidence(sentences, idx):
                    score = 2
            row[f"{key}__score"] = score
            row[f"{key}__sim"] = round(float(sim[idx,j]),4)
            row[f"{key}__snippet"] = sentences[idx][:200]
            row[f"{key}__reason"] = "semantic"
            total += score

        row["Total_Score"] = total
        row["Processed_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        row["Status"] = "OK"
        rows.append(row)

    except Exception as e:
        quarantine(pdf, f"RUNTIME_FAIL: {e}")

# =========================
# SAVE
# =========================
df_new = pd.DataFrame(rows).applymap(excel_safe)
df_final = pd.concat([df_existing, df_new], ignore_index=True)

for attempt in range(WRITE_RETRIES):
    try:
        df_final.to_excel(EXCEL_PATH, index=False)
        log("Excel saved successfully")
        break
    except (IllegalCharacterError, PermissionError) as e:
        log(f"[WARN] Excel write failed ({attempt+1}) :: {e}")
        time.sleep(2)
else:
    raise RuntimeError("Excel write failed after retries")

log("[⚡] SAFE TO CLOSE — RESUME READY")
