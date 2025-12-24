"""
⚡ ISO27001 PATH A2 — AUTHORITATIVE REBUILD ENGINE ⚡
--------------------------------------------------
MODE        : ACCURACY-FIRST
SOURCE      : ORIGINAL PDFs ONLY
BATCH SIZE  : 20
RESUME      : AUTOMATIC
SCHEMA      : HARD-LOCKED
ZERO FILL   : NEVER
"""

# =========================
# BOOT SEQUENCE
# =========================
print("\n[⚡] >>> PATH A2 REBUILD ENGINE INITIALIZING <<<\n")

import os
import re
import fitz
import pandas as pd
import numpy as np
from datetime import datetime
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import nltk
from nltk.tokenize import sent_tokenize

# =========================
# CONFIG — EDIT ONLY IF NEEDED
# =========================
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
INPUT_EXCEL = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection Path A.xlsx"
OUTPUT_EXCEL = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection Path A_REBUILT.xlsx"

BATCH_SIZE = 20

SIM_MENTION = 0.60
SIM_HIGH = 0.72
WINDOW = 1

ISO_DOMAINS = {
    "A.5":  "Information security policies",
    "A.6":  "Organization of information security",
    "A.7":  "Human resource security",
    "A.8":  "Asset management",
    "A.9":  "Access control",
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
ISO_KEYS = list(ISO_DOMAINS.keys())

# =========================
# NLP INIT
# =========================
print("[⚡] Loading NLTK tokenizer")
nltk.download("punkt", quiet=True)

print("[⚡] Loading embedding model (offline)")
model = SentenceTransformer(
    "all-mpnet-base-v2",
    local_files_only=True,
    device="cpu"
)
iso_embeddings = model.encode(list(ISO_DOMAINS.values()), show_progress_bar=False)
print("[⚡] Embedding model ONLINE\n")

# =========================
# HELPERS
# =========================
def extract_text(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        chunks = []
        for page in doc:
            try:
                chunks.append(page.get_text("text") or "")
            except Exception:
                pass
        doc.close()
        text = "\n".join(chunks)
        return text.strip() if len(text.strip()) > 50 else None
    except Exception:
        return None

def split_sentences(txt):
    txt = re.sub(r"\s+", " ", txt)
    return [s for s in sent_tokenize(txt) if len(s.strip()) > 15]

def has_evidence(sents, idx):
    keywords = [
        "implemented", "established", "maintained", "audit",
        "certified", "monitored", "trained", "reviewed",
        "tested", "assessed"
    ]
    lo = max(0, idx - WINDOW)
    hi = min(len(sents) - 1, idx + WINDOW)
    return any(
        any(k in sents[i].lower() for k in keywords)
        for i in range(lo, hi + 1)
    )

# =========================
# LOAD INPUT
# =========================
df = pd.read_excel(INPUT_EXCEL)
print(f"[⚡] Loaded INPUT Excel :: {len(df)} rows")

# =========================
# SCHEMA LOCK (CRITICAL)
# =========================
BASE_COLUMNS = [
    "Year", "Company", "File"
]

FINAL_COLUMNS = (
    BASE_COLUMNS +
    ISO_KEYS +
    ["Total_Score", "Status", "Processed_On"]
)

df = df[BASE_COLUMNS].copy()

for k in ISO_KEYS:
    df[k] = np.nan

df["Total_Score"] = np.nan
df["Status"] = ""
df["Processed_On"] = ""

# =========================
# RESUME LOGIC
# =========================
if os.path.exists(OUTPUT_EXCEL):
    prev = pd.read_excel(OUTPUT_EXCEL)
    done_files = set(prev.loc[prev["Status"] == "OK", "File"])
    df.loc[df["File"].isin(done_files), prev.columns] = prev.set_index("File").loc[
        df["File"], prev.columns
    ].values
    print(f"[⚡] Resume detected :: {len(done_files)} rows already completed")
else:
    done_files = set()

# =========================
# SELECT BATCH
# =========================
remaining = df[~df["File"].isin(done_files)]
batch = remaining.head(BATCH_SIZE)

print(f"[⚡] Remaining rows :: {len(remaining)}")
print(f"[⚡] Processing batch :: {len(batch)} rows\n")

# =========================
# REBUILD LOOP
# =========================
processed = 0

for idx, row in batch.iterrows():
    pdf_name = row["File"]
    pdf_path = os.path.join(PDF_FOLDER, pdf_name)

    print(f"[⚡] Scanning PDF :: {pdf_name}")

    if not os.path.exists(pdf_path):
        df.at[idx, "Status"] = "PDF_MISSING"
        continue

    text = extract_text(pdf_path)
    if not text:
        df.at[idx, "Status"] = "NO_TEXT"
        continue

    sentences = split_sentences(text)
    if not sentences:
        df.at[idx, "Status"] = "NO_SENTENCES"
        continue

    sent_emb = model.encode(sentences, show_progress_bar=False)
    sims = cosine_similarity(sent_emb, iso_embeddings)

    scores = {}
    for j, key in enumerate(ISO_KEYS):
        best_idx = int(np.argmax(sims[:, j]))
        best_sim = sims[best_idx, j]

        score = 0
        if best_sim >= SIM_MENTION:
            score = 1
            if best_sim >= SIM_HIGH or has_evidence(sentences, best_idx):
                score = 2

        scores[key] = score
        df.at[idx, key] = score

    df.at[idx, "Total_Score"] = sum(scores.values())
    df.at[idx, "Status"] = "OK"
    df.at[idx, "Processed_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    processed += 1

# =========================
# SAVE
# =========================
df.to_excel(OUTPUT_EXCEL, index=False)

print("\n[💾] Batch committed safely")
print(f"[⚡] Rows processed this run :: {processed}")
print(f"[⚠️] Rows remaining :: {len(remaining) - processed}")
print("\n[⚡] SAFE TO CLOSE — RE-RUN TO CONTINUE\n")
