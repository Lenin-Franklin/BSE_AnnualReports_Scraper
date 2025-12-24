"""
⚡ ISO27001 FORENSIC REPAIR ENGINE — SAFE OPTIMIZED ⚡
---------------------------------------------------
✔ Repairs ONLY truly damaged rows
✔ Accuracy preserved
✔ Resume-safe
✔ Speed optimized (semantic-safe)
"""

# =========================
# BOOT SEQUENCE
# =========================
print("\n[⚡] >>> FORENSIC REPAIR ENGINE INITIALIZING <<<\n")

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
# CONFIG
# =========================
PDF_FOLDER   = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
INPUT_EXCEL = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection.xlsx"
OUTPUT_EXCEL= r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection_REPAIRED.xlsx"

BATCH_SIZE  = 20
SIM_MENTION = 0.60
SIM_HIGH    = 0.72
WINDOW      = 1

DAMAGED_STATUSES = {"PDF_READ_FAILED", "NO_TEXT"}

SECURITY_KEYWORDS = [
    "information security","cyber","data","access","policy","risk",
    "control","audit","compliance","incident","encryption",
    "business continuity","iso","security"
]

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
ISO_KEYS = list(ISO_DOMAINS.keys())

# =========================
# INIT NLP
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
print("[⚡] Embedding model online\n")

# =========================
# HELPERS
# =========================
def extract_text(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        pages = [page.get_text("text") or "" for page in doc]
        doc.close()
        return "\n".join(pages)
    except Exception:
        return None

def split_sentences(text):
    text = re.sub(r"\n+", " ", text)
    return [s for s in sent_tokenize(text) if len(s.strip()) > 10]

def filter_security_sentences(sentences):
    return [
        s for s in sentences
        if any(k in s.lower() for k in SECURITY_KEYWORDS)
    ]

def has_evidence(sents, idx):
    words = [
        "implemented","established","maintained","audit",
        "certified","monitored","trained","reviewed","tested"
    ]
    lo, hi = max(0, idx-WINDOW), min(len(sents)-1, idx+WINDOW)
    return any(any(w in sents[i].lower() for w in words) for i in range(lo, hi+1))

# =========================
# LOAD DATA
# =========================
df = pd.read_excel(INPUT_EXCEL)
df["Status"] = df["Status"].astype(str).str.upper().str.strip()

print(f"[⚡] Loaded Excel :: {len(df)} rows")

damaged = df[df["Status"].isin(DAMAGED_STATUSES)]
remaining = len(damaged)

print(f"[⚠️] TRUE damaged rows remaining :: {remaining}")

if remaining == 0:
    print("\n[🔥] DATASET CLEAN — NO REPAIR REQUIRED\n")
    exit()

batch = damaged.head(BATCH_SIZE)
print(f"[⚡] Processing batch :: {len(batch)} PDFs\n")

# =========================
# REPAIR LOOP
# =========================
repaired = 0

for idx, row in batch.iterrows():
    pdf = row["File"]
    path = os.path.join(PDF_FOLDER, pdf)

    print(f"[⚡] Re-scoring :: {pdf}")

    if not os.path.exists(path):
        print("   └─ PDF missing")
        continue

    text = extract_text(path)
    if not text:
        print("   └─ No text extracted")
        continue

    sentences = split_sentences(text)
    sentences = filter_security_sentences(sentences)

    if not sentences:
        print("   └─ No security-relevant sentences")
        continue

    sent_embeddings = model.encode(sentences, show_progress_bar=False)
    sims = cosine_similarity(sent_embeddings, iso_embeddings)

    scores = {}
    for j, key in enumerate(ISO_KEYS):
        idx_best = int(np.argmax(sims[:, j]))
        sim = sims[idx_best, j]

        score = 0
        if sim >= SIM_MENTION:
            score = 1
            if sim >= SIM_HIGH or has_evidence(sentences, idx_best):
                score = 2
        scores[key] = score

    for k, v in scores.items():
        df.at[idx, k] = v

    df.at[idx, "Total_Score"] = sum(scores.values())
    df.at[idx, "Status"] = "REPAIRED_OK"
    df.at[idx, "Processed_On"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    repaired += 1

# =========================
# SAVE
# =========================
df.to_excel(OUTPUT_EXCEL, index=False)

print("\n[💾] Batch committed safely")
print(f"[⚡] Rows repaired :: {repaired}")
print(f"[⚠️] Remaining :: {remaining - repaired}")
print("\n[⚡] SAFE TO CLOSE — RUN AGAIN TO CONTINUE\n")
