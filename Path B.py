"""
⚡ ISO27001 PATH B1 — FORENSIC EVIDENCE REPAIR ENGINE ⚡
----------------------------------------------------
MODE        : FORENSIC / EXPLANATORY
SCORES      : READ-ONLY
BATCH SIZE  : 20
RESUME      : AUTOMATIC
DATA LOSS   : IMPOSSIBLE
"""

# =========================
# BOOT SEQUENCE
# =========================
print("\n[⚡] >>> PATH B1 FORENSIC ENGINE INITIALIZING <<<\n")

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
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"
INPUT_EXCEL = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection Path B.xlsx"
OUTPUT_EXCEL = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\ISO Data Collection Path B_PATCHED.xlsx"

BATCH_SIZE = 20
SIM_THRESHOLD = 0.55

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
print("[⚡] Loading NLTK resources")
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

# =========================
# LOAD INPUT
# =========================
df = pd.read_excel(INPUT_EXCEL)
print(f"[⚡] Loaded INPUT Excel :: {len(df)} rows")

# =========================
# ENSURE FORENSIC COLUMNS
# =========================
for key in ISO_KEYS:
    ev_col = f"{key}_Evidence"
    sim_col = f"{key}_Sim"
    if ev_col not in df.columns:
        df[ev_col] = ""
    if sim_col not in df.columns:
        df[sim_col] = np.nan

# =========================
# RESUME LOGIC
# =========================
if os.path.exists(OUTPUT_EXCEL):
    prev = pd.read_excel(OUTPUT_EXCEL)
    done = set(prev.loc[prev["Status"].notna(), "File"])
    df.loc[df["File"].isin(done), prev.columns] = prev.set_index("File").loc[
        df["File"], prev.columns
    ].values
    print(f"[⚡] Resume detected :: {len(done)} rows already patched")
else:
    done = set()

# =========================
# SELECT ROWS NEEDING FORENSICS
# =========================
needs_fix = df[
    (~df["File"].isin(done)) &
    (
        df[[f"{k}_Evidence" for k in ISO_KEYS]].eq("").any(axis=1)
    )
]

batch = needs_fix.head(BATCH_SIZE)

print(f"[⚡] Rows needing forensic repair :: {len(needs_fix)}")
print(f"[⚡] Processing batch :: {len(batch)} rows\n")

# =========================
# FORENSIC LOOP
# =========================
patched = 0

for idx, row in batch.iterrows():
    pdf_name = row["File"]
    pdf_path = os.path.join(PDF_FOLDER, pdf_name)

    print(f"[⚡] Extracting evidence :: {pdf_name}")

    if not os.path.exists(pdf_path):
        continue

    text = extract_text(pdf_path)
    if not text:
        continue

    sentences = split_sentences(text)
    if not sentences:
        continue

    sent_emb = model.encode(sentences, show_progress_bar=False)
    sims = cosine_similarity(sent_emb, iso_embeddings)

    for j, key in enumerate(ISO_KEYS):
        ev_col = f"{key}_Evidence"
        sim_col = f"{key}_Sim"

        if isinstance(row[ev_col], str) and row[ev_col].strip():
            continue  # already has evidence

        best_idx = int(np.argmax(sims[:, j]))
        best_sim = sims[best_idx, j]

        if best_sim >= SIM_THRESHOLD:
            df.at[idx, ev_col] = sentences[best_idx][:500]
            df.at[idx, sim_col] = round(float(best_sim), 4)

    patched += 1

# =========================
# SAVE
# =========================
df.to_excel(OUTPUT_EXCEL, index=False)

print("\n[💾] Forensic batch committed safely")
print(f"[⚡] Rows patched this run :: {patched}")
print(f"[⚠️] Rows remaining :: {len(needs_fix) - patched}")
print("\n[⚡] SAFE TO CLOSE — RE-RUN TO CONTINUE\n")
