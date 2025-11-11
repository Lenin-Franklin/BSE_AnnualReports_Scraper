"""
iso27001_scoring_excel_progress.py
Hybrid semantic + keyword ISO27001 scoring (0/1/2) per domain per PDF.
Outputs: disclosure_scoring_detailed.xlsx
"""

import os
import re
import warnings
import logging
import fitz  # PyMuPDF
import pandas as pd
import numpy as np
from tqdm import tqdm
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import nltk
from nltk.tokenize import sent_tokenize

# -------------------------
# CLEANUP / QUIET MODE
# -------------------------
warnings.filterwarnings("ignore")
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
logging.getLogger("transformers").setLevel(logging.ERROR)
logging.getLogger("sentence_transformers").setLevel(logging.ERROR)
logging.getLogger("huggingface_hub").setLevel(logging.ERROR)

# -------------------------
# NLTK SETUP (with punkt_tab fix)
# -------------------------
def ensure_nltk_data():
    try:
        nltk.data.find("tokenizers/punkt")
    except LookupError:
        nltk.download("punkt")
    try:
        nltk.data.find("tokenizers/punkt_tab")
    except LookupError:
        nltk.download("punkt_tab")

ensure_nltk_data()

# -------------------------
# CONFIG / PATHS
# -------------------------
PDF_FOLDER = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\Company_PDF"

OUTPUT_XLSX = r"C:\Users\lenin\OneDrive\Desktop\BSE_Scraper\disclosure_scoring_detailed.xlsx"

EMBEDDING_MODEL = "all-mpnet-base-v2"
SIMILARITY_THRESHOLD_MENTION = 0.60
SIMILARITY_THRESHOLD_HIGH = 0.72
EVIDENCE_PROXIMITY_WINDOW = 1

EVIDENCE_WORDS = [
    "implemented", "established", "maintained", "program", "procedure", "policy",
    "audit", "audited", "reviewed", "tested", "assessed", "certified", "enforced",
    "monitored", "trained", "performed", "investigated", "reported", "resolved"
]

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

# -------------------------
# Helper functions
# -------------------------
def extract_text_from_pdf(pdf_path):
    text_chunks = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text_chunks.append(page.get_text("text"))
    return "\n".join(text_chunks)

def split_to_sentences(text):
    normalized = re.sub(r"\n+", " ", text)
    sentences = sent_tokenize(normalized)
    return [s.strip() for s in sentences if len(s.strip()) > 10]

def find_evidence_in_window(sentences, idx, window, evidence_words):
    lo = max(0, idx - window)
    hi = min(len(sentences) - 1, idx + window)
    for i in range(lo, hi + 1):
        s = sentences[i].lower()
        for ev in evidence_words:
            if re.search(r"\b" + re.escape(ev) + r"\b", s):
                return True, i, s
    return False, None, None

# -------------------------
# Main processing
# -------------------------
print("Loading embedding model:", EMBEDDING_MODEL)
model = SentenceTransformer(EMBEDDING_MODEL)

iso_keys = list(iso_descriptions.keys())
iso_texts = [iso_descriptions[k] for k in iso_keys]
iso_embeddings = model.encode(iso_texts, show_progress_bar=False)

pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
if not pdf_files:
    print("⚠️ No PDF files found in:", PDF_FOLDER)
    exit()

rows = []

# Progress bar for all PDFs
for pdf_file in tqdm(pdf_files, desc="Processing PDFs", unit="PDF"):
    pdf_path = os.path.join(PDF_FOLDER, pdf_file)
    base = os.path.splitext(pdf_file)[0]
    try:
        company, year = base.split("_", 1)
    except ValueError:
        company = base
        year = ""

    full_text = extract_text_from_pdf(pdf_path)
    sentences = split_to_sentences(full_text)
    if not sentences:
        print(f"\n⚠️ Skipping {pdf_file}: no readable text found.")
        continue

    sentence_embeddings = model.encode(sentences, show_progress_bar=False)
    sim = cosine_similarity(sentence_embeddings, iso_embeddings)

    doc_row = {"Company": company, "Year": year, "File": pdf_file}
    details = {}

    for j, domain in enumerate(iso_keys):
        best_idx = int(np.argmax(sim[:, j]))
        best_sim = float(sim[best_idx, j])
        best_sentence = sentences[best_idx]

        score, reason, method = 0, "", "embedding"

        if best_sim >= SIMILARITY_THRESHOLD_MENTION:
            score = 1
            reason = f"semantic_match (sim={best_sim:.3f})"
            ev_found, _, _ = find_evidence_in_window(sentences, best_idx, EVIDENCE_PROXIMITY_WINDOW, EVIDENCE_WORDS)
            if ev_found or best_sim >= SIMILARITY_THRESHOLD_HIGH:
                score = 2
                reason = f"demonstrated (sim={best_sim:.3f})" if ev_found else f"strong_semantic (sim={best_sim:.3f})"
        else:
            core_keywords = iso_core_keywords.get(domain, [])
            found_keyword = any(re.search(r"\b" + re.escape(kw.lower()) + r"\b", full_text.lower()) for kw in core_keywords)
            if found_keyword:
                score, reason, method = 1, "keyword_found_fallback", "keyword"
                if any(re.search(r"\b" + re.escape(ev) + r"\b", full_text.lower()) for ev in EVIDENCE_WORDS):
                    score, reason = 2, "keyword_with_evidence"
            else:
                score, reason = 0, "no_match"

        doc_row[domain] = score
        details[domain] = {
            "score": score,
            "best_sim": best_sim,
            "best_sentence": best_sentence,
            "reason": reason,
            "method": method
        }

    for d in iso_keys:
        doc_row[f"{d}__score"] = details[d]["score"]
        doc_row[f"{d}__sim"] = round(details[d]["best_sim"], 4)
        snippet = details[d]["best_sentence"]
        doc_row[f"{d}__snippet"] = snippet[:200] + "..." if len(snippet) > 200 else snippet
        doc_row[f"{d}__reason"] = details[d]["reason"]

    rows.append(doc_row)

# -------------------------
# Save results to Excel
# -------------------------
df = pd.DataFrame(rows)
domain_cols = [k for k in iso_keys]
df["Total_Score"] = df[[c for c in domain_cols]].sum(axis=1)

df.to_excel(OUTPUT_XLSX, index=False)
print("\n✅ All PDFs processed successfully.")
print("📁 Results saved to:", OUTPUT_XLSX)
