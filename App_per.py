import io
import re
from typing import Dict

import streamlit as st
import pypdf  # Changed from pdfplumber to pypdf
from docx import Document
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ═══════════════════════════════════════════════════════
# PDF PARSING — Using pypdf (PyPDF2 successor)
# ═══════════════════════════════════════════════════════

def extract_text_from_pdf(file_bytes: bytes) -> str:
    """
    Extract all text from PDF using pypdf.
    pypdf produces cleaner text output than pdfplumber for this use case.
    """
    text = ""
    pdf_reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    
    return text


def parse_issue_briefing_pdf(text: str) -> Dict[str, str]:
    """
    Extract issue fields from PDF text using regex patterns.
    This approach works reliably with pypdf's text extraction.
    """
    fields = {
        "Title": "",
        "Issue ID": "",
        "Description": "",
        "Issue Impact": "",
        "Issue Root Cause": "",
    }

    # Extract Issue ID (Format: ISSUE-00077693)
    issue_id_match = re.search(r'ISSUE-\d+', text, re.IGNORECASE)
    if issue_id_match:
        fields["Issue ID"] = issue_id_match.group(0).strip()


    title_match = re.search(r'Title\s+(.+?)\s+Issue\s*ID', text, re.IGNORECASE)
    if title_match:
        fields["Title"] = title_match.group(1).strip()

    # Extract Description (Between "Description" and "Issue Impact")
    desc_match = re.search(
        r'Description\s*\n\s*(.+?)(?=\s*Issue Impact)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    if desc_match:
        fields["Description"] = desc_match.group(1).strip()

    # Extract Issue Impact (Between "Issue Impact" and "Issue Root Cause")
    impact_match = re.search(
        r'Issue Impact\s*\n\s*(.+?)(?=\s*Issue Root Cause)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    if impact_match:
        fields["Issue Impact"] = impact_match.group(1).strip()

    # Extract Issue Root Cause (Between "Issue Root Cause" and "Overall Issue Rating")
    root_cause_match = re.search(
        r'Issue Root Cause\s*\n\s*(.+?)(?=\s*Overall Issue Rating)',
        text,
        re.IGNORECASE | re.DOTALL
    )
    if root_cause_match:
        fields["Issue Root Cause"] = root_cause_match.group(1).strip()

    return fields


# ═══════════════════════════════════════════════════════
# DOCX PARSING (UNCHANGED, WORKING)
# ═══════════════════════════════════════════════════════

def parse_icp_docx_from_file(file_bytes: bytes) -> Dict[str, str]:
    fields = {
        "Title": "",
        "Issue ID": "",
        "Description": "",
        "Issue Impact": "",
        "Issue Root Cause": "",
    }

    doc = Document(io.BytesIO(file_bytes))
    kv_map = {}

    for table in doc.tables:
        for row in table.rows:
            cells = [" ".join(cell.text.split()) for cell in row.cells]
            i = 0
            while i < len(cells) - 1:
                key = cells[i].rstrip(":").strip()
                value = cells[i + 1].strip()
                if key:
                    kv_map[key] = value
                i += 2

    normalized_kv = {" ".join(k.split()).lower(): v for k, v in kv_map.items()}
    fields["Title"] = normalized_kv.get("issue title", "")
    fields["Issue ID"] = normalized_kv.get("source system issue reference", "")

    para_text = "\n".join(p.text for p in doc.paragraphs)
    table_text = "\n".join(
        cell.text for table in doc.tables
        for row in table.rows
        for cell in row.cells
    )
    combined = " ".join((para_text + "\n" + table_text).split())

    m = re.search(r"Issue Description:\s*(.+?)\s*Issue Root Cause:", combined)
    if m:
        fields["Description"] = m.group(1).strip()

    m = re.search(r"Issue Root Cause:\s*(.+?)\s*Issue Impact:", combined)
    if m:
        fields["Issue Root Cause"] = m.group(1).strip()

    m = re.search(r"Issue Impact:\s*(.+?)\s*Background Context:", combined)
    if m:
        fields["Issue Impact"] = m.group(1).strip()
    else:
        m = re.search(r"Issue Impact:\s*(.+?)\s*Section C:", combined)
        if m:
            fields["Issue Impact"] = m.group(1).strip()

    return fields


# ═══════════════════════════════════════════════════════
# SIMILARITY
# ═══════════════════════════════════════════════════════

def text_similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    vectorizer = TfidfVectorizer().fit([a, b])
    tfidf = vectorizer.transform([a, b])
    return float(cosine_similarity(tfidf[0:1], tfidf[1:2])[0][0])


def compute_similarity(f1: Dict[str, str], f2: Dict[str, str]) -> Dict[str, float]:
    scores = {
        "Issue ID":         1.0 if f1.get("Issue ID") == f2.get("Issue ID") else 0.0,
        "Title":            text_similarity(f1.get("Title", ""),            f2.get("Title", "")),
        "Description":      text_similarity(f1.get("Description", ""),      f2.get("Description", "")),
        "Issue Root Cause": text_similarity(f1.get("Issue Root Cause", ""), f2.get("Issue Root Cause", "")),
        "Issue Impact":     text_similarity(f1.get("Issue Impact", ""),     f2.get("Issue Impact", "")),
    }
    scores["Overall"] = sum(scores.values()) / len(scores)
    return scores


def match_label(score: float, threshold: float) -> str:
    return "✅ Match" if score >= threshold else "❌ Mismatch"


# ═══════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="IBF vs ICP Comparison", layout="wide")
    st.title("📄 IBF vs ICP Comparison")
    st.write("Upload the Issue Briefing Form (PDF) and Issue Closure Pack (DOCX) to compare key fields.")

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader("Upload Issue Briefing Form (PDF)", type=["pdf"])
    with col2:
        docx_file = st.file_uploader("Upload Issue Closure Pack ICP (DOCX)", type=["docx"])

    threshold = st.slider("Match threshold", 0.0, 1.0, 0.8, 0.05)

    if pdf_file and docx_file and st.button("🔍 Compare"):
        pdf_bytes = pdf_file.read()
        docx_bytes = docx_file.read()

        # Parse using pypdf approach
        pdf_text = extract_text_from_pdf(pdf_bytes)
        f1 = parse_issue_briefing_pdf(pdf_text)
        f2 = parse_icp_docx_from_file(docx_bytes)

        with st.expander("🔍 Raw PDF text (first 3000 chars)"):
            st.text(pdf_text[:3000])
        
        with st.expander("🔍 Parsed fields — IBF (PDF)"):
            st.json(f1)

        with st.expander("🔍 Parsed fields — ICP (DOCX)"):
            st.json(f2)

        # Key fields comparison
        st.subheader("📋 Key Fields Comparison")
        df_kv = pd.DataFrame([
            {"Attribute": "Title",            "IBF (PDF)": f1.get("Title", ""),            "ICP (DOCX)": f2.get("Title", "")},
            {"Attribute": "Issue ID",         "IBF (PDF)": f1.get("Issue ID", ""),         "ICP (DOCX)": f2.get("Issue ID", "")},
            {"Attribute": "Description",      "IBF (PDF)": f1.get("Description", ""),      "ICP (DOCX)": f2.get("Description", "")},
            {"Attribute": "Issue Root Cause", "IBF (PDF)": f1.get("Issue Root Cause", ""), "ICP (DOCX)": f2.get("Issue Root Cause", "")},
            {"Attribute": "Issue Impact",     "IBF (PDF)": f1.get("Issue Impact", ""),     "ICP (DOCX)": f2.get("Issue Impact", "")},
        ])
        st.dataframe(df_kv, use_container_width=True)
        st.download_button(
            "⬇️ Download key fields CSV",
            df_kv.to_csv(index=False).encode("utf-8"),
            "ibf_icp_key_fields.csv", "text/csv"
        )

        # Similarity scores
        st.subheader("📊 Similarity Scores")
        scores = compute_similarity(f1, f2)
        sim_rows = []
        for field in ["Issue ID", "Title", "Description", "Issue Root Cause", "Issue Impact"]:
            score = scores[field]
            sim_rows.append({
                "Field": field,
                "Similarity Score": round(score, 3),
                "Result": match_label(score, 1.0 if field == "Issue ID" else threshold),
            })
        sim_rows.append({
            "Field": "Overall",
            "Similarity Score": round(scores["Overall"], 3),
            "Result": match_label(scores["Overall"], threshold),
        })
        df_sim = pd.DataFrame(sim_rows)
        st.dataframe(df_sim, use_container_width=True)
        st.download_button(
            "⬇️ Download similarity scores CSV",
            df_sim.to_csv(index=False).encode("utf-8"),
            "ibf_icp_similarity_scores.csv", "text/csv"
        )


if __name__ == "__main__":
    main()
