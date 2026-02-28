import io
import re
from typing import Dict

import streamlit as st
import pdfplumber
from docx import Document
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# ---------- TEXT EXTRACTION HELPERS ----------

def extract_text_from_pdf(file_bytes: bytes) -> str:
    text = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text.append(page.extract_text() or "")
    return "\n".join(text)


def extract_text_from_docx(file_bytes: bytes) -> str:
    file_stream = io.BytesIO(file_bytes)
    doc = Document(file_stream)
    text = []
    for para in doc.paragraphs:
        text.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text for cell in row.cells]
            text.append(" | ".join(cells))
    return "\n".join(text)


# ---------- FIELD PARSING FROM RAW TEXT ----------

def parse_issue_briefing_pdf_from_file(file_bytes: bytes) -> Dict[str, str]:
    """
    Parse IBF PDF by extracting tables cell-by-cell using pdfplumber,
    matching the actual table layout of the document.
    """
    fields = {
        "Title": "",
        "Issue ID": "",
        "Description": "",
        "Issue Impact": "",
        "Issue Root Cause": "",
    }

    kv_map = {}
    paragraph_text_parts = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            # ---- Extract tables: walk rows and pair cells as key→value ----
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Clean cells: strip whitespace, replace None
                    cells = [
                        (c.strip().replace("\n", " ") if c else "")
                        for c in row
                    ]
                    # Walk cells in pairs
                    i = 0
                    while i < len(cells) - 1:
                        key = cells[i].rstrip(":").strip()
                        value = cells[i + 1].strip()
                        if key:
                            kv_map[key] = value
                        i += 2

            # ---- Also collect plain text for paragraph-style fields ----
            raw = page.extract_text()
            if raw:
                paragraph_text_parts.append(raw)

    # ---- Map key-value pairs ----
    fields["Title"] = kv_map.get("Title", "")
    fields["Issue ID"] = kv_map.get("Issue ID", kv_map.get("Issue\nID", ""))

    # ---- Paragraph fields from plain text ----
    combined = "\n".join(paragraph_text_parts)

    desc_match = re.search(
        r"Description\s+(.+?)Issue Impact", combined, flags=re.DOTALL
    )
    if desc_match:
        fields["Description"] = " ".join(desc_match.group(1).split())

    impact_match = re.search(
        r"Issue Impact\s+(.+?)Issue Root Cause", combined, flags=re.DOTALL
    )
    if impact_match:
        fields["Issue Impact"] = " ".join(impact_match.group(1).split())

    root_match = re.search(
        r"Issue Root Cause\s+(.+?)Overall Issue Rating", combined, flags=re.DOTALL
    )
    if root_match:
        fields["Issue Root Cause"] = " ".join(root_match.group(1).split())

    return fields



def parse_icp_docx_from_file(file_bytes: bytes) -> Dict[str, str]:
    """
    Parse ICP DOCX by walking tables cell-by-cell.
    This correctly handles fields in separate table cells.
    """
    fields = {
        "Title": "",
        "Issue ID": "",
        "Description": "",
        "Issue Impact": "",
        "Issue Root Cause": "",
    }

    doc = Document(io.BytesIO(file_bytes))

    # ---- Step 1: Extract all table cells as a flat key→value map ----
    # Many tables are key-value pairs side by side in the same row.
    # e.g. | Issue Title: | Investment Banking... | Source System Issue Reference: | ISSUE-00077693 |
    kv_map = {}
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells]
            # Walk cells in pairs (label, value, label, value ...)
            i = 0
            while i < len(cells) - 1:
                key = cells[i].rstrip(":")
                value = cells[i + 1]
                if key and value:
                    kv_map[key.strip()] = value.strip()
                i += 2

    # ---- Step 2: Map known keys to fields ----
    fields["Title"] = kv_map.get("Issue Title", "")
    fields["Issue ID"] = kv_map.get("Source System Issue Reference", "")

    # ---- Step 3: For paragraph-style fields, scan paragraphs ----
    # These are in merged single-cell rows (long text blocks)
    full_text = "\n".join([p.text for p in doc.paragraphs])

    # Also collect all table cell text for paragraph blocks
    table_text_parts = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                table_text_parts.append(cell.text)
    table_text = "\n".join(table_text_parts)
    combined = full_text + "\n" + table_text

    desc_match = re.search(
        r"Issue Description:\s*(.+?)Issue Root Cause:", combined, flags=re.DOTALL
    )
    if desc_match:
        fields["Description"] = " ".join(desc_match.group(1).split())

    root_match = re.search(
        r"Issue Root Cause:\s*(.+?)Issue Impact:", combined, flags=re.DOTALL
    )
    if root_match:
        fields["Issue Root Cause"] = " ".join(root_match.group(1).split())

    impact_match = re.search(
        r"Issue Impact:\s*(.+?)Background Context:", combined, flags=re.DOTALL
    )
    if impact_match:
        fields["Issue Impact"] = " ".join(impact_match.group(1).split())
    else:
        impact_match2 = re.search(
            r"Issue Impact:\s*(.+?)Section C:", combined, flags=re.DOTALL
        )
        if impact_match2:
            fields["Issue Impact"] = " ".join(impact_match2.group(1).split())

    return fields



# ---------- SIMILARITY COMPUTATION ----------

def text_similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    vectorizer = TfidfVectorizer().fit([a, b])
    tfidf = vectorizer.transform([a, b])
    sim = cosine_similarity(tfidf[0:1], tfidf[1:2])[0][0]
    return float(sim)


def compute_similarity(f1: Dict[str, str], f2: Dict[str, str]) -> Dict[str, float]:
    scores = {}
    scores["Issue ID"] = 1.0 if f1.get("Issue ID") == f2.get("Issue ID") else 0.0
    scores["Title"] = text_similarity(f1.get("Title", ""), f2.get("Title", ""))
    scores["Description"] = text_similarity(
        f1.get("Description", ""), f2.get("Description", "")
    )
    scores["Issue Root Cause"] = text_similarity(
        f1.get("Issue Root Cause", ""), f2.get("Issue Root Cause", "")
    )
    scores["Issue Impact"] = text_similarity(
        f1.get("Issue Impact", ""), f2.get("Issue Impact", "")
    )
    scores["Overall"] = sum(scores.values()) / len(scores)
    return scores


def to_match_label(score: float, threshold: float) -> str:
    return "Match" if score >= threshold else "Mismatch"


# ---------- STREAMLIT APP ----------

def main():
    st.title("IBF vs ICP Comparison")

    st.write(
        "Upload the Issue Briefing Form (PDF) and Issue Closure Pack ICP (DOCX) "
        "to view key fields, similarity scores, and match/mismatch flags."
    )

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader(
            "Upload Issue Briefing Form (PDF)",
            type=["pdf"],
            key="pdf",
        )
    with col2:
        docx_file = st.file_uploader(
            "Upload Issue Closure Pack ICP (DOCX)",
            type=["docx"],
            key="docx",
        )

    threshold = st.slider(
        "Match threshold (for similarity scores)",
        min_value=0.0,
        max_value=1.0,
        value=0.8,
        step=0.05,
    )

    if pdf_file and docx_file and st.button("Compare"):
        pdf_bytes = pdf_file.read()
        docx_bytes = docx_file.read()

        pdf_text = extract_text_from_pdf(pdf_bytes)
        docx_text = extract_text_from_docx(docx_bytes)

        # Optional: debug view
        with st.expander("Raw PDF text (first 2000 chars)"):
            st.text(pdf_text[:2000])

        with st.expander("Raw DOCX text (first 2000 chars)"):
            st.text(docx_text[:2000])

        # Parse structured fields
        f1 = parse_issue_briefing_pdf_from_file(pdf_bytes)
        f2 = parse_icp_docx_from_file(docx_bytes)

        st.subheader("Parsed fields from IBF (PDF)")
        st.json(f1)

        st.subheader("Parsed fields from ICP (DOCX)")
        st.json(f2)

        # ---------- Key-value table ----------
        rows_kv = [
            {
                "Attribute": "Title",
                "IBF (PDF)": f1.get("Title", ""),
                "ICP (DOCX)": f2.get("Title", ""),
            },
            {
                "Attribute": "Issue ID",
                "IBF (PDF)": f1.get("Issue ID", ""),
                "ICP (DOCX)": f2.get("Issue ID", ""),
            },
            {
                "Attribute": "Description",
                "IBF (PDF)": f1.get("Description", ""),
                "ICP (DOCX)": f2.get("Description", ""),
            },
            {
                "Attribute": "Issue Root Cause",
                "IBF (PDF)": f1.get("Issue Root Cause", ""),
                "ICP (DOCX)": f2.get("Issue Root Cause", ""),
            },
            {
                "Attribute": "Issue Impact",
                "IBF (PDF)": f1.get("Issue Impact", ""),
                "ICP (DOCX)": f2.get("Issue Impact", ""),
            },
        ]
        df_kv = pd.DataFrame(rows_kv)

        st.subheader("Key fields from IBF and ICP")
        st.table(df_kv)

        # CSV download for key-value table
        csv_kv = df_kv.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download key fields as CSV",
            data=csv_kv,
            file_name="ibf_icp_key_fields.csv",
            mime="text/csv",
        )

        # ---------- Similarity scores ----------
        scores = compute_similarity(f1, f2)

        sim_rows = []
        for field in ["Issue ID", "Title", "Description", "Issue Root Cause", "Issue Impact"]:
            score = scores[field]
            label = to_match_label(
                score,
                threshold if field != "Issue ID" else 1.0,
            )
            sim_rows.append(
                {
                    "Field": field,
                    "Similarity Score": round(score, 3),
                    "Result": label,
                }
            )
        sim_rows.append(
            {
                "Field": "Overall",
                "Similarity Score": round(scores["Overall"], 3),
                "Result": to_match_label(scores["Overall"], threshold),
            }
        )
        df_sim = pd.DataFrame(sim_rows)

        st.subheader("Similarity results")
        st.table(df_sim)

        csv_sim = df_sim.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download similarity scores as CSV",
            data=csv_sim,
            file_name="ibf_icp_similarity_scores.csv",
            mime="text/csv",
        )


if __name__ == "__main__":
    main()
