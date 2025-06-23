import streamlit as st
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from sentence_transformers import SentenceTransformer, util
import re

model = SentenceTransformer('all-MiniLM-L6-v2')

def extract_text_from_docx(file):
    doc = Document(file)
    return " ".join(para.text for para in doc.paragraphs)

def extract_text_from_pdf(file):
    try:
        reader = PdfReader(file)
        text = []
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text.append(extracted)
        return " ".join(text) if text else "[EMPTY]"
    except Exception:
        return "[UNREADABLE]"

def extract_experience(text):
    experience_matches = re.findall(r'(\d+)[+]?\s+years?', text.lower())
    years = [int(y) for y in experience_matches if y.isdigit()]
    return max(years) if years else 0

def infer_location(text):
    lines = text.split('\n')
    for line in lines[:5]:
        if any(city in line.lower() for city in ["pune", "bangalore", "delhi", "mumbai", "hyderabad"]):
            return line.lower()
    return ""

def semantic_qr_match(text, qr_list, threshold):
    matched_qrs = []
    text_embedding = model.encode([text], convert_to_tensor=True)
    for qr in qr_list:
        qr_embedding = model.encode([qr], convert_to_tensor=True)
        similarity = util.pytorch_cos_sim(qr_embedding, text_embedding).item()
        if similarity >= threshold:
            matched_qrs.append(qr.strip())
    match_percent = len(matched_qrs) / len(qr_list) if qr_list else 0
    return matched_qrs, match_percent

def match_resume_to_position(text, position_row, threshold):
    location_score = 0
    experience_score = 0
    qr_score = 0

    expected_location = str(position_row['Location']).lower()
    inferred_location = infer_location(text)
    if expected_location in text.lower() or expected_location in inferred_location:
        location_score = 1

    actual_experience = extract_experience(text)
    expected_experience = int(re.findall(r'\d+', str(position_row['Experience']))[0]) if re.findall(r'\d+', str(position_row['Experience'])) else 0
    if abs(actual_experience - expected_experience) <= 1:
        experience_score = 1

    qr_keywords = str(position_row['Essential QRs']).split(',')
    matched_qrs, match_percent = semantic_qr_match(text, qr_keywords, threshold)
    if match_percent >= 0.6:
        qr_score = 1

    total_score = location_score + experience_score + qr_score
    decision = "Use" if total_score == 3 else "Do Not Use"

    return location_score, experience_score, qr_score, decision, matched_qrs, int(match_percent * 100)

st.title("Resume to Position Matcher (Semantic Enhanced)")

threshold = st.slider("Set QR Matching Threshold (Cosine Similarity)", 0.0, 1.0, 0.6, 0.01)

checklist_file = st.file_uploader("Upload Position Checklist (Excel with columns: Position Title, Essential QRs, Experience, Location)", type=["xlsx"])
resume_files = st.file_uploader("Upload Resumes (PDF/DOCX)", type=["pdf", "docx"], accept_multiple_files=True)

if checklist_file and resume_files:
    df_positions = pd.read_excel(checklist_file)

    st.write("### Job Positions Loaded")
    st.dataframe(df_positions)

    results = []

    for resume in resume_files:
        if resume.name.lower().endswith(".pdf"):
            text = extract_text_from_pdf(resume)
        elif resume.name.lower().endswith(".docx"):
            text = extract_text_from_docx(resume)
        else:
            continue

        if text in ["[EMPTY]", "[UNREADABLE]"]:
            continue

        best_match = None
        best_score = -1
        best_data = None

        for _, row in df_positions.iterrows():
            loc_score, exp_score, qr_score, decision, matched_qrs, match_percent = match_resume_to_position(text, row, threshold)
            score = loc_score + exp_score + qr_score
            if score > best_score:
                best_score = score
                best_match = row['Position Title']
                best_data = (loc_score, exp_score, qr_score, decision, matched_qrs, match_percent)

        if best_data:
            loc_score, exp_score, qr_score, decision, matched_qrs, match_percent = best_data
            results.append({
                "File Name": resume.name,
                "Best Match Position": best_match,
                "Location Match": "✅" if loc_score else "❌",
                "Experience Match": "✅" if exp_score else "❌",
                "QR Match": "✅" if qr_score else "❌",
                "Decision": decision,
                "QR Match %": match_percent,
                "Matched QRs": ", ".join(matched_qrs)
            })

    results_df = pd.DataFrame(results)

    st.write("### Resume to Position Matching Results")
    st.dataframe(results_df)

    csv = results_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download Results as CSV", data=csv, file_name="semantic_resume_matches.csv", mime="text/csv")
