import streamlit as st
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from fuzzywuzzy import fuzz
import re

# Extract text content from DOCX resumes
def extract_text_from_docx(file):
    doc = Document(file)
    return " ".join(para.text for para in doc.paragraphs)

# Extract text content from PDF resumes
def extract_text_from_pdf(file):
    reader = PdfReader(file)
    text = []
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text.append(extracted)
    return " ".join(text)

# Extract years of experience from text, ignoring internships
def extract_experience(text):
    experience_matches = re.findall(r'(\d+)[+]?\s+years?', text.lower())
    years = [int(y) for y in experience_matches if y.isdigit()]
    return max(years) if years else 0

# Match resume content to essential QRs and calculate match percentage
def match_essential_qrs(text, qr_list):
    matched_qrs = [qr.strip() for qr in qr_list if fuzz.token_set_ratio(qr.strip().lower(), text.lower()) > 70]
    match_percent = len(matched_qrs) / len(qr_list) if qr_list else 0
    return matched_qrs, match_percent >= 0.6

# Try to infer location from the top lines of the resume
def infer_location(text):
    lines = text.split('\n')
    for line in lines[:5]:
        if any(city in line.lower() for city in ["pune", "bangalore", "delhi", "mumbai", "hyderabad"]):
            return line.lower()
    return ""

# Score a resume against a job position
def match_resume_to_position(text, position_row):
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
    matched_qrs, qr_met = match_essential_qrs(text, qr_keywords)
    if qr_met:
        qr_score = 1

    total_score = location_score + experience_score + qr_score
    decision = "Use" if total_score == 3 else "Do Not Use"

    return location_score, experience_score, qr_score, decision, matched_qrs

# App title
st.title("📄 Resume to Position Matcher (Enhanced)")

# File uploader for Excel checklist
checklist_file = st.file_uploader(
    "Upload Position Checklist (Excel with columns: Position Title, Essential QRs, Experience, Location)",
    type=["xlsx"]
)

# File uploader for multiple resume files
resume_files = st.file_uploader(
    "Upload Resumes (PDF/DOCX)",
    type=["pdf", "docx"],
    accept_multiple_files=True
)

# Main matching logic
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

        best_match = None
        best_score = -1
        best_data = None

        # Compare resume to each job position and keep the best match
        for _, row in df_positions.iterrows():
            loc_score, exp_score, qr_score, decision, matched_qrs = match_resume_to_position(text, row)
            score = loc_score + exp_score + qr_score
            if score > best_score:
                best_score = score
                best_match = row['Position Title']
                best_data = (loc_score, exp_score, qr_score, decision, matched_qrs)

        # Store final decision and component scores
        if best_data:
            loc_score, exp_score, qr_score, decision, matched_qrs = best_data
            results.append({
                "File Name": resume.name,
                "Best Match Position": best_match,
                "Location Match": "✅" if loc_score else "❌",
                "Experience Match": "✅" if exp_score else "❌",
                "QR Match": "✅" if qr_score else "❌",
                "Decision": decision,
                "Matched QRs": ", ".join(matched_qrs)
            })

    # Display and download final results
    results_df = pd.DataFrame(results)

    st.write("### Resume to Position Matching Results")
    st.dataframe(results_df)

    csv = results_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download Results as CSV",
        data=csv,
        file_name="resume_position_matches.csv",
        mime="text/csv"
    )
