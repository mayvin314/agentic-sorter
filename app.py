import streamlit as st
import pandas as pd
from docx import Document
from PyPDF2 import PdfReader
from fuzzywuzzy import fuzz

# Extract text from a DOCX resume file
def extract_text_from_docx(file):
    doc = Document(file)
    return " ".join(para.text for para in doc.paragraphs)

# Extract text from a PDF resume file
def extract_text_from_pdf(file):
    reader = PdfReader(file)
    text = []
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text.append(extracted)
    return " ".join(text)

# Evaluate how well a resume matches a job position using fuzzy matching
def match_resume_to_position(text, position_row):
    qr_score = 0
    exp_score = 0
    loc_score = 0
    position_score = 0

    # Split essential qualifications into individual keywords/phrases
    qr_keywords = str(position_row['Essential QRs']).split(',')
    for keyword in qr_keywords:
        # Increase QR score if the keyword matches with the resume text
        if fuzz.token_set_ratio(keyword.lower().strip(), text.lower()) > 70:
            qr_score += 1

    # Score experience match
    if fuzz.partial_ratio(str(position_row['Experience']).lower(), text.lower()) > 70:
        exp_score = 1

    # Score location match
    if fuzz.partial_ratio(str(position_row['Location']).lower(), text.lower()) > 70:
        loc_score = 1

    # Score job title match
    if fuzz.token_set_ratio(str(position_row['Position Title']).lower(), text.lower()) > 70:
        position_score = 1

    # Calculate total match score
    total_score = qr_score + exp_score + loc_score + position_score
    return total_score

# Set up the Streamlit app interface
title_text = "📄 Resume to Position Matcher"
st.title(title_text)

# Upload Excel file containing job positions
checklist_file = st.file_uploader(
    "Upload Position Checklist (Excel with columns: Position Title, Essential QRs, Experience, Location)",
    type=["xlsx"]
)

# Upload multiple resumes
resume_files = st.file_uploader(
    "Upload Resumes (PDF/DOCX)",
    type=["pdf", "docx"],
    accept_multiple_files=True
)

# If both uploads are provided, start processing
if checklist_file and resume_files:
    # Load job positions data from Excel
    df_positions = pd.read_excel(checklist_file)

    # Display job positions table
    st.write("### Job Positions Loaded")
    st.dataframe(df_positions)

    results = []

    # Loop through each uploaded resume
    for resume in resume_files:
        # Extract text depending on file type
        if resume.name.lower().endswith(".pdf"):
            text = extract_text_from_pdf(resume)
        elif resume.name.lower().endswith(".docx"):
            text = extract_text_from_docx(resume)
        else:
            continue

        best_match = None
        best_score = -1

        # Compare resume against each position
        for _, row in df_positions.iterrows():
            score = match_resume_to_position(text, row)
            if score > best_score:
                best_score = score
                best_match = row['Position Title']

        # Save results for this resume
        results.append({
            "File Name": resume.name,
            "Best Match Position": best_match,
            "Score": best_score
        })

    # Create DataFrame of all results
    results_df = pd.DataFrame(results)

    # Display the matching results
    st.write("### Resume to Position Matching Results")
    st.dataframe(results_df)

    # Allow user to download the results as a CSV file
    csv = results_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download Results as CSV",
        data=csv,
        file_name="resume_position_matches.csv",
        mime="text/csv"
    )
