import streamlit as st
from docx import Document
from io import BytesIO
import pdfplumber
import spacy

# Load SpaCy English model
nlp = spacy.load("en_core_web_sm")

def extract_text_from_first_page(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
    return text

def parse_pdf_text(text):
    doc = nlp(text)
    user_data = {
        "[NAME]": "",
        "[ADDRESS]": "",
        "[PHONE]": "",
        "[EMAIL]": "",
        "[SUMMARY]": "",
        "[EXPERIENCE]": "",
        "[EDUCATION]": "",
        "[SKILLS]": "",
    }
    
    for ent in doc.ents:
        if ent.label_ == "PERSON":
            user_data["[NAME]"] = ent.text
        elif ent.label_ == "GPE":  # Geopolitical entity, used for addresses
            user_data["[ADDRESS]"] = ent.text
        elif ent.label_ == "ORG":  # Organizations, might be useful for experience
            user_data["[EXPERIENCE]"] += ent.text + "\n"
        elif ent.label_ == "DATE":  # Dates, might be useful for education and experience
            user_data["[EDUCATION]"] += ent.text + "\n"
        elif "@" in ent.text:
            user_data["[EMAIL]"] = ent.text
        elif ent.label_ == "CARDINAL":  # Basic way to capture phone numbers
            user_data["[PHONE]"] = ent.text

    # Additional heuristic parsing for summary and skills
    summary_start = text.lower().find("summary:")
    if summary_start != -1:
        summary_end = text.find("\n", summary_start)
        user_data["[SUMMARY]"] = text[summary_start + 8:summary_end].strip()

    skills_start = text.lower().find("skills:")
    if skills_start != -1:
        skills_end = text.find("\n", skills_start)
        user_data["[SKILLS]"] = text[skills_start + 7:skills_end].strip()
    
    return user_data

def fill_template(doc, user_data):
    for paragraph in doc.paragraphs:
        for key, value in user_data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    return doc

def save_document(doc):
    byte_io = BytesIO()
    doc.save(byte_io)
    byte_io.seek(0)
    return byte_io

# Streamlit application
st.title("Three60 CV Generator")

uploaded_pdf = st.file_uploader("Upload the candidate's PDF CV", type="pdf")
if uploaded_pdf:
    # Extract text from PDF
    text = extract_text_from_first_page(uploaded_pdf)
    st.text("Extracted Text:")
    st.text(text)
    
    # Parse the extracted text using NLP
    user_data = parse_pdf_text(text)
    st.json(user_data)
    
    # Load the template
    template_path = "CV_Martin_Boddy_THREE60_2024.docx"  # Path to the template
    doc = Document(template_path)
    
    # Fill the template
    filled_doc = fill_template(doc, user_data)
    
    # Save the document and provide download link
    docx_file = save_document(filled_doc)
    st.download_button(
        label="Download the completed CV",
        data=docx_file,
        file_name="Completed_CV.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
