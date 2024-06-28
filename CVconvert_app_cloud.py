import streamlit as st
from docx import Document
from io import BytesIO
import pdfplumber
import re

def extract_text_from_first_page(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
    return text

def parse_pdf_text(text):
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

    name_match = re.search(r"Name:\s*(.*)", text)
    if name_match:
        user_data["[NAME]"] = name_match.group(1).strip()

    address_match = re.search(r"Address:\s*(.*)", text)
    if address_match:
        user_data["[ADDRESS]"] = address_match.group(1).strip()

    phone_match = re.search(r"Phone:\s*(.*)", text)
    if phone_match:
        user_data["[PHONE]"] = phone_match.group(1).strip()

    email_match = re.search(r"Email:\s*(.*)", text)
    if email_match:
        user_data["[EMAIL]"] = email_match.group(1).strip()

    summary_match = re.search(r"Summary:\s*(.*)", text, re.DOTALL)
    if summary_match:
        user_data["[SUMMARY]"] = summary_match.group(1).strip()

    experience_match = re.search(r"Experience:\s*(.*)", text, re.DOTALL)
    if experience_match:
        user_data["[EXPERIENCE]"] = experience_match.group(1).strip()

    education_match = re.search(r"Education:\s*(.*)", text, re.DOTALL)
    if education_match:
        user_data["[EDUCATION]"] = education_match.group(1).strip()

    skills_match = re.search(r"Skills:\s*(.*)", text, re.DOTALL)
    if skills_match:
        user_data["[SKILLS]"] = skills_match.group(1).strip()

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
    
    # Parse the extracted text
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
