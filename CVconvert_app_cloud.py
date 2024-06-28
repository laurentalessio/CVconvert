import streamlit as st
from docx import Document
from io import BytesIO
import pdfplumber
import openai
import re
import os

def extract_text_from_first_page(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
    return text

def extract_entities(text, api_key):
    openai.api_key = api_key
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=f"Extract the following entities from the text: Name, Address, Phone, Email, Summary, Experience, Education, Skills. Text: {text}",
        max_tokens=500,
        n=1,
        stop=None,
        temperature=0.7,
    )
    return response.choices[0].text

def parse_entities(extracted_text):
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
    for key in user_data.keys():
        pattern = re.compile(rf"{key}:(.*?)(\n|$)")
        match = pattern.search(extracted_text)
        if match:
            user_data[key] = match.group(1).strip()
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

api_key = st.text_input("Enter your OpenAI API key:", type="password")

uploaded_pdf = st.file_uploader("Upload the candidate's PDF CV", type="pdf")
if uploaded_pdf and api_key:
    # Extract text from PDF
    text = extract_text_from_first_page(uploaded_pdf)
    st.text("Extracted Text:")
    st.text(text)
    
    # Extract entities using OpenAI
    extracted_text = extract_entities(text, api_key)
    st.text("Extracted Entities Text:")
    st.text(extracted_text)
    
    # Parse the extracted text
    user_data = parse_entities(extracted_text)
    st.json(user_data)
    
    # Load the template
    template_path = os.path.join(os.path.dirname(__file__), "CV_Martin_Boddy_THREE60_2024.docx")
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
