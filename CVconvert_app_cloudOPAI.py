import streamlit as st
from docx import Document
from io import BytesIO
import pdfplumber
import openai
import re
import os

def extract_text_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def extract_entities(text, api_key):
    openai.api_key = api_key
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": "You are a helpful assistant that extracts information from CVs."},
            {"role": "user", "content": f"Extract the following entities from the text: Name, Address, Phone, Email, Summary, Experience, Education, Skills. Text: {text}"}
        ],
        max_tokens=1000,
        n=1,
        temperature=0.7,
    )
    return response.choices[0].message['content']

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
        pattern = re.compile(rf"{key.strip('[]')}:(.*?)(\n\n|\Z)", re.IGNORECASE | re.DOTALL)
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

api_key = st.text_input("Enter your OpenAI API key (optional):", type="password")
api_key = api_key or os.getenv("OPENAI_API_KEY")

uploaded_pdf = st.file_uploader("Upload the candidate's PDF CV", type="pdf")

if uploaded_pdf and api_key:
    if st.button("Generate CV"):
        try:
            with st.spinner("Processing CV..."):
                # Extract text from PDF
                text = extract_text_from_pdf(uploaded_pdf)
                st.text("Extracted Text from PDF:")
                st.text(text[:500] + "...") # Show first 500 characters
                
                # Extract entities using OpenAI
                extracted_text = extract_entities(text, api_key)
                st.text("Extracted Entities:")
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
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
elif not uploaded_pdf:
    st.warning("Please upload a PDF CV to process.")
elif not api_key:
    st.warning("Please enter an OpenAI API key to use this application.")
