import streamlit as st
import PyPDF2
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from openai import OpenAI
import io

def read_pdf(file):
    pdf_reader = PyPDF2.PdfReader(file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

def read_docx(file):
    doc = Document(file)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def process_cv(consultant_cv, template_cv, api_key):
    client = OpenAI(api_key=api_key)
    try:
        prompt = f"""
        Format the following consultant CV according to the THREE60 template provided below. 
        Maintain the exact structure, headings, and order of sections as in the template.
        Do not add any sections that are not in the template.
        If information for a section is not available in the consultant's CV, leave that section empty or write 'Information not provided'.
        
        Consultant CV to format:
        {consultant_cv}

        Template CV:
        {template_cv}

        Please provide the formatted CV content, preserving the structure and headings of the template.
        """

        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a CV formatting assistant."},
                {"role": "user", "content": prompt}
            ]
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Error processing CV: {str(e)}")
        return None

def create_word_document(formatted_cv, template_file):
    # Load the template
    doc = Document(template_file)
    
    # Clear the content of the template, keeping styles and formatting
    for paragraph in doc.paragraphs:
        if paragraph.text and not paragraph.style.name.startswith('Heading'):
            paragraph.clear()
    
    # Split the formatted CV into lines
    lines = formatted_cv.split('\n')
    
    current_paragraph = None
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        if line.startswith('**') and line.endswith('**'):
            # This is a heading
            text = line.strip('**').strip()
            for paragraph in doc.paragraphs:
                if paragraph.text.strip() == text:
                    current_paragraph = paragraph
                    break
            else:
                current_paragraph = doc.add_paragraph(text)
                current_paragraph.style = 'Heading 1'
        elif current_paragraph is not None:
            current_paragraph.add_run('\n' + line)
        else:
            current_paragraph = doc.add_paragraph(line)
    
    # Save the document to a bytes buffer
    doc_buffer = io.BytesIO()
    doc.save(doc_buffer)
    doc_buffer.seek(0)
    return doc_buffer

def main():
    st.title("CV Formatter for Three60 Template")

    # Add input for OpenAI API key
    api_key = st.text_input("Enter your OpenAI API key", type="password")

    consultant_cv_file = st.file_uploader("Upload Consultant CV (PDF)", type="pdf")
    template_cv_file = st.file_uploader("Upload Three60 Template Example (DOCX)", type="docx")

    if consultant_cv_file and template_cv_file and api_key:
        consultant_cv = read_pdf(consultant_cv_file)
        template_cv = read_docx(template_cv_file)

        if st.button("Process CV"):
            formatted_cv = process_cv(consultant_cv, template_cv, api_key)
            if formatted_cv:
                doc_buffer = create_word_document(formatted_cv, template_cv_file)
                st.success("CV formatted successfully!")
                st.download_button(
                    label="Download Formatted CV",
                    data=doc_buffer,
                    file_name="formatted_cv.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.warning("Please upload both CV files and enter your OpenAI API key to proceed.")

if __name__ == "__main__":
    main()
