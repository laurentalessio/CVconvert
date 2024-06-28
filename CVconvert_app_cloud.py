import streamlit as st
import PyPDF2
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
import openai

def read_pdf(file):
    pdf_reader = PyPDF2.PdfReader(file)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return text

from openai import OpenAI

def process_cv(consultant_cv, template_cv, api_key):
    client = OpenAI(api_key=api_key)
    try:
        prompt = f"""
        You are a CV formatting assistant. Format the following consultant CV according to the THREE60 template provided below. 
        Ensure that the formatted CV includes the following sections in this order:

        1. Name
        2. Current position or role
        3. Years of experience
        4. Discipline
        5. Role
        6. Technical skills
        7. Professional skills
        8. A brief professional summary paragraph
        9. Work Experience - Summary (bullet points)
        10. Work Experience - Detailed (for each position: company, dates, position, and highlights)
        11. Education and training
        12. Personal skills and competencies (including Nationality, Languages, Software, and Professional Affiliations)

        Use the style and formatting from the template, including bold text for headings and section titles.
        Maintain a similar level of detail as seen in the template.

        Consultant CV to format:
        {consultant_cv}

        Template CV:
        {template_cv}

        Please provide the formatted CV in a form that can be easily copied into a Word document.
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

def create_word_document(formatted_cv):
    doc = Document()
    
    # Create styles
    styles = doc.styles
    style = styles.add_style('Bold', WD_STYLE_TYPE.PARAGRAPH)
    style.font.bold = True
    style.font.size = Pt(12)

    # Split the formatted CV into lines
    lines = formatted_cv.split('\n')

    for line in lines:
        if line.strip().startswith('**') and line.strip().endswith('**'):
            # This is a heading, use the Bold style
            p = doc.add_paragraph(line.strip('*'), style='Bold')
        else:
            # Regular text
            p = doc.add_paragraph(line)

    doc.save("formatted_cv.docx")

def main():
    st.title("CV Formatter for Three60 Template")

    # Add input for OpenAI API key
    api_key = st.text_input("Enter your OpenAI API key", type="password")

    consultant_cv_file = st.file_uploader("Upload Consultant CV (PDF)", type="pdf")
    template_cv_file = st.file_uploader("Upload Three60 Template Example (PDF)", type="pdf")

    if consultant_cv_file and template_cv_file and api_key:
        consultant_cv = read_pdf(consultant_cv_file)
        template_cv = read_pdf(template_cv_file)

        if st.button("Process CV"):
            formatted_cv = process_cv(consultant_cv, template_cv, api_key)
            if formatted_cv:
                create_word_document(formatted_cv)
                st.success("CV formatted successfully!")
                st.download_button(
                    label="Download Formatted CV",
                    data=open("formatted_cv.docx", "rb"),
                    file_name="formatted_cv.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.warning("Please upload both CV files and enter your OpenAI API key to proceed.")

if __name__ == "__main__":
    main()
