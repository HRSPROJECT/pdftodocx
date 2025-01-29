import streamlit as st
import io
from docx import Document
from docx.shared import Inches
from PyPDF2 import PdfReader
# For using pdfplumber, uncomment this:
# import pdfplumber


def extract_text_from_pdf(pdf_file):
    """Extracts text from a PDF file."""
    text = ""
    try:
        pdf_reader = PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    except Exception as e:
       st.error(f"Error reading PDF: {e}")
       return None

    return text
# For using pdfplumber, uncomment this:
# def extract_text_from_pdf(pdf_file):
#     """Extracts text from a PDF file using pdfplumber."""
#     text = ""
#     try:
#        with pdfplumber.open(pdf_file) as pdf:
#             for page in pdf.pages:
#                 text += page.extract_text()
#     except Exception as e:
#        st.error(f"Error reading PDF: {e}")
#        return None
#
#     return text


def create_docx(text):
    """Creates a DOCX document from the given text."""
    document = Document()
    document.add_paragraph(text)
    return document


def main():
    st.title("PDF to DOCX Converter")
    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")
    if uploaded_file is not None:
        with st.spinner("Processing..."):
            pdf_text = extract_text_from_pdf(uploaded_file)
            if pdf_text:
                docx_document = create_docx(pdf_text)

                # Save DOCX file to BytesIO for download
                docx_stream = io.BytesIO()
                docx_document.save(docx_stream)
                docx_stream.seek(0)

                st.download_button(
                    label="Download DOCX",
                    data=docx_stream,
                    file_name="converted.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
    st.markdown("Created by [Your Name] with Streamlit")
if __name__ == "__main__":
    main()
