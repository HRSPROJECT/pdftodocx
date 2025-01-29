import streamlit as st
import io
from docx import Document
from PyPDF2 import PdfReader
import base64

# CSS to hide Streamlit elements and set a default theme
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """

st.markdown(hide_st_style, unsafe_allow_html=True)

# Function to create the navigation bar
def create_navbar(theme):
    navbar_style = f"""
        <nav style="background-color: {'#333' if theme == 'dark' else '#f0f0f0'}; padding: 10px; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 20px; font-weight: bold; color: {'white' if theme == 'dark' else 'black'};">PDF to DOCX Converter</span>
            <a href="https://hrsproject.github.io/home/" target="_blank" style="color: {'white' if theme == 'dark' else 'black'}; text-decoration: none; padding: 8px 12px; border: 1px solid {'white' if theme == 'dark' else 'black'}; border-radius: 4px;">Explore</a>
        </nav>
    """
    st.markdown(navbar_style, unsafe_allow_html=True)

# Function to create a download button that includes preview
def create_download_button(docx_stream, file_name, preview_text, theme):
    preview_html = f"""
            <div style="padding: 10px; margin-bottom: 10px; border: 1px solid {'#555' if theme == 'dark' else '#ccc'}; background-color: {'#222' if theme == 'dark' else '#fff'}; color: {'white' if theme == 'dark' else 'black'}; border-radius: 5px; max-height: 300px; overflow-y: auto;">
                <p style="font-size: 14px; white-space: pre-wrap;">{preview_text}</p>
            </div>
    """
    st.markdown(preview_html, unsafe_allow_html=True)
    st.download_button(
        label="Download DOCX",
        data=docx_stream,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# Function to extract text from PDF
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


# Function to create a DOCX document
def create_docx(text):
    """Creates a DOCX document from the given text."""
    document = Document()
    document.add_paragraph(text)
    return document

# Function to encode an image
def get_base64_encoded_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode()

def main():

    # Initialize theme in session state
    if "theme" not in st.session_state:
       st.session_state.theme = "light"

    # Theme toggle
    col1, col2 = st.columns(2)
    with col2:
        if st.button("Toggle Dark Mode" if st.session_state.theme == "light" else "Toggle Light Mode"):
            st.session_state.theme = "dark" if st.session_state.theme == "light" else "light"

    create_navbar(st.session_state.theme)

    st.title("PDF to DOCX Converter")
    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")


    if uploaded_file is not None:
        with st.spinner("Processing..."):
            pdf_text = extract_text_from_pdf(uploaded_file)
            if pdf_text:
                docx_document = create_docx(pdf_text)
                docx_stream = io.BytesIO()
                docx_document.save(docx_stream)
                docx_stream.seek(0)

                # Provide preview and download button
                create_download_button(docx_stream, "converted.docx", pdf_text, st.session_state.theme)

    st.markdown("Created by [Your Name] with Streamlit")
if __name__ == "__main__":
    main()
