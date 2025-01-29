import streamlit as st
import fitz
import os
from docx import Document
from docx.shared import Inches
import tempfile

def pdf_to_docx(pdf_path, docx_path, searchable=True):
    """Converts a PDF to DOCX.

    Args:
        pdf_path: Path to the PDF file.
        docx_path: Path to the output DOCX file.
        searchable: Whether to make the text searchable (OCR). Defaults to True.
    """
    doc = fitz.open(pdf_path)

    if searchable:
        # If searchable is True, extract text with OCR and insert images
        document = Document()
        for page_num in range(doc.page_count):
            page = doc[page_num]
            text = page.get_text("text")  # Extract text with OCR
            document.add_paragraph(text)  # Add the extracted text

            # Convert page to image and insert
            zoom_factor = 2  # Adjust for higher resolution if needed
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            
            # Save the image in memory instead of using filesystem
            image_bytes = pix.tobytes("png")
            
            document.add_picture(image_bytes, width=Inches(6.5))

        document.save(docx_path)  # Save the DOCX

    else:
        # If searchable is False, insert images only (non-searchable)
        document = Document()
        for page_num in range(doc.page_count):
            page = doc[page_num]
            zoom_factor = 2  # Adjust for higher resolution if needed
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            
            # Save the image in memory instead of using filesystem
            image_bytes = pix.tobytes("png")
            
            document.add_picture(image_bytes, width=Inches(6.5))

        document.save(docx_path)  # Save the DOCX
        
def main():
    st.title("PDF to DOCX Converter")
    
    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])
    
    if uploaded_file is not None:
        searchable = st.checkbox("Make Text Searchable (OCR)")
        
        if st.button("Convert to DOCX"):
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_file:
                tmp_file.write(uploaded_file.read())
                pdf_path = tmp_file.name
            
            docx_path = os.path.splitext(pdf_path)[0] + ".docx"
            
            with st.spinner("Converting..."):
                pdf_to_docx(pdf_path, docx_path, searchable=searchable)

            with open(docx_path, "rb") as file:
                docx_bytes = file.read()
            
            os.unlink(pdf_path)  # Delete the temporary PDF file
            os.unlink(docx_path)  # Delete the temporary DOCX file

            st.download_button(
                label="Download DOCX",
                data=docx_bytes,
                file_name=os.path.splitext(uploaded_file.name)[0] + ".docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )


if __name__ == "__main__":
    main()
