import streamlit as st
import fitz
import os
from docx import Document
from docx.shared import Inches
import tempfile
import io


def pdf_to_docx(pdf_path, docx_path, searchable=True):
    """Converts a PDF to DOCX."""
    print(f"Starting PDF conversion: {pdf_path}")
    try:
        doc = fitz.open(pdf_path)
        print(f"PDF opened successfully: {pdf_path}")
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return

    if searchable:
        print("Searchable conversion enabled.")
        document = Document()
        for page_num in range(doc.page_count):
            print(f"Processing page: {page_num}")
            page = doc[page_num]
            text = page.get_text("text")  # Extract text with OCR
            document.add_paragraph(text)  # Add the extracted text

            # Convert page to image and insert
            zoom_factor = 2
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            
            # Save the image in memory using io.BytesIO
            image_bytes = pix.tobytes("png")
            image_stream = io.BytesIO(image_bytes)
            
            try:
                document.add_picture(image_stream, width=Inches(6.5))
                print(f"Image added for page: {page_num}")
            except Exception as e:
                print(f"Error adding image for page {page_num}: {e}")
                return

        document.save(docx_path)  # Save the DOCX
        print(f"Searchable DOCX saved: {docx_path}")

    else:
        print("Searchable conversion disabled.")
        document = Document()
        for page_num in range(doc.page_count):
            print(f"Processing page (non-searchable): {page_num}")
            page = doc[page_num]
            zoom_factor = 2  # Adjust for higher resolution if needed
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            
            # Save the image in memory using io.BytesIO
            image_bytes = pix.tobytes("png")
            image_stream = io.BytesIO(image_bytes)
            
            try:
              document.add_picture(image_stream, width=Inches(6.5))
              print(f"Image added for page (non-searchable): {page_num}")
            except Exception as e:
              print(f"Error adding image for page (non-searchable) {page_num}: {e}")
              return

        document.save(docx_path)  # Save the DOCX
        print(f"Non-searchable DOCX saved: {docx_path}")

hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

def main():
    st.title("PDF to DOCX Converter")
    
    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])
    
    if uploaded_file is not None:
        searchable = st.checkbox("Make Text Searchable (OCR)")
        
        if st.button("Convert to DOCX"):
          print("Conversion button pressed")
          try:
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
            print("Download button displayed")
          except Exception as e:
             print(f"Error during conversion or download: {e}")


if __name__ == "__main__":
    main()
