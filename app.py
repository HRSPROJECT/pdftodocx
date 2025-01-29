import streamlit as st
import fitz
import os
from docx import Document
from docx.shared import Inches
import tempfile
import io
import requests  # To interact with the cloud storage API
import json
import urllib.parse

# --- Function to upload a file and generate a shareable link
def upload_to_cloud(file_path, filename):
    """Placeholder for actual cloud upload logic."""
    # Replace with your actual API endpoint, headers, parameters, etc.
    # This is just an example; you need your own service
    
    cloud_api_endpoint = "YOUR_CLOUD_UPLOAD_API_ENDPOINT"
    # IMPORTANT: Add your api key to avoid unauthorized usage.
    api_key = "YOUR_API_KEY"

    try:
      with open(file_path, 'rb') as f:
          files = {'file': (filename, f)}  
          headers = {'Authorization': f'Bearer {api_key}'}
          
          response = requests.post(cloud_api_endpoint, files=files, headers=headers)
          
          response.raise_for_status() # Raise exception for bad status codes
          
          response_json = response.json()
          shareable_link = response_json.get('url')
          
          if not shareable_link:
              raise ValueError("No 'url' found in the response from the cloud API.")
          
          print(f"Successfully uploaded to cloud. Link:{shareable_link}")
          return shareable_link
    
    except requests.exceptions.RequestException as e:
          print(f"Error connecting to cloud service: {e}")
          return None

    except json.JSONDecodeError as e:
        print(f"Error decoding JSON: {e}")
        return None
    
    except ValueError as e:
        print(f"Error on response from cloud service: {e}")
        return None


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


def main():
    st.title("PDF to DOCX")
    
    st.markdown(
        """
        <a href="https://hrsproject.github.io/home/" target="_blank">Explore More</a>
        """,
        unsafe_allow_html=True,
    )
    
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

            # Upload the file and get the shareable link
            shareable_link = upload_to_cloud(docx_path, os.path.basename(docx_path))

            # Delete the local DOCX file
            os.unlink(docx_path)  
            os.unlink(pdf_path)
            
            if shareable_link:
                # Create the WhatsApp share link
                whatsapp_message = f"Download your DOCX file here: {shareable_link}"
                whatsapp_link = f"https://wa.me/?text={urllib.parse.quote(whatsapp_message)}"
                
                st.markdown(
                    f"""
                    <a href="{whatsapp_link}" target="_blank">Share on WhatsApp</a>
                    """,
                   unsafe_allow_html=True
                )
            
            else:
                st.error("Error creating shareable link.")

          except Exception as e:
             print(f"Error during conversion or upload: {e}")


if __name__ == "__main__":
    main()
