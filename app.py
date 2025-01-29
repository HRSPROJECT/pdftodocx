import fitz
import os
from docx import Document
from docx.shared import Inches
import tempfile
import io


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
            
            # Save the image in memory using io.BytesIO
            image_bytes = pix.tobytes("png")
            image_stream = io.BytesIO(image_bytes)
            
            document.add_picture(image_stream, width=Inches(6.5))

        document.save(docx_path)  # Save the DOCX

    else:
        # If searchable is False, insert images only (non-searchable)
        document = Document()
        for page_num in range(doc.page_count):
            page = doc[page_num]
            zoom_factor = 2  # Adjust for higher resolution if needed
            mat = fitz.Matrix(zoom_factor, zoom_factor)
            pix = page.get_pixmap(matrix=mat)
            
            # Save the image in memory using io.BytesIO
            image_bytes = pix.tobytes("png")
            image_stream = io.BytesIO(image_bytes)
            
            document.add_picture(image_stream, width=Inches(6.5))

        document.save(docx_path)  # Save the DOCX
