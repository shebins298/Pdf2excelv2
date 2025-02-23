import streamlit as st
import pdfplumber
import io
from docx import Document
import pandas as pd

def pdf_to_word(pdf_bytes):
    doc = Document()
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                doc.add_paragraph(text)
    return doc

def pdf_to_excel(pdf_bytes):
    text_content = []
    all_tables = []
    
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            # Extract text
            text = page.extract_text()
            if text:
                text_content.append(text)
            
            # Extract tables
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)
    
    # Create Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Add text content
        pd.DataFrame({'Text': text_content}).to_excel(
            writer, sheet_name='Text Content', index=False)
        
        # Add tables
        for i, table in enumerate(all_tables):
            pd.DataFrame(table).to_excel(
                writer, sheet_name=f'Table_{i+1}', index=False)
    
    return output

# Streamlit UI
st.set_page_config(page_title="PDF Converter", layout="wide")
st.title("PDF to Word/Excel Converter")

uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"])

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Convert to Word"):
            doc = pdf_to_word(pdf_bytes)
            bio = io.BytesIO()
            doc.save(bio)
            st.download_button(
                label="Download Word Document",
                data=bio.getvalue(),
                file_name="converted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    
    with col2:
        if st.button("Convert to Excel"):
            excel_file = pdf_to_excel(pdf_bytes)
            st.download_button(
                label="Download Excel File",
                data=excel_file.getvalue(),
                file_name="converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
