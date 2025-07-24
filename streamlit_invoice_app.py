# streamlit_app.py

import streamlit as st
import pandas as pd
from fpdf import FPDF
from zipfile import ZipFile
import io
from datetime import datetime

st.set_page_config(page_title="Invoice/Statement Generator", layout="centered")
st.title("📄 Invoice / Statement PDF Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

def create_pdf(row, doc_type):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "ALM Media Billing", 0, 1, 'C')
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, doc_type.upper(), 0, 1, 'C')
    pdf.ln(10)

    pdf.set_font("Arial", size=11)
    pdf.cell(100, 8, f"Contact Name: {row.get('Contact Name', '')}", ln=0)
    pdf.cell(90, 8, f"Sub Ref: {row.get('Sub Ref Number', '')}", ln=1)
    pdf.cell(100, 8, f"Ship To: {row.get('Ship To Address', '')}", ln=0)
    pdf.cell(90, 8, f"Mail To: {row.get('Mail To Address', '')}", ln=1)
    pdf.cell(100, 8, f"Expire Date: {row.get('Expire Date', '')}", ln=1)

    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "Billing Details", ln=1)

    pdf.set_font("Arial", size=11)
    pdf.cell(0, 8, f"Product: {row.get('Product', 'N/A')} - ${row.get('Amount', '0.00')}", ln=1)
    pdf.cell(0, 8, f"Status: {doc_type}", ln=1)
    pdf.ln(5)
    pdf.cell(0, 8, f"Generated On: {datetime.now().strftime('%Y-%m-%d')}", ln=1)

    return pdf

if uploaded_file:
    st.success("✅ File uploaded. Generating PDFs...")
    df = pd.read_excel(uploaded_file)
    zip_buffer = io.BytesIO()

    with ZipFile(zip_buffer, 'w') as zip_file:
        for index, row in df.iterrows():
            doc_type = row.get("BQ", "Invoice")
            pdf = create_pdf(row, doc_type)

            customer_name = row.get("Contact Name", f"customer_{index}")
            safe_name = customer_name.replace(" ", "_").replace("/", "-")
            pdf_filename = f"{doc_type}_{safe_name}.pdf"

            pdf_str = pdf.output(dest='S').encode('latin1')
zip_file.writestr(pdf_filename, pdf_str)


zip_buffer.seek(0)
st.download_button(
        label="📥 Download ZIP of PDFs",
        data=zip_buffer,
        file_name="invoices_statements.zip",
        mime="application/zip"
    )
