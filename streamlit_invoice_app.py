import streamlit as st
import pandas as pd
from fpdf import FPDF
from zipfile import ZipFile
import io
from datetime import datetime, timedelta
import tempfile
import os
from PIL import Image

st.set_page_config(page_title="ALM Invoice Generator", layout="centered")
st.title("📄 ALM Invoice PDF Generator")

# Add logo uploader with format validation
logo_file = st.file_uploader("Upload Company Logo", type=["jpg", "jpeg", "png"])
uploaded_file = st.file_uploader("Upload Subscription Report CSV", type=["csv"])

class ALMInvoice(FPDF):
    def __init__(self, *args, **kwargs):
        self.logo = kwargs.pop('logo', None)
        super().__init__(*args, **kwargs)
        
    def header(self):
        # Add logo if provided and valid
        if self.logo and self.logo['valid']:
            try:
                self.image(self.logo['path'], x=10, y=8, w=30)
            except:
                pass  # Silently fail if image can't be loaded
        
        # Invoice title
        self.set_font('Arial', 'B', 20)
        self.cell(0, 25, 'INVOICE', 0, 1, 'C')
        self.ln(10)

def process_logo(uploaded_file):
    """Validate and prepare logo for PDF"""
    if not uploaded_file:
        return None
        
    try:
        # Verify image is valid
        img = Image.open(uploaded_file)
        img.verify()
        
        # Save to temp file
        temp_dir = tempfile.mkdtemp()
        logo_path = os.path.join(temp_dir, 'logo.jpg')
        uploaded_file.seek(0)
        with open(logo_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
            
        return {
            'path': logo_path,
            'valid': True,
            'temp_dir': temp_dir
        }
    except Exception as e:
        st.warning(f"Invalid logo image: {str(e)}")
        return None

def convert_excel_date(excel_date):
    """Robust Excel date conversion"""
    try:
        if pd.isna(excel_date) or excel_date == '':
            return None
        return datetime(1899, 12, 30) + timedelta(days=int(float(excel_date)))
    except:
        return None

def create_invoice(row, logo):
    try:
        pdf = ALMInvoice(orientation="L", logo=logo)
        pdf.add_page()
        
        # --- Invoice Content ---
        # 1. Header Section
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 10, f"Invoice #: {row.get('Sub_Ref_No', 'N/A')}", 0, 1)
        pdf.cell(0, 10, f"Date: {datetime.now().strftime('%m/%d/%Y')}", 0, 1)
        pdf.ln(10)
        
        # 2. Bill To / Ship To
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(95, 10, "BILL TO:", 0, 0)
        pdf.cell(95, 10, "SHIP TO:", 0, 1)
        
        pdf.set_font('Arial', '', 12)
        pdf.cell(95, 6, str(row.get('Bill_To_Contact_name', '')), 0, 0)
        pdf.cell(95, 6, str(row.get('Ship_To_Contact_name', '')), 0, 1)
        
        # Continue with all other sections...
        # Add your product table, payment info, etc here
        # Make sure to use row.get() with default values
        
        return pdf
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

if uploaded_file:
    try:
        # Process logo first
        logo = process_logo(logo_file)
        
        # Read CSV
        df = pd.read_csv(uploaded_file)
        zip_buffer = io.BytesIO()
        success_count = 0
        
        with ZipFile(zip_buffer, 'w') as zip_file:
            for _, row in df.iterrows():
                pdf = create_invoice(row, logo)
                if pdf:
                    filename = f"Invoice_{row.get('Sub_Ref_No', '')}.pdf"
                    try:
                        pdf_bytes = pdf.output(dest='S').encode('latin1')
                        zip_file.writestr(filename, pdf_bytes)
                        success_count += 1
                    except:
                        continue
        
        # Clean up logo temp files
        if logo and 'temp_dir' in logo:
            try:
                os.remove(logo['path'])
                os.rmdir(logo['temp_dir'])
            except:
                pass
        
        if success_count > 0:
            zip_buffer.seek(0)
            st.success(f"✅ Generated {success_count} invoices successfully!")
            st.download_button(
                label="📥 Download Invoices",
                data=zip_buffer,
                file_name="invoices.zip",
                mime="application/zip"
            )
        else:
            st.error("Failed to generate any invoices")
            
    except Exception as e:
        st.error(f"Processing failed: {str(e)}")
