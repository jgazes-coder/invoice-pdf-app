import streamlit as st
import pandas as pd
from fpdf import FPDF
from zipfile import ZipFile
import io
from datetime import datetime, timedelta
import tempfile
import base64

st.set_page_config(page_title="ALM Invoice Generator", layout="centered")
st.title("📄 ALM Invoice PDF Generator")

# Add logo uploader
logo_file = st.file_uploader("Upload Company Logo (JPG/PNG)", type=["jpg", "jpeg", "png"])
uploaded_file = st.file_uploader("Upload Subscription Report CSV", type=["csv"])

class ALMInvoice(FPDF):
    def __init__(self, *args, **kwargs):
        self.logo_data = kwargs.pop('logo_data', None)
        super().__init__(*args, **kwargs)
        
    def header(self):
        # Add logo if provided
        if self.logo_data:
            try:
                # Create temp file for logo
                with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp:
                    tmp.write(self.logo_data)
                    tmp_path = tmp.name
                
                self.image(tmp_path, x=10, y=8, w=30)
                try:
                    os.unlink(tmp_path)
                except:
                    pass
            except Exception as e:
                st.warning(f"Logo not loaded: {str(e)}")
        
        # Invoice title
        self.set_font('Arial', 'B', 20)
        self.cell(0, 25, 'INVOICE', 0, 1, 'C')
        self.ln(10)

def convert_excel_date(excel_date):
    """Robust Excel date conversion"""
    try:
        if pd.isna(excel_date) or excel_date == '':
            return None
        return datetime(1899, 12, 30) + timedelta(days=int(float(excel_date)))
    except:
        return None

def create_invoice(row, logo_data=None):
    try:
        pdf = ALMInvoice(orientation="L", logo_data=logo_data)
        pdf.add_page()
        
        # Convert dates
        dates = {
            'order_date': convert_excel_date(row['Order_date']),
            'expire_date': convert_excel_date(row['Expire_Date']),
            'due_date': convert_excel_date(row.get('DueDate'))
        }
        
        # --- Your Invoice Content Here ---
        # Example: Basic info section
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, f"Invoice #: {row['Sub_Ref_No']}", 0, 1)
        
        # Bill To section
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, "BILL TO:", 0, 1)
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 6, f"{row['Bill_To_Contact_name']}", 0, 1)
        pdf.cell(0, 6, f"{row['Bill_to_Company']}", 0, 1)
        
        # Add all your other content sections here...
        # Make sure to use proper error handling for each field
        
        return pdf
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file)
        logo_data = logo_file.read() if logo_file else None
        
        zip_buffer = io.BytesIO()
        success_count = 0
        
        with ZipFile(zip_buffer, 'w') as zip_file:
            for _, row in df.iterrows():
                pdf = create_invoice(row, logo_data)
                if pdf:
                    filename = f"Invoice_{row['Sub_Ref_No']}.pdf"
                    pdf_bytes = pdf.output(dest='S').encode('latin1')
                    zip_file.writestr(filename, pdf_bytes)
                    success_count += 1
        
        if success_count > 0:
            zip_buffer.seek(0)
            st.success(f"✅ Generated {success_count}/{len(df)} invoices!")
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
