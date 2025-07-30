import streamlit as st
import pandas as pd
from fpdf import FPDF
from zipfile import ZipFile
import io
from datetime import datetime, timedelta
import os
import tempfile

st.set_page_config(page_title="ALM Invoice Generator", layout="centered")
st.title("📄 ALM Invoice PDF Generator")

# Add logo uploader
logo_file = st.file_uploader("Upload Company Logo (JPG/PNG)", type=["jpg", "jpeg", "png"])
uploaded_file = st.file_uploader("Upload Subscription Report CSV", type=["csv"])

class ALMInvoice(FPDF):
    def __init__(self, logo_data=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.logo_data = logo_data
        
    def header(self):
        # Add logo if provided
        if self.logo_data:
            try:
                # Create a temporary file for the logo
                with tempfile.NamedTemporaryFile(delete=False, suffix='.jpg') as tmp_logo:
                    tmp_logo.write(self.logo_data)
                    tmp_logo_path = tmp_logo.name
                
                self.image(tmp_logo_path, x=10, y=8, w=30)
                os.unlink(tmp_logo_path)  # Delete temp file immediately after use
            except Exception as e:
                st.warning(f"Could not load logo: {str(e)}")
        
        # Invoice title (centered)
        self.set_font('Arial', 'B', 20)
        self.cell(0, 25, 'INVOICE', 0, 1, 'C')
        self.ln(10)

def convert_excel_date(excel_date):
    """Convert Excel serial date to datetime object"""
    try:
        if pd.isna(excel_date) or excel_date == '':
            return None
        return datetime(1899, 12, 30) + timedelta(days=int(float(excel_date)))
    except:
        return None

def create_invoice(row, logo_data=None):
    pdf = ALMInvoice(logo_data=logo_data, orientation="L")  # Landscape mode
    pdf.add_page()
    
    # Convert all date fields
    order_date = convert_excel_date(row['Order_date'])
    expire_date = convert_excel_date(row['Expire_Date'])
    due_date = convert_excel_date(row.get('DueDate'))  # Using get() in case column doesn't exist
    
    # Format dates for display
    formatted_order_date = order_date.strftime('%m/%d/%Y') if order_date else "N/A"
    formatted_expire_date = expire_date.strftime('%m/%d/%Y') if expire_date else "N/A"
    formatted_due_date = due_date.strftime('%m/%d/%Y') if due_date else "Due Upon Receipt"

    # [Rest of your invoice generation code...]
    # Use the formatted dates wherever needed in your PDF
    
    return pdf

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file)
        
        # Get logo data if uploaded
        logo_data = logo_file.getvalue() if logo_file else None
        
        # Generate filename for each invoice
        def generate_filename(row):
            today = datetime.now().strftime('%m_%d_%Y')
            company = row['Ship_to_Company'].replace(' ', '_').upper()
            sub_ref = row['Sub_Ref_No']
            group_id = row['Group_ID']
            account_num = row['Customer_Account_Number']
            return f"{company}_{sub_ref}_{today}_INV_{group_id}_{account_num}.pdf"
        
        zip_buffer = io.BytesIO()
        
        with ZipFile(zip_buffer, 'w') as zip_file:
            for _, row in df.iterrows():
                pdf = create_invoice(row, logo_data=logo_data)
                filename = generate_filename(row)
                pdf_bytes = pdf.output(dest='S').encode('latin1')
                zip_file.writestr(filename, pdf_bytes)
        
        zip_buffer.seek(0)
        st.success(f"✅ Successfully generated {len(df)} invoices!")
        
        st.download_button(
            label="📥 Download All Invoices (ZIP)",
            data=zip_buffer,
            file_name="alm_invoices.zip",
            mime="application/zip"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
