import streamlit as st
import pandas as pd
from fpdf import FPDF
from zipfile import ZipFile
import io
from datetime import datetime
import os

st.set_page_config(page_title="ALM Invoice Generator", layout="centered")
st.title("📄 ALM Invoice PDF Generator")

# Add logo uploader
logo_file = st.file_uploader("Upload Company Logo (JPG/PNG)", type=["jpg", "jpeg", "png"])

uploaded_file = st.file_uploader("Upload Subscription Report CSV", type=["csv"])

class ALMInvoice(FPDF):
    def __init__(self, logo_path=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.logo_path = logo_path
        
    def header(self):
        # Add logo if provided
        if self.logo_path and os.path.exists(self.logo_path):
            self.image(self.logo_path, x=10, y=8, w=30)
        
        # Invoice title (centered)
        self.set_font('Arial', 'B', 20)
        self.cell(0, 25, 'INVOICE', 0, 1, 'C')
        self.ln(10)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_invoice(row, logo_path=None):
    pdf = ALMInvoice(logo_path=logo_path, orientation="L")  # Landscape mode
    pdf.add_page()
    
    # Three-column table (no borders)
    pdf.set_font('Arial', '', 12)
    
    # Row 1: Invoice Number and Date
    pdf.cell(95, 10, "", 0, 0)  # First column empty
    pdf.cell(95, 10, f"Invoice #: {row['Sub_Ref_No']}", 0, 0)
    pdf.cell(0, 10, f"Date: {datetime.now().strftime('%m/%d/%Y')}", 0, 1)
    
    # Row 2: Headers with gray background
    pdf.set_fill_color(230, 230, 230)  # 20% gray
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(95, 10, "Bill To", 0, 0, 'L', fill=True)
    pdf.cell(95, 10, "Ship To", 0, 0, 'L', fill=True)
    pdf.cell(0, 10, "", 0, 1, fill=True)  # Empty third column
    
    # Row 3: Addresses
    pdf.set_font('Arial', '', 12)
    # Bill To Address
    pdf.cell(95, 6, f"{row['Bill_To_Contact_name']}", 0, 0, 'L')
    pdf.cell(95, 6, f"{row['Ship_To_Contact_name']}", 0, 0, 'L')
    pdf.cell(0, 6, "", 0, 1)
    
    pdf.cell(95, 6, f"{row['Bill_to_Company']}", 0, 0, 'L')
    pdf.cell(95, 6, f"{row['Ship_to_Company']}", 0, 0, 'L')
    pdf.cell(0, 6, "", 0, 1)
    
    pdf.cell(95, 6, f"{row['Bill_to_St_Address']}", 0, 0, 'L')
    pdf.cell(95, 6, f"{row['Ship_to_St_Address']}", 0, 0, 'L')
    pdf.cell(0, 6, "", 0, 1)
    
    city_state_zip_bill = f"{row['Bill_to_City']} {row['Bill_to_State']} {row['Bill_to_Zip']}"
    city_state_zip_ship = f"{row['Ship_to_City']} {row['Ship_to_State']} {row['Ship_to_Zip']}"
    pdf.cell(95, 6, city_state_zip_bill, 0, 0, 'L')
    pdf.cell(95, 6, city_state_zip_ship, 0, 0, 'L')
    
    # Promo and Sales Codes (third column)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 6, f"PROMO: {row['Curr_Promo_Code']}", 0, 1, 'L')
    pdf.cell(95, 6, "", 0, 0)  # Empty first column
    pdf.cell(95, 6, "", 0, 0)  # Empty second column
    pdf.cell(0, 6, f"SALES: {row['SalesCode']}", 0, 1, 'L')
    pdf.ln(5)  # Line space before next section
    
    # 6-column table
    col_widths = [30, 25, 40, 25, 30, 30]  # Adjust as needed
    
    # Header Row
    pdf.set_fill_color(230, 230, 230)  # 20% gray
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(col_widths[0], 10, "Cust. Acct. #", 1, 0, 'C', fill=True)
    pdf.cell(col_widths[1], 10, "Order #", 1, 0, 'C', fill=True)
    pdf.cell(col_widths[2], 10, "Purchase Order", 1, 0, 'C', fill=True)
    pdf.cell(col_widths[3], 10, "Term", 1, 0, 'C', fill=True)
    pdf.cell(col_widths[4], 10, "Order Date", 1, 0, 'C', fill=True)
    pdf.cell(col_widths[5], 10, "Due Date", 1, 1, 'C', fill=True)
    
    # Data Row
    pdf.set_font('Arial', '', 12)
    # Convert Excel date to proper format
    order_date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(row['Order_date'])) - 2
    formatted_order_date = order_date.strftime('%m/%d/%Y')
    
    pdf.cell(col_widths[0], 10, str(row['Customer_Account_Number']), 1, 0, 'C')
    pdf.cell(col_widths[1], 10, str(row['Order']), 1, 0, 'C')
    pdf.cell(col_widths[2], 10, str(row['PO_Num']) if pd.notna(row['PO_Num']) else "", 1, 0, 'C')
    pdf.cell(col_widths[3], 10, f"{row['Term']} days", 1, 0, 'C')
    pdf.cell(col_widths[4], 10, formatted_order_date, 1, 0, 'C')
    pdf.cell(col_widths[5], 10, "Due Upon Receipt", 1, 1, 'C')
    
    # [Rest of your existing invoice content...]
    # Installment Information
    pdf.ln(10)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, f"Installment Effort #: {int(row['Effort_No'])}", 0, 1)
    pdf.ln(5)
    
    # [Continue with all your existing content...]
    
    return pdf

if uploaded_file:
    try:
        # Save logo temporarily if uploaded
        logo_path = None
        if logo_file:
            logo_path = "temp_logo.jpg"
            with open(logo_path, "wb") as f:
                f.write(logo_file.getbuffer())
        
        df = pd.read_csv(uploaded_file)
        
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
                pdf = create_invoice(row, logo_path=logo_path)
                filename = generate_filename(row)
                pdf_bytes = pdf.output(dest='S').encode('latin1')
                zip_file.writestr(filename, pdf_bytes)
        
        # Clean up temporary logo file
        if logo_path and os.path.exists(logo_path):
            os.remove(logo_path)
        
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
        # Clean up temporary logo file if error occurs
        if logo_path and os.path.exists(logo_path):
            os.remove(logo_path)
