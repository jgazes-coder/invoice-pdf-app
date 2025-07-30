import streamlit as st
import pandas as pd
from fpdf import FPDF
from zipfile import ZipFile
import io
from datetime import datetime

st.set_page_config(page_title="ALM Invoice Generator", layout="centered")
st.title("📄 ALM Invoice PDF Generator")

uploaded_file = st.file_uploader("Upload Subscription Report CSV", type=["csv"])

class ALMInvoice(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'ALM Global, LLC', 0, 1, 'C')
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'INVOICE', 0, 1, 'C')
        self.ln(10)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_invoice(row):
    pdf = ALMInvoice()
    pdf.add_page()
    
    # Date and Invoice Number
    invoice_date = datetime.now().strftime('%B %d, %Y')
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 8, invoice_date, 0, 1, 'R')
    
    # Billing Information
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 8, 'BILL TO:', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"{row['Bill_To_Contact_name']}", 0, 1)
    pdf.cell(0, 6, f"{row['Bill_to_Company']}", 0, 1)
    pdf.cell(0, 6, f"{row['Bill_to_St_Address']}", 0, 1)
    if pd.notna(row['Bill_to_Division']):
        pdf.cell(0, 6, f"{row['Bill_to_Division']}", 0, 1)
    pdf.cell(0, 6, f"{row['Bill_to_City']} {row['Bill_to_State']} {row['Bill_to_Zip']}", 0, 1)
    pdf.ln(5)
    
    # Customer Account Info
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(40, 6, 'CUST. ACCT. #:', 0, 0)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"{row['Customer_Account_Number']}", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(40, 6, 'ORDER #:', 0, 0)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"{row['Order']}", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(40, 6, 'PURCHASE ORDER:', 0, 0)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"{row['PO_Num'] if pd.notna(row['PO_Num']) else ''}", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(40, 6, 'TERM:', 0, 0)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"{row['Term']} days", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(40, 6, 'ORDER DATE:', 0, 0)
    pdf.set_font('Arial', '', 10)
    order_date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(row['Order_date']) - 2)
    pdf.cell(0, 6, order_date.strftime('%m/%d/%Y'), 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(40, 6, 'DUE DATE:', 0, 0)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, "Due Upon Receipt", 0, 1)
    pdf.ln(10)
    
    # Installment Information
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, f"Installment Effort #: {int(row['Effort_No'])}", 0, 1)
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "CUMULATIVE QUARTERLY BILLINGS:", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"${float(row['Amount_Due']):,.2f}", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "PAYMENTS:", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"${float(row['Paid_Amount']):,.2f}", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "CURRENT QUARTERLY BILLINGS:", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"${float(row['Instalment']):,.2f}", 0, 1)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "MINIMUM BAL. DUE TODAY:", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"${float(row['Instalment']):,.2f}", 0, 1)
    pdf.ln(10)
    
    # Subscription Details
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "SHIP TO:", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"{row['Ship_To_Contact_name']}", 0, 1)
    pdf.cell(0, 6, f"{row['Ship_to_Company']}", 0, 1)
    pdf.cell(0, 6, f"{row['Ship_to_St_Address']}", 0, 1)
    if pd.notna(row['Ship_to_Division']):
        pdf.cell(0, 6, f"{row['Ship_to_Division']}", 0, 1)
    pdf.cell(0, 6, f"{row['Ship_to_City']} {row['Ship_to_State']} {row['Ship_to_Zip']}", 0, 1)
    pdf.ln(5)
    
    # Product Table
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 6, 'Sub. Ref. #', 1, 0, 'C')
    pdf.cell(25, 6, 'Product', 1, 0, 'C')
    pdf.cell(15, 6, 'Copies', 1, 0, 'C')
    pdf.cell(50, 6, 'Full Journal Name', 1, 0, 'C')
    pdf.cell(15, 6, 'Seats', 1, 0, 'C')
    pdf.cell(30, 6, 'Description', 1, 0, 'C')
    pdf.cell(20, 6, 'End Date', 1, 0, 'C')
    pdf.cell(20, 6, 'Sales', 1, 0, 'C')
    pdf.cell(15, 6, 'S&H', 1, 0, 'C')
    pdf.cell(15, 6, 'Tax', 1, 0, 'C')
    pdf.cell(20, 6, 'Payment', 1, 0, 'C')
    pdf.cell(20, 6, 'Total Due', 1, 1, 'C')
    
    pdf.set_font('Arial', '', 8)
    # Convert Excel date to datetime
    expire_date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(row['Expire_Date']) - 2)
    
    pdf.cell(30, 6, str(row['Sub_Ref_No']), 1)
    pdf.cell(25, 6, row['Pub_Code'], 1)
    pdf.cell(15, 6, str(int(row['Quantity'])), 1)
    pdf.cell(50, 6, row['Pub_desc'], 1)
    pdf.cell(15, 6, str(int(row['Num_of_Seats'])), 1)
    pdf.cell(30, 6, 'Online+EntSite', 1)
    pdf.cell(20, 6, expire_date.strftime('%m/%d/%Y'), 1)
    pdf.cell(20, 6, f"${float(row['Material_Amount']):,.2f}", 1)
    pdf.cell(15, 6, f"${float(row['Postage']):,.2f}", 1)
    pdf.cell(15, 6, f"${float(row['Tax']):,.2f}", 1)
    pdf.cell(20, 6, f"${float(row['Paid_Amount']):,.2f}", 1)
    pdf.cell(20, 6, f"${float(row['Amount_Due']):,.2f}", 1)
    pdf.ln(10)
    
    # Payment Instructions
    pdf.set_font('Arial', '', 10)
    pdf.multi_cell(0, 6, "To Pay by Check: Make checks payable to ALM Global, LLC and reference your subscription number on your check. Allow 14-21 days for your check to credit to your account. Disregard invoices you may receive once you have paid.")
    pdf.multi_cell(0, 6, "Send your check along with this form to the address below:")
    pdf.multi_cell(0, 6, "US: ALM Global, LLC, PO BOX 70162, Philadelphia, PA, 19176-9628")
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "Payment Options", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.multi_cell(0, 6, "To Pay by EFT:\nBANK NAME: WELLS FARGO BANK, N.A.\nADDRESS: 420 Montgomery Street, San Francisco, CA 94104\nACCOUNT NUMBER: 2000005971161\nABA NUMBER: 121000248\nBANK ACCOUNT NAME: ALM Global, LLC\nSWIFT: WFBIUS6S\nCHIPS: 0407")
    pdf.ln(5)
    
    pdf.multi_cell(0, 6, "Please include your invoice # and copy with any check payment or email ar.remit.advice@alm.com for electronic payments.")
    pdf.ln(5)
    
    pdf.multi_cell(0, 6, "Thank you for your business.")
    pdf.multi_cell(0, 6, "For general inquiries and customer support, contact us by phone 1-877-256-2472, email: customercare@alm.com, or fax 646-822-5050")
    pdf.multi_cell(0, 6, "If you have already made payment in full, please disregard this notification.")
    pdf.multi_cell(0, 6, f"Call your Account Manager, {row['SalesCode']} at for information about our publications")
    
    return pdf

if uploaded_file:
    try:
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
                pdf = create_invoice(row)
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
