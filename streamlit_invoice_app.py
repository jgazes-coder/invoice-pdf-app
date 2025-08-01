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
                pass
        
        # Invoice title
        self.set_font('Arial', 'B', 20)
        self.cell(0, 25, 'INVOICE', 0, 1, 'C')
        self.ln(10)

def process_logo(uploaded_file):
    if not uploaded_file:
        return None
    try:
        img = Image.open(uploaded_file)
        img.verify()
        temp_dir = tempfile.mkdtemp()
        logo_path = os.path.join(temp_dir, 'logo.jpg')
        uploaded_file.seek(0)
        with open(logo_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        return {'path': logo_path, 'valid': True, 'temp_dir': temp_dir}
    except Exception as e:
        st.warning(f"Invalid logo image: {str(e)}")
        return None

def convert_excel_date(excel_date):
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
        
        # Set strict layout control
        pdf.set_margins(left=15, top=15, right=15)
        pdf.set_auto_page_break(True, margin=15)
        
        # 1. Three-column header
        pdf.set_font('Arial', '', 12)
        pdf.cell(95, 10, "", 0, 0)
        pdf.cell(95, 10, f"Invoice #: {row.get('Sub_Ref_No', 'N/A')}", 0, 0)
        pdf.cell(0, 10, f"Date: {datetime.now().strftime('%m/%d/%Y')}", 0, 1)
        
        # Bill To / Ship To section
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(95, 10, "Bill To", 0, 0, 'L', fill=True)
        pdf.cell(95, 10, "Ship To", 0, 0, 'L', fill=True)
        pdf.cell(0, 10, "", 0, 1, fill=True)
        
        pdf.set_font('Arial', '', 12)
        # [Rest of your address section remains the same...]

        # 2. Six-column account info table - NEW APPROACH
        col_widths = [46, 37, 35, 26, 30, 37]  # Your exact requested widths
        
        # Calculate starting position to center the table
        total_width = sum(col_widths)
        start_x = (297 - total_width) / 2  # Center on landscape A4 (297mm wide)
        
        # Header Row - FORCED WIDTHS
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        headers = ["Cust. Acct. #", "Order #", "PO", "Term", "Order Date", "Due Date"]
        
        pdf.set_x(start_x)
        for width, header in zip(col_widths, headers):
            # Use cell() instead of multi_cell() for strict width control
            pdf.cell(width, 10, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data Row - FORCED WIDTHS
        pdf.set_font('Arial', '', 12)  # Restored original font size
        order_date = convert_excel_date(row.get('Order_date'))
        due_date = convert_excel_date(row.get('DueDate'))
        
        data = [
            str(row.get('Customer_Account_Number', ''))[:12],
            str(row.get('Order', ''))[:8],
            str(row.get('PO_Num', ''))[:6] if pd.notna(row.get('PO_Num')) else "",
            f"{row.get('Term', '')} days",
            order_date.strftime('%m/%d/%Y') if order_date else "N/A",
            due_date.strftime('%m/%d/%Y') if due_date else "Due Now"
        ]
        
        pdf.set_x(start_x)
        for width, value in zip(col_widths, data):
            pdf.cell(width, 10, value, 1, 0, 'C')
        pdf.ln()
        pdf.ln(10)

        # 3. Twelve-column product table - NEW APPROACH
        col_widths_product = [20, 14, 12, 24, 11, 22, 17, 21, 13, 16, 20, 24]
        total_product_width = sum(col_widths_product)
        start_x_product = (297 - total_product_width) / 2
        
        # Header Row
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        headers = ["Sub.Ref#", "Product", "Copies", "Journal", "Seats", "Desc", 
                  "End Date", "Sales", "S&H", "Tax", "Payment", "Total Due"]
        
        pdf.set_x(start_x_product)
        for width, header in zip(col_widths_product, headers):
            pdf.cell(width, 10, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data Row
        pdf.set_font('Arial', '', 12)
        expire_date = convert_excel_date(row.get('Expire_Date'))
        data = [
            str(row.get('Sub_Ref_No', ''))[:8],
            str(row.get('Pub_Code', ''))[:6],
            str(int(row.get('Quantity', 0))),
            str(row.get('Pub_desc', ''))[:12],
            str(int(row.get('Num_of_Seats', 0))),
            str(row.get('Delivery_Code', ''))[:8],
            expire_date.strftime('%m/%d/%Y') if expire_date else "N/A",
            f"${float(row.get('Material_Amount', 0)):,.2f}",
            f"${float(row.get('Postage', 0)):,.2f}",
            f"${float(row.get('Tax', 0)):,.2f}",
            f"${float(row.get('Paid_Amount', 0)):,.2f}",
            f"${float(row.get('Amount_Due', 0)):,.2f}"
        ]
        
        pdf.set_x(start_x_product)
        for width, value in zip(col_widths_product, data):
            pdf.cell(width, 10, value, 1, 0, 'C')
        pdf.ln()
        
        return pdf
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None

if uploaded_file:
    try:
        logo = process_logo(logo_file)
        df = pd.read_csv(uploaded_file)
        zip_buffer = io.BytesIO()
        success_count = 0
        
        with ZipFile(zip_buffer, 'w') as zip_file:
            for _, row in df.iterrows():
                pdf = create_invoice(row, logo)
                if pdf:
                    filename = f"Invoice_{row.get('Sub_Ref_No', '')}.pdf"
                    pdf_bytes = pdf.output(dest='S').encode('latin1')
                    zip_file.writestr(filename, pdf_bytes)
                    success_count += 1
        
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
    except Exception as e:
        st.error(f"Processing failed: {str(e)}")
