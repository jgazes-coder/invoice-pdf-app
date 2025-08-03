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
        # Initialize PDF with portrait orientation and tight margins
        pdf = ALMInvoice(orientation="P", logo=logo)
        pdf.add_page()
        pdf.set_margins(left=10, top=15, right=10)
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Font sizes
        base_font_size = 9
        header_font_size = 16

        # Invoice Header
        pdf.set_font('Arial', 'B', header_font_size)
        pdf.cell(0, 15, 'INVOICE', 0, 1, 'C')
        pdf.ln(5)

        # 1. Three-column header (invoice#, date)
        pdf.set_font('Arial', '', base_font_size)
        pdf.cell(60, 6, "", 0, 0)
        pdf.cell(60, 6, f"Invoice #: {row.get('Invoic', 'N/A')}", 0, 0)
        pdf.cell(0, 6, f"Date: {datetime.now().strftime('%m/%d/%Y')}", 0, 1)
        pdf.ln(5)

        # 2. Three-column Bill To / Ship To / Promo-Sales section
        col_widths = [70, 70, 50]  # Adjusted widths for three columns
        
        # Header row with gray background
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.cell(col_widths[0], 8, "Bill To", 1, 0, 'L', fill=True)
        pdf.cell(col_widths[1], 8, "Ship To", 1, 0, 'L', fill=True)
        pdf.cell(col_widths[2], 8, "", 1, 1, 'L', fill=True)
        
        # Content rows
        pdf.set_font('Arial', '', base_font_size)
        
        # First row: Contact names and Promo code
        pdf.cell(col_widths[0], 6, str(row.get('Bill_To_Contact_name', '')), 1, 0, 'L')
        pdf.cell(col_widths[1], 6, str(row.get('Ship_To_Contact_name', '')), 1, 0, 'L')
        pdf.cell(col_widths[2], 6, f"PROMO: {row.get('Curr_Promo_Code', '')}", 1, 1, 'L')
        
        # Second row: Companies and Sales code
        pdf.cell(col_widths[0], 6, str(row.get('Bill_to_Company', '')), 1, 0, 'L')
        pdf.cell(col_widths[1], 6, str(row.get('Ship_to_Company', '')), 1, 0, 'L')
        pdf.cell(col_widths[2], 6, f"SALES: {row.get('SalesCode', '')}", 1, 1, 'L')
        
        # Third row: Street addresses
        pdf.cell(col_widths[0], 6, str(row.get('Bill_to_St_Address', '')), 1, 0, 'L')
        pdf.cell(col_widths[1], 6, str(row.get('Ship_to_St_Address', '')), 1, 0, 'L')
        pdf.cell(col_widths[2], 6, "", 1, 1, 'L')
        
        # Fourth row: City/State/Zip
        bill_csz = f"{row.get('Bill_to_City', '')} {row.get('Bill_to_State', '')} {row.get('Bill_to_Zip', '')}"
        ship_csz = f"{row.get('Ship_to_City', '')} {row.get('Ship_to_State', '')} {row.get('Ship_to_Zip', '')}"
        pdf.cell(col_widths[0], 6, bill_csz, 1, 0, 'L')
        pdf.cell(col_widths[1], 6, ship_csz, 1, 0, 'L')
        pdf.cell(col_widths[2], 6, "", 1, 1, 'L')
        
        pdf.ln(8)

        # 3. Six-column account info table
        col_widths = [25, 25, 35, 20, 25, 25]
        headers = ["Cust. Acct.", "Order #", "PO Num", "Term", "Order Date", "Due Date"]
        
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.set_x(10)
        for width, header in zip(col_widths, headers):
            pdf.cell(width, 8, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        pdf.set_font('Arial', '', base_font_size)
        pdf.set_x(10)
        order_date = convert_excel_date(row.get('Order_date'))
        due_date = convert_excel_date(row.get('DueDate'))
        
        cells = [
            str(row.get('Customer_Account_Number', '')),
            str(row.get('Order', '')),
            str(row.get('PO_Num', '')) if pd.notna(row.get('PO_Num')) else "",
            f"{row.get('Term', '')} days",
            order_date.strftime('%m/%d/%Y') if order_date else "N/A",
            due_date.strftime('%m/%d/%Y') if due_date else "Upon Receipt"
        ]
        
        for width, value in zip(col_widths, cells):
            pdf.cell(width, 8, value, 1, 0, 'C')
        pdf.ln(10)

        # 4. Twelve-column product table
        col_widths = [18, 15, 12, 25, 10, 20, 18, 15, 12, 12, 15, 18]
        headers = [
            "Sub Ref", "Product", "Copies", "Journal", 
            "Seats", "Desc", "End Date", "Sales", 
            "S&H", "Tax", "Payment", "Total Due"
        ]
        
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', base_font_size-1)
        pdf.set_x(10)
        for width, header in zip(col_widths, headers):
            pdf.cell(width, 8, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        pdf.set_font('Arial', '', base_font_size)
        pdf.set_x(10)
        expire_date = convert_excel_date(row.get('Expire_Date'))
        
        cells = [
            str(row.get('Sub_Ref_No', '')),
            str(row.get('Pub_Code', '')),
            str(int(row.get('Quantity', 0))),
            str(row.get('Pub_desc', ''))[:15],
            str(int(row.get('Num_of_Seats', 0))),
            str(row.get('Delivery_Code', '')),
            expire_date.strftime('%m/%d/%Y') if expire_date else "N/A",
            f"${float(row.get('Material_Amount', 0)):,.2f}",
            f"${float(row.get('Postage', 0)):,.2f}",
            f"${float(row.get('Tax', 0)):,.2f}",
            f"${float(row.get('Paid_Amount', 0)):,.2f}",
            f"${float(row.get('Amount_Due', 0)):,.2f}"
        ]
        
        for width, value in zip(col_widths, cells):
            pdf.cell(width, 8, value, 1, 0, 'C')
        pdf.ln()
        
        # Total row
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.set_x(10)
        for i in range(5):
            pdf.cell(col_widths[i], 8, "", 1, 0, 'C', fill=True)
        pdf.cell(col_widths[5], 8, "Total Due", 1, 0, 'C', fill=True)
        pdf.cell(col_widths[6], 8, "", 1, 0, 'C', fill=True)
        for i in range(7, 12):
            pdf.cell(col_widths[i], 8, cells[i], 1, 0, 'C', fill=True)
        pdf.ln(10)

        # 5. Installment and Payment Info
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.cell(40, 6, "Installment Effort #:", 0, 0, 'L')
        pdf.set_font('Arial', '', base_font_size)
        pdf.cell(0, 6, str(row.get('Effort_No', 'N/A')), 0, 1, 'L')
        pdf.ln(5)

        # Payment table
        table_data = [
            ("CUMULATIVE BILLINGS:", "GroupOutst"),
            ("PAYMENTS:", "Paid_Amount"),
            ("CURRENT BILLINGS:", "Instalment_Due"),
            ("BAL. DUE TODAY:", "Instalment")
        ]
        
        for label, field in table_data:
            pdf.set_x(100)
            pdf.set_font('Arial', 'B', base_font_size)
            pdf.cell(60, 6, label, 0, 0, 'L')
            pdf.set_font('Arial', '', base_font_size)
            value = row.get(field, 0)
            pdf.cell(30, 6, f"${float(value):,.2f}", 0, 1, 'R')
        pdf.ln(8)

        # 6. Payment instructions
        pdf.set_font('Arial', 'BU', base_font_size)
        pdf.cell(0, 6, "Payment Options", 0, 1, 'L')
        pdf.ln(3)
        
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.cell(25, 6, "By Check:", 0, 0, 'L')
        pdf.set_font('Arial', '', base_font_size)
        pdf.multi_cell(0, 6, "Make payable to ALM Global, LLC. Mail to: PO BOX 70162, Philadelphia, PA 19176-9628")
        pdf.ln(3)
        
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.cell(25, 6, "By EFT:", 0, 0, 'L')
        pdf.set_font('Arial', '', base_font_size)
        pdf.multi_cell(0, 6, "Wells Fargo Bank | ABA: 121000248 | Acct: 2000005971161 | SWIFT: WFBIUS6S")
        pdf.ln(3)
        
        pdf.multi_cell(0, 6, "Email remittance to: ar.remit.advice@alm.com")
        pdf.ln(5)
        
        # 7. Footer
        pdf.set_font('Arial', '', base_font_size-1)
        pdf.cell(0, 6, "Thank you for your business. Contact: 1-877-256-2472 | customercare@alm.com", 0, 1, 'C')
        pdf.set_font('Arial', 'B', base_font_size-1)
        pdf.cell(0, 6, "If already paid, please disregard this notice.", 0, 1, 'C')

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
