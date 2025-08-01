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
        
        # Set document margins
        pdf.set_left_margin(15)
        pdf.set_right_margin(15)
        
        # 1. Three-column header
        pdf.set_font('Arial', '', 12)
        pdf.cell(95, 10, "", 0, 0)  # Empty first column
        pdf.cell(95, 10, f"Invoice #: {row.get('Invoic', 'N/A')}", 0, 0)
        pdf.cell(0, 10, f"Date: {datetime.now().strftime('%m/%d/%Y')}", 0, 1)
        
        # 2. Bill To / Ship To section
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(95, 10, "Bill To", 0, 0, 'L', fill=True)
        pdf.cell(95, 10, "Ship To", 0, 0, 'L', fill=True)
        pdf.cell(0, 10, "", 0, 1, fill=True)
        
        pdf.set_font('Arial', '', 12)
        # Bill To Address
        pdf.cell(95, 6, f"{row.get('Bill_To_Contact_name', '')}", 0, 0, 'L')
        pdf.cell(95, 6, f"{row.get('Ship_To_Contact_name', '')}", 0, 0, 'L')
        pdf.cell(0, 6, f"PROMO: {row.get('Curr_Promo_Code', '')}", 0, 1, 'L')
        
        pdf.cell(95, 6, f"{row.get('Bill_to_Company', '')}", 0, 0, 'L')
        pdf.cell(95, 6, f"{row.get('Ship_to_Company', '')}", 0, 0, 'L')
        pdf.cell(0, 6, f"SALES: {row.get('SalesCode', '')}", 0, 1, 'L')
        
        pdf.cell(95, 6, f"{row.get('Bill_to_St_Address', '')}", 0, 0, 'L')
        pdf.cell(95, 6, f"{row.get('Ship_to_St_Address', '')}", 0, 0, 'L')
        pdf.cell(0, 6, "", 0, 1)
        
        city_state_zip_bill = f"{row.get('Bill_to_City', '')} {row.get('Bill_to_State', '')} {row.get('Bill_to_Zip', '')}"
        city_state_zip_ship = f"{row.get('Ship_to_City', '')} {row.get('Ship_to_State', '')} {row.get('Ship_to_Zip', '')}"
        pdf.cell(95, 6, city_state_zip_bill, 0, 0, 'L')
        pdf.cell(95, 6, city_state_zip_ship, 0, 0, 'L')
        pdf.cell(0, 6, "", 0, 1)
        
        pdf.ln(5)

        # 3. Six-column account info table - FORCE EXACT WIDTHS
        col_widths = [50, 40, 40, 30, 35, 40]  # Your exact requested widths
        
        # Header Row
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        headers = ["Cust. Acct. #", "Order #", "Purchase Order", "Term", "Order Date", "Due Date"]
        
        # Reset to left margin
        pdf.set_x(15)
        
        # Draw each header cell with exact width
        for width, header in zip(col_widths, headers):
            pdf.cell(width, 10, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data Row
        pdf.set_font('Arial', '', 12)
        order_date = convert_excel_date(row.get('Order_date'))
        formatted_order_date = order_date.strftime('%m/%d/%Y') if order_date else "N/A"
        
        due_date = convert_excel_date(row.get('DueDate'))
        formatted_due_date = due_date.strftime('%m/%d/%Y') if due_date else "Due Upon Receipt"
        
        data_values = [
            str(row.get('Customer_Account_Number', '')),
            str(row.get('Order', '')),
            str(row.get('PO_Num', '')) if pd.notna(row.get('PO_Num')) else "",
            f"{row.get('Term', '')} days",
            formatted_order_date,
            formatted_due_date
        ]
        
        # Reset to left margin
        pdf.set_x(15)
        
        # Draw each data cell with exact width
        for width, value in zip(col_widths, data_values):
            pdf.cell(width, 10, value, 1, 0, 'C')
        pdf.ln()
        pdf.ln(10)

        # 4. Twelve-column product table - FORCE EXACT WIDTHS
        col_widths_product = [27, 18, 18, 25, 15, 28, 27, 27, 18, 18, 25, 25]
        
        # Header Row
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        product_headers = [
            "Sub. Ref #", "Product", "Copies", "Full Journal Name", "Seats", 
            "Description", "End Date", "Sales", "S&H", "Tax", "Payment", "Total Due"
        ]
        
        # Reset to left margin
        pdf.set_x(15)
        
        # Draw each header cell with exact width
        for width, header in zip(col_widths_product, product_headers):
            pdf.cell(width, 10, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data Row
        pdf.set_font('Arial', '', 12)
        expire_date = convert_excel_date(row.get('Expire_Date'))
        formatted_expire_date = expire_date.strftime('%m/%d/%Y') if expire_date else "N/A"
        
        product_values = [
            str(row.get('Sub_Ref_No', '')),
            str(row.get('Pub_Code', '')),
            str(int(row.get('Quantity', 0))),
            str(row.get('Pub_desc', '')),
            str(int(row.get('Num_of_Seats', 0))),
            str(row.get('Delivery_Code', '')),
            formatted_expire_date,
            f"${float(row.get('Material_Amount', 0)):,.2f}",
            f"${float(row.get('Postage', 0)):,.2f}",
            f"${float(row.get('Tax', 0)):,.2f}",
            f"${float(row.get('Paid_Amount', 0)):,.2f}",
            f"${float(row.get('Amount_Due', 0)):,.2f}"
        ]
        
        # Reset to left margin
        pdf.set_x(15)
        
        # Draw each data cell with exact width
        for width, value in zip(col_widths_product, product_values):
            pdf.cell(width, 10, value, 1, 0, 'C')
        pdf.ln()
        
        # Total Row
        pdf.set_font('Arial', 'B', 12)
        pdf.set_fill_color(230, 230, 230)
        
        # Reset to left margin
        pdf.set_x(15)
        
        # First 5 empty columns
        for _ in range(5):
            pdf.cell(col_widths_product[_], 10, "", 1, 0, 'C', fill=True)
        
        # "Total Due" label
        pdf.cell(col_widths_product[5], 10, "Total Due", 1, 0, 'C', fill=True)
        
        # Empty column
        pdf.cell(col_widths_product[6], 10, "", 1, 0, 'C', fill=True)
        
        # Values for last 5 columns
        for i in range(7, 12):
            pdf.cell(col_widths_product[i], 10, product_values[i], 1, 0, 'C', fill=True)
        pdf.ln()

                # [...] (Previous code remains exactly the same until after the twelve-column table)

        # Installment Effort line with gray background
        pdf.set_fill_color(230, 230, 230)  # 20% gray
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(40, 10, "Installment Effort #:", 0, 0, 'L', fill=True)
        pdf.set_font('Arial', '', 10)
        pdf.cell(0, 10, f" {row.get('Effort_No', 'N/A')}", 0, 1, 'L', fill=True)
        pdf.ln(5)

        # 4th Quarter Payment section
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(60, 10, "4th Quarter Payment", 0, 0, 'L')
        
        # Create the 2x4 table
        table_width = 120  # Total width for the right side table
        col1_width = 80    # Width for label column
        col2_width = 40    # Width for value column
        
        # Table data - labels and corresponding CSV columns
        table_data = [
            ("CUMULATIVE QUARTERLY BILLINGS:", "GroupOutst"),
            ("PAYMENTS:", "Paid_Amount"),
            ("CURRENT QUARTERLY BILLINGS:", "Instalment_Due"),
            ("MINIMUM BAL. DUE TODAY:", "Instalment")
        ]
        
        # Draw the table
        for label, col_name in table_data:
            pdf.set_x(75)  # Position for the table
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(col1_width, 10, label, 1, 0, 'L')
            pdf.set_font('Arial', '', 10)
            value = row.get(col_name, 'N/A')
            if isinstance(value, (int, float)):
                value = f"${float(value):,.2f}"
            pdf.cell(col2_width, 10, str(value), 1, 1, 'R')
        
        pdf.ln(10)

        # Account Manager line
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 10, "Call your Account Manager, Heather Granger at 973-854-2932 for information about our publications.", 0, 1, 'C')
        pdf.ln(5)

        # Payment Options section
        pdf.set_font('Arial', 'BU', 12)
        pdf.cell(0, 10, "Payment Options", 0, 1, 'C')
        pdf.ln(10)

        # Check payment instructions
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(30, 6, "To Pay by Check:", 0, 0, 'L')
        pdf.set_font('Arial', '', 12)
        pdf.multi_cell(0, 6, "Make checks payable to ALM Global, LLC and reference your subscription number on your check. Allow 14-21 days for your check to credit to your account. Disregard invoices you may receive once you have paid.")
        pdf.ln(5)
        
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 6, "Send your check along with this form to the address below:", 0, 1, 'L')
        pdf.cell(0, 6, "US: ALM Global, LLC, PO BOX 70162, Philadelphia, PA, 19176-9628", 0, 1, 'L')
        pdf.ln(5)

        # EFT payment instructions
        pdf.set_font('Arial', 'BU', 12)
        pdf.cell(0, 10, "To Pay by EFT:", 0, 1, 'L')
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 6, "BANK NAME: WELLS FARGO BANK, N.A.", 0, 1, 'L')
        pdf.cell(0, 6, "ADDRESS: 420 Montgomery Street, San Francisco, CA 94104", 0, 1, 'L')
        pdf.cell(0, 6, "ACCOUNT NUMBER: 2000005971161", 0, 1, 'L')
        pdf.cell(0, 6, "ABA NUMBER: 121000248", 0, 1, 'L')
        pdf.cell(0, 6, "BANK ACCOUNT NAME: ALM Global, LLC", 0, 1, 'L')
        pdf.cell(0, 6, "SWIFT: WFBIUS6S", 0, 1, 'L')
        pdf.cell(0, 6, "CHIPS: 0407", 0, 1, 'L')
        pdf.ln(5)

        # Remittance advice
        pdf.cell(0, 6, "Please include your invoice # and copy with any check payment or", 0, 1, 'L')
        pdf.cell(0, 6, "email ar.remit.advice@alm.com for electronic payments.", 0, 1, 'L')
        pdf.ln(10)

        # Footer section
        pdf.set_font('Arial', '', 12)
        pdf.cell(0, 6, "Thank you for your business.", 0, 1, 'C')
        pdf.cell(0, 6, "For general inquiries and customer support, contact us by", 0, 1, 'C')
        pdf.cell(0, 6, "phone 1-877-256-2472, email: customercare@alm.com, or fax 646-822-5050", 0, 1, 'C')
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0, 6, "If you have already made payment in full, please disregard this notification.", 0, 1, 'C')

        return pdf
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")
        return None
        
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
