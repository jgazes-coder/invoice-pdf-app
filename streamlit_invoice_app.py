def create_invoice(row, logo):
    try:
        # Initialize PDF with portrait orientation and tight margins
        pdf = ALMInvoice(orientation="P", logo=logo)
        pdf.add_page()
        pdf.set_margins(left=10, top=15, right=10)  # Slightly wider margins for safety
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Font sizes
        base_font_size = 9  # Slightly smaller base font
        header_font_size = 16  # Smaller invoice header
        
        # 0. Logo and Invoice Header
        pdf.set_font('Arial', 'B', header_font_size)
        pdf.cell(0, 15, 'INVOICE', 0, 1, 'C')
        pdf.ln(5)

        # 1. Three-column header (invoice#, date)
        pdf.set_font('Arial', '', base_font_size)
        col1_width = 60
        col2_width = 60
        pdf.cell(col1_width, 6, "", 0, 0)
        pdf.cell(col2_width, 6, f"Invoice #: {row.get('Invoic', 'N/A')}", 0, 0)
        pdf.cell(0, 6, f"Date: {datetime.now().strftime('%m/%d/%Y')}", 0, 1)
        pdf.ln(8)

        # 2. Bill To / Ship To section (compact version)
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', base_font_size)
        
        # Header row
        pdf.cell(95, 6, "Bill To", 0, 0, 'L', fill=True)
        pdf.cell(95, 6, "Ship To", 0, 1, 'L', fill=True)
        
        # Content rows
        pdf.set_font('Arial', '', base_font_size)
        fields = [
            ('Bill_To_Contact_name', 'Ship_To_Contact_name'),
            ('Bill_to_Company', 'Ship_to_Company'),
            ('Bill_to_St_Address', 'Ship_to_St_Address'),
            ('', '')  # City/State/Zip will be combined
        ]
        
        for bill_field, ship_field in fields:
            if bill_field:
                bill_value = str(row.get(bill_field, ''))
                ship_value = str(row.get(ship_field, ''))
            else:
                # Combine city/state/zip
                bill_value = f"{row.get('Bill_to_City', '')} {row.get('Bill_to_State', '')} {row.get('Bill_to_Zip', '')}"
                ship_value = f"{row.get('Ship_to_City', '')} {row.get('Ship_to_State', '')} {row.get('Ship_to_Zip', '')}"
            
            pdf.cell(95, 5, bill_value, 0, 0, 'L')
            pdf.cell(95, 5, ship_value, 0, 1, 'L')
        
        # Promo and sales code on same line
        pdf.cell(95, 5, f"PROMO: {row.get('Curr_Promo_Code', '')}", 0, 0, 'L')
        pdf.cell(95, 5, f"SALES: {row.get('SalesCode', '')}", 0, 1, 'L')
        pdf.ln(8)

        # 3. Six-column account info table (compact)
        col_widths = [25, 25, 35, 20, 25, 25]  # Total: 155 (fits in 190mm width with margins)
        headers = ["Cust. Acct.", "Order #", "PO Num", "Term", "Order Date", "Due Date"]
        
        # Header
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.set_x(10)
        for width, header in zip(col_widths, headers):
            pdf.cell(width, 8, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data
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

        # 4. Twelve-column product table (FIXED version)
        col_widths = [18, 15, 12, 25, 10, 20, 18, 15, 12, 12, 15, 18]  # Total: 190mm
        headers = [
            "Sub Ref", "Product", "Copies", "Journal", 
            "Seats", "Desc", "End Date", "Sales", 
            "S&H", "Tax", "Payment", "Total Due"
        ]
        
        # Header row - using single line cells
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', base_font_size-1)  # Slightly smaller
        pdf.set_x(10)
        for width, header in zip(col_widths, headers):
            pdf.cell(width, 8, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data row
        pdf.set_font('Arial', '', base_font_size)
        pdf.set_x(10)
        expire_date = convert_excel_date(row.get('Expire_Date'))
        
        cells = [
            str(row.get('Sub_Ref_No', '')),
            str(row.get('Pub_Code', '')),
            str(int(row.get('Quantity', 0))),
            str(row.get('Pub_desc', ''))[:15],  # Truncate long journal names
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

        # 5. Installment and Payment Info (compact)
        pdf.set_font('Arial', 'B', base_font_size)
        pdf.cell(40, 6, "Installment Effort #:", 0, 0, 'L')
        pdf.set_font('Arial', '', base_font_size)
        pdf.cell(0, 6, str(row.get('Effort_No', 'N/A')), 0, 1, 'L')
        pdf.ln(5)

        # Payment table (right-aligned)
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

        # 6. Payment instructions (compact)
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
