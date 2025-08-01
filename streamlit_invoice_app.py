        # 2. Six-column account info table
        pdf.set_left_margin(15)  # Set left margin
        pdf.set_x(15)  # Reset X position to margin start
        
        col_widths = [46, 37, 35, 26, 30, 37]  # Your custom widths (sum=211mm)
        total_width = sum(col_widths)
        
        # Debug output (visible in Streamlit)
        st.write(f"Debug - Table Widths:")
        st.write(f"• Available width: 267mm (297mm page - 15mm margins)")
        st.write(f"• Your columns sum: {total_width}mm")
        st.write(f"• Column widths: {col_widths}")
        
        if total_width > 267:
            st.warning("Warning: Column widths exceed available space!")
        
        # Header Row
        pdf.set_fill_color(230, 230, 230)
        pdf.set_font('Arial', 'B', 12)
        headers = ["Cust. Acct. #", "Order #", "Purchase Order", "Term", "Order Date", "Due Date"]
        
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 10, header, 1, 0, 'C', fill=True)
        pdf.ln()
        
        # Data Row - Add content trimming to prevent overflow
        pdf.set_font('Arial', '', 10)  # Slightly smaller font
        data_values = [
            str(row.get('Customer_Account_Number', ''))[:12],  # Trim long values
            str(row.get('Order', ''))[:8],
            (str(row.get('PO_Num', ''))[:6] if pd.notna(row.get('PO_Num')) else "",
            f"{row.get('Term', '')}d",  # "d" instead of "days" to save space
            formatted_order_date,
            formatted_due_date
        ]
        
        for i, value in enumerate(data_values):
            pdf.cell(col_widths[i], 10, value, 1, 0, 'C')
        pdf.ln()
