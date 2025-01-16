import streamlit as st
import pandas as pd
import qrcode
import os
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage
from reportlab.lib import colors

def generate_qr_codes(file, sheet_name, column_name, output_dir):
    # Save the uploaded file to a temporary path
    temp_file_path = os.path.join(output_dir, 'temp_file.xlsx')
    with open(temp_file_path, 'wb') as f:
        f.write(file.getbuffer())

    # Load the Excel file and read the specified column
    df = pd.read_excel(temp_file_path, sheet_name=sheet_name)
    column_data = df[column_name]

    # Generate and save QR codes for each entry in the column
    qr_paths = []
    for i, data in enumerate(column_data):
        # Create a QR code instance
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )

        # Add data to the QR code
        qr.add_data(data)
        qr.make(fit=True)

        # Create an image from the QR code
        img = qr.make_image(fill_color="black", back_color="white")

        # Save the image
        img_path = os.path.join(output_dir, f'qrcode_{i+1}.png')
        img.save(img_path)
        qr_paths.append(img_path)

    # Load the Excel file for writing
    wb = load_workbook(temp_file_path)
    ws = wb[sheet_name]

    # Insert QR code images into the corresponding rows
    for i, img_path in enumerate(qr_paths):
        img = Image(img_path)
        cell = f'B{i+2}'  # Assuming the QR codes will be inserted starting from the second row in column B
        ws.add_image(img, cell)

    # Save the updated Excel file
    updated_file_path = os.path.join(output_dir, 'updated_file.xlsx')
    wb.save(updated_file_path)

    # Convert the updated Excel DataFrame to a list of lists for the Table
    table_data = [list(df.columns) + ['QR Code']]  # Add a header for the QR code column

    # Add data rows to the table, including the QR code images
    for i, row in df.iterrows():
        row_data = list(row.values)
        img_path = qr_paths[i]
        row_data.append(img_path)
        table_data.append(row_data)

    # Create a PDF document
    pdf_path = os.path.join(output_dir, 'output_with_qr.pdf')
    pdf = SimpleDocTemplate(pdf_path, pagesize=letter)

    # Create a Table with the data and QR code images
    table_elements = []
    for row_data in table_data:
        row_elements = []
        for item in row_data:
            if isinstance(item, str) and item.endswith('.png'):
                row_elements.append(ReportLabImage(item, width=50, height=50))
            else:
                row_elements.append(item)
        table_elements.append(row_elements)

    table = Table(table_elements)

    # Apply some basic styling to the table
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])
    table.setStyle(style)

    # Build the PDF
    elements = [table]
    pdf.build(elements)

    return updated_file_path, pdf_path

def main():
    st.title("Excel to QR Code Generator")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file is not None:
        sheet_name = st.text_input("Sheet name", "Sheet1")
        column_name = st.text_input("Column name", "YourColumnName")

        if st.button("Generate QR Codes"):
            with st.spinner("Generating QR codes..."):
                output_dir = "qr_codes"
                updated_file_path, pdf_path = generate_qr_codes(uploaded_file, sheet_name, column_name, output_dir)
                st.success("QR codes generated successfully!")
                st.download_button(label="Download updated Excel file", data=open(updated_file_path, 'rb'), file_name=updated_file_path)
                st.download_button(label="Download PDF file", data=open(pdf_path, 'rb'), file_name=pdf_path)

if __name__ == "__main__":
    main()
