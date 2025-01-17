import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage
from reportlab.lib import colors

def generate_qr_codes(df, selected_columns, qr_column):
    # Read the specified column for QR codes
    column_data = df[qr_column]

    # Generate and store QR codes for each entry in the column
    qr_images = []
    for data in column_data:
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
        
        # Save the QR code to a BytesIO object
        img_buffer = BytesIO()
        img.save(img_buffer, format="PNG")
        img_buffer.seek(0)
        qr_images.append(img_buffer)

    # Convert the DataFrame to a list of lists for the Table
    table_data = [selected_columns + ['QR Code']]  # Add a header for the QR code column

    # Add data rows to the table, including the QR code images
    for i, row in df.iterrows():
        row_data = list(row[selected_columns].values)
        qr_img = qr_images[i]
        row_data.append(qr_img)
        table_data.append(row_data)

    # Create a PDF document
    pdf_buffer = BytesIO()
    pdf = SimpleDocTemplate(pdf_buffer, pagesize=letter)

    # Create a Table with the data and QR code images
    table_elements = []
    for row_data in table_data:
        row_elements = []
        for item in row_data:
            if isinstance(item, BytesIO):
                img = ReportLabImage(item, width=50, height=50)
                row_elements.append(img)
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
        ('FONTSIZE', (0, 0), 12),
        ('BOTTOMPADDING', (0, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])
    table.setStyle(style)

    # Build the PDF
    elements = [table]
    pdf.build(elements)

    # Return the PDF as a BytesIO object
    pdf_buffer.seek(0)
    return pdf_buffer

def main():
    st.title("QR Code Generator")

    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["xlsx", "csv"])
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)

        columns = df.columns.tolist()
        qr_column = st.selectbox("Select the column for QR codes", columns)
        selected_columns = st.multiselect("Select the columns to include in the PDF", columns, default=qr_column)

        if st.button("Generate QR Codes"):
            with st.spinner("Generating QR codes..."):
                pdf_buffer = generate_qr_codes(df, selected_columns, qr_column)
                st.success("QR codes generated successfully!")
                st.download_button(label="Download PDF file", data=pdf_buffer, file_name='output_with_qr.pdf', mime='application/pdf')

if __name__ == "__main__":
    main()
