import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as ReportLabImage, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from PIL import Image
import fitz  # PyMuPDF

def generate_qr_codes(df, qr_columns, layout):
    qr_images = {}

    # Generate and store QR codes for each selected QR column
    for col in qr_columns:
        qr_images[col] = []
        for data in df[col]:
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
            qr_images[col].append(img_buffer)

    styles = getSampleStyleSheet()

    # Convert the DataFrame to a list of lists for the Table
    if layout == "All columns + Qr code":
        table_data = [qr_columns + [f'QR Code ({col})' for col in qr_columns]]
    elif layout == "Column-QR Code/Column-QR Code":
        table_data = [[item for col in qr_columns for item in (col, f'QR Code ({col})')]]
    else:
        table_data = [[f'{col} & QR Code' for col in qr_columns]]

    # Add data rows to the table, including the QR code images
    for i, row in df.iterrows():
        if layout == "All columns + Qr code":
            row_data = list(row[qr_columns].values)
            for col in qr_columns:
                qr_img = qr_images[col][i]
                row_data.append(qr_img)
        elif layout == "Column-QR Code/Column-QR Code":
            row_data = []
            for col in qr_columns:
                row_data.append(row[col])
                if col in qr_columns:
                    row_data.append(qr_images[col][i])
        else:
            row_data = []
            for col in qr_columns:
                if col in qr_columns:
                    qr_img = qr_images[col][i]
                    text = str(row[col])
                    combined_content = [Paragraph(text, styles['Normal']), Spacer(1, 0.2*inch), ReportLabImage(qr_img, width=50, height=50)]
                    row_data.append(combined_content)
                else:
                    row_data.append(str(row[col]))
        table_data.append(row_data)

    # Create a PDF document
    pdf_buffer = BytesIO()
    pdf = SimpleDocTemplate(pdf_buffer, pagesize=letter)

    # Create a Table with the data and QR code images
    table_elements = []
    for row_data in table_data:
        row_elements = []
        for item in row_data:
            if isinstance(item, list):
                cell_content = []
                for sub_item in item:
                    cell_content.append(sub_item)
                row_elements.append(cell_content)
            elif isinstance(item, BytesIO):
                img = ReportLabImage(item, width=50, height=50)
                row_elements.append(img)
            else:
                row_elements.append(str(item))
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
        ('GRID', (0, 1), (-1, -1), 1, colors.black),
    ])
    table.setStyle(style)

    # Build the PDF
    elements = [table]
    pdf.build(elements)

    # Return the PDF as a BytesIO object
    pdf_buffer.seek(0)
    return pdf_buffer

def read_csv_with_encodings(file):
    encodings = ['utf-8', 'latin1', 'utf-16', 'ISO-8859-1']
    for encoding in encodings:
        try:
            return pd.read_csv(file, encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("Unable to decode file with given encodings.")

def main():
    st.title("QR Code Generator")

    uploaded_file = st.file_uploader("Upload your Excel or CSV file", type=["xlsx", "csv"])
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith('.csv'):
            df = read_csv_with_encodings(uploaded_file)

        columns = df.columns.tolist()
        qr_columns = st.multiselect("Select the columns for QR codes", columns)
        layout = st.selectbox("Select the layout", ["Column-QR Code/Column-QR Code", "All columns + Qr code", "Content & QR Code in same cell"])

        if st.button("Generate QR Codes"):
            with st.spinner("Generating QR codes..."):
                pdf_buffer = generate_qr_codes(df, qr_columns, layout)
                
                # Generate a preview image from the PDF
                pdf_bytes = pdf_buffer.getvalue()
                pdf_document = fitz.open("pdf", pdf_bytes)
                pdf_page = pdf_document.load_page(0)
                pdf_image = pdf_page.get_pixmap()
                img_buffer = BytesIO(pdf_image.tobytes())
                img_buffer.seek(0)
                image = Image.open(img_buffer)
                
                st.success("QR codes generated successfully!")
                st.download_button(label="Download PDF file", data=pdf_buffer, file_name='output_with_qr.pdf', mime='application/pdf')
                st.image(image, caption="PDF Preview", use_container_width=True)

if __name__ == "__main__":
    main()
