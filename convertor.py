import os
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

def excel_to_pdf(excel_file, pdf_file):
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Create a PDF file
    pdf = SimpleDocTemplate(pdf_file, pagesize=landscape(letter))
    elements = []

    # Create a Table with the data
    data = [df.columns.values.tolist()] + df.values.tolist()
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    elements.append(table)
    pdf.build(elements)

def convert_all_excels_to_pdfs(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            excel_path = os.path.join(directory, filename)
            pdf_path = os.path.splitext(excel_path)[0] + '.pdf'
            excel_to_pdf(excel_path, pdf_path)
            print(f'Converted {excel_path} to {pdf_path}')

# Set the directory to the current directory
current_directory = os.getcwd()

# Convert all Excel files in the current directory to PDFs
convert_all_excels_to_pdfs(current_directory)
